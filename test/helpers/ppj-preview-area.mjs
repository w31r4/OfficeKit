import assert from "node:assert/strict";
import path from "node:path";
import { writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import JSZip from "jszip";
import sharp from "sharp";

// Real codec inputs; fixed chart geometry is an independent fixture oracle.
export async function checkNativeAreas({pairBase,sourceWorkspace,compilePpjWorkspace,projectPptxToPpj,
  withoutAuthoredSnapshot,createPpjSceneView,savePaint,assertProductionEntry,artifacts}) {
  const cases=[],failures=[],rejections=[];
  const hash=bytes=>createHash("sha256").update(bytes).digest("hex");
  const program=structuredClone(pairBase);
  program.pages[0].elements=[{id:"area",type:"chart",chartType:"area",frame:{x:60,y:80,width:450,height:300},
    xAxis:{visible:false},yAxis:{min:-4,max:8,visible:false},style:{stacking:"none"},data:{
      categories:["A","B","Missing","Zero","Negative"],series:[
        {id:"first",name:"First",values:[4,2,null,0,-4],fill:{type:"solid",color:"#CC2200"}},
        {id:"second",name:"Second",values:[null,null,4,0,null],fill:{type:"solid",color:"#0044CC"}},
      ]}},
    {id:"untouched",type:"shape",frame:{x:600,y:440,width:40,height:30},geometry:{kind:"preset",preset:"rect"},style:{fill:{type:"solid",color:"#00AA44"}}}];
  const workspace=p=>({...sourceWorkspace,program:Buffer.from(JSON.stringify(p))});
  async function check(name,input,receipt,{first=4,reverse=false}={}) {
    const candidateFile=`${name}.pptx`,inputFile=`${name}.ppj`;
    await writeFile(path.join(artifacts,candidateFile),receipt.file,{flag:"wx"});
    await writeFile(path.join(artifacts,inputFile),input.program,{flag:"wx"});
    const chart=createPpjSceneView(receipt).pages[0].nodes.find(n=>n.kind==="chart").native;
    assert.equal(chart.type,4);assert.equal(chart.grouping,"none");
    assert.deepEqual(chart.series.map(s=>s.values),[[first,2,0,0,-4],[0,0,4,0,0]]);
    assert.deepEqual(chart.series.map(s=>s.missingValueIndexes),[[2],[0,1,4]]);
    assert.equal(chart.xAxis.reverse??false,reverse);assert.equal(chart.yAxis.reverse??false,reverse);
    const original=input.program.slice(),sourceBefore=input.source?.slice();
    const ordinary=await compilePpjWorkspace(input),painted=await savePaint(name,receipt),svg=painted.pages[0].svg;
    assert.deepEqual(receipt.file,ordinary.file);assert.equal(painted.reliability.status,"requires-review");
    assert.match(svg,/data-officekit-chart="area"/);assert.doesNotMatch(svg,/data-officekit-line-segment=/);
    const polygons=[...svg.matchAll(/<path data-officekit-area-segment="([^"]+)" d="([^"]+)"[^>]*fill="([^"]+)"/g)].map(m=>m.slice(1));
    const literal=[[[105,265],[105,265-first*17.5],[195,230],[195,265]],
      [[375,265],[375,265],[465,335],[465,265]],[[285,265],[285,195],[375,265],[375,265]]];
    const expected=literal.map((points,i)=>[["0:1","3:4","2:3"][i],points.map(([x,y],j)=>`${j?"L":"M"} ${reverse?570-x:x} ${reverse?460-y:y}`).join(" ")+" Z",["#CC2200","#CC2200","#0044CC"][i]]);
    assert.deepEqual(polygons,[expected[2],expected[0],expected[1]],"native AREA has independently bounded filled regions, first series in front");
    assert.equal((svg.match(/data-officekit-missing-point=/g)||[]).length,4);
    assert.equal((svg.match(/data-officekit-review-point="area-zero"/g)||[]).length,2);
    const raster=await sharp(Buffer.from(svg)).removeAlpha().raw().toBuffer({resolveWithObject:true});
    for(const [px,py,color] of [[150,240,[204,34,0]],[240,220,[255,255,255]],[435,300,[204,34,0]],
      [330,250,[0,68,204]],[115,185,first===4?[255,255,255]:[204,34,0]],[620,450,[0,170,68]]]) {
      // Reflect plot samples around fixed fixture axes; sibling stays put.
      const x=reverse&&px<510?570-px-1:px,y=reverse&&px<510?460-py-1:py;
      const offset=(y*raster.info.width+x)*raster.info.channels;
      assert.deepEqual([...raster.data.subarray(offset,offset+3)],color,`${name}/${px},${py}`);
    }
    const published=await assertProductionEntry(name,input,receipt,painted);
    assert.deepEqual(input.program,original);if(sourceBefore)assert.deepEqual(input.source,sourceBefore);
    const row={name,first,reverse,candidateFile,candidateSha256:hash(receipt.file),inputFile,inputSha256:hash(input.program),
      sceneSha256:receipt.previewScene.sha256,positiveNegativeMissingZeroPixels:true,independentPolygonGeometry:true,
      sceneOnOffCandidateEqual:true,inputsPreserved:true,productionReliability:published.reliability.status};
    cases.push(row);return row;
  }
  try {
    const input=workspace(program),authored=await compilePpjWorkspace(input,{includePreviewScene:true});
    await check("area-authored",input,authored);
    const reversed=structuredClone(program);reversed.pages[0].elements[0].xAxis.reverse=true;reversed.pages[0].elements[0].yAxis.reverse=true;
    const reversedInput=workspace(reversed);
    await check("area-reversed",reversedInput,await compilePpjWorkspace(reversedInput,{includePreviewScene:true}),{reverse:true});
    const source=await withoutAuthoredSnapshot(authored.file),sourceBefore=source.slice();
    const sourceFile="area-original.pptx";await writeFile(path.join(artifacts,sourceFile),source,{flag:"wx"});
    const project=()=>projectPptxToPpj(source,{sourceUri:sourceFile,assetRootUri:"assets"});
    const projected=await project(),sourceInput={program:projected.programJson,source,assets:projected.assets};
    const noop=await compilePpjWorkspace(sourceInput,{includePreviewScene:true});assert.deepEqual(noop.file,source);
    await check("area-source-noop",sourceInput,noop);
    const freshInput=await project(),changed=JSON.parse(new TextDecoder().decode(freshInput.programJson));
    changed.pages[0].elements.find(e=>e.type==="chart").data.series[0].values[0]=6;
    const editInput={program:Buffer.from(JSON.stringify(changed)),source,assets:freshInput.assets};
    await writeFile(path.join(artifacts,"area-value.request.ppj"),editInput.program,{flag:"wx"});
    const candidate=await compilePpjWorkspace(editInput,{includePreviewScene:true});
    const row=await check("area-source-value",editInput,candidate,{first:6});
    const fresh=await projectPptxToPpj(candidate.file,{sourceUri:"area-value.pptx",assetRootUri:"assets"});
    const chart=JSON.parse(new TextDecoder().decode(fresh.programJson)).pages[0].elements.find(e=>e.type==="chart");
    assert.equal(chart.chartType,"area");assert.equal(chart.style.stacking,"none");
    assert.deepEqual(chart.data.series.map(s=>s.values),[[6,2,null,0,-4],[null,null,4,0,null]]);
    const oldZip=await JSZip.loadAsync(source),newZip=await JSZip.loadAsync(candidate.file),changedParts=[];
    assert.deepEqual(Object.keys(oldZip.files).sort(),Object.keys(newZip.files).sort());
    for(const part of Object.keys(oldZip.files))if(!oldZip.files[part].dir&&
      !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array"))))changedParts.push(part);
    assert.deepEqual(changedParts,["ppt/slides/charts/chart1.xml"]);assert.deepEqual(source,sourceBefore);
    const oldXml=await oldZip.file(changedParts[0]).async("string"),newXml=await newZip.file(changedParts[0]).async("string");
    const target='<c:pt idx="0"><c:v>4</c:v></c:pt>';
    assert.equal(oldXml.split(target).length,2,"fixture has one exact target native value");
    assert.equal(newXml,oldXml.replace(target,'<c:pt idx="0"><c:v>6</c:v></c:pt>'),"all non-target ChartML bytes preserved");
    row.nonTargetChartXmlPreserved=true;
    row.reprojection=true;row.changedParts=changedParts;row.sourceFile=sourceFile;row.sourceSha256=hash(source);
    row.reprojectionFile="area-source-value.reprojected.ppj";row.reprojectionSha256=hash(fresh.programJson);
    await writeFile(path.join(artifacts,row.reprojectionFile),fresh.programJson,{flag:"wx"});
  }catch(error){failures.push({name:"ordinary-area",code:error.code,message:error.message});}
  for(const grouping of ["stacked","percent-stacked"])try {
    const value=structuredClone(program);value.pages[0].elements[0].style.stacking=grouping;
    const input=workspace(value),before=input.program.slice();
    await writeFile(path.join(artifacts,`area-${grouping}.request.ppj`),input.program,{flag:"wx"});
    const receipt=await compilePpjWorkspace(input,{includePreviewScene:true}),painted=await savePaint(`area-${grouping}-rejected`,receipt);
    assert.deepEqual(input.program,before);
    assert.equal(painted.reliability.status,"failed");assert.doesNotMatch(painted.pages[0].svg,/data-officekit-area-segment=/);
    assert.ok(painted.diagnostics.some(d=>d.reason==="preview.scene.paint.chart-semantics"&&d.scenePath.endsWith(".chart.grouping")));
    rejections.push({grouping,stage:"paint",inputPreserved:true,candidateSha256:hash(receipt.file)});
  }catch(error){failures.push({name:`area-${grouping}-rejection`,code:error.code,message:error.message});}
  return {cases,failures,rejections,expectedCount:4};
}
