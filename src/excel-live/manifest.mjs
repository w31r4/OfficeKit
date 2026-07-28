import { excelLiveError } from "./errors.mjs";
import { writePrivateText } from "./state.mjs";

export function excelBridgeOrigin(port) {
  if (!Number.isSafeInteger(port) || port < 1024 || port > 65535) {
    throw excelLiveError("invalid-state", "Excel bridge port is invalid.");
  }
  return `https://localhost:${port}`;
}

export function renderExcelManifest({ addinId, port, packageVersion }) {
  assertUuid(addinId, "Add-in ID");
  const origin = excelBridgeOrigin(port);
  const version = manifestVersion(packageVersion);
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${addinId}</Id>
  <Version>${version}</Version>
  <ProviderName>OfficeKit</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="OfficeKit"/>
  <Description DefaultValue="Connect an open Excel workbook to local OfficeKit automation."/>
  <IconUrl DefaultValue="${origin}/assets/officekit-excel-32.png"/>
  <SupportUrl DefaultValue="${origin}/support.html"/>
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="ExcelApi" MinVersion="1.8"/>
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${origin}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription"/>
    <Requirements>
      <bt:Sets DefaultMinVersion="1.1">
        <bt:Set Name="SharedRuntime" MinVersion="1.1"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <AllFormFactors>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="OfficeKit.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="OfficeKit.ShowTaskpane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Description"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>OfficeKitTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${origin}/assets/officekit-excel-32.png"/>
        <bt:Image id="Icon.32" DefaultValue="${origin}/assets/officekit-excel-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${origin}/assets/officekit-excel-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${origin}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="OfficeKit"/>
        <bt:String id="Button.Label" DefaultValue="OfficeKit"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residDescription" DefaultValue="Connect an open Excel workbook to local OfficeKit automation."/>
        <bt:String id="Button.Description" DefaultValue="Connect this workbook to OfficeKit."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export async function writeExcelManifest(paths, config, packageVersion) {
  const manifest = renderExcelManifest({
    addinId: config.addinId,
    port: config.port,
    packageVersion,
  });
  await writePrivateText(paths.manifest, manifest);
  return manifest;
}

function manifestVersion(packageVersion) {
  if (typeof packageVersion !== "string" || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/u.test(packageVersion)) {
    throw excelLiveError("invalid-manifest", "OfficeKit package version cannot be used in an Office manifest.");
  }
  const [major, minor, patch] = packageVersion.split("-", 1)[0].split(".");
  return `${major}.${minor}.${patch}.0`;
}

function assertUuid(value, label) {
  if (typeof value !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/iu.test(value)) {
    throw excelLiveError("invalid-manifest", `${label} must be a UUID.`);
  }
}
