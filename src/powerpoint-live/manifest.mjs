import { officeLiveError } from "../live/errors.mjs";
import { writePrivateText } from "../excel-live/state.mjs";
import { POWERPOINT_ADDIN_ID } from "./state.mjs";

export function powerpointBridgeOrigin(port) {
  if (!Number.isSafeInteger(port) || port < 1024 || port > 65535) {
    throw officeLiveError("invalid-state", "PowerPoint bridge port is invalid.");
  }
  return `https://localhost:${port}`;
}

export function renderPowerPointManifest({ addinId, port, packageVersion }) {
  assertUuid(addinId, "PowerPoint add-in ID");
  const origin = powerpointBridgeOrigin(port);
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
  <Description DefaultValue="Connect an open PowerPoint presentation to local OfficeKit automation."/>
  <IconUrl DefaultValue="${origin}/powerpoint/assets/officekit-powerpoint-32.png"/>
  <SupportUrl DefaultValue="${origin}/powerpoint/support.html"/>
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.8"/>
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${origin}/powerpoint/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription"/>
    <Requirements>
      <bt:Sets DefaultMinVersion="1.1">
        <bt:Set Name="PowerPointApi" MinVersion="1.8"/>
        <bt:Set Name="SharedRuntime" MinVersion="1.1"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="Presentation">
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
                    <TaskpaneId>OfficeKitPowerPointTaskpane</TaskpaneId>
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
        <bt:Image id="Icon.16" DefaultValue="${origin}/powerpoint/assets/officekit-powerpoint-32.png"/>
        <bt:Image id="Icon.32" DefaultValue="${origin}/powerpoint/assets/officekit-powerpoint-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${origin}/powerpoint/assets/officekit-powerpoint-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${origin}/powerpoint/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="OfficeKit"/>
        <bt:String id="Button.Label" DefaultValue="OfficeKit"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residDescription" DefaultValue="Connect an open PowerPoint presentation to local OfficeKit automation."/>
        <bt:String id="Button.Description" DefaultValue="Connect this presentation to OfficeKit."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export async function writePowerPointManifest(paths, config, packageVersion) {
  const manifest = renderPowerPointManifest({ addinId: POWERPOINT_ADDIN_ID, port: config.port, packageVersion });
  await writePrivateText(paths.manifest, manifest);
  return manifest;
}

function manifestVersion(packageVersion) {
  if (typeof packageVersion !== "string" || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/u.test(packageVersion)) {
    throw officeLiveError("invalid-manifest", "OfficeKit package version cannot be used in a PowerPoint manifest.");
  }
  const [major, minor, patch] = packageVersion.split("-", 1)[0].split(".");
  return `${major}.${minor}.${patch}.0`;
}

function assertUuid(value, label) {
  if (typeof value !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/iu.test(value)) {
    throw officeLiveError("invalid-manifest", `${label} must be a UUID.`);
  }
}
