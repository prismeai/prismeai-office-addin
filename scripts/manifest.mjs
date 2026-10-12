// Generates the two add-in manifests into dist/ :
//   manifest.xml          Word + Excel + PowerPoint (TaskPaneApp)
//   manifest-outlook.xml  Outlook (MailApp — Office requires a separate file)
//
//   node scripts/manifest.mjs [--host https://addin.example.com]
//        [--instance https://studio.example.com] [--api https://api.example.com/v2]
//        [--agent <id>] [--org <slug>] [--tools slug:office-mcp]
//        [--auth anonymous|sso] [--sso-provider <slug>] [--sso-client-id <guid>] [--sso-tenant <guid>]
//        [--version 1.0.1.0] [--name "Prisme.ai"]
//
// --host is where taskpane.html is served (HTTPS). The other options are baked
// into the task pane URL as managed config: they pre-fill and lock the connect
// form for every user of the deployment. Never pass a secret here.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

const args = {};
for (let i = 2; i < process.argv.length; i += 2) {
  const key = process.argv[i];
  const value = process.argv[i + 1];
  if (!key.startsWith('--') || value === undefined) {
    console.error(`Bad argument: ${key}`);
    process.exit(1);
  }
  args[key.slice(2)] = value;
}

const host = (args.host || 'https://localhost:3100').replace(/\/+$/, '');
if (!host.startsWith('https://')) {
  console.error('--host must be an https:// URL (Office refuses plain HTTP).');
  process.exit(1);
}
const name = args.name || 'Prisme.ai';
// Office wants a 4-part version >= 1.0 (0.x is rejected), so the package
// version is prefixed: 0.1.0 → 1.0.1.0. Bump it on every manifest change.
const version = args.version || `1.${pkg.version}`;

// Stable add-in identities — changing an id installs a separate add-in.
const IDS = {
  document: '675f35ff-9be5-4838-99cd-f26c5042b091',
  outlook: 'a09205ec-0aec-4633-9e14-f95ebc0b42c7',
};

const xml = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const managed = new URLSearchParams();
// CLI option → task pane query parameter (read by taskpane.js as managed config).
const MANAGED_PARAMS = {
  instance: 'instance',
  api: 'api',
  agent: 'agent',
  org: 'org',
  tools: 'tools',
  auth: 'auth',
  'sso-provider': 'sso_provider',
  'sso-client-id': 'sso_client',
  'sso-tenant': 'sso_tenant',
};
if (args.auth && !['anonymous', 'sso'].includes(args.auth)) {
  console.error('--auth must be "anonymous" or "sso".');
  process.exit(1);
}
if (args.auth === 'sso') {
  const missing = ['api', 'sso-provider', 'sso-client-id', 'sso-tenant'].filter((k) => !args[k]);
  if (missing.length) {
    console.error(`--auth sso also needs: ${missing.map((k) => '--' + k).join(', ')}`);
    process.exit(1);
  }
}
for (const [option, param] of Object.entries(MANAGED_PARAMS)) {
  if (args[option]) managed.set(param, args[option]);
}
const query = managed.toString();
const taskpaneUrl = xml(`${host}/taskpane.html${query ? '?' + query : ''}`);
const icon = (size) => xml(`${host}/icons/icon-${size}.png`);

// The chat iframe and the sign-in page live on the instance origin.
const appDomains = args.instance
  ? `\n  <AppDomains>\n    <AppDomain>${xml(new URL(args.instance).origin)}</AppDomain>\n  </AppDomains>`
  : '';

const description = 'Chat with a Prisme.ai agent that always knows the document you have open.';

const resources = `<Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${icon(16)}"/>
        <bt:Image id="Icon.32" DefaultValue="${icon(32)}"/>
        <bt:Image id="Icon.80" DefaultValue="${icon(80)}"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${taskpaneUrl}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="${xml(name)}"/>
        <bt:String id="Button.Label" DefaultValue="${xml(name)}"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Button.Tooltip" DefaultValue="Open the Prisme.ai agent next to your document."/>
      </bt:LongStrings>
    </Resources>`;

const iconSet = `<Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>`;

const button = (id, action) => `<Control xsi:type="Button" id="Prisme.${id}.Button">
                <Label resid="Button.Label"/>
                <Supertip>
                  <Title resid="Button.Label"/>
                  <Description resid="Button.Tooltip"/>
                </Supertip>
                ${iconSet}
                ${action}
              </Control>`;

// ── Word / Excel / PowerPoint ───────────────────────────────────────────────
const documentHost = (type) => `<Host xsi:type="${type}">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Prisme.${type}.Group">
                <Label resid="Group.Label"/>
                ${iconSet}
                ${button(
                  type,
                  `<Action xsi:type="ShowTaskpane">
                  <TaskpaneId>PrismeTaskpane</TaskpaneId>
                  <SourceLocation resid="Taskpane.Url"/>
                </Action>`
                )}
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>`;

const documentManifest = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${IDS.document}</Id>
  <Version>${version}</Version>
  <ProviderName>Prisme.ai</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${xml(name)}"/>
  <Description DefaultValue="${description}"/>
  <IconUrl DefaultValue="${icon(32)}"/>
  <HighResolutionIconUrl DefaultValue="${icon(64)}"/>
  <SupportUrl DefaultValue="https://www.prisme.ai"/>${appDomains}
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${taskpaneUrl}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      ${['Document', 'Workbook', 'Presentation'].map(documentHost).join('\n      ')}
    </Hosts>
    ${resources}
  </VersionOverrides>
</OfficeApp>
`;

// ── Outlook ─────────────────────────────────────────────────────────────────
// SupportsPinning (keep the pane open across messages) only exists in the
// V1_1 overrides, which must be nested inside a complete V1_0 block.
const mailSurface = (surface, id, pinning) => `<ExtensionPoint xsi:type="${surface}">
            <OfficeTab id="TabDefault">
              <Group id="Prisme.${id}.Group">
                <Label resid="Group.Label"/>
                ${button(
                  id,
                  `<Action xsi:type="ShowTaskpane">
                  <SourceLocation resid="Taskpane.Url"/>${pinning ? '\n                  <SupportsPinning>true</SupportsPinning>' : ''}
                </Action>`
                )}
              </Group>
            </OfficeTab>
          </ExtensionPoint>`;

const mailHosts = (pinning) => `<Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          ${mailSurface('MessageReadCommandSurface', 'MailRead', pinning)}
          ${mailSurface('MessageComposeCommandSurface', 'MailCompose', false)}
        </DesktopFormFactor>
      </Host>
    </Hosts>`;

const mailRequirements = (min) => `<Requirements>
      <bt:Sets DefaultMinVersion="${min}">
        <bt:Set Name="Mailbox"/>
      </bt:Sets>
    </Requirements>`;

const outlookManifest = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>${IDS.outlook}</Id>
  <Version>${version}</Version>
  <ProviderName>Prisme.ai</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${xml(name)}"/>
  <Description DefaultValue="Chat with a Prisme.ai agent that always knows the message you have open."/>
  <IconUrl DefaultValue="${icon(64)}"/>
  <HighResolutionIconUrl DefaultValue="${icon(128)}"/>
  <SupportUrl DefaultValue="https://www.prisme.ai"/>${appDomains}
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${taskpaneUrl}"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ${mailRequirements('1.3')}
    ${mailHosts(false)}
    ${resources}
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ${mailRequirements('1.5')}
    ${mailHosts(true)}
    ${resources}
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`;

const dist = join(root, 'dist');
mkdirSync(dist, { recursive: true });
writeFileSync(join(dist, 'manifest.xml'), documentManifest);
writeFileSync(join(dist, 'manifest-outlook.xml'), outlookManifest);
console.log(`dist/manifest.xml          Word, Excel, PowerPoint`);
console.log(`dist/manifest-outlook.xml  Outlook`);
console.log(`Task pane: ${host}/taskpane.html${query ? '?' + query : ''}`);
