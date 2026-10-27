# Prisme.ai Office Add-in

> Status: **Preview** (v0.x).

A Microsoft 365 add-in that opens a Prisme.ai agent in a pane next to the
document or email you are working on, in **Word, Excel, PowerPoint and Outlook**
(Windows, Mac and web). The agent knows the open document, the selection or the
open message, and — when its tools are enabled — works in it: find a passage in
a long document, compute over a large sheet, rewrite a paragraph or a slide,
fill an Excel range, prepare a reply or a meeting.

This guide is for administrators installing the add-in against a **self-hosted
Prisme.ai instance**.

- [How it works](#how-it-works)
- [Compatibility](#compatibility)
- [Before you begin](#before-you-begin)
- [Installation](#installation)
  - [Step 1 — Host the task pane](#step-1--host-the-task-pane)
  - [Step 2 — Prepare the Prisme.ai instance](#step-2--prepare-the-prismeai-instance)
  - [Step 3 — Generate the manifests](#step-3--generate-the-manifests)
  - [Optional — Single sign-on](#optional--single-sign-on)
  - [Step 4 — Pilot with one user](#step-4--pilot-with-one-user)
  - [Step 5 — Deploy to the organization](#step-5--deploy-to-the-organization)
  - [Step 6 — Verify](#step-6--verify)
- [Updating and removing](#updating-and-removing)
- [Troubleshooting](#troubleshooting)
- [Security and data flow](#security-and-data-flow)
- [Agent tools](#agent-tools)
- [Known limitations](#known-limitations)
- [Local development](#local-development)

## How it works

An Office add-in is two things: a **manifest** (an XML file that tells Office
where to load the pane from) and a **web page** served over HTTPS.

```
Office (Word / Excel / PowerPoint / Outlook)
 └─ task pane ── https://<addin-host>/taskpane.html      static files, this repo
     └─ iframe ── https://<platform-url>/embed/chat      your Prisme.ai instance
                     └─ API calls ── https://<api-url>   your Prisme.ai instance
```

- The **task pane** is a thin shell. It reads the document through Office.js,
  passes that context to the chat, and executes the agent's tools in the
  document.
- The **chat** is the hosted Prisme.ai widget, loaded in an iframe from your
  instance. The session token stays inside that iframe; the task pane never
  sees it.
- Nothing is sent to Prisme.ai SAS or to any third party. The add-in has no
  backend of its own.

You deploy **two manifests**: `manifest.xml` (Word, Excel, PowerPoint) and
`manifest-outlook.xml` (Outlook requires its own).

## Compatibility

| Client | Supported |
|---|---|
| Word, Excel, PowerPoint on the web | Yes |
| Word, Excel, PowerPoint on Windows (Microsoft 365 subscription, version 1704 or later) | Yes |
| Word, Excel, PowerPoint on Mac (version 15.34 or later) | Yes |
| Outlook on the web, new Outlook and classic Outlook on Windows, Outlook on Mac (Microsoft 365 subscription, Exchange Online mailbox) | Yes |
| Office 2016 / 2019 perpetual or volume-licensed | No |
| Outlook with POP3 / IMAP accounts, or Exchange on-premises mailboxes | No |
| Office and Outlook mobile apps | No |

The add-in has been validated against Microsoft's manifest schema. Test it in
your own environment with a pilot user ([Step 4](#step-4--pilot-with-one-user))
before a wide rollout.

## Before you begin

Collect or confirm every item below. Each one maps to a failure described in
[Troubleshooting](#troubleshooting).

**Prisme.ai instance**

- [ ] **Platform URL**, e.g. `https://studio.your-company.com`
- [ ] **API URL**, e.g. `https://api.your-company.com/v2`
      (both are shown in the console under **Settings → Access Tokens**)
- [ ] **Agent ID** of the agent users will talk to (on the agent's page)
- [ ] A sign-in mode:
      - **Anonymous** — the agent accepts anonymous access. If your
        organization restricts it, note the **org slug**.
      - **Single sign-on** — users are identified by their Microsoft account.
        Requires that your instance already signs users in with Microsoft
        Entra ID through an OIDC provider. See
        [Optional — Single sign-on](#optional--single-sign-on).

**Hosting**

- [ ] An **HTTPS host** for the task pane, with a certificate trusted by user
      devices, e.g. `https://office.your-company.com`. Any static web server
      works. The host name must be **stable**: it is written into the manifest.

**Microsoft 365 tenant**

- [ ] An admin account with **Global Administrator**, or **Exchange
      Administrator** plus **Application Administrator**. The role must be
      permanently active: roles activated through Privileged Identity
      Management (PIM) are not recognized by the Integrated apps page.
- [ ] Users have a licence that includes **Exchange Online** and an active
      mailbox: Microsoft 365 Business Basic / Standard / Premium, Office 365
      E1 / E3 / E5 / F3, Microsoft 365 E3 / E5 / F3, or the Education /
      Government equivalents. Microsoft 365 Apps for business / enterprise
      alone are not enough.
- [ ] The directory is in, or federated to, **Microsoft Entra ID**.
- [ ] Target groups are **top-level** groups. Nested groups and
      non-mail-enabled security groups are not supported by Microsoft's
      deployment service.

**Network** — from every user device, over HTTPS (443):

- [ ] `appsforoffice.microsoft.com` (Microsoft's Office.js runtime)
- [ ] the task pane host
- [ ] the Platform URL and the API URL, including WebSocket connections

**Workstation used to generate the manifests**

- [ ] Node.js 18 or later

## Installation

### Step 1 — Host the task pane

Publish these files from this repository, unchanged, at the root of your HTTPS
host:

```
taskpane.html
taskpane.js
tools/office.js  word.js  excel.js  powerpoint.js  outlook.js  index.js
icons/icon-16.png  icon-32.png  icon-64.png  icon-80.png  icon-128.png
popup-relay.html  popup-relay.js  auth-redirect.html  auth-redirect.js
vendor/msal-browser.min.js  msal-popup-relay.min.js  msal-redirect-bridge.min.js
```

The last two lines are used by single sign-on only: Microsoft's sign-in
library, and the sign-in windows of Office on the web.

`npm run package` builds a zip containing exactly these files in `dist/`.

Server requirements:

- **HTTPS only**, with a certificate trusted by user devices. Office refuses
  plain HTTP and untrusted certificates.
- **Do not send `X-Frame-Options`**, and do not send a
  `Content-Security-Policy` with a restrictive `frame-ancestors`. Office on the
  web loads the pane inside an iframe.
- **Icons must be cacheable**: no `Cache-Control: no-cache` or `no-store` on
  `/icons/`.
- Do not cache `taskpane.html`, `taskpane.js` and `tools/` for long, so updates reach
  users quickly.

Example nginx server block:

```nginx
server {
  listen 443 ssl;
  server_name office.your-company.com;
  ssl_certificate     /etc/ssl/office.your-company.com.crt;
  ssl_certificate_key /etc/ssl/office.your-company.com.key;

  root /var/www/prismeai-office-addin;

  location /icons/ { add_header Cache-Control "public, max-age=86400"; }
  location /       { add_header Cache-Control "no-cache"; }
}
```

Check it:

```bash
curl -sI https://office.your-company.com/taskpane.html
# Expect: HTTP 200, no X-Frame-Options header
curl -sI https://office.your-company.com/icons/icon-32.png
# Expect: HTTP 200, content-type: image/png
```

### Step 2 — Prepare the Prisme.ai instance

1. **Agent access.** For single sign-on, confirm your users have access to the
   agent. For anonymous sign-in, all of the following must be true:

   - The agent is **published** and its visibility is **Public**.
   - **Anonymous access is enabled on the agent's organization**, with a role.
     Note the organization slug: it is the `--org` option in Step 3. Without
     it, the anonymous visitor belongs to no organization and the agent is
     refused.
   - If the organization restricts anonymous access to a list of **allowed
     domains**, that list contains the task pane host (for example
     `prisme-addin.your-company.com`). The instance checks the host of the page
     that embeds the chat, which is the task pane.

2. **Framing.** The chat page `/embed/chat` must be allowed to load inside the
   task pane. Check what your console currently sends:

   ```bash
   curl -sI https://studio.your-company.com/embed/chat | grep -i content-security-policy
   ```

   - `frame-ancestors *` (the default): nothing to do.
   - `frame-ancestors 'self' …`: your instance sets `EMBED_ALLOWED_ORIGINS` on
     the console (platform) service. Add the task pane host **and** the Office
     on the web origins, because the browser checks every page in the framing
     chain, not only the task pane:

     ```
     EMBED_ALLOWED_ORIGINS=https://office.your-company.com,https://*.officeapps.live.com,https://*.office.com,https://*.cloud.microsoft,https://outlook.office.com,https://outlook.office365.com,https://your-tenant.sharepoint.com,https://your-tenant-my.sharepoint.com
     ```

     Restart the console service after changing it. The Microsoft origins above
     are a starting point: if the chat stays blank in Office on the web, the
     browser console names the exact origin that was refused (see
     [Troubleshooting](#troubleshooting)). Office desktop only needs the task
     pane host.

3. **Document tools (optional).** For the agent to read beyond the beginning
   of a long document, search it, compute over a sheet and write into the
   document, install the `office-mcp` workspace on the instance and attach it
   to the agent (slug `slug:office-mcp`). See [Agent tools](#agent-tools).
   Without it, the add-in is a chat that knows the beginning of the open
   document.

### Step 3 — Generate the manifests

From a clone of this repository:

```bash
npm run manifest -- \
  --host https://office.your-company.com \
  --instance https://studio.your-company.com \
  --api https://api.your-company.com/v2 \
  --agent <agent-id>

npm run validate
```

This writes `dist/manifest.xml` and `dist/manifest-outlook.xml`, and validates
both against Microsoft's schema. Both commands must end without error.

| Option | Required | Meaning |
|---|---|---|
| `--host` | Yes | HTTPS origin serving the task pane (Step 1). |
| `--instance` | Recommended | Platform URL. |
| `--api` | Recommended | API URL. |
| `--agent` | Recommended | Agent ID. |
| `--org` | No | Org slug, if anonymous access requires it. |
| `--tools` | No | Office tools workspace slug. Omit to disable the document tools. |
| `--auth` | No | `anonymous` (default) or `sso`. |
| `--sso-provider` | With `--auth sso` | Slug of the Entra ID SSO provider on the instance. |
| `--sso-client-id` | With `--auth sso` | Application (client) ID of the Entra app. |
| `--sso-tenant` | With `--auth sso` | Directory (tenant) ID. |
| `--name` | No | Display name in Office (default `Prisme.ai`). |
| `--version` | No | Manifest version, four numbers, `1.0.0.0` or higher. |

The instance options are written into the task pane URL inside the manifest.
They pre-fill the connection form and **lock** it for every user. If you omit
them, each user is asked for the Platform URL, API URL and Agent ID on first
launch.

> These values are readable by every user the add-in is assigned to. **Never
> put a token, password or API key in them.**

### Optional — Single sign-on

With single sign-on, the add-in uses the Microsoft account the user is already
signed in with in Office. In desktop Office there is no sign-in window: Office
hands the add-in a Microsoft ID token. In Office on the web, clicking
**Connect** opens a small Prisme.ai window: the user clicks **Continue**, and
Microsoft's sign-in window opens (it closes by itself if the user is already
signed in to Microsoft). The sign-in is then reused as long as the browser tab
stays open. Your instance exchanges the ID token for a
Prisme.ai session for that user, and the agent runs with the user's own
identity and permissions.

**Requirements**

- Your instance signs users in with **Microsoft Entra ID** through an **OIDC**
  SSO provider, declared in AI Governance or in the api-gateway's
  `authProviders.config.yml`. Other identity providers, SAML, and the Azure
  sign-in configured through the api-gateway's `AZURE_AD_*` environment
  variables are not supported by this mode.
- A Prisme.ai version that provides the token exchange endpoint
  (`POST /v2/login/token-exchange`) and the **Allow Token Exchange** option on
  SSO providers.
- An up-to-date Microsoft 365 build of Office. Single sign-on relies on
  Office's Nested App Authentication; on a build without it, the pane reports
  that Office must be updated.

**1. Entra ID — extend the app registration your instance already uses**

Use the **same app registration** as the instance's SSO provider. Entra issues
a different user identifier to each app registration, and your instance matches
users on that identifier first: with a separate registration, users would only
be found by their email address.

1. In the [Microsoft Entra admin center](https://entra.microsoft.com), open
   **App registrations** and select the application used by Prisme.ai SSO.
2. Under **Authentication**, add the platform **Single-page application** with
   these two redirect URIs — the task pane host, without `https://` for desktop
   Office, and the sign-in landing page for Office on the web:

   ```
   brk-multihub://office.your-company.com
   https://office.your-company.com/auth-redirect.html
   ```

   They must be of type **Single-page application**, not **Web**.

3. Under **API permissions**, confirm the Microsoft Graph delegated permission
   **User.Read** is present and grant admin consent, so users are not prompted.
4. Under **Token configuration**, add the optional **ID** token claims
   `email`, `given_name` and `family_name` if they are not there yet. Your
   instance reads the user's email and name from the ID token.
5. Note the **Application (client) ID** and the **Directory (tenant) ID**.

**2. Prisme.ai — allow token exchange on the SSO provider**

Open the SSO provider where it is declared: in the console (AI Governance →
SSO), or in the api-gateway's `authProviders.config.yml`, under the provider's
`config`. A change to the file takes effect when the api-gateway restarts.

1. Check that the issuer (**Issuer** in the console, `issuer` in the file) is
   `https://login.microsoftonline.com/<tenant-id>/v2.0`, with your tenant ID.
   This is the issuer of the ID tokens the add-in obtains. If your provider is
   configured with a different issuer (for example the older
   `https://sts.windows.net/<tenant-id>/`), do not change it on a production
   instance without testing web sign-in first: contact Prisme.ai.
2. Turn on **Allow Token Exchange** (`allowTokenExchange: true` in the file).
   The issuer is then required: in the console the provider cannot be saved
   without it, and a provider in the file is skipped at startup, which also
   stops web sign-in through it.
3. Leave the token exchange audiences empty (no `tokenExchangeAudiences` in the
   file): the provider's own client ID is accepted, which is the app
   registration above.
4. Note the provider's **slug** (in the file, its key under `providers:`).

See [Token Exchange](https://docs.prisme.ai/self-hosting/entreprise/authentication#token-exchange-headless-clients)
in the Prisme.ai documentation.

**3. API gateway — allow the task pane host to call the API**

The exchange is called from the task pane, so its origin must be allowed by
the API. Add it to the `CORS_ADDITIONAL_ALLOWED_ORIGINS` environment variable
of the **api-gateway** service (comma-separated), then restart the service:

```
CORS_ADDITIONAL_ALLOWED_ORIGINS=https://office.your-company.com
```

**4. Generate the manifests with the SSO options**

```bash
npm run manifest -- \
  --host https://office.your-company.com \
  --instance https://studio.your-company.com \
  --api https://api.your-company.com/v2 \
  --agent <agent-id> \
  --auth sso \
  --sso-provider <provider-slug> \
  --sso-client-id <application-client-id> \
  --sso-tenant <directory-tenant-id>
```

None of these values is a secret.

### Step 4 — Pilot with one user

Deploy to yourself first, using the same procedure as
[Step 5](#step-5--deploy-to-the-organization) with only your own account as the
audience. Then run the checks in [Step 6](#step-6--verify) in each application
your users rely on, on desktop and on the web.

To try a manifest without an admin deployment, sideload it yourself. The
procedure below uses Office on the web, which needs nothing installed. The
task pane must be reachable from your browser (your hosted URL, or
`npm run serve` for a local manifest).

#### Word, Excel and PowerPoint

These three applications share one file: `dist/manifest.xml`.

1. Open a document in Word on the web, a workbook in Excel on the web, or a
   presentation in PowerPoint on the web.
2. On the **Home** tab, select **Add-ins → More Add-ins → My Add-ins → Upload
   My Add-in**.
3. Choose `dist/manifest.xml`.
4. The **Prisme.ai** button appears on the **Home** tab.

Repeat these steps **in each application**: a manifest uploaded in Word is not
available in Excel or PowerPoint. The upload is kept in the browser, so it has
to be done again in another browser or after clearing site data.

#### Outlook

Outlook uses its own file: `dist/manifest-outlook.xml`.

1. Open <https://aka.ms/olksideload>. Outlook on the web opens with the
   **Add-Ins for Outlook** dialog.
2. Select **My add-ins → Custom Addins → Add a custom add-in → Add from File**.
3. Choose `dist/manifest-outlook.xml`.
4. Open an email: the **Prisme.ai** button is in the message ribbon, or under
   **Apps**.

This requires an Exchange Online mailbox (Microsoft 365). It does not work with
IMAP or POP accounts.

#### What to try in each application

| Application | Document context | Agent tools |
|---|---|---|
| Word | *"Summarize this document."* | *"Give me the outline of this document"*, *"Where is the penalty clause?"*, then *"Add a conclusion at the end."* |
| Excel | *"Summarize this sheet."* | *"Which sheets does this workbook contain?"*, *"What is the total amount per region?"*, then *"Write this table starting at B2."* |
| PowerPoint | *"Summarize this presentation."* | *"What is on slide 2?"*, *"Rewrite the title of slide 2"*, then *"Add a slide after slide 2 with the same design."* |
| Outlook, reading an email | *"What is this email about?"* | *"Prepare a reply"*, then *"Create a meeting tomorrow at 10 with the sender"*: an unsent form must open each time. |
| Outlook, writing an email | Select some text, then the **quote** button. | *"Insert a closing sentence."* |

In every application, also select some text or cells, select the **quote**
button at the top of the pane, and check the selection appears in the message
box.

Keep in mind while testing:

- **Keep the pane open** while a tool runs. Otherwise the agent waits 30 seconds,
  then reports that no add-in is attached.
- **In Outlook, pin the pane.** Otherwise the agent keeps answering about the
  previous email when you open another one.
- **The agent tools need the `office-mcp` workspace** (see
  [Agent tools](#agent-tools)). If the document context works but a tool
  fails, check first that `office-mcp` is imported on the instance and added to
  the agent as an MCP server.
- **In PowerPoint, name the slide** ("slide 4"): the add-in cannot tell which
  slide is on screen.
- Only these four applications are supported. The pane behaves as a plain chat
  in any other Office application, and the Office mobile apps are not
  supported.

For the desktop applications, follow Microsoft's guides:
[sideload an Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
and
[sideload an Outlook add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).
The wording of these menus changes over time; Microsoft's guides are the
reference.

### Step 5 — Deploy to the organization

Do this **twice**: once with `manifest.xml`, once with `manifest-outlook.xml`.

1. Sign in to the [Microsoft 365 admin center](https://admin.microsoft.com).
2. Go to **Settings → Integrated apps**.
3. Select **Upload custom apps**, and choose **Office Add-in** as the app type.
4. Choose the option to upload a manifest file from your device, and select
   the file. Upload the file rather than giving its URL.
5. Choose who gets the add-in: yourself only, specific users or groups, or the
   entire organization.
6. Review the summary and finish the deployment.

The wording of these screens changes over time; Microsoft's reference is
[Get started with the integrated apps portal](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).

The add-in can take **up to 24 hours** to reach every assigned user. Users may
need to restart the Office application.

**Sovereign and government clouds** (GCC, GCC High, DoD, 21Vianet) do not have
the Integrated apps page. Use **Settings → Add-ins → Deploy Add-in → Upload
custom apps** instead
([Centralized Deployment](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)).

### Step 6 — Verify

As an assigned user:

1. **Word, Excel or PowerPoint** — open a document. A **Prisme.ai** button
   appears on the **Home** tab. Select it: the pane opens and shows the chat
   (or the connection form, if the manifest carries no instance options).
2. Ask *"Summarize this document."* The answer must reflect the content of the
   open file.
3. Select some text, select the **quote** button at the top of the pane, and
   check the selection appears in the message box.
4. **Outlook** — open an email, select the **Prisme.ai** button in the message
   ribbon (or under **Apps**), and ask *"What is this email about?"*
5. If single sign-on is enabled, the chat opens without any sign-in prompt
   (Microsoft may ask for consent once), and the conversation appears under the
   user's own account on the instance.
6. If the document tools are enabled, ask the agent to insert a sentence at the
   cursor and confirm it appears in the document.

## Updating and removing

**New version of the task pane** (`taskpane.html`, `taskpane.js`, `tools/`): replace the
files on the host. No change in the admin center. Users get the new version the
next time they open the pane.

**Change of host, instance, agent or any manifest option**: regenerate the
manifests with a **higher `--version`**, then update each one in **Integrated
apps** (select the add-in, then update it with the new file). A manifest
uploaded with an unchanged version number is not picked up.

**Change of host name**: the host is part of the add-in's configuration, so
plan it as a redeployment and keep the old host serving until the new manifest
has propagated.

**Remove**: in **Integrated apps**, select the add-in and remove it. Do this
for both manifests. See
[Manage deployment of add-ins](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins).

## Troubleshooting

| Symptom | Likely cause | What to do |
|---|---|---|
| The Prisme.ai button does not appear | Propagation delay | Wait up to 24 hours, then restart the Office application. |
| | User not eligible | Check the user has a supported licence, an Exchange Online mailbox, and is assigned directly or through a top-level group. |
| | Add-ins disabled in the tenant | In Exchange Online PowerShell, `Get-OrganizationConfig \| fl AppsForOfficeEnabled` must return `True`. |
| | Unsupported Office version | See [Compatibility](#compatibility). |
| Deployment fails in Integrated apps | Admin role activated through PIM | Use an account whose role is permanently active. |
| The pane opens but is blank, or shows "add-in error" | Task pane host unreachable or certificate not trusted | Open `https://<addin-host>/taskpane.html` in a browser on the same device. Fix the network path or certificate. |
| | `X-Frame-Options` on the task pane host (Office on the web) | Remove the header ([Step 1](#step-1--host-the-task-pane)). |
| The pane shows the header buttons but the chat area is blank | The instance refuses to be framed | Open the browser developer tools (Office on the web) and look for a `frame-ancestors` error. Add the refused origin to `EMBED_ALLOWED_ORIGINS` ([Step 2](#step-2--prepare-the-prismeai-instance)). |
| | Platform URL wrong or unreachable | Open `<platform-url>/embed/chat` in a browser on the same device. |
| The chat loads but messages fail | API URL wrong or blocked, or WebSocket blocked by a proxy | Check the API URL and the proxy rules. |
| | Agent not available anonymously | Check that the agent is published and Public, and the `--org` slug. |
| The message is sent and nothing comes back (anonymous) | The instance refused the anonymous session, and this version of the chat does not display the refusal | In the browser developer tools, read the response of the `messages/stream` and `login/anonymous` requests. Most often the task pane host is missing from the organization's allowed domains for anonymous access. |
| The pane returns to the connection form with "Could not reach the Prisme.ai API" (SSO) | The task pane host is not allowed by the API | Add it to `CORS_ADDITIONAL_ALLOWED_ORIGINS` on the api-gateway and restart it. |
| "The Prisme.ai instance refused the sign-in" (SSO) | Token exchange not enabled, wrong provider slug, or issuer mismatch | Check **Allow Token Exchange**, the slug, and that the provider's issuer is exactly `https://login.microsoftonline.com/<tenant-id>/v2.0`. |
| | The ID token comes from another app registration | Use the client ID of the app registration configured on the SSO provider. |
| "This version of Office does not support single sign-on" | Office build without Nested App Authentication | Update Office, or use Office on the web. |
| Microsoft shows an error about the redirect URI or the application (SSO) | The `brk-multihub://` or `auth-redirect.html` redirect URI is missing, does not match the task pane host, or is declared under **Web** | Fix the Single-page application redirect URIs on the app registration. |
| "Error opening popup window" (SSO, Office on the web) | The browser blocks the sign-in windows | Allow pop-ups for the Office site and the task pane host, then click **Connect** again. |
| "Office.js did not load" (SSO) | `appsforoffice.microsoft.com` is blocked by a proxy or firewall | Open `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` on the same device and allow it. |
| Users are signed in but appear as new, empty accounts (SSO) | The client ID belongs to a different app registration than the SSO provider's, or the ID token has no `email` claim | Use the provider's app registration and add the optional claims. |
| The connection form appears although the manifest was configured | The manifest was generated without `--agent` | Regenerate with the instance options and a higher `--version`. |
| The agent does not know the document | Empty document, or content not loaded yet | Click in the document, then ask again. Context is refreshed when the selection changes. |
| The agent answers about the previous email | The pane is not pinned (Outlook) | Pin the pane so it follows the selected message. |
| The agent says it cannot edit the document | No tools workspace configured | Install it on the instance and regenerate the manifests with `--tools`. |
| The agent does not use a tool added to `office-mcp`, or calls one that no longer exists | The agent keeps the tool list of an MCP server in a cache | Save the agent again in Agent Factory, then start a new conversation. |
| A tool fails with "unknown tool" or "Tool not found" | The task pane files and the `office-mcp` workspace are not from the same version | Update both: the files on the task pane host, and the workspace on the instance. |
| "No Office add-in attached" | The pane was closed, reloaded or still loading when the tool was called | Open the pane, wait for the chat, and ask again. |
| PowerPoint slides come back flagged `incomplete`, or a read is slow, right after opening a long presentation | PowerPoint on the web is still loading the far slides | Ask again a moment later. |
| PowerPoint on the web shows another slide for a moment while the agent reads | The read brings a slide PowerPoint has not loaded yet on screen, so that it loads | Nothing to do: the slides the user had selected are selected back once the read is done. |
| Users still see an old version after an update | Office cached the manifest or the pane | Check the manifest version was raised. Then [clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache) on the device. |
| A new conversation starts every time the pane is opened | The chat runs in a third-party iframe, and browsers isolate or clear its storage | Can happen in this version, depending on the browser or webview. |

On Windows, the pane runs in Microsoft Edge WebView2; on Mac, in Safari's
WebKit. On a device where WebView2 is managed by Windows Information Protection,
`msedgewebview2.exe` must be on the protected apps list.

## Security and data flow

- **What the add-in reads**: as context for each message, the beginning of the
  open document (about 8,000 characters: the first paragraphs in Word, the top
  of the active sheet in Excel, the first slides in PowerPoint), the current
  selection, and in Outlook the open message (subject, sender, recipients,
  body). When the tools are enabled, also any part of the open document the
  agent asks for, including the other worksheets of the workbook. It does not
  read other files, other mailbox folders, or the calendar.
- **Where it goes**: only to the Prisme.ai instance in the manifest: the
  context when the user sends a message, and the result of a tool when the
  agent calls one. The task pane host serves static files and receives no
  document content.
- **Permissions requested from Office**: `ReadWriteDocument` (Word, Excel,
  PowerPoint) and `ReadWriteItem` (Outlook, the open item only). The add-in
  does not use Microsoft Graph to read content. With anonymous sign-in it needs
  no Entra app registration.
- **Single sign-on**: the Microsoft ID token is obtained through Office and
  sent only to your instance's API, which verifies its signature, issuer and
  audience before opening a session. The resulting session token is passed to
  the chat iframe and kept in memory only. No client secret is involved.
- **Writing**: document edits happen only when the agent calls a tool, under
  the tool permissions configured on the agent. The add-in never sends an
  email or an invitation: a reply, a new email or a meeting is opened in
  Outlook's own form for the user to review and send.
- **Untrusted content**: a document or email can contain text written to
  manipulate the agent. Keep confirmation enabled on the agent for tools that
  modify content, and have users review changes.

## Agent tools

The tools are served by the **`office-mcp`** workspace, an MCP server that runs
on the Prisme.ai instance. It holds the tool catalog, checks the caller, and
relays each call to the user's task pane, which executes it through Office.js on
the open document. Nothing is read or written through Microsoft Graph.

`office-mcp` is a regular Prisme.ai workspace: it is imported, configured and
updated like any other. It is maintained in the Prisme.ai workspaces repository
(`workspaces/office-mcp`), not in this one.

To enable the tools:

1. **Import `office-mcp`** on the instance.
2. **Check its configuration.** The workspace configuration must contain the
   `mcpTools` catalog (the tools listed below). If the agent sees no tools after
   the import, the catalog is missing: copy it from the workspace's `index.yml`.
3. **Add the MCP server to the agent**, with the URL
   `<API URL>/workspaces/slug:office-mcp/webhooks/mcp`.
4. **Point the add-in at the workspace**: generate the manifests with
   `--tools slug:office-mcp`, or type it in **Advanced → Office tools
   workspace** on the connection form.

A call only succeeds while the user has the Prisme.ai pane open: the workspace
waits 30 seconds for the pane to answer, then reports that no add-in is
attached. Set the agent's tool permissions so that the tools that modify
content ask for confirmation.

### Reading: three tools, the same in every application

A document can be a three-hundred-page contract or a workbook of several
hundred thousand rows, so no tool returns "the document". The agent works the
way one works on a large file: look at the structure, read one part, search.

| Tool | Arguments | What it returns |
|---|---|---|
| `office_outline` | `{ from? }` | What the document is made of and how big it is, without its content. Word: number of paragraphs and the headings with their position. Excel: each worksheet with its used range, size and first row. PowerPoint: number of slides and the title of each. Outlook: subject, sender, recipients, length. |
| `office_read` | `{ target? }` | One part of the document, from the start when `target` is omitted. |
| `office_search` | `{ query, in?, limit? }` | The number of matches and, for the first ones (20 by default, 50 at most), their position and the surrounding text, or the whole row in Excel. Plain text, case-insensitive. |

A position is a short string, returned by one tool and accepted as is by the
others:

| Application | Position | Meaning |
|---|---|---|
| Word | `p:120`, `p:120-180` | Paragraph number, or a span of paragraphs |
| Excel | `Sales!A1:L50`, `A1:C20`, `Sales` | A range on a sheet, a range on the active sheet, the top of a sheet |
| PowerPoint | `slide:12`, `slide:12-20` | Slide number, or a span of slides |
| Outlook | `c:12001` | Character position in the body of the email |

Every answer is bounded: a read returns about 12,000 characters of content at
most, and no tool result exceeds 40,000 characters. When there is more, the
result carries `next`, the position to pass to the same tool to continue. A
search or an aggregation that runs out of time returns what it has, flagged
`partial`, with the position to resume from.

### Selection and writing

| Tool | Arguments | Applications |
|---|---|---|
| `office_get_selection` | — | All (Outlook: compose mode only) |
| `office_insert_text` | `{ text, location?: 'replace' \| 'start' \| 'end' }` | All (`location` applies to Word; elsewhere the selection is replaced) |
| `word_replace_text` | `{ search, replace, paragraph?, all? }` — `search` is the exact text, one line of 255 characters at most; `paragraph` limits the change to that paragraph | Word |
| `word_insert_paragraph` | `{ text, after_paragraph?, after?, heading? }` — after a paragraph number, after the paragraph containing `after`, or at the end | Word |
| `word_add_comment` | `{ anchor, comment, paragraph? }` | Word |
| `excel_compute` | `{ op, value?, group_by?, filter?, sheet? }` — `op` is sum, count, average, min or max; columns are named by header or by letter; `filter` is `{ column, equals }`. Computes over the whole table without handing its rows to the agent | Excel |
| `excel_write_range` | `{ address, values, sheet? }` — `address` is the top-left cell, `values` a 2D array | Excel |
| `excel_add_sheet` | `{ name }` | Excel |
| `excel_add_chart` | `{ address, type?, title?, sheet? }` — `type` is column, bar, line, pie, area or scatter | Excel |
| `powerpoint_write_slide` | `{ text, slide, shape? }` — `shape` is a shape `id` (from `office_read` or `office_search`) whose text is replaced; without it a new text box is added | PowerPoint |
| `powerpoint_add_slide` | `{ like?, position?, duplicate? }` — adds an empty slide with the layout of slide `like` (or a full copy with `duplicate`), at slide number `position` (defaults to the end); returns the new slide's shapes | PowerPoint |
| `powerpoint_delete_slide` | `{ slide }` | PowerPoint |
| `powerpoint_move_slide` | `{ slide, position }` | PowerPoint |
| `outlook_draft_reply` | `{ body, replyAll? }` — opens an unsent draft | Outlook (read mode) |
| `outlook_draft_new` | `{ body, to?, cc?, subject? }` — opens an unsent new message | Outlook |
| `outlook_draft_meeting` | `{ subject, attendees?, optional?, start?, end?, location?, body? }` — opens an unsent meeting; `start` and `end` are local times such as `2026-10-14T10:00` | Outlook |

After the catalog of `office-mcp` changes, save the agent again in Agent
Factory: the agent keeps the list of tools of an MCP server in a cache that is
only refreshed when the agent is saved.

## Known limitations

- **Single sign-on works with Microsoft Entra ID only** (OIDC provider on the
  instance). With any other identity provider, users are anonymous to the
  instance.
- **The Office tools workspace (`office-mcp`) is not shipped in this
  repository.** It lives in the Prisme.ai workspaces repository and is deployed
  separately; the add-in and the workspace must expose the same list of tools.
- **Text only.** The tools read and write text, cell values and formulas.
  Images, charts, tables in slides, speaker notes, formatting, headers, footers
  and footnotes are not read. In Word, the cells of a table are read as
  paragraphs.
- **Large documents are read in parts**, never as a whole: about 12,000
  characters per read. The agent is given the beginning of the document as
  context and uses the outline, read and search tools for the rest.
- **Search is plain text**: case-insensitive, no wildcard, no regular
  expression, one line of 255 characters at most. In Word and PowerPoint a
  match cannot span two paragraphs or two shapes.
- **`excel_compute` handles one table per sheet**: the first row of the used
  range is taken as the headers, the filter is a single equality, and the 50
  largest groups are returned.
- **A tool call has 25 seconds.** A search or an aggregation that needs more
  returns a partial result with the position to resume from. Timings depend on
  the host: PowerPoint on the web is the slowest, at a few hundred
  milliseconds per round trip.
- **PowerPoint: the slide on screen is not detected.** Writing takes a slide
  number; the user has to name the slide ("slide 12").
- **PowerPoint: a long presentation that was just opened is read slowly at
  first.** For a few minutes PowerPoint on the web can refuse the text of the
  far slides (observed past slide 100 of a 150-slide deck), or list no shape
  at all for them. In the second case, a read or an outline brings the first
  such slide on screen for a moment (PowerPointApi 1.5) so that PowerPoint
  loads it, reads again, then selects the user's slides back. Slides still not
  loaded come back flagged `incomplete` and read normally a moment later.
- **PowerPoint on the web: a blank slide looks like a slide not loaded yet.**
  It is flagged `incomplete` too, and a read that covers it takes about
  1.5 s more. Desktop clients load every slide: a blank slide is just blank.
- **The tools have been tested on Office on the web.** The desktop clients load the same code;
  pilot them before a rollout ([Step 4](#step-4--pilot-with-one-user)).
- **Word paragraph numbers shift when paragraphs are inserted or deleted**:
  after such a change, the agent has to read or search again before using a
  position.
- **Outlook works on the open message only**: no search across the mailbox, no
  reading of the calendar (a meeting can be prepared, but free slots are not
  looked up). PDF and image attachments are not read.
- **Several panes open at once** (Word, Excel and PowerPoint side by side):
  each pane answers the tool calls of its own conversation only, from the
  second message of that conversation on. During the first message of a new
  conversation a pane cannot tell yet: a tool of another application is left
  to the right pane, but a tool shared by all applications (`office_read`,
  `office_outline`, `office_search`) can be answered by another open pane, from
  another document. Close the panes that are not in use, or send a first
  message before asking about the document.
- **Not available in Teams.** Teams does not load Office add-ins; it needs its
  own app package.
- **No mobile clients.**

## Local development

```bash
npm run certs      # once: install a trusted localhost certificate
npm run manifest   # manifests pointing at https://localhost:3100
npm run serve      # serve the task pane over HTTPS on port 3100
npm run check      # syntax check
npm test           # unit tests (Node's built-in runner, no dependency)
npm run validate   # validate the manifests against Microsoft's schema
```

Then sideload `dist/manifest.xml` (see [Step 4](#step-4--pilot-with-one-user)),
or run `npx office-addin-debugging start dist/manifest.xml` for Office desktop.

There is no build step: the task pane is plain HTML and JavaScript.

Served from `localhost` or `127.0.0.1`, or with `debug=1` added to the task
pane URL, the pane runs in debug mode: each tool result carries a `diag` entry
with the time spent in Office. It is off on a production host.

Layout:

- `taskpane.html`, `taskpane.js` — the pane: connection form, sign-in, the chat
  iframe, the document context pushed to the agent, and the reception of the
  agent's tool calls (one at a time, with a time limit).
- `tools/` — what the tools do, one file per application (`word.js`,
  `excel.js`, `powerpoint.js`, `outlook.js`), `office.js` for what they share
  and `index.js` for the tools common to all applications and the table of
  every tool by name. Plain scripts loaded by `taskpane.html`, in that order.
- `scripts/manifest.mjs` — generates both XML manifests.
- `test/` — unit tests. `harness.mjs` loads the pane's scripts against in-memory
  documents that behave like the Office hosts (data only after `load` and
  `sync`, a rejected PowerPoint batch, a host that stops answering) and counts
  what is pulled out of the document, so a test can check that a tool stays
  cheap on a large file, not only that its answer is right.
- `vendor/` — `@azure/msal-browser` 5.26.0, unmodified: `lib/msal-browser.min.js`,
  `lib/popup-relay/msal-popup-relay.min.js` and
  `lib/redirect-bridge/msal-redirect-bridge.min.js` from the npm package, all
  three of the same version (licence alongside). Loaded only when single
  sign-on is enabled.
- `popup-relay.html`, `popup-relay.js`, `auth-redirect.html`, `auth-redirect.js` — sign-in on
  Office on the web, where Office does not broker it. A popup opened by the
  task pane's iframe cannot hand Microsoft's answer back (browsers partition
  its storage, and Microsoft's pages isolate their opener), so MSAL opens the
  relay page top-level, which opens Microsoft's window and relays the answer
  to the pane. `auth-redirect.html` is the redirect URI of that window.
- `PREREQUIS.md` — internal design notes (French).

## License

See [LICENSE](LICENSE).
