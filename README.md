# Prisme.ai Office Add-in

> Status: **Preview** (v0.x).

A Microsoft 365 add-in that opens a Prisme.ai agent in a pane next to the
document or email you are working on, in **Word, Excel, PowerPoint and Outlook**
(Windows, Mac and web). The agent reads the open document, the selection or the
open message, and — when its tools are enabled — writes back: insert text, fill
an Excel range, prepare a reply draft.

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
icons/icon-16.png  icon-32.png  icon-64.png  icon-80.png  icon-128.png
vendor/msal-browser.min.js        (Microsoft's sign-in library, used by single sign-on only)
```

`npm run package` builds a zip containing exactly these files in `dist/`.

Server requirements:

- **HTTPS only**, with a certificate trusted by user devices. Office refuses
  plain HTTP and untrusted certificates.
- **Do not send `X-Frame-Options`**, and do not send a
  `Content-Security-Policy` with a restrictive `frame-ancestors`. Office on the
  web loads the pane inside an iframe.
- **Icons must be cacheable**: no `Cache-Control: no-cache` or `no-store` on
  `/icons/`.
- Do not cache `taskpane.html` and `taskpane.js` for long, so updates reach
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

1. **Agent access.** For anonymous sign-in, confirm the agent can be used
   anonymously (and note the org slug if anonymous access is scoped to an
   organization). For single sign-on, confirm your users have access to the
   agent.

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

3. **Document actions (optional).** For the agent to write into documents,
   install the `office-mcp` workspace on the instance and attach it to the
   agent (slug `slug:office-mcp`). See [Agent tools](#agent-tools). Without it, the
   add-in is a chat that knows the open document.

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
| `--tools` | No | Office tools workspace slug. Omit to disable document actions. |
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
signed in with in Office. There is no sign-in window: Office hands the add-in a
Microsoft ID token, and your instance exchanges it for a Prisme.ai session for
that user. The agent then runs with the user's own identity and permissions.

**Requirements**

- Your instance signs users in with **Microsoft Entra ID** through an **OIDC**
  SSO provider. Other identity providers and SAML are not supported by this
  mode.
- A Prisme.ai version that provides the token exchange endpoint
  (`POST /v2/login/token-exchange`) and the **Allow Token Exchange** option on
  SSO providers.
- An up-to-date Microsoft 365 build of Office. Single sign-on relies on
  Office's Nested App Authentication; on a build without it, the pane reports
  that Office must be updated.

**1. Entra ID — extend the app registration your instance already uses**

Use the **same app registration** as the instance's SSO provider. Entra issues
a different user identifier to each app registration, and your instance matches
users on that identifier: a separate registration would not resolve to the same
Prisme.ai accounts.

1. In the [Microsoft Entra admin center](https://entra.microsoft.com), open
   **App registrations** and select the application used by Prisme.ai SSO.
2. Under **Authentication**, add the platform **Single-page application** with
   this redirect URI — the task pane host, without `https://`:

   ```
   brk-multihub://office.your-company.com
   ```

3. Under **API permissions**, confirm the Microsoft Graph delegated permission
   **User.Read** is present and grant admin consent, so users are not prompted.
4. Under **Token configuration**, add the optional **ID** token claims
   `email`, `given_name` and `family_name` if they are not there yet. Your
   instance reads the user's email and name from the ID token.
5. Note the **Application (client) ID** and the **Directory (tenant) ID**.

**2. Prisme.ai — allow token exchange on the SSO provider**

In the console, open the SSO provider (AI Governance → SSO):

1. Check that **Issuer** is `https://login.microsoftonline.com/<tenant-id>/v2.0`,
   with your tenant ID. This is the issuer of the ID tokens the add-in obtains.
   If your provider is configured with a different issuer (for example the
   older `https://sts.windows.net/<tenant-id>/`), do not change it on a
   production instance without testing web sign-in first: contact Prisme.ai.
2. Turn on **Allow Token Exchange**.
3. Leave the token exchange audiences empty: the provider's own client ID is
   accepted, which is the app registration above.
4. Note the provider's **slug**.

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

To try a manifest without an admin deployment, you can also
[sideload it](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
(Outlook: [sideload an Outlook add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)).

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
6. If document actions are enabled, ask the agent to insert a sentence at the
   cursor and confirm it appears in the document.

## Updating and removing

**New version of the task pane** (`taskpane.html`, `taskpane.js`): replace the
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
| | Agent not available anonymously | Check the agent's access settings and the `--org` slug. |
| The pane returns to the connection form with "Could not reach the Prisme.ai API" (SSO) | The task pane host is not allowed by the API | Add it to `CORS_ADDITIONAL_ALLOWED_ORIGINS` on the api-gateway and restart it. |
| "The Prisme.ai instance refused the sign-in" (SSO) | Token exchange not enabled, wrong provider slug, or issuer mismatch | Check **Allow Token Exchange**, the slug, and that the provider's issuer is exactly `https://login.microsoftonline.com/<tenant-id>/v2.0`. |
| | The ID token comes from another app registration | Use the client ID of the app registration configured on the SSO provider. |
| "This version of Office does not support single sign-on" | Office build without Nested App Authentication | Update Office, or use Office on the web. |
| Microsoft shows an error about the redirect URI or the application (SSO) | The `brk-multihub://` redirect URI is missing or does not match the task pane host | Fix the Single-page application redirect URI on the app registration. |
| Users are signed in but appear as new, empty accounts (SSO) | The client ID belongs to a different app registration than the SSO provider's, or the ID token has no `email` claim | Use the provider's app registration and add the optional claims. |
| The connection form appears although the manifest was configured | The manifest was generated without `--agent` | Regenerate with the instance options and a higher `--version`. |
| The agent does not know the document | Empty document, or content not loaded yet | Click in the document, then ask again. Context is refreshed when the selection changes. |
| The agent answers about the previous email | The pane is not pinned (Outlook) | Pin the pane so it follows the selected message. |
| The agent says it cannot edit the document | No tools workspace configured | Install it on the instance and regenerate the manifests with `--tools`. |
| Users still see an old version after an update | Office cached the manifest or the pane | Check the manifest version was raised. Then [clear the Office cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache) on the device. |
| A new conversation starts every time the pane is opened | The chat runs in a third-party iframe, and browsers isolate or clear its storage | Can happen in this version, depending on the browser or webview. |

On Windows, the pane runs in Microsoft Edge WebView2; on Mac, in Safari's
WebKit. On a device where WebView2 is managed by Windows Information Protection,
`msedgewebview2.exe` must be on the protected apps list.

## Security and data flow

- **What the add-in reads**: the open document (Word body, the active Excel
  sheet's used range capped at 200 × 50 cells, PowerPoint slide text capped at
  50 slides), the current selection, and in Outlook the open message (subject,
  sender, recipients, body). It does not read other files, other mailbox
  folders, or the calendar.
- **Where it goes**: only to the Prisme.ai instance in the manifest, when the
  user sends a message. The task pane host serves static files and receives no
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
  the tool permissions configured on the agent. The add-in never sends email:
  a reply is opened as a draft for the user to review and send.
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
   `mcpTools` catalog (the eight tools below). If the agent sees no tools after
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

| Tool | Arguments | Applications |
|---|---|---|
| `office_read` | — | All |
| `office_get_selection` | — | All (Outlook: compose mode only) |
| `office_insert_text` | `{ text, location?: 'replace' \| 'start' \| 'end' }` | All (`location` applies to Word; elsewhere the selection is replaced) |
| `excel_read_range` | `{ address?, sheet? }` — defaults to the selection, 5,000 cells maximum | Excel |
| `excel_write_range` | `{ address, values, sheet? }` — `address` is the top-left cell, `values` a 2D array | Excel |
| `powerpoint_read_slides` | — | PowerPoint |
| `outlook_read_item` | — | Outlook |
| `outlook_draft_reply` | `{ body, replyAll? }` — opens an unsent draft | Outlook (read mode) |

## Known limitations

- **Single sign-on works with Microsoft Entra ID only** (OIDC provider on the
  instance). With any other identity provider, users are anonymous to the
  instance.
- **The Office tools workspace (`office-mcp`) is not shipped in this repository**
  and has not been validated on a live instance yet.
- **No mailbox-wide or calendar features** in Outlook: only the open message.
- **No mobile clients.**
- PDF and image attachments are not read from the open email.

## Local development

```bash
npm run certs      # once: install a trusted localhost certificate
npm run manifest   # manifests pointing at https://localhost:3100
npm run serve      # serve the task pane over HTTPS on port 3100
npm run check      # syntax check
npm run validate   # validate the manifests against Microsoft's schema
```

Then sideload `dist/manifest.xml` (see [Step 4](#step-4--pilot-with-one-user)),
or run `npx office-addin-debugging start dist/manifest.xml` for Office desktop.

There is no build step: the task pane is plain HTML and JavaScript.

Layout:

- `taskpane.html`, `taskpane.js` — the pane: hosts the chat iframe, pushes the
  document context, executes the agent's tools.
- `scripts/manifest.mjs` — generates both XML manifests.
- `vendor/msal-browser.min.js` — `@azure/msal-browser` 5.25.0, unmodified
  (`lib/msal-browser.min.js` from the npm package; licence alongside). Loaded
  only when single sign-on is enabled.
- `PREREQUIS.md` — internal design notes (French).

## License

See [LICENSE](LICENSE).
