<div align="center">

<img src="web/assets/icon-128.png" alt="Report Missed Spam" width="112" height="112">

# Report Missed Spam

**One click in Outlook. One clean report to your email security team.**

![Outlook](https://img.shields.io/badge/Outlook-new%20%7C%20classic%20%7C%20web-0078D4?logo=microsoftoutlook&logoColor=white)
![Office Add-in](https://img.shields.io/badge/Office%20Add--in-Mailbox%201.9-2B579A)
![Deploy](https://img.shields.io/badge/deploy-M365%20Integrated%20Apps-5E5E5E)
![Made by](https://img.shields.io/badge/made%20by-TR%20Computer%20Sales-2E8B57)

<sub>Outlook add-in &nbsp;·&nbsp; function command &nbsp;·&nbsp; no task pane &nbsp;·&nbsp; `ReadItem` only</sub>

</div>

---

One ribbon button on any opened message. Click it and:

1. A dialog asks **"Did you release this message from your quarantine?"**
   - **Yes**: shows "Mailprotector already quarantined this message... Please block the sender in your
     digest report." with a Close button. Nothing is reported.
   - **No**: a new compose window opens, addressed to `Spam@emailservice.io`, with the original message
     **attached as an item** (not forwarded) and a customizable body. The user reviews and clicks Send.
2. Closing the dialog with the X also reports nothing.
Before the question, the add-in reads the message's internet headers. If it has an `X-PHISHTEST` or
`X-PHISHTEST-Curricula` header (list and wording in `phishTest` in `config.js`), the dialog says
**You caught a phishing simulation! Good Job! Please use the "Report a Phish" or "Phish Alert" button instead!** and nothing is reported. If the headers can't be
read, the normal flow continues.

Function command only, no task pane, `ReadItem` permission, Mailbox 1.9 requirement set.

## Layout

| Path | Purpose |
|---|---|
| `web/config.js` | **Edit this**: recipient, subject, email body, and all dialog wording |
| `web/commands.js` | Button logic (dialog, then compose) |
| `web/cmd3.html` | Page the manifest points at; loads Office.js, config.js, commands.js |
| `web/commands.html` | Same page, kept so older manifests (1.0.0.1/.2) still work |
| `web/dialog.html`, `web/dialog.js` | The quarantine question dialog |
| `web/assets/` | Ribbon icons (from your PSD) |
| `manifest.template.xml` | Manifest; `{{BASE_URL}}` is filled in by the script |
| `Deploy-AddIn.ps1` | Builds `dist/manifest.xml`, optionally installs per tenant |

## 1. Host the `web/` folder over HTTPS

Any static host works (Azure Static Web Apps, GitHub Pages, Cloudflare Pages, an existing web server).
Serve `web/` at the root so `https://<host>/cmd3.html` and `https://<host>/assets/icon-32.png` resolve.
One host serves every client tenant.

## 2. Build the manifest

```powershell
.\Deploy-AddIn.ps1 -BaseUrl https://spamreport.example.com -ManifestOnly
```

Writes `dist\manifest.xml`.

## 3. Deploy

**Admin center (known-good path):** Settings > Integrated apps > Upload custom apps > Office Add-in >
upload `dist\manifest.xml` > Everyone (or a group) > Deploy. Microsoft says it can take up to 24 hours for it to appear
in all clients.

**Scripted across tenants (untested, try one tenant first):**

```powershell
.\Deploy-AddIn.ps1 -BaseUrl https://spamreport.example.com `
    -UserPrincipalName you@yourmsp.com `
    -TenantDomains contoso.onmicrosoft.com, fabrikam.onmicrosoft.com
```

Uses Exchange Online PowerShell over GDAP (`New-App -OrganizationApp`). Needs an Exchange-capable
GDAP role. I did not find a native add-in deployment in CIPP's docs, so this is the scripted fallback.

## Updating an already-deployed add-in

- **Hosted files only** (`config.js`, `commands.js`, `dialog.*`, icons): update the repo/host. No tenant action.
  Outlook may cache the old files; to force a refresh, do a manifest update as below.
- **Manifest changes, or to force clients to refetch:** raise `<Version>` in `manifest.template.xml`
  (and `manifest.xml`), keep `<Id>` unchanged, and update the existing add-in in
  Integrated apps by uploading the new manifest. Don't remove and re-add it; that drops the
  deployment and its assignments. `Deploy-AddIn.ps1` appends the version to the page URL as `?v=`,
  and the pages forward it to `config.js`, `commands.js` and the dialog, so a version bump
  refreshes all of them. If you edit `manifest.xml` by hand, keep the two `?v=` values equal to `<Version>`.
- Only the admin who deployed the add-in, or a Global Admin, can update or remove it.

## Customizing the message

Edit `web/config.js` and re-host the folder. No manifest change or tenant redeploy is needed.
Email tokens: `{reporter}`, `{sender}`, `{subject}`, `{date}`. Body is HTML (32 KB max); token values
are escaped. Dialog text is plain text. Set `quarantinePrompt.enabled` to `false` to skip the dialog
and go straight to compose.

## Test checklist (do this before rolling out)

1. Classic Outlook (Windows): open a message, button appears on the ribbon, compose opens with the
   message attached as an item.
2. New Outlook (Windows) and Outlook on the web: same checks.
3. Send one to the mailbox and confirm the vendor's pipeline accepts it as a valid sample (attached
   item, not forwarded). Ask them whether they need `.eml` specifically; an item attachment is
   delivered as an attached message, which may or may not satisfy their intake.
4. Dialog: Yes shows the quarantine message and reports nothing; No opens compose; the X reports
   nothing. In Outlook on the web, the first dialog may show a browser pop-up/allow prompt.
5. Check the button also shows when a message is opened in its own window, and for shared mailboxes.

## Known limits

- If the dialog never opens within 15 seconds (for example the user ignores the "wants to display a new
  window" prompt), the report is cancelled with a notice and compose does not open. Errors opening the dialog still fall through to compose.
- Not verified: how each Outlook client handles a dialog launched from a ribbon function command
  (Microsoft documents dialogs from add-in commands but not this exact timing). Test all three clients.
- Two clicks: Send is the user's click. Fully silent sending would need Graph + a multi-tenant app.
- Item attachments need the message's EWS ID; a few contexts (some shared/delegate scenarios) don't
  provide one. The add-in shows an error notification instead of opening an empty compose window.
- Mobile Outlook is not configured (desktop form factor only).
