# oc-m365

Microsoft 365 for OpenCode agents: Outlook mail and calendar, Teams chats and
channels, OneDrive and SharePoint files, and people — through Microsoft Graph,
as the signed-in user. Anything other people see (sending mail, replies, meeting
invitations, Teams messages, sharing links) asks first by default, and the
built-in `m365` skill teaches agents to draft first and keep requests small.
macOS and Linux; OpenCode 2.2.6+ and Pixi.

## Sign-in

A Microsoft sign-in as the user (SSO, MFA, Conditional Access): in the browser on
the OpenCode computer, or with a code on any device. The plugin keeps and
refreshes its own tokens (encrypted; never shown to agents).

It signs in with Microsoft's own **Microsoft Graph Command Line Tools** app
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`, the app Graph PowerShell uses), or with
an app registration from your IT (`clientId`).

**Many organizations require an administrator to approve** mail, Teams and file
access ("Approval required" during sign-in). The GUI and the agent show a ready
request for IT naming the app and the delegated permissions. Access presets:

| Preset | Permissions (delegated) |
|---|---|
| `full` (default) | User.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Chat.ReadWrite, ChannelMessage.Send, Team.ReadBasic.All, Channel.ReadBasic.All, Files.ReadWrite.All, Sites.ReadWrite.All, People.Read |
| `read` | User.Read, Mail.Read, Calendars.Read, Chat.Read, Team.ReadBasic.All, Channel.ReadBasic.All, Files.Read.All, Sites.Read.All, People.Read |
| `custom` | your list |

## What agents get

| Tool | |
|---|---|
| `m365_list` | Accounts, who is signed in, unread mail, granted permissions, what the agent may do |
| `m365_graph` | Any Microsoft Graph request (v1.0 or beta), including `$batch` |
| `m365_download` / `m365_upload` | Files to and from OneDrive/SharePoint (any size), mail attachments |
| `m365_account_upsert` / `_remove`, `m365_signin`, `m365_refresh` | Setup |

## Permissions

Per account, for mail, calendar, Teams & chat, OneDrive & SharePoint, people &
directory and other: read, change, **send / share (others see it)**, delete,
admin; plus edit/remove account. Each request is classified by method and path
(every request in a `$batch`). Defaults: reading and changing your own things
run; sending and sharing, deletes, admin and unknown writes ask. Editable in the
GUI (`/m365`, the Review-pane **+** menu, or **Settings → Microsoft 365**);
**Allow always** on a prompt writes allow.

## Setup

```sh
pixi run --locked install
pixi run signin work --tenant contoso.com --email you@contoso.com      # --device, --read-only, --client-id <id>
```

## Development

```sh
pixi run --locked test            # fake Microsoft identity platform + Graph
pixi run --locked smoke-host      # real OpenCode 2.2.6 CLI, scripted model, fake Microsoft
```
