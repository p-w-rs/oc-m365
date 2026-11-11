---
name: m365
description: Work effectively in Microsoft 365 through the oc-m365 tools and Microsoft Graph — Outlook mail (search, read, draft, reply, send, attachments), calendar (events, availability, meetings), Teams (chats, channels, messages), OneDrive and SharePoint (find, read, upload, share files), and people. Load before non-trivial Microsoft 365 work.
---

# Working in Microsoft 365

## Tools

- `m365_graph` — any Microsoft Graph request as the signed-in user: `path` (`/v1.0` by default, `/beta/...` when
  needed), `method`, JSON `body`, `headers` (only `Prefer`, `ConsistencyLevel`, `If-Match`, `If-None-Match`).
  Output over 60k characters is cut.
- `m365_download path=… to=<local path>` / `m365_upload from=… to=<drive path> conflict=replace|rename|fail` — file
  content to and from OneDrive/SharePoint, mail attachments. Upload **replaces** by default: pass `conflict: "fail"` or
  `"rename"` unless overwriting is intended.
- `m365_list` — accounts, who is signed in, granted permissions (`granted`), what you may do. Setup:
  `m365_account_upsert` (tenant, email, preset `full`/`read`/`custom`) → `m365_signin` (`device=true` for a code);
  `m365_refresh` afterwards. "Approval required" at sign-in = the tenant needs admin consent; give the user the
  request for IT that the tool returns.

Permissions (defaults): reading and changing your own things (drafts, folders, flags, events without attendees, your
files) run; **anything other people see — send, reply, forward, invitations, meeting responses, Teams posts, sharing
links/invites — asks**, as do deletes and directory changes. A 403 means the user or the granted permissions do not
allow it: say which, do not work around it.

The `full` preset grants User.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Chat.ReadWrite, ChannelMessage.Send,
Team.ReadBasic.All, Channel.ReadBasic.All, Files.ReadWrite.All, Sites.ReadWrite.All, People.Read (`read`: the read
versions). Not included, so expect 403 unless `granted` lists them: reading channel messages (ChannelMessage.Read.All),
`findMeetingTimes` (Calendars.Read.Shared), listing `/users` (User.ReadBasic.All), other people's mailboxes.

## Keep responses small

Always `$select` the fields you need and `$top` a small page; follow `@odata.nextLink` only when needed. For mail,
ask for text bodies: header `Prefer: outlook.body-content-type="text"`. Use `$batch` (POST /$batch, up to 20
requests; body `{requests:[{id:"1",method:"GET",url:"/me/…"}]}`) to fetch several things at once.

## Mail

- Recent: `GET /me/messages?$top=10&$select=subject,from,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime desc`
- Unread in inbox: `GET /me/mailFolders/inbox/messages?$filter=isRead eq false&$select=…` (combining `$filter` and
  `$orderby`: the ordered property must also be in the filter, first).
- Search: `GET /me/messages?$search="budget review"&$select=…` (KQL like `"from:pat subject:budget"` works; no
  `ConsistencyLevel` needed; `$search` cannot combine with `$orderby` — results come newest first, max 1,000).
- One message: `GET /me/messages/{id}?$select=subject,from,toRecipients,body,hasAttachments` (+ Prefer text).
- Draft, then send — the safe way: `POST /me/messages` with `{subject, body:{contentType:"Text",content}, toRecipients:[{emailAddress:{address}}]}`
  creates a draft; show it to the user; `POST /me/messages/{id}/send` sends it.
  Replies: `POST /me/messages/{id}/createReply` (or `/createReplyAll`, `/createForward`) creates a draft — pass either
  `{comment}` or `{message:{…}}`, not both → `PATCH` to edit → `/send`. `/reply` with `{comment}` sends at once.
- Attachments: `GET /me/messages/{id}/attachments?$select=id,name,size,contentType`, then
  `m365_download path=/me/messages/{id}/attachments/{aid}/$value to=…`.
- Move/flag/read state: `PATCH /me/messages/{id}` `{isRead:true}`, `POST /me/messages/{id}/move {destinationId}`
  (well-known names like `archive`, `deleteditems` work as ids).
- Write as the user would: their language and tone, short, no invented facts. Never send without the user seeing the
  final text unless they explicitly said to.

## Calendar

- Next days: `GET /me/calendarView?startDateTime=2026-11-10T00:00:00Z&endDateTime=2026-11-17T00:00:00Z&$select=subject,start,end,location,attendees,isOnlineMeeting&$orderby=start/dateTime`
  (header `Prefer: outlook.timezone="America/Denver"` for local times — use the user's zone from
  `GET /me/mailboxSettings?$select=timeZone`).
- Free/busy: `POST /me/calendar/getSchedule` `{schedules:["a@x.com"], startTime:{dateTime,timeZone}, endTime:{…},
  availabilityViewInterval:30}` (a read; no one is notified). `POST /me/findMeetingTimes` needs Calendars.Read.Shared.
- Create: `POST /me/events` with subject, start/end `{dateTime, timeZone}`, attendees (sends invitations — asks),
  `isOnlineMeeting:true, onlineMeetingProvider:"teamsForBusiness"` for a Teams link. Without attendees it is just
  a block on the user's calendar.
- Respond: `POST /me/events/{id}/accept|decline|tentativelyAccept {comment, sendResponse}` (asks).

## Teams

- Teams and channels: `GET /me/joinedTeams?$select=id,displayName`, `GET /teams/{id}/channels?$select=id,displayName`.
- Chats: `GET /me/chats?$expand=members&$top=20`, `GET /chats/{id}/messages?$top=20` (max 50, newest first).
- Channel messages: `GET /teams/{id}/channels/{cid}/messages?$top=20` (replies: `/messages/{mid}/replies`) — needs
  ChannelMessage.Read.All, which most tenants only grant with admin consent; a 403 there is that.
- Post: `POST /chats/{id}/messages` or `/teams/{id}/channels/{cid}/messages` with `{body:{contentType:"text",content}}`
  — seen by others, asks. Mention people only when asked.

## OneDrive and SharePoint

- Find: `GET /me/drive/root/search(q='quarterly report')?$select=id,name,webUrl,parentReference,size,lastModifiedDateTime`,
  across everything you can access: `POST /search/query` `{requests:[{entityTypes:["driveItem"],query:{queryString:"…"}}]}`.
- Browse: `GET /me/drive/root/children`, `GET /me/drive/root:/Folder/Sub:/children`.
- SharePoint: `GET /sites?search=Engineering`, `GET /sites/{siteId}/drives`, `GET /drives/{driveId}/root/children`.
- Read content: `m365_download path=/drives/{d}/items/{i}/content to=…`, then read the local file.
- Write: `m365_upload from=./report.pdf to=/me/drive/root:/Reports/report.pdf conflict=fail` (any size). Never
  overwrite a colleague's document without asking; prefer a new file or version note.
- Share: `POST /drives/{d}/items/{i}/createLink {type:"view", scope:"organization"}` or `/invite` — asks; prefer
  organization-scoped view links over anonymous ones.

## People

`GET /me/people?$search="pat"&$select=displayName,scoredEmailAddresses` to resolve names to addresses (People.Read).
`GET /users?$filter=startswith(displayName,'Pat')&$select=displayName,mail,jobTitle` needs `User.ReadBasic.All`
(not in the presets). Confirm the right person when a name is ambiguous.

## Errors

401 = sign-in expired (`m365_signin`); 403 = permission (compare with `granted` in `m365_list`) or the user's own
access; 404 = wrong id or no access; 429 = throttled (respect Retry-After); `ErrorInvalidRecipients` = bad address;
412 = changed since read (re-read, then `If-Match`).
