---
name: m365
description: Work effectively in Microsoft 365 through the oc-m365 tools and Microsoft Graph — Outlook mail (search, read, draft, reply, send, attachments), calendar (events, availability, meetings), Teams (chats, channels, messages), OneDrive and SharePoint (find, read, upload, share files), and people. Load before non-trivial Microsoft 365 work.
---

# Working in Microsoft 365

## Tools

- `m365_graph` — any Microsoft Graph request as the signed-in user (`/v1.0` by default, `/beta/...` when needed).
- `m365_download` / `m365_upload` — file content to and from OneDrive/SharePoint, mail attachments.
- `m365_list` — accounts, who is signed in, granted permissions, what you may do.

Permissions: reading and changing your own things (drafts, folders, events without attendees, your files) usually run;
**anything other people see — sending mail, replying, meeting invitations, Teams messages, sharing links — asks**,
as do deletes. A 403 means the user or the granted permissions do not allow it: say which, do not work around it.

## Keep responses small

Always `$select` the fields you need and `$top` a small page; follow `@odata.nextLink` only when needed. For mail,
ask for text bodies: header `Prefer: outlook.body-content-type="text"`. Use `$batch` (POST /$batch, up to 20
requests) to fetch several things at once.

## Mail

- Recent: `GET /me/messages?$top=10&$select=subject,from,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime desc`
- Unread in inbox: `GET /me/mailFolders/inbox/messages?$filter=isRead eq false&$select=…`
- Search: `GET /me/messages?$search="budget review"&$select=…` (header `ConsistencyLevel: eventual` not needed for
  messages; `$search` cannot combine with `$orderby`).
- One message: `GET /me/messages/{id}?$select=subject,from,toRecipients,body,hasAttachments` (+ Prefer text).
- Draft, then send — the safe way: `POST /me/messages` with `{subject, body:{contentType:"Text",content}, toRecipients:[{emailAddress:{address}}]}`
  creates a draft; show it to the user; `POST /me/messages/{id}/send` sends it.
  Replies: `POST /me/messages/{id}/createReply` (draft) → edit → `/send`; or `/reply` with `{comment}` to send at once.
- Attachments: `GET /me/messages/{id}/attachments?$select=id,name,size,contentType`, then
  `m365_download path=/me/messages/{id}/attachments/{aid}/$value to=…`.
- Move/flag/read state: `PATCH /me/messages/{id}` `{isRead:true}`, `POST /me/messages/{id}/move {destinationId}`.
- Write as the user would: their language and tone, short, no invented facts. Never send without the user seeing the
  final text unless they explicitly said to.

## Calendar

- Next days: `GET /me/calendarView?startDateTime=2026-11-10T00:00:00Z&endDateTime=2026-11-17T00:00:00Z&$select=subject,start,end,location,attendees,isOnlineMeeting&$orderby=start/dateTime`
  (header `Prefer: outlook.timezone="America/Denver"` for local times — use the user's zone from
  `GET /me/mailboxSettings?$select=timeZone`).
- Free times: `POST /me/findMeetingTimes` or `POST /me/calendar/getSchedule` (reads; no one is notified).
- Create: `POST /me/events` with subject, start/end `{dateTime, timeZone}`, attendees (sends invitations — asks),
  `isOnlineMeeting:true, onlineMeetingProvider:"teamsForBusiness"` for a Teams link. Without attendees it is just
  a block on the user's calendar.
- Respond: `POST /me/events/{id}/accept|decline|tentativelyAccept {comment, sendResponse}`.

## Teams

- Teams and channels: `GET /me/joinedTeams?$select=id,displayName`, `GET /teams/{id}/channels?$select=id,displayName`.
- Channel messages: `GET /teams/{id}/channels/{cid}/messages?$top=20` (replies: `/messages/{mid}/replies`).
- Chats: `GET /me/chats?$expand=members&$top=20`, `GET /chats/{id}/messages?$top=20`.
- Post: `POST /chats/{id}/messages` or `/teams/{id}/channels/{cid}/messages` with `{body:{contentType:"text",content}}`
  — seen by others, asks. Mention people only when asked.
- Some Teams reads need permissions that require admin approval (e.g. ChannelMessage.Read.All); a 403 there is a
  tenant setting.

## OneDrive and SharePoint

- Find: `GET /me/drive/root/search(q='quarterly report')?$select=id,name,webUrl,parentReference,size,lastModifiedDateTime`,
  across everything: `POST /search/query` `{requests:[{entityTypes:["driveItem"],query:{queryString:"…"}}]}`.
- Browse: `GET /me/drive/root/children`, `GET /me/drive/root:/Folder/Sub:/children`.
- SharePoint: `GET /sites?search=Engineering`, `GET /sites/{siteId}/drives`, `GET /drives/{driveId}/root/children`.
- Read content: `m365_download` then read the local file (Office files: convert or extract text locally). Small text
  files can also be fetched with `GET …/content`.
- Write: `m365_upload from=./report.pdf to=/me/drive/root:/Reports/report.pdf` (any size). Never overwrite a
  colleague's document without asking; prefer a new file or version note.
- Share: `POST /drives/{d}/items/{i}/createLink {type:"view", scope:"organization"}` or `/invite` — asks; prefer
  organization-scoped view links over anonymous ones.

## People

`GET /me/people?$search="pat"&$select=displayName,scoredEmailAddresses` to resolve names to addresses;
`GET /users?$filter=startswith(displayName,'Pat')&$select=displayName,mail,jobTitle` (directory search needs
`User.ReadBasic.All`, sometimes unavailable). Confirm the right person when a name is ambiguous.

## Errors

401 = sign-in expired (`m365_signin`); 403 = permission (granted list in `m365_list`) or the user's own access;
404 = wrong id or no access; 429 = throttled (respect Retry-After); `ErrorInvalidRecipients` = bad address.
