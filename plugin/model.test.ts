import { test } from "node:test"
import assert from "node:assert/strict"
import { adminRequest, classifyGraph, defaultEffect, explainSignInError, graphCapabilities, normalizeAccount, requestedScopes } from "./model.js"

test("accounts: defaults, presets and validation", () => {
  const account = normalizeAccount({ id: "work", tenant: "asirobots.com", email: "Pat@ACME.example" }, new Date(0))
  assert.deepEqual(account, { id: "work", auth: "browser", tenant: "asirobots.com", preset: "full", email: "pat@acme.example", updatedAt: "1970-01-01T00:00:00.000Z" })
  assert.ok(requestedScopes(account).includes("Mail.Send"))
  assert.ok(requestedScopes(account).includes("offline_access"))
  assert.ok(!requestedScopes(normalizeAccount({ id: "r", tenant: "x.com", preset: "read" })).includes("Mail.Send"))
  assert.deepEqual(requestedScopes(normalizeAccount({ id: "c", tenant: "x.com", preset: "custom", scopes: ["Mail.Read"] })), ["Mail.Read", "offline_access"])
  assert.throws(() => normalizeAccount({ id: "c", tenant: "x.com", preset: "custom" }), /scopes/)
  assert.throws(() => normalizeAccount({ id: "c", tenant: "x.com", scopes: ["Mail.Read; drop"] }), /permission names/)
  assert.throws(() => normalizeAccount({ id: "c", tenant: "not a tenant" }), /tenant/)
  assert.equal(normalizeAccount({ id: "o" }).tenant, "organizations")
})

test("Graph requests are classified; anything others see is separate", () => {
  const cases: [string, string, unknown, string][] = [
    ["GET", "/me/messages?$top=5", undefined, "mail-read"], ["POST", "/me/messages", { subject: "x" }, "mail-write"], ["PATCH", "/me/messages/1", { isRead: true }, "mail-write"],
    ["POST", "/me/messages/1/send", undefined, "mail-send"], ["POST", "/me/sendMail", {}, "mail-send"], ["POST", "/me/messages/1/reply", {}, "mail-send"],
    ["POST", "/me/messages/1/createReply", undefined, "mail-write"], ["DELETE", "/me/messages/1", undefined, "mail-delete"],
    ["GET", "/me/calendarView?startDateTime=x", undefined, "calendar-read"], ["POST", "/me/calendar/getSchedule", {}, "calendar-read"],
    ["POST", "/me/events", { subject: "focus" }, "calendar-write"], ["POST", "/me/events", { subject: "sync", attendees: [{}] }, "calendar-send"],
    ["POST", "/me/events/1/accept", {}, "calendar-send"], ["GET", "/me/joinedTeams", undefined, "chat-read"],
    ["POST", "/chats/1/messages", {}, "chat-send"], ["POST", "/teams/t/channels/c/messages/m/replies", {}, "chat-send"], ["POST", "/teams/t/channels", {}, "chat-admin"],
    ["GET", "/me/drive/root/children", undefined, "files-read"], ["PUT", "/me/drive/root:/a.txt:/content", "x", "files-write"],
    ["POST", "/drives/d/items/i/createLink", {}, "files-send"], ["POST", "/drives/d/items/i/invite", {}, "files-send"], ["DELETE", "/drives/d/items/i", undefined, "files-delete"],
    ["GET", "/me", undefined, "directory-read"], ["GET", "/users?$filter=x", undefined, "directory-read"], ["PATCH", "/me", {}, "directory-admin"],
    ["POST", "/search/query", {}, "other-read"], ["POST", "/subscriptions", {}, "other-write"],
    ["GET", "https://graph.microsoft.com/beta/me/messages", undefined, "mail-read"], ["GET", "/v1.0/me/mailFolders/inbox/messages", undefined, "mail-read"],
  ]
  for (const [method, path, body, expected] of cases) assert.equal(classifyGraph(method, path, body), expected, `${method} ${path}`)
  assert.deepEqual(graphCapabilities("POST", "/v1.0/$batch", { requests: [{ id: "1", method: "GET", url: "/me/messages" }, { id: "2", method: "POST", url: "/me/sendMail" }, { id: "3", method: "GET", url: "/me/events" }] }), ["mail-read", "calendar-read", "mail-send"])
})

test("defaults: reads and own changes allowed; sending, sharing, deletes, admin and unknown writes ask", () => {
  for (const capability of ["mail-read", "mail-write", "calendar-write", "files-write", "manage"]) assert.equal(defaultEffect(capability), "allow")
  for (const capability of ["mail-send", "calendar-send", "chat-send", "files-send", "mail-delete", "directory-admin", "other-write", "remove"]) assert.equal(defaultEffect(capability), "ask")
})

test("sign-in errors explained; admin request names the app and permissions", () => {
  const account = normalizeAccount({ id: "w", tenant: "asirobots.com" })
  assert.match(explainSignInError("access_denied", "AADSTS90094: Admin approval required", account), /requires an administrator.*Microsoft Graph Command Line Tools.*14d82eec/)
  assert.match(adminRequest(account), /Mail\.Send.*Delegated only/)
  assert.match(adminRequest(normalizeAccount({ id: "w", tenant: "x.com", clientId: "11111111-2222-3333-4444-555555555555" })), /app 11111111/)
})
