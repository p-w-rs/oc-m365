import type { ToolContext } from "@opencode/plugin/promise/tool"
import { isAbsolute, resolve } from "node:path"
import { adminRequest, graphCapabilities, mergeAccount, normalizeAccount, requestedScopes, validAccountId, type Account, type Discovery } from "./model.js"
import type { M365Service } from "./service.js"

export const TOOL_NAMES = [
  "m365_list", "m365_graph", "m365_download", "m365_upload",
  "m365_account_upsert", "m365_account_remove", "m365_signin", "m365_refresh",
] as const

export interface ToolDefinition {
  readonly name: (typeof TOOL_NAMES)[number]
  readonly description: () => string
  readonly input: Record<string, unknown>
  readonly execute: (input: any, context: ToolContext) => Promise<{ content: string }>
}

const MAX_OUTPUT = 60_000
const json = (value: unknown) => ({ content: JSON.stringify(value, null, 2) })
const ACCOUNT = { type: "string", description: "Account id (optional when there is only one account)" }
const ALLOWED_HEADERS = /^(prefer|consistencylevel|if-match|if-none-match)$/i

export const USAGE = [
  "Load the 'm365' skill before non-trivial Microsoft 365 work (mail, calendar, Teams, OneDrive/SharePoint).",
  "m365_graph calls any Microsoft Graph endpoint as the user (e.g. GET /me/messages?$top=10&$select=subject,from). Use $select/$top/$filter to keep results small; $batch for several requests. m365_download / m365_upload move files.",
  "Sending mail, invitations, posting in Teams and sharing files are seen by others: they ask the user by default. Draft first and show the user.",
].join("\n")

function describe(account: Account, discovery: Discovery | undefined): string {
  const who = discovery?.name ? `${discovery.name} <${discovery.email ?? discovery.upn}>` : "not signed in"
  return `  ${account.id}: ${who} in ${account.tenant}${discovery?.unread !== undefined ? `; ${discovery.unread} unread in Inbox` : ""}${account.label ? ` (${account.label})` : ""}${discovery?.error ? ` [problem: ${discovery.error.slice(0, 160)}]` : ""}`
}

export function inventory(accounts: readonly Account[], discovery: Record<string, Discovery>): string {
  if (!accounts.length) return "No Microsoft 365 accounts are configured yet. Create one with m365_account_upsert (tenant, email), then m365_signin."
  return ["Microsoft 365 accounts (Outlook mail and calendar, Teams, OneDrive, SharePoint):", ...accounts.map(account => describe(account, discovery[account.id]))].join("\n")
}

const clip = (text: string) => text.length > MAX_OUTPUT ? `${text.slice(0, MAX_OUTPUT)}\n…[truncated ${text.length - MAX_OUTPUT} characters; use $select, $top or a narrower request]` : text
const summarize = (value: unknown) => { const text = typeof value === "string" ? value : JSON.stringify(value); return text.length > 600 ? `${text.slice(0, 600)}…` : text }

export function createTools(service: M365Service, sessionDirectory: (sessionID: string) => Promise<string | undefined> = async () => undefined): ToolDefinition[] {
  const caller = (context: ToolContext) => ({ sessionID: context.sessionID, agent: context.agent, source: { type: "tool" as const, messageID: context.messageID, id: context.id } })
  const check = (context: ToolContext, capability: string, entity: string, metadata?: Record<string, unknown>) =>
    service.bridge.check({ ...caller(context), capability, entities: [entity], ...(metadata && { metadata }) })
  const localPath = async (path: string, context: ToolContext) => {
    if (isAbsolute(path)) return path
    const base = await sessionDirectory(context.sessionID).catch(() => undefined)
    if (!base) throw new Error("Pass an absolute local path (the session directory is unknown)")
    return resolve(base, path)
  }

  return [
    {
      name: "m365_list",
      description: () => {
        const { accounts, discovery } = service.store.cached()
        return `${inventory(accounts, discovery)}\n${USAGE}\nCall this tool for details: granted permissions and what you may do (allow / ask = the user is prompted / deny).`
      },
      input: { type: "object", properties: {}, additionalProperties: false },
      execute: async (_input, context) => {
        const settings = await service.store.settings()
        const accounts = await Promise.all((await service.store.list()).map(async account => {
          const signedIn = await service.signedIn(account)
          const discovery = signedIn ? await service.rediscover(account) : await service.store.discovery(account.id)
          const may = Object.fromEntries(await Promise.all(["mail-read", "mail-write", "mail-send", "calendar-send", "chat-send", "files-write", "files-send", "directory-read"].map(async capability =>
            [capability, (await service.bridge.effect({ ...caller(context), capability, entity: account.id })).effect] as const)))
          const signIn = service.signIns.get(account.id)
          return {
            id: account.id, tenant: account.tenant, ...(account.email && { email: account.email }), ...(account.label && { label: account.label }), signedIn,
            signedInAs: discovery?.name ? `${discovery.name} <${discovery.email ?? discovery.upn}>` : null, ...(discovery?.unread !== undefined && { unreadInbox: discovery.unread }),
            grantedPermissions: await service.granted(account), requested: requestedScopes(account),
            ...(signIn?.status === "waiting" && { signIn: signIn.code ? `waiting: code ${signIn.code} at ${signIn.url}` : `waiting for the user: ${signIn.url}` }),
            ...(signIn?.status === "failed" && { signInFailed: signIn.error }), ...(discovery?.error && { problem: discovery.error }),
            ...(settings.defaultAccount === account.id && { default: true }), youMay: may,
          }
        }))
        return json({ accounts, usage: USAGE })
      },
    },
    {
      name: "m365_graph",
      description: () => "Call Microsoft Graph as the signed-in user. path: e.g. /me/messages?$top=10&$select=subject,from,receivedDateTime (v1.0 by default; /beta/... for beta). method GET/POST/PATCH/PUT/DELETE; body as JSON; headers: Prefer, ConsistencyLevel, If-Match only. POST /$batch runs up to 20 requests. Anything others see (sending, invitations, Teams posts, sharing) asks the user by default.",
      input: {
        type: "object",
        properties: {
          account: ACCOUNT, method: { type: "string", enum: ["GET", "POST", "PATCH", "PUT", "DELETE"] }, path: { type: "string", minLength: 1, maxLength: 4000 },
          body: {}, headers: { type: "object", additionalProperties: { type: "string" } },
        },
        required: ["path"], additionalProperties: false,
      },
      execute: async (input: { account?: string; method?: string; path: string; body?: unknown; headers?: Record<string, string> }, context) => {
        const account = await service.pick(input.account)
        const method = (input.method ?? "GET").toUpperCase()
        for (const name of Object.keys(input.headers ?? {})) if (!ALLOWED_HEADERS.test(name)) throw new Error(`Header ${name} is not allowed (allowed: Prefer, ConsistencyLevel, If-Match, If-None-Match)`)
        const href = service.graphHref(input.path)
        for (const capability of graphCapabilities(method, href, input.body)) {
          await check(context, capability, account.id, { request: `${method} ${input.path}`.slice(0, 500), ...(input.body !== undefined && { body: summarize(input.body) }), account: account.email ?? account.id })
        }
        const result = await service.graph(account, method, input.path, { ...(input.body !== undefined && { body: input.body }), ...(input.headers && { headers: input.headers }) })
        const data = result.data === undefined ? { status: result.status, ...(result.headers.location && { location: result.headers.location }) } : result.data
        return { content: clip(typeof data === "string" ? data : JSON.stringify(data, null, 2)) }
      },
    },
    {
      name: "m365_download",
      description: () => "Download file content from Microsoft Graph to a local file: a OneDrive/SharePoint item (/me/drive/root:/Folder/file.docx:/content, /drives/{id}/items/{id}/content) or a mail attachment (/me/messages/{id}/attachments/{id}/$value). to: local path (relative to the session directory).",
      input: { type: "object", properties: { account: ACCOUNT, path: { type: "string", minLength: 1, maxLength: 4000 }, to: { type: "string", minLength: 1, maxLength: 1000 } }, required: ["path", "to"], additionalProperties: false },
      execute: async (input: { account?: string; path: string; to: string }, context) => {
        const account = await service.pick(input.account)
        const capability = graphCapabilities("GET", service.graphHref(input.path))[0]!
        await check(context, capability, account.id, { request: `download ${input.path}`, to: input.to })
        const to = await localPath(input.to, context)
        return json({ saved: to, ...(await service.download(account, input.path, to)) })
      },
    },
    {
      name: "m365_upload",
      description: () => "Upload a local file to OneDrive or SharePoint (any size). to: a drive path like /me/drive/root:/Reports/q4.xlsx or /drives/{driveId}/root:/Shared/plan.docx or /sites/{siteId}/drive/root:/Docs/a.pdf. conflict: replace (default), rename or fail.",
      input: {
        type: "object",
        properties: { account: ACCOUNT, from: { type: "string", minLength: 1, maxLength: 1000 }, to: { type: "string", minLength: 1, maxLength: 4000 }, conflict: { type: "string", enum: ["replace", "rename", "fail"] } },
        required: ["from", "to"], additionalProperties: false,
      },
      execute: async (input: { account?: string; from: string; to: string; conflict?: "replace" | "rename" | "fail" }, context) => {
        const account = await service.pick(input.account)
        if (!/(^|\/)(me\/drive|drives\/[^/]+|sites\/[^/]+\/drive|groups\/[^/]+\/drive|users\/[^/]+\/drive)\/(root|items\/[^/]+)/i.test(input.to)) throw new Error("to must be a drive path, e.g. /me/drive/root:/Folder/file.ext")
        await check(context, "files-write", account.id, { request: `upload ${input.from} → ${input.to}`, conflict: input.conflict ?? "replace" })
        const item = await service.upload(account, await localPath(input.from, context), input.to, input.conflict) as { id?: string; name?: string; webUrl?: string; size?: number } | undefined
        return json({ uploaded: { id: item?.id, name: item?.name, size: item?.size, webUrl: item?.webUrl } })
      },
    },
    {
      name: "m365_account_upsert",
      description: () => "Create or update a Microsoft 365 account (fields merge; null clears). tenant: the organization's domain (contoso.com) or directory id. email: pre-fills sign-in. preset: full (read and write mail, calendar, Teams, files) or read; custom with scopes=[...]. clientId: an app registration from IT, if the default Microsoft app is not approved. Never pass secrets.",
      input: {
        type: "object",
        properties: {
          id: { type: "string", pattern: "^[A-Za-z0-9][A-Za-z0-9._-]*$", maxLength: 64, description: "Name chosen by the user, e.g. work" },
          rename: { type: "string", pattern: "^[A-Za-z0-9][A-Za-z0-9._-]*$", maxLength: 64 },
          tenant: { type: "string" }, email: { type: ["string", "null"] }, auth: { type: "string", enum: ["browser", "device"] },
          preset: { type: "string", enum: ["full", "read", "custom"] }, scopes: { type: ["array", "null"], items: { type: "string" } },
          clientId: { type: ["string", "null"] }, label: { type: ["string", "null"] }, notes: { type: ["string", "null"] },
        },
        required: ["id"], additionalProperties: false,
      },
      execute: async (input: Record<string, unknown> & { id: string; rename?: string }, context) => {
        const { id, rename, ...patch } = input
        validAccountId(id)
        const existing = await service.store.get(id)
        await check(context, "manage", id, { operation: existing ? "update" : "create" })
        if (rename !== undefined && !existing) throw new Error(`Cannot rename: no account ${id}`)
        const target = rename ?? id
        if (target !== id) await check(context, "manage", validAccountId(target), { operation: "rename", from: id })
        const saved = await service.store.put(normalizeAccount({ ...mergeAccount(existing, patch), id: target }), existing && target !== id ? id : undefined)
        if (existing && target !== id) { await service.vault.renameEntity(id, target); await service.policy.renameEntity(id, target) }
        if (existing && (existing.tenant !== saved.tenant || existing.clientId !== saved.clientId || requestedScopes(existing).join() !== requestedScopes(saved).join())) await service.signOut(saved)
        const signedIn = await service.signedIn(saved)
        return json({ saved, signedIn, ...(!signedIn && { next: `sign in: m365_signin account=${saved.id}` }) })
      },
    },
    {
      name: "m365_account_remove",
      description: () => "Delete a Microsoft 365 account from this plugin with its stored sign-in and permission settings. Asks the user by default. Does not change anything in Microsoft 365.",
      input: { type: "object", properties: { account: { type: "string" } }, required: ["account"], additionalProperties: false },
      execute: async (input: { account: string }, context) => {
        const account = await service.pick(input.account)
        await check(context, "remove", account.id)
        service.signIns.get(account.id)?.flow.cancel()
        service.signIns.delete(account.id)
        await service.store.remove(account.id)
        await service.vault.deleteEntity(account.id)
        await service.policy.removeEntity(account.id)
        return json({ removed: account.id })
      },
    },
    {
      name: "m365_signin",
      description: () => "Start the Microsoft sign-in. browser (default): a link to open on this computer; device=true: a code to enter at microsoft.com/devicelogin on any device. Completes in the background; wait blocks up to that many seconds. If Microsoft shows 'Approval required', the organization must approve the permissions first: the result includes a request for IT.",
      input: { type: "object", properties: { account: ACCOUNT, device: { type: "boolean" }, wait: { type: "integer", minimum: 0, maximum: 600 } }, additionalProperties: false },
      execute: async (input: { account?: string; device?: boolean; wait?: number }, context) => {
        const account = await service.pick(input.account)
        await check(context, "manage", account.id, { operation: "sign in" })
        const state = await service.startSignIn(account, { ...(input.device !== undefined && { device: input.device }) })
        const deadline = Date.now() + (input.wait ?? 0) * 1000
        while (state.status === "waiting" && Date.now() < deadline) {
          context.signal?.throwIfAborted()
          await new Promise(resolve => setTimeout(resolve, 1000))
        }
        const discovery = state.status === "done" ? await service.store.discovery(account.id) : undefined
        return json({
          account: account.id, status: state.status,
          ...(state.status === "waiting" && { tellUser: state.code
            ? `On any device, open ${state.url} and enter the code ${state.code}, then sign in${account.email ? ` as ${account.email}` : ""}.`
            : `Open this link on this computer and sign in${account.email ? ` as ${account.email}` : ""}: ${state.url}  — or click Sign in in OpenCode's Microsoft 365 view. If Microsoft shows "Approval required", stop there and tell me.` }),
          ifApprovalRequired: `Ask IT: ${adminRequest(account)}`,
          ...(state.error && { error: state.error }),
          ...(discovery && { signedInAs: discovery.name ? `${discovery.name} <${discovery.email ?? discovery.upn}>` : null, granted: discovery.granted, ...(discovery.error && { problem: discovery.error }) }),
        })
      },
    },
    {
      name: "m365_refresh",
      description: () => "Check an account again: who is signed in, granted permissions, unread mail.",
      input: { type: "object", properties: { account: ACCOUNT }, additionalProperties: false },
      execute: async (input: { account?: string }) => {
        const account = await service.pick(input.account)
        return json({ account: account.id, ...(await service.rediscover(account)) })
      },
    },
  ]
}
