import { validName } from "oc-secure-kit"

export const ACCOUNTS_FORMAT = "oc-m365-accounts/1"

/** Microsoft's own "Microsoft Graph Command Line Tools" app (used by Graph PowerShell): public client, any tenant. */
export const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
export const DEFAULT_AUTHORITY = "https://login.microsoftonline.com"
export const DEFAULT_GRAPH = "https://graph.microsoft.com"

/** Delegated permissions per preset. offline_access (refresh) is always added. */
export const SCOPE_PRESETS = {
  full: ["User.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite", "Chat.ReadWrite", "ChannelMessage.Send", "Team.ReadBasic.All", "Channel.ReadBasic.All", "Files.ReadWrite.All", "Sites.ReadWrite.All", "People.Read"],
  read: ["User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "Files.Read.All", "Sites.Read.All", "People.Read"],
} as const
export type Preset = keyof typeof SCOPE_PRESETS | "custom"

/**
 * How an account signs in (both are Microsoft sign-ins as the user, with SSO and MFA):
 * - browser: a sign-in page in the browser on the OpenCode computer.
 * - device: a code entered at microsoft.com/devicelogin on any device (remote OpenCode servers).
 */
export type Auth = "browser" | "device"

export interface Account {
  readonly id: string
  readonly label?: string
  readonly auth: Auth
  /** Tenant: domain (contoso.com), directory id, or "organizations" (any work account). */
  readonly tenant: string
  /** App (client) id to sign in with; default Microsoft Graph Command Line Tools. */
  readonly clientId?: string
  /** Which permissions to request. */
  readonly preset: Preset
  /** For preset=custom: the delegated scopes, e.g. ["Mail.Read", "Files.Read.All"]. */
  readonly scopes?: readonly string[]
  /** Pre-fills sign-in. */
  readonly email?: string
  readonly notes?: string
  readonly updatedAt: string
}

export interface Discovery {
  readonly name?: string
  readonly email?: string
  readonly upn?: string
  readonly granted?: readonly string[]
  readonly unread?: number
  readonly checkedAt?: string
  readonly error?: string
}

export const AREAS = ["mail", "calendar", "chat", "files", "directory", "other"] as const
export type Area = (typeof AREAS)[number]
/** send: something other people see (sending mail, invitations, posting messages, sharing links). */
export const VERBS = ["read", "write", "send", "delete", "admin"] as const
export type Verb = (typeof VERBS)[number]
export const CAPABILITIES = [...AREAS.flatMap(area => VERBS.map(verb => `${area}-${verb}` as const)), "manage", "remove"] as const
export type Capability = (typeof CAPABILITIES)[number]

/** Starting point only; users change it in the table. Anything others see, deletes, admin and unknown writes ask. */
export function defaultEffect(capability: string): "allow" | "ask" {
  return /-(send|delete|admin)$/.test(capability) || ["other-write", "remove"].includes(capability) ? "ask" : "allow"
}

export const ACCESS = "access-token"
export const REFRESH = "refresh-token"
/** { expiresAt, scopes } for the stored tokens. */
export const TOKEN_META = "token-meta"
export const SECRET_NAMES = [ACCESS, REFRESH, TOKEN_META] as const

const ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/
const CONTROL = /[\u0000-\u001f\u007f]/
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
const SCOPE = /^[A-Za-z][A-Za-z0-9.]*$/
const fail = (message: string): never => { throw new Error(message) }
const text = (value: unknown, field: string, max: number): string | undefined => {
  if (value === undefined || value === null || value === "") return undefined
  if (typeof value !== "string" || value.length > max || CONTROL.test(value)) fail(`${field} must be a single-line string up to ${max} characters`)
  return (value as string).trim()
}

export function validAccountId(id: unknown): string {
  if (typeof id !== "string" || !ID.test(id)) return fail(`account id ${JSON.stringify(id)} must be 1-64 letters, digits, '.', '_' or '-'`)
  return validName("account", id)
}

export function normalizeAccount(input: unknown, now = new Date()): Account {
  if (!input || typeof input !== "object" || Array.isArray(input)) return fail("account must be an object")
  const raw = input as Record<string, unknown>
  for (const key of Object.keys(raw)) if (!["id", "label", "auth", "tenant", "clientId", "preset", "scopes", "email", "notes", "updatedAt"].includes(key)) fail(`Unknown account field ${JSON.stringify(key)}`)
  const id = validAccountId(raw.id)
  const auth = (raw.auth ?? "browser") as Auth
  if (!["browser", "device"].includes(auth)) fail("auth must be browser or device")
  const tenant = text(raw.tenant, "tenant", 200) ?? "organizations"
  if (!GUID.test(tenant) && !/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(tenant) && !["organizations", "common"].includes(tenant)) fail("tenant must be a domain (contoso.com), a directory id, or organizations")
  const clientId = text(raw.clientId, "clientId", 64)
  if (clientId !== undefined && !GUID.test(clientId)) fail("clientId must be an app (client) id, a GUID")
  const preset = (raw.preset ?? "full") as Preset
  if (!["full", "read", "custom"].includes(preset)) fail("preset must be full, read or custom")
  let scopes: string[] | undefined
  if (raw.scopes !== undefined && raw.scopes !== null) {
    if (!Array.isArray(raw.scopes) || raw.scopes.length > 40 || raw.scopes.some(scope => typeof scope !== "string" || !SCOPE.test(scope))) fail("scopes must be a list of permission names like Mail.Read")
    scopes = [...new Set(raw.scopes as string[])]
  }
  if (preset === "custom" && !scopes?.length) fail("preset custom needs scopes")
  const email = text(raw.email, "email", 254)?.toLowerCase()
  if (email !== undefined && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail("email is not an email address")
  const label = text(raw.label, "label", 120)
  const notes = typeof raw.notes === "string" && raw.notes.length <= 4000 ? raw.notes : raw.notes == null ? undefined : fail("notes must be text up to 4000 characters")
  return {
    id, auth, tenant, preset, ...(clientId && { clientId }), ...(preset === "custom" && scopes && { scopes }), ...(email && { email }),
    ...(label && { label }), ...(notes && { notes }), updatedAt: now.toISOString(),
  }
}

export function mergeAccount(existing: Account | undefined, patch: Record<string, unknown>): Record<string, unknown> {
  const base: Record<string, unknown> = existing ? { ...existing } : {}
  delete base.updatedAt
  for (const [key, value] of Object.entries(patch)) {
    if (value === null) delete base[key]
    else base[key] = value
  }
  return base
}

/** The delegated scopes an account requests. */
export function requestedScopes(account: Account): string[] {
  return [...(account.preset === "custom" ? account.scopes ?? [] : SCOPE_PRESETS[account.preset]), "offline_access"]
}

/** A request an IT admin can act on when the tenant requires admin approval. */
export function adminRequest(account: Account): string {
  const scopes = requestedScopes(account).join(", ")
  return account.clientId && account.clientId !== DEFAULT_CLIENT_ID
    ? `Please grant admin consent in Entra ID (${account.tenant}) for the app ${account.clientId}, delegated permissions: ${scopes}.`
    : `Please grant admin consent in Entra ID (${account.tenant}) for Microsoft's first-party app "Microsoft Graph Command Line Tools" (app ID ${DEFAULT_CLIENT_ID}, used by Graph PowerShell), delegated permissions: ${scopes}. Delegated only: it acts as the signed-in user with that user's own access. Alternatively, register a dedicated app (public client, redirect URI http://localhost, allow public client flows) with those delegated permissions, grant consent, and send me its app ID.`
}

/** Explains Microsoft sign-in errors in plain words. */
export function explainSignInError(code: string | undefined, description: string | undefined, account: Account): string {
  const text = `${code ?? ""} ${description ?? ""}`
  if (/AADSTS90094|AADSTS65001|admin approval|consent_required|AADSTS900941|AADSTS90095/i.test(text)) return `Your organization requires an administrator to approve these permissions. Ask IT: ${adminRequest(account)}`
  if (/AADSTS50105|AADSTS50011|AADSTS700016|AADSTS7000112/i.test(text)) return `The sign-in app is not available in this tenant (${(/AADSTS\d+/.exec(text) ?? ["?"])[0]}): ask IT to enable it, or use an app registration from IT (clientId). ${adminRequest(account)}`
  if (/access_denied|AADSTS65004/i.test(text)) return "The sign-in was declined."
  return `Microsoft sign-in failed: ${(description ?? code ?? "unknown error").split("\n")[0]!.slice(0, 300)}`
}

// ------------------------------------------------------------------ what a Graph request does

const verbOrder: Verb[] = ["read", "write", "send", "delete", "admin"]

function areaOf(path: string): Area {
  if (/\/(chats?|teams|joinedteams|channels|teamwork|presences?|primarychannel|installedapps)(\/|$|\(|\?)/.test(path)) return "chat"
  if (/\/(onlinemeetings|events|calendars?|calendarview|calendargroups|findmeetingtimes|getschedule|reminderview)(\/|$|\(|\?)/.test(path)) return "calendar"
  if (/\/(drives?|sites|shares|items|lists|driveitems?)(\/|$|\(|\?)/.test(path)) return "files"
  if (/\/(messages|mailfolders|sendmail|mailboxsettings|inferenceclassification|contacts|contactfolders|outlook)(\/|$|\(|\?)/.test(path)) return "mail"
  if (/^\/(me|users(\/[^/]+)?|groups(\/[^/]+)?|people|organization|directory\w*|domains|serviceprincipals|applications|me\/people|me\/manager|me\/directreports|me\/memberof|me\/photo)(\/|$|\?)/.test(path) || /\/(manager|directreports|memberof|members|owners|photo|people)(\/|$)/.test(path)) return "directory"
  return "other"
}

/** Permission a single Graph request needs, from its method, path and (for a few cases) body. */
export function classifyGraph(method: string, url: string, body?: unknown): Capability {
  const m = method.toUpperCase()
  const path = url.replace(/^https?:\/\/[^/]+/i, "").replace(/^\/(v1\.0|beta)(?=\/)/i, "").split("?")[0]!.toLowerCase().replace(/\/+$/, "") || "/"
  const area = path === "/search/query" ? "other" : areaOf(path)
  if (m === "GET" || m === "HEAD") return `${area}-read`
  if (m === "POST" && /\/(getschedule|findmeetingtimes|getbyids|search\/query|getmembergroups|checkmembergroups|getmemberobjects|delta|preview|validateproperties|microsoft\.graph\.getoffice365activeuserdetail)$/.test(path)) return `${area}-read`
  if (m === "DELETE" || /\/(permanentdelete)$/.test(path)) return `${area}-delete`
  if (area === "mail" && m === "POST" && /\/(sendmail|send|reply|replyall|forward)$/.test(path)) return "mail-send"
  if (area === "calendar") {
    if (m === "POST" && /\/(accept|decline|tentativelyaccept|cancel|forward|proposenewtime)$/.test(path)) return "calendar-send"
    const attendees = (body as { attendees?: unknown[] } | undefined)?.attendees
    if ((m === "POST" || m === "PATCH") && Array.isArray(attendees) && attendees.length) return "calendar-send"
  }
  if (area === "chat") {
    if (m === "POST" && /\/(messages|replies|sendactivitynotification)$/.test(path)) return "chat-send"
    if (m === "POST" && (/^\/teams$/.test(path) || /\/(channels|members|installedapps|clone|archive|unarchive)$/.test(path))) return "chat-admin"
    if (/\/members(\/|$)/.test(path)) return "chat-admin"
  }
  if (area === "files" && m === "POST" && /\/(createlink|invite|permissions|grant)$/.test(path)) return "files-send"
  if (area === "files" && /\/permissions(\/|$)/.test(path)) return "files-send"
  if (area === "directory") return "directory-admin"
  return `${area}-write`
}

/** The permissions a request needs; a $batch needs those of each request inside. Sorted, unique. */
export function graphCapabilities(method: string, url: string, body?: unknown): Capability[] {
  const path = url.replace(/^https?:\/\/[^/]+/i, "").replace(/^\/(v1\.0|beta)(?=\/)/i, "").split("?")[0]!.toLowerCase()
  if (method.toUpperCase() === "POST" && path === "/$batch") {
    const requests = (body as { requests?: { method?: string; url?: string; body?: unknown }[] } | undefined)?.requests
    if (!Array.isArray(requests) || !requests.length) return ["other-write"]
    const all = requests.map(request => classifyGraph(String(request.method ?? "GET"), `/${String(request.url ?? "").replace(/^\//, "")}`, request.body))
    return [...new Set(all)].sort((a, b) => verbOrder.indexOf(a.split("-")[1] as Verb) - verbOrder.indexOf(b.split("-")[1] as Verb))
  }
  return [classifyGraph(method, url, body)]
}
