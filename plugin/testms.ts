// Test helper: a fake Microsoft identity platform (auth code + PKCE with a localhost redirect, device code, refresh,
// "approval required") and a fake Microsoft Graph (me, mail, send, calendar, Teams, drive upload/download, $batch).
import { createHash, randomBytes } from "node:crypto"
import { createServer, type IncomingMessage, type ServerResponse } from "node:http"

export interface FakeMicrosoft {
  readonly authority: string
  readonly graphUrl: string
  /** The browser: follow the authorize URL to the localhost redirect. */
  approve(authorizeUrl: string): Promise<void>
  /** The user enters the device code. */
  approveDevice(): void
  /** Tenants that require admin approval for user sign-ins. */
  readonly needsAdmin: Set<string>
  readonly sent: { to: string; subject: string }[]
  readonly files: Map<string, Buffer>
  readonly requests: { method: string; path: string }[]
  expireTokens(): void
  stop(): Promise<void>
}

export async function startFakeMicrosoft(): Promise<FakeMicrosoft> {
  const codes = new Map<string, { challenge: string; redirect: string; scope: string }>()
  const access = new Map<string, number>()
  const refresh = new Map<string, string>()
  const devices = new Map<string, { scope: string; approved: boolean }>()
  const needsAdmin = new Set<string>()
  const sent: FakeMicrosoft["sent"] = []
  const files = new Map<string, Buffer>([["Reports/q3.txt", Buffer.from("Q3 revenue up 12%\n")]])
  const requests: FakeMicrosoft["requests"] = []
  const drafts = new Map<string, { subject: string; to: string }>()
  let lastDevice = ""
  let base = ""
  const send = (res: ServerResponse, status: number, body: unknown, headers: Record<string, string> = {}) => { res.writeHead(status, { "content-type": "application/json", ...headers }); res.end(JSON.stringify(body)) }
  const read = async (req: IncomingMessage) => { const chunks: Buffer[] = []; for await (const chunk of req) chunks.push(chunk as Buffer); return Buffer.concat(chunks) }
  const grant = (scope: string) => {
    const token = `eyJ0eXAi.${randomBytes(16).toString("hex")}.sig`; access.set(token, Date.now() + 3600_000)
    const r = `0.AR8A${randomBytes(24).toString("hex")}`; refresh.set(r, scope)
    return { token_type: "Bearer", access_token: token, refresh_token: r, expires_in: 3599, scope: scope.split(" ").filter(s => s !== "offline_access").map(s => /^https?:/.test(s) ? s : s).join(" ") }
  }
  const message = { id: "msg-1", subject: "Budget review", from: { emailAddress: { address: "boss@acme.example", name: "Boss" } }, isRead: false, bodyPreview: "Please review", hasAttachments: true }

  const graph = async (req: IncomingMessage, res: ServerResponse, method: string, path: string, body: unknown): Promise<{ status: number; body?: unknown; raw?: Buffer }> => {
    requests.push({ method, path })
    const p = path.split("?")[0]!
    if (p === "/v1.0/me") return { status: 200, body: { displayName: "Pat Lee", mail: "pat@acme.example", userPrincipalName: "pat@acme.example" } }
    if (p === "/v1.0/me/mailFolders/inbox") return { status: 200, body: { unreadItemCount: 3 } }
    if (p === "/v1.0/me/messages" && method === "GET") return { status: 200, body: { value: [message] } }
    if (p === "/v1.0/me/messages" && method === "POST") { const b = body as { subject: string; toRecipients: { emailAddress: { address: string } }[] }; const id = `draft-${drafts.size + 1}`; drafts.set(id, { subject: b.subject, to: b.toRecipients[0]!.emailAddress.address }); return { status: 201, body: { id, isDraft: true, subject: b.subject } } }
    const sendDraft = /^\/v1\.0\/me\/messages\/([^/]+)\/send$/.exec(p)
    if (sendDraft && method === "POST") { const draft = drafts.get(sendDraft[1]!); if (!draft) return { status: 404, body: { error: { code: "ErrorItemNotFound", message: "not found" } } }; sent.push(draft); return { status: 202 } }
    if (p === "/v1.0/me/sendMail") { const b = body as { message: { subject: string; toRecipients: { emailAddress: { address: string } }[] } }; sent.push({ subject: b.message.subject, to: b.message.toRecipients[0]!.emailAddress.address }); return { status: 202 } }
    if (p === "/v1.0/me/messages/msg-1/attachments/att-1/$value") return { status: 200, raw: Buffer.from("%PDF-fake\n") }
    if (p === "/v1.0/me/events" && method === "POST") return { status: 201, body: { id: "ev-1", ...(body as object) } }
    if (p === "/v1.0/me/calendarView") return { status: 200, body: { value: [{ subject: "Standup", start: { dateTime: "2026-11-10T16:00:00" } }] } }
    if (p === "/v1.0/me/joinedTeams") return { status: 200, body: { value: [{ id: "team-1", displayName: "Robotics" }] } }
    if (/^\/v1\.0\/teams\/team-1\/channels\/[^/]+\/messages$/.test(p) && method === "POST") return { status: 201, body: { id: "post-1" } }
    if (p === "/v1.0/me/chats") return { status: 403, body: { error: { code: "Forbidden", message: "Missing scope permissions on the request. API requires one of 'Chat.Read'." } } }
    const drive = /^\/v1\.0\/me\/drive\/root:\/(.+?):\/content$/.exec(decodeURIComponent(p))
    if (drive && method === "GET") { const file = files.get(drive[1]!); return file ? { status: 200, raw: file } : { status: 404, body: { error: { code: "itemNotFound", message: "not found" } } } }
    if (drive && method === "PUT") { files.set(drive[1]!, body as Buffer); return { status: 201, body: { id: `item-${files.size}`, name: drive[1]!.split("/").at(-1), size: (body as Buffer).length, webUrl: `https://acme-my.sharepoint.com/${drive[1]}` } } }
    const session = /^\/v1\.0\/me\/drive\/root:\/(.+?):\/createUploadSession$/.exec(decodeURIComponent(p))
    if (session) return { status: 200, body: { uploadUrl: `${base}/upload/${encodeURIComponent(session[1]!)}` } }
    if (/\/createLink$/.test(p)) return { status: 201, body: { link: { webUrl: "https://acme.sharepoint.com/s/x" } } }
    return { status: 404, body: { error: { code: "notFound", message: `fake graph: ${method} ${p}` } } }
  }

  const uploads = new Map<string, Buffer[]>()
  const server = createServer(async (req, res) => {
    const url = new URL(req.url!, base)
    const path = url.pathname
    const auth = /^\/([^/]+)\/oauth2\/v2\.0\/(authorize|token|devicecode)$/.exec(path)
    if (auth) {
      const [, tenant, endpoint] = auth
      if (endpoint === "authorize") {
        const p = url.searchParams
        if (needsAdmin.has(tenant!)) { res.writeHead(302, { location: `${p.get("redirect_uri")}?error=access_denied&error_description=${encodeURIComponent("AADSTS90094: Admin approval required.")}&state=${p.get("state")}` }); res.end(); return }
        if (!/^http:\/\/localhost:\d+$/.test(p.get("redirect_uri") ?? "") || p.get("code_challenge_method") !== "S256") return send(res, 400, { error: "invalid_request" })
        const code = randomBytes(8).toString("hex")
        codes.set(code, { challenge: p.get("code_challenge")!, redirect: p.get("redirect_uri")!, scope: p.get("scope")! })
        res.writeHead(302, { location: `${p.get("redirect_uri")}?code=${code}&state=${p.get("state")}` }); res.end(); return
      }
      const form = new URLSearchParams((await read(req)).toString())
      if (endpoint === "devicecode") {
        if (needsAdmin.has(tenant!)) return send(res, 400, { error: "invalid_grant", error_description: "AADSTS65001: The user or administrator has not consented to use the application." })
        lastDevice = randomBytes(8).toString("hex"); devices.set(lastDevice, { scope: form.get("scope")!, approved: false })
        return send(res, 200, { device_code: lastDevice, user_code: "FAKE-CODE", verification_uri: "https://microsoft.com/devicelogin", expires_in: 900, interval: 1 })
      }
      const type = form.get("grant_type")
      if (type === "authorization_code") {
        const entry = codes.get(form.get("code") ?? "")
        if (!entry || entry.redirect !== form.get("redirect_uri") || createHash("sha256").update(form.get("code_verifier") ?? "").digest("base64url") !== entry.challenge) return send(res, 400, { error: "invalid_grant" })
        codes.delete(form.get("code")!)
        return send(res, 200, grant(entry.scope))
      }
      if (type === "refresh_token") { const scope = refresh.get(form.get("refresh_token") ?? ""); if (!scope) return send(res, 400, { error: "invalid_grant", error_description: "AADSTS70008: expired" }); refresh.delete(form.get("refresh_token")!); return send(res, 200, grant(scope)) }
      if (type === "urn:ietf:params:oauth:grant-type:device_code") {
        const device = devices.get(form.get("device_code") ?? "")
        if (!device) return send(res, 400, { error: "expired_token" })
        if (!device.approved) return send(res, 400, { error: "authorization_pending" })
        devices.delete(form.get("device_code")!)
        return send(res, 200, grant(device.scope))
      }
      return send(res, 400, { error: "unsupported_grant_type" })
    }
    if (path.startsWith("/upload/")) {
      const name = decodeURIComponent(path.slice(8))
      const chunk = await read(req)
      const range = /bytes (\d+)-(\d+)\/(\d+)/.exec(String(req.headers["content-range"]))!
      const parts = uploads.get(name) ?? []; parts.push(chunk); uploads.set(name, parts)
      if (Number(range[2]) + 1 < Number(range[3])) return send(res, 202, { nextExpectedRanges: [`${Number(range[2]) + 1}-`] })
      files.set(name, Buffer.concat(parts)); uploads.delete(name)
      return send(res, 201, { id: "item-big", name: name.split("/").at(-1), size: Number(range[3]) })
    }
    if (path.startsWith("/v1.0/") || path === "/v1.0" || path.startsWith("/beta/")) {
      const token = req.headers.authorization?.replace(/^Bearer /, "") ?? ""
      if (!(access.get(token)! > Date.now())) return send(res, 401, { error: { code: "InvalidAuthenticationToken", message: "Access token has expired" } })
      const raw = await read(req)
      const isJson = /json/.test(String(req.headers["content-type"] ?? ""))
      const body = raw.length ? (isJson ? JSON.parse(raw.toString()) : raw) : undefined
      if (path === "/v1.0/$batch") {
        const responses = await Promise.all((body as { requests: { id: string; method: string; url: string; body?: unknown }[] }).requests.map(async item => { const r = await graph(req, res, item.method, `/v1.0${item.url.startsWith("/") ? item.url : `/${item.url}`}`, item.body); return { id: item.id, status: r.status, body: r.body } }))
        return send(res, 200, { responses })
      }
      const result = await graph(req, res, req.method!, `${path}${url.search}`, body)
      if (result.raw) { res.writeHead(result.status, { "content-type": "application/octet-stream" }); res.end(result.raw); return }
      if (result.body === undefined) { res.writeHead(result.status); res.end(); return }
      return send(res, result.status, result.body)
    }
    send(res, 404, { error: "not found" })
  })
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve))
  base = `http://127.0.0.1:${(server.address() as { port: number }).port}`
  return {
    authority: base, graphUrl: base, needsAdmin, sent, files, requests,
    async approve(authorizeUrl) {
      const redirect = await fetch(authorizeUrl, { redirect: "manual" })
      const location = redirect.headers.get("location")
      if (!location) throw new Error(`authorize failed: ${redirect.status} ${await redirect.text()}`)
      await fetch(location.replace("http://localhost:", "http://127.0.0.1:"))
    },
    approveDevice() { const device = devices.get(lastDevice); if (device) device.approved = true },
    expireTokens() { for (const token of access.keys()) access.set(token, 0) },
    async stop() { await new Promise(resolve => server.close(resolve)) },
  }
}
