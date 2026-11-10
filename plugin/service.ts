import { createReadStream, createWriteStream } from "node:fs"
import { mkdir, stat } from "node:fs/promises"
import { dirname } from "node:path"
import { Readable } from "node:stream"
import { pipeline } from "node:stream/promises"
import { PolicyStore, Redactor, Serial, Vault, type PermissionBridge } from "oc-secure-kit"
import { ACCESS, DEFAULT_AUTHORITY, DEFAULT_CLIENT_ID, DEFAULT_GRAPH, explainSignInError, REFRESH, requestedScopes, SECRET_NAMES, TOKEN_META, type Account, type Discovery } from "./model.js"
import { refreshGrant, SignInError, startBrowser, startDevice, type Endpoint, type Grant, type PendingSignIn } from "./oauth.js"
import { AccountStore } from "./store.js"

export const PLUGIN_ID = "oc-m365"
const UPLOAD_SIMPLE_LIMIT = 4 * 1024 * 1024
const UPLOAD_CHUNK = 10 * 320 * 1024

export interface SignInState {
  readonly account: string
  readonly url: string
  readonly code?: string
  readonly expiresAt: number
  status: "waiting" | "done" | "failed"
  error?: string
  readonly flow: PendingSignIn
}

export class GraphError extends Error {
  constructor(message: string, readonly status: number, readonly code?: string) { super(message) }
}

export interface GraphResponse { readonly status: number; readonly data: unknown; readonly headers: Record<string, string> }

export class M365Service {
  readonly redactor = new Redactor()
  readonly signIns = new Map<string, SignInState>()
  #bridge?: PermissionBridge
  #redactorStale = false
  #refreshing = new Map<string, Serial>()

  private constructor(
    readonly dir: string,
    readonly store: AccountStore,
    readonly vault: Vault,
    readonly policy: PolicyStore,
    readonly authority: string,
    readonly graphUrl: string,
  ) {}

  static async open(dir: string, options: { authority?: string; graphUrl?: string } = {}): Promise<M365Service> {
    const [store, vault, policy] = await Promise.all([AccountStore.open(dir), Vault.open({ dir, plugin: PLUGIN_ID }), PolicyStore.open(dir)])
    const service = new M365Service(dir, store, vault, policy, (options.authority ?? DEFAULT_AUTHORITY).replace(/\/+$/, ""), (options.graphUrl ?? DEFAULT_GRAPH).replace(/\/+$/, ""))
    await service.refreshRedactor()
    vault.onChange(() => { service.#redactorStale = true })
    return service
  }

  attach(bridge: PermissionBridge): void { this.#bridge = bridge }
  get bridge(): PermissionBridge {
    if (!this.#bridge) throw new Error(`${PLUGIN_ID}: permissions are not set up`)
    return this.#bridge
  }

  async pick(id: string | undefined): Promise<Account> {
    const accounts = await this.store.list()
    if (id) {
      const found = accounts.find(account => account.id === id || account.email === id.toLowerCase())
      if (!found) throw new Error(`No Microsoft 365 account ${JSON.stringify(id)}. Accounts: ${accounts.map(account => account.id).join(", ") || "(none; create one with m365_account_upsert)"}`)
      return found
    }
    if (accounts.length === 1) return accounts[0]!
    const fallback = (await this.store.settings()).defaultAccount
    if (fallback) { const found = accounts.find(account => account.id === fallback); if (found) return found }
    if (!accounts.length) throw new Error("No Microsoft 365 account is configured yet. Create one with m365_account_upsert (tenant, email), then m365_signin.")
    throw new Error(`Several Microsoft 365 accounts exist (${accounts.map(account => account.id).join(", ")}). Pass account=<id>; decide from the conversation or ask the user.`)
  }

  endpoint(account: Account): Endpoint { return { authority: this.authority, tenant: account.tenant, clientId: account.clientId ?? DEFAULT_CLIENT_ID } }

  // ------------------------------------------------------------------ tokens

  async saveGrant(account: Account, grant: Grant): Promise<void> {
    await this.vault.set(account.id, ACCESS, grant.accessToken)
    if (grant.refreshToken) await this.vault.set(account.id, REFRESH, grant.refreshToken)
    await this.vault.set(account.id, TOKEN_META, JSON.stringify({ expiresAt: grant.expiresAt, scopes: grant.scopes }))
    await this.refreshRedactor()
  }

  async signedIn(account: Account): Promise<boolean> { return this.vault.has(account.id, REFRESH) || this.vault.has(account.id, ACCESS) }

  async granted(account: Account): Promise<string[]> {
    const meta = await this.vault.get(account.id, TOKEN_META)
    return meta ? (JSON.parse(meta) as { scopes?: string[] }).scopes ?? [] : []
  }

  /** A valid access token, refreshed when it is about to expire (or `force`). */
  async accessToken(account: Account, force = false): Promise<string> {
    const serial = this.#refreshing.get(account.id) ?? new Serial()
    this.#refreshing.set(account.id, serial)
    return serial.run(async () => {
      const token = await this.vault.get(account.id, ACCESS)
      const meta = await this.vault.get(account.id, TOKEN_META)
      const expiresAt = meta ? (JSON.parse(meta) as { expiresAt?: number }).expiresAt ?? 0 : 0
      if (token && !force && expiresAt - Date.now() > 5 * 60_000) return token
      const refresh = await this.vault.get(account.id, REFRESH)
      if (!refresh) {
        if (token) return token
        throw new Error(`${account.id} is not signed in to Microsoft 365. Call m365_signin account=${account.id} and give the user the link or code.`)
      }
      try {
        const grant = await refreshGrant(this.endpoint(account), refresh, requestedScopes(account))
        await this.saveGrant(account, { ...grant, refreshToken: grant.refreshToken ?? refresh })
        return grant.accessToken
      } catch (error) {
        if (error instanceof SignInError) throw new Error(`${account.id}'s sign-in expired or was revoked: sign in again (m365_signin). ${explainSignInError(error.code, error.description, account)}`)
        throw error
      }
    })
  }

  // ------------------------------------------------------------------ Graph

  /** Resolves a Graph path ("/me", "v1.0/me", "/beta/…", or a full graph URL) to a URL. */
  graphHref(path: string): string {
    if (/^https?:\/\//i.test(path)) {
      const url = new URL(path)
      if (url.origin !== new URL(this.graphUrl).origin) throw new Error(`Only Microsoft Graph URLs (${this.graphUrl}) are allowed`)
      return url.href
    }
    const clean = path.startsWith("/") ? path : `/${path}`
    return `${this.graphUrl}${/^\/(v1\.0|beta)\//i.test(clean) ? clean : `/v1.0${clean}`}`
  }

  async graph(account: Account, method: string, path: string, init: { body?: unknown; headers?: Record<string, string> } = {}): Promise<GraphResponse> {
    const href = this.graphHref(path)
    const send = async (force: boolean) => fetch(href, {
      method: method.toUpperCase(), redirect: "follow",
      headers: { authorization: `Bearer ${await this.accessToken(account, force)}`, accept: "application/json", ...(init.body !== undefined && { "content-type": "application/json" }), ...init.headers },
      ...(init.body !== undefined && { body: typeof init.body === "string" ? init.body : JSON.stringify(init.body) }),
    })
    let response = await send(false)
    if (response.status === 401) response = await send(true)
    const type = response.headers.get("content-type") ?? ""
    const text = await response.text()
    const data = /json/.test(type) && text ? (() => { try { return JSON.parse(text) } catch { return text } })() : text || undefined
    if (!response.ok) {
      const error = (data as { error?: { code?: string; message?: string } } | undefined)?.error
      const hint = response.status === 403 ? " (the signed-in user or the granted permissions do not allow this; granted: " + (await this.granted(account)).join(" ") + ")" : ""
      throw new GraphError(`Graph ${method.toUpperCase()} ${new URL(href).pathname} → ${response.status} ${error?.code ?? ""}: ${(error?.message ?? (typeof data === "string" ? data : "")).slice(0, 800)}${hint}`, response.status, error?.code)
    }
    return { status: response.status, data, headers: Object.fromEntries(["content-type", "location", "retry-after"].flatMap(name => response.headers.get(name) ? [[name, response.headers.get(name)!]] : [])) }
  }

  /** Downloads Graph content (a driveItem's /content, an attachment's /$value, …) to a local file. */
  async download(account: Account, path: string, to: string): Promise<{ bytes: number; contentType?: string }> {
    const href = this.graphHref(path)
    const response = await fetch(href, { headers: { authorization: `Bearer ${await this.accessToken(account)}` }, redirect: "follow" })
    if (!response.ok || !response.body) throw new GraphError(`Download ${new URL(href).pathname} → ${response.status}: ${(await response.text().catch(() => "")).slice(0, 400)}`, response.status)
    await mkdir(dirname(to), { recursive: true })
    await pipeline(Readable.fromWeb(response.body as never), createWriteStream(to))
    return { bytes: (await stat(to)).size, ...(response.headers.get("content-type") && { contentType: response.headers.get("content-type")! }) }
  }

  /** Uploads a local file to a drive path like /me/drive/root:/Folder/name.ext: (any size; large files in chunks). */
  async upload(account: Account, from: string, target: string, conflict: "replace" | "rename" | "fail" = "replace"): Promise<unknown> {
    const size = (await stat(from)).size
    const item = target.replace(/:?\/content$/i, "").replace(/:?$/, ":")
    if (size <= UPLOAD_SIMPLE_LIMIT) {
      const href = this.graphHref(`${item}/content?@microsoft.graph.conflictBehavior=${conflict}`)
      const body = await new Response(Readable.toWeb(createReadStream(from)) as never).arrayBuffer()
      const response = await fetch(href, { method: "PUT", headers: { authorization: `Bearer ${await this.accessToken(account)}`, "content-type": "application/octet-stream" }, body })
      const data = await response.json().catch(() => ({}))
      if (!response.ok) throw new GraphError(`Upload → ${response.status}: ${JSON.stringify(data).slice(0, 400)}`, response.status)
      return data
    }
    const session = (await this.graph(account, "POST", `${item}/createUploadSession`, { body: { item: { "@microsoft.graph.conflictBehavior": conflict } } })).data as { uploadUrl: string }
    const handle = await import("node:fs/promises").then(fs => fs.open(from, "r"))
    try {
      let offset = 0, result: unknown
      while (offset < size) {
        const length = Math.min(UPLOAD_CHUNK, size - offset)
        const chunk = Buffer.alloc(length)
        await handle.read(chunk, 0, length, offset)
        const response = await fetch(session.uploadUrl, { method: "PUT", headers: { "content-length": String(length), "content-range": `bytes ${offset}-${offset + length - 1}/${size}` }, body: chunk })
        if (!response.ok) throw new GraphError(`Upload chunk at ${offset} → ${response.status}: ${(await response.text()).slice(0, 300)}`, response.status)
        result = await response.json().catch(() => undefined)
        offset += length
      }
      return result
    } finally { await handle.close() }
  }

  // ------------------------------------------------------------------ discovery and sign-in

  async rediscover(account: Account): Promise<Discovery> {
    const previous = await this.store.discovery(account.id)
    let discovery: Discovery
    try {
      const me = (await this.graph(account, "GET", "/me?$select=displayName,mail,userPrincipalName")).data as { displayName?: string; mail?: string; userPrincipalName?: string }
      const granted = await this.granted(account)
      let unread: number | undefined
      if (granted.some(scope => /^Mail\./i.test(scope))) {
        try { unread = ((await this.graph(account, "GET", "/me/mailFolders/inbox?$select=unreadItemCount")).data as { unreadItemCount?: number }).unreadItemCount } catch {}
      }
      const wanted = account.email
      const problem = wanted && me.mail && me.mail.toLowerCase() !== wanted && me.userPrincipalName?.toLowerCase() !== wanted
        ? `Signed in as ${me.mail} (${me.userPrincipalName}), but this account is set to ${wanted}. Sign in again with ${wanted}, or change the account's email.` : undefined
      discovery = {
        ...(me.displayName && { name: me.displayName }), ...(me.mail && { email: me.mail }), ...(me.userPrincipalName && { upn: me.userPrincipalName }),
        granted, ...(unread !== undefined && { unread }), checkedAt: new Date().toISOString(), ...(problem && { error: problem }),
      }
    } catch (error) {
      discovery = { ...(previous ?? {}), error: (error as Error).message.slice(0, 500), checkedAt: new Date().toISOString() }
    }
    await this.store.setDiscovery(account.id, discovery)
    return discovery
  }

  async startSignIn(account: Account, options: { device?: boolean } = {}): Promise<SignInState> {
    const running = this.signIns.get(account.id)
    if (running?.status === "waiting" && running.expiresAt > Date.now()) return running
    running?.flow.cancel()
    const device = options.device ?? account.auth === "device"
    const scopes = requestedScopes(account)
    let flow: PendingSignIn
    try { flow = device ? await startDevice(this.endpoint(account), scopes) : await startBrowser(this.endpoint(account), scopes, { ...(account.email && { loginHint: account.email }) }) }
    catch (error) { throw new Error(error instanceof SignInError ? explainSignInError(error.code, error.description, account) : (error as Error).message) }
    const state: SignInState = { account: account.id, url: flow.url, ...(flow.code && { code: flow.code }), expiresAt: flow.expiresAt, status: "waiting", flow }
    this.signIns.set(account.id, state)
    flow.done
      .then(async grant => { await this.saveGrant(account, grant); await this.rediscover(account); state.status = "done" })
      .catch(error => {
        if (state.status !== "waiting") return
        state.status = "failed"
        state.error = error instanceof SignInError ? explainSignInError(error.code, error.description, account) : (error as Error).message
      })
    return state
  }

  async signOut(account: Account): Promise<void> {
    for (const name of SECRET_NAMES) await this.vault.delete(account.id, name)
    await this.store.setDiscovery(account.id, undefined)
    await this.refreshRedactor()
  }

  async syncRedactor(): Promise<void> {
    await this.vault.list()
    if (this.#redactorStale) { this.#redactorStale = false; await this.refreshRedactor() }
  }
  async refreshRedactor(): Promise<void> {
    this.redactor.replaceAll((await this.vault.values()).filter(item => item.name !== TOKEN_META).map(item => [`${item.entity}/${item.name}`, item.value] as const))
  }

  dispose(): void { for (const state of this.signIns.values()) state.flow.cancel() }
}
