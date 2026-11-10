import { createHash, randomBytes } from "node:crypto"
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http"

/** Microsoft identity platform (Entra ID, v2 endpoints) for a public client: auth code + PKCE, device code, refresh. */

export interface Grant {
  readonly accessToken: string
  readonly refreshToken?: string
  readonly expiresAt: number
  readonly scopes: string[]
}

export class SignInError extends Error {
  constructor(readonly code: string | undefined, readonly description: string | undefined) { super(description ?? code ?? "sign-in failed") }
}

export interface Endpoint { readonly authority: string; readonly tenant: string; readonly clientId: string }
const base = (endpoint: Endpoint) => `${endpoint.authority.replace(/\/+$/, "")}/${encodeURIComponent(endpoint.tenant)}/oauth2/v2.0`

async function token(endpoint: Endpoint, params: Record<string, string>): Promise<Grant> {
  const response = await fetch(`${base(endpoint)}/token`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: endpoint.clientId, ...params }) })
  const data = await response.json().catch(() => ({})) as Record<string, unknown>
  if (!data.access_token) throw new SignInError(String(data.error ?? response.status), typeof data.error_description === "string" ? data.error_description : undefined)
  return {
    accessToken: String(data.access_token), ...(typeof data.refresh_token === "string" && { refreshToken: data.refresh_token }),
    expiresAt: Date.now() + (Number(data.expires_in) || 3600) * 1000,
    scopes: typeof data.scope === "string" ? data.scope.split(/\s+/).filter(Boolean) : [],
  }
}

export const refreshGrant = (endpoint: Endpoint, refreshToken: string, scopes: readonly string[]) =>
  token(endpoint, { grant_type: "refresh_token", refresh_token: refreshToken, scope: scopes.join(" ") })

const b64 = (buffer: Buffer) => buffer.toString("base64url")

export interface PendingSignIn {
  /** browser: open this; device: where to enter the code. */
  readonly url: string
  readonly code?: string
  readonly expiresAt: number
  readonly done: Promise<Grant>
  cancel(): void
}

/** Starts a browser sign-in with a loopback redirect (http://localhost:<port>, as Microsoft requires for desktop apps). */
export async function startBrowser(endpoint: Endpoint, scopes: readonly string[], options: { loginHint?: string; timeoutMs?: number } = {}): Promise<PendingSignIn> {
  let settle!: (value: URLSearchParams) => void, fail!: (error: Error) => void
  const result = new Promise<URLSearchParams>((resolve, reject) => { settle = resolve; fail = reject })
  const verifier = b64(randomBytes(32)), state = b64(randomBytes(16))
  const handler = (req: IncomingMessage, res: ServerResponse) => {
    const url = new URL(req.url ?? "/", "http://localhost")
    if (url.searchParams.get("state") !== state) { res.writeHead(404); res.end(); return }
    const ok = !!url.searchParams.get("code")
    res.writeHead(200, { "content-type": "text/html; charset=utf-8" })
    res.end(`<!doctype html><title>Microsoft 365 sign-in</title><body style="font:16px system-ui;margin:3em">${ok ? "<h2>Signed in to Microsoft 365.</h2><p>You can close this tab and return to OpenCode.</p>" : "<h2>Sign-in did not complete.</h2><p>Return to OpenCode for details.</p>"}</body>`)
    settle(url.searchParams)
  }
  // "localhost" may resolve to IPv4 or IPv6: listen on both loopback addresses with one port.
  const servers: Server[] = [createServer(handler)]
  await new Promise<void>(resolve => servers[0]!.listen(0, "127.0.0.1", resolve))
  const port = (servers[0]!.address() as { port: number }).port
  const v6 = createServer(handler)
  await new Promise<void>(resolve => { v6.once("error", () => resolve()); v6.listen(port, "::1", () => { servers.push(v6); resolve() }) })
  const redirect = `http://localhost:${port}`
  const url = new URL(`${base(endpoint)}/authorize`)
  for (const [key, value] of Object.entries({
    client_id: endpoint.clientId, response_type: "code", redirect_uri: redirect, response_mode: "query", scope: scopes.join(" "), state,
    code_challenge: b64(createHash("sha256").update(verifier).digest()), code_challenge_method: "S256", prompt: "select_account",
    ...(options.loginHint && { login_hint: options.loginHint }),
  })) url.searchParams.set(key, value)
  const timeoutMs = options.timeoutMs ?? 15 * 60_000
  const timer = setTimeout(() => fail(new SignInError("timeout", "The sign-in was not finished in time. If Microsoft showed 'Approval required', an administrator must approve the permissions first.")), timeoutMs)
  const done = result.then(params => {
    const code = params.get("code")
    if (!code) throw new SignInError(params.get("error") ?? undefined, params.get("error_description") ?? undefined)
    return token(endpoint, { grant_type: "authorization_code", code, redirect_uri: redirect, code_verifier: verifier, scope: scopes.join(" ") })
  }).finally(() => { clearTimeout(timer); for (const server of servers) server.close() })
  done.catch(() => {})
  return { url: url.href, expiresAt: Date.now() + timeoutMs, done, cancel: () => fail(new SignInError("cancelled", "Sign-in cancelled")) }
}

/** Starts a device-code sign-in: the user enters a code at Microsoft's page on any device. */
export async function startDevice(endpoint: Endpoint, scopes: readonly string[]): Promise<PendingSignIn> {
  const response = await fetch(`${base(endpoint)}/devicecode`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: endpoint.clientId, scope: scopes.join(" ") }) })
  const data = await response.json().catch(() => ({})) as Record<string, unknown>
  if (!data.device_code) throw new SignInError(String(data.error ?? response.status), typeof data.error_description === "string" ? data.error_description : undefined)
  let cancelled = false
  const expiresAt = Date.now() + (Number(data.expires_in) || 900) * 1000
  const done = (async () => {
    let interval = (Number(data.interval) || 5) * 1000
    while (!cancelled && Date.now() < expiresAt) {
      await new Promise(resolve => setTimeout(resolve, interval))
      if (cancelled) break
      try { return await token(endpoint, { grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: String(data.device_code) }) }
      catch (error) {
        const code = (error as SignInError).code
        if (code === "authorization_pending") continue
        if (code === "slow_down") { interval += 5000; continue }
        throw error
      }
    }
    throw new SignInError(cancelled ? "cancelled" : "expired_token", cancelled ? "Sign-in cancelled" : "The code expired; start again")
  })()
  done.catch(() => {})
  return { url: String(data.verification_uri ?? "https://microsoft.com/devicelogin"), code: String(data.user_code), expiresAt, done, cancel: () => { cancelled = true } }
}
