// Microsoft sign-in for a Microsoft 365 account from a terminal (the GUI and m365_signin do the same).
//   pixi run signin <account> [--tenant contoso.com] [--email you@contoso.com] [--device] [--read-only] [--client-id <app id>]
// Never prints tokens.
import { spawn } from "node:child_process"
import { pluginDataDir } from "oc-secure-kit"
import { adminRequest, mergeAccount, normalizeAccount } from "../plugin/model.ts"
import { M365Service } from "../plugin/service.ts"

const args = process.argv.slice(2)
const option = name => { const i = args.indexOf(name); return i >= 0 && args[i + 1] && !args[i + 1].startsWith("--") ? args[i + 1] : undefined }
const id = args[0] && !args[0].startsWith("--") ? args[0] : undefined
if (!id) { console.error("Usage: pixi run signin <account> [--tenant contoso.com] [--email you@contoso.com] [--device] [--read-only] [--client-id <app id>]"); process.exit(2) }
const service = await M365Service.open(pluginDataDir("oc-m365", process.env.OC_M365_DATA_DIR || undefined))
let account = await service.store.get(id)
const patch = { ...(option("--tenant") && { tenant: option("--tenant") }), ...(option("--email") && { email: option("--email") }), ...(option("--client-id") && { clientId: option("--client-id") }), ...(args.includes("--read-only") && { preset: "read" }) }
if (!account || Object.keys(patch).length) account = await service.store.put(normalizeAccount({ ...mergeAccount(account, patch), id }))
let state
try { state = await service.startSignIn(account, { ...(args.includes("--device") && { device: true }) }) }
catch (error) { console.error(error.message); process.exit(1) }
if (state.code) console.log(`\nOn any device, open ${state.url} and enter the code ${state.code}.\n`)
else {
  console.log(`\nOpening your browser. Sign in${account.email ? ` as ${account.email}` : ""} and accept. If it does not open, visit:\n\n${state.url}\n`)
  spawn(process.platform === "darwin" ? "open" : "xdg-open", [state.url], { stdio: "ignore", detached: true }).on("error", () => {}).unref()
}
console.log(`If Microsoft shows "Approval required", stop (Ctrl-C) and ask IT:\n${adminRequest(account)}\n`)
while (state.status === "waiting") await new Promise(resolve => setTimeout(resolve, 500))
if (state.status !== "done") { console.error(state.error); process.exit(1) }
const d = await service.store.discovery(id)
console.log(`Signed in as ${d?.name} <${d?.email ?? d?.upn}>. Granted: ${(d?.granted ?? []).join(" ")}.`)
service.dispose(); process.exit(0)
