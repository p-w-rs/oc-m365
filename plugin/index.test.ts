import { after, before, test } from "node:test"
import assert from "node:assert/strict"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join } from "node:path"
import type { Plugin } from "@opencode/plugin"
import { PolicyStore, Vault } from "oc-secure-kit"
import plugin, { createPlugin, parseOptions, PLUGIN_ID } from "./index.js"
import { startFakeMicrosoft, type FakeMicrosoft } from "./testms.js"

const dirs: string[] = []
let fake: FakeMicrosoft
before(async () => { fake = await startFakeMicrosoft() })
after(async () => { await fake.stop(); await Promise.all(dirs.map(dir => rm(dir, { recursive: true, force: true }))) })

type Effect = "allow" | "ask" | "deny"
type Tool = { name: string; description: string; execute: (input: unknown, context: unknown) => Promise<{ content: string }> }

async function host(reply: "approve" | "reject" = "approve", options: Record<string, unknown> = {}) {
  const dir = await mkdtemp(join(tmpdir(), "ocms-"))
  dirs.push(dir)
  const hooks = { evaluate: [] as any[], after: [] as any[] }
  const transforms: ((editor: any) => void)[] = []
  const tools = new Map<string, Tool>()
  const skills: { id: string; content: string }[] = []
  const prompts: { action: string; metadata?: Record<string, unknown> }[] = []
  let rpc: any
  const registration = (list: unknown[], item: unknown) => { list.push(item); return { dispose: async () => { const i = list.indexOf(item); if (i >= 0) list.splice(i, 1) } } }
  const rebuild = () => { tools.clear(); for (const transform of transforms) transform({ add: (tool: Tool) => tools.set(tool.name, tool) }) }
  const ctx = {
    options: { dataDirectory: dir, authority: fake.authority, graphUrl: fake.graphUrl, ...options },
    session: { get: async () => ({ parentID: undefined, location: { directory: dir } }) },
    permission: {
      hook: async (_: string, callback: any) => registration(hooks.evaluate, callback),
      assert: async (input: any) => {
        const event = { ...input, effect: "allow" as Effect }
        for (const hook of hooks.evaluate) await hook(event)
        if (event.effect === "deny") throw new Error(event.message ?? "denied")
        if (event.effect === "ask") { prompts.push({ action: input.action, metadata: input.metadata }); if (reply === "reject") throw new Error("rejected by user") }
      },
    },
    tool: {
      transform: async (callback: any) => { transforms.push(callback); rebuild(); return { dispose: async () => { transforms.splice(transforms.indexOf(callback), 1); rebuild() } } },
      reload: async () => rebuild(),
      hook: async (_: string, callback: any) => registration(hooks.after, callback),
    },
    skill: { transform: async (callback: any) => { callback({ add: (skill: any) => skills.push(skill) }); return { dispose: async () => { skills.length = 0 } } } },
    event: { subscribe: () => (async function* () {})() },
    agent: { list: async () => ({ data: [{ id: "build", name: "build", mode: "primary", hidden: false }] }) },
    rpc: { register: async (_: unknown, handlers: any) => { rpc = handlers; return { dispose: async () => { rpc = undefined } } } },
  }
  const cleanup = await createPlugin().setup(ctx as unknown as Plugin.Context) as () => Promise<void>
  let n = 0
  const call = async (name: string, input: unknown, agent = "build") => {
    const tool = tools.get(name)
    if (!tool) throw new Error(`no tool ${name}`)
    const c = { sessionID: "s", agent, messageID: "m", id: `c${++n}`, signal: new AbortController().signal, progress: async () => {} }
    let event: any
    try { event = { tool: name, id: c.id, input, status: "completed", result: await tool.execute(input, c) } }
    catch (error) { event = { tool: name, id: c.id, input, status: "error", error } }
    for (const hook of hooks.after) await hook(event)
    if (event.status === "error") throw event.error
    const text = event.result.content
    try { return JSON.parse(text) } catch { return text }
  }
  const signIn = async (id = "work", input: Record<string, unknown> = {}) => {
    await call("m365_account_upsert", { id, tenant: "acme.example", email: "pat@acme.example", ...input })
    const started = await call("m365_signin", { account: id })
    if (started.status === "failed") return started
    await fake.approve(/(http\S+\/authorize\S+?)(\s|$)/.exec(started.tellUser)![1]!)
    for (let i = 0; i < 100; i++) { const list = await call("m365_list", {}); const account = list.accounts.find((a: { id: string }) => a.id === id); if (account.signedInAs || account.signInFailed) return account; await new Promise(r => setTimeout(r, 30)) }
    throw new Error("sign-in did not settle")
  }
  return { dir, tools, skills, prompts, call, signIn, cleanup, rpc: () => rpc, policy: () => PolicyStore.open(dir), vault: () => Vault.open({ dir, plugin: PLUGIN_ID }) }
}

test("identity, tools, skill, and safe failure", async () => {
  assert.equal(plugin.id, PLUGIN_ID)
  const h = await host()
  assert.equal(h.tools.size, 8)
  assert.equal(h.skills[0]!.id, "m365")
  assert.match(h.skills[0]!.content, /Draft, then send/)
  assert.match(h.tools.get("m365_list")!.description, /No Microsoft 365 accounts are configured yet/)
  await h.cleanup()
  assert.equal(h.tools.size, 0)
  const broken = await host("approve", { graphUrl: "ftp://x" })
  assert.equal(broken.tools.size, 0)
  assert.throws(() => parseOptions({ authority: 1 }), /string/)
})

test("browser sign-in (PKCE, localhost redirect); inventory shows who and unread", async () => {
  const h = await host()
  const work = await h.signIn()
  assert.equal(work.signedInAs, "Pat Lee <pat@acme.example>")
  assert.equal(work.unreadInbox, 3)
  assert.ok(work.grantedPermissions.includes("Mail.Send"))
  await new Promise(r => setTimeout(r, 30))
  assert.match(h.tools.get("m365_list")!.description, /work: Pat Lee <pat@acme\.example> in acme\.example; 3 unread in Inbox/)
  const token = (await (await h.vault()).get("work", "access-token"))!
  assert.ok(!JSON.stringify(await h.call("m365_list", {})).includes(token), "tokens never reach the agent")
})

test("admin approval required: explained with a ready request for IT", async () => {
  fake.needsAdmin.add("strict.example")
  const h = await host()
  await h.call("m365_account_upsert", { id: "strict", tenant: "strict.example" })
  const started = await h.call("m365_signin", { account: "strict" })
  await fake.approve(/(http\S+\/authorize\S+?)(\s|$)/.exec(started.tellUser)![1]!)
  for (let i = 0; i < 100 && !(await h.rpc().state({})).accounts[0].signIn?.error; i++) await new Promise(r => setTimeout(r, 30))
  const state = await h.rpc().state({})
  assert.match(state.accounts[0].signIn.error, /requires an administrator to approve.*Microsoft Graph Command Line Tools/)
  assert.match(state.accounts[0].adminRequest, /Mail\.Send/)
  await assert.rejects(h.call("m365_signin", { account: "strict", device: true }), /requires an administrator/)
})

test("mail: read freely, drafts freely, sending asks; Teams posts and sharing ask; $batch checks every part", async () => {
  const h = await host()
  await h.signIn()
  const mail = await h.call("m365_graph", { path: "/me/messages?$top=5&$select=subject,from" })
  assert.equal(mail.value[0].subject, "Budget review")
  const draft = await h.call("m365_graph", { method: "POST", path: "/me/messages", body: { subject: "Re: budget", body: { contentType: "Text", content: "Looks good" }, toRecipients: [{ emailAddress: { address: "boss@acme.example" } }] } })
  assert.deepEqual(h.prompts, [], "reading and drafting do not ask")
  await h.call("m365_graph", { method: "POST", path: `/me/messages/${draft.id}/send` })
  assert.deepEqual(fake.sent.at(-1), { subject: "Re: budget", to: "boss@acme.example" })
  await h.call("m365_graph", { method: "POST", path: "/me/events", body: { subject: "Focus time" } })
  await h.call("m365_graph", { method: "POST", path: "/me/events", body: { subject: "Sync", attendees: [{ emailAddress: { address: "a@acme.example" } }] } })
  await h.call("m365_graph", { method: "POST", path: "/teams/team-1/channels/c1/messages", body: { body: { content: "hi" } } })
  await h.call("m365_graph", { method: "POST", path: "/drives/d/items/i/createLink", body: { type: "view", scope: "organization" } })
  assert.deepEqual((h.prompts as { action: string }[]).map(p => p.action), ["oc-m365.mail-send", "oc-m365.calendar-send", "oc-m365.chat-send", "oc-m365.files-send"])
  h.prompts.length = 0
  const batch = await h.call("m365_graph", { method: "POST", path: "/$batch", body: { requests: [{ id: "1", method: "GET", url: "/me/joinedTeams" }, { id: "2", method: "GET", url: "/me/calendarView?startDateTime=a&endDateTime=b" }] } })
  assert.equal(batch.responses.length, 2)
  assert.deepEqual(h.prompts, [])
  const rejecting = await host("reject")
  await rejecting.signIn()
  await assert.rejects(rejecting.call("m365_graph", { method: "POST", path: "/$batch", body: { requests: [{ id: "1", method: "GET", url: "/me/messages" }, { id: "2", method: "POST", url: "/me/sendMail", body: { message: { subject: "x", toRecipients: [{ emailAddress: { address: "x@y.z" } }] } } }] } }), /rejected by user/)
  assert.equal(fake.sent.filter(item => item.subject === "x").length, 0, "nothing sent when the batch was rejected")
  await (await h.policy()).setCells([{ entity: "work", row: { agent: "plan" }, capability: "mail-write", effect: "deny" }])
  await assert.rejects(h.call("m365_graph", { method: "POST", path: "/me/messages", body: { subject: "x", toRecipients: [{ emailAddress: { address: "a@b.c" } }] } }, "plan"), /Denied by the oc-m365 permissions table \(mail-write on work\)/)
  await assert.rejects(h.call("m365_graph", { path: "/me/chats" }), /403 Forbidden.*granted:/)
  await assert.rejects(h.call("m365_graph", { path: "https://evil.example/v1.0/me" }), /Only Microsoft Graph URLs/)
  await assert.rejects(h.call("m365_graph", { path: "/me", headers: { Authorization: "x" } }), /not allowed/)
})

test("tokens refresh transparently on expiry", async () => {
  const h = await host()
  await h.signIn()
  const before = await (await h.vault()).get("work", "refresh-token")
  fake.expireTokens()
  assert.equal((await h.call("m365_graph", { path: "/me" })).displayName, "Pat Lee")
  assert.notEqual(await (await h.vault()).get("work", "refresh-token"), before)
})

test("files: download a drive item and an attachment, upload small and large files", async () => {
  const h = await host()
  await h.signIn()
  const got = await h.call("m365_download", { path: "/me/drive/root:/Reports/q3.txt:/content", to: "out/q3.txt" })
  assert.equal(await readFile(join(h.dir, "out", "q3.txt"), "utf8"), "Q3 revenue up 12%\n")
  assert.equal(got.bytes, 18)
  await h.call("m365_download", { path: "/me/messages/msg-1/attachments/att-1/$value", to: join(h.dir, "a.pdf") })
  assert.match(await readFile(join(h.dir, "a.pdf"), "utf8"), /%PDF/)
  await writeFile(join(h.dir, "small.txt"), "hello\n")
  const small = await h.call("m365_upload", { from: "small.txt", to: "/me/drive/root:/Notes/small.txt" })
  assert.equal(small.uploaded.size, 6)
  const big = Buffer.alloc(5 * 1024 * 1024, 7)
  await writeFile(join(h.dir, "big.bin"), big)
  const large = await h.call("m365_upload", { from: "big.bin", to: "/me/drive/root:/Data/big.bin" })
  assert.equal(large.uploaded.size, big.length)
  assert.ok(fake.files.get("Data/big.bin")!.equals(big), "chunked upload reassembled")
  await assert.rejects(h.call("m365_upload", { from: "small.txt", to: "/me/messages" }), /drive path/)
})

test("GUI admin: device code, read-only preset, export and import", async () => {
  const h = await host()
  await h.rpc().createAccount({ id: "work", tenant: "acme.example", auth: "device", preset: "read" })
  const started = await h.rpc().signIn({ id: "work" })
  assert.equal(started.code, "FAKE-CODE")
  fake.approveDevice()
  for (let i = 0; i < 200 && !(await h.rpc().state({})).accounts[0].who; i++) await new Promise(r => setTimeout(r, 50))
  const state = await h.rpc().state({ sessionID: "s" })
  assert.equal(state.accounts[0].who, "Pat Lee")
  assert.ok(!state.accounts[0].granted.includes("Mail.Send"), "read preset")
  const token = (await (await h.vault()).get("work", "access-token"))!
  assert.ok(!JSON.stringify(state).includes(token))
  await h.rpc().setCells({ changes: [{ entity: "work", capability: "mail-send", effect: "deny", agent: "*" }] })
  const exported = await h.rpc().exportBundle({ passphrase: "export passphrase 1" })
  const other = await host()
  await other.rpc().applyImport({ content: exported.content, passphrase: "export passphrase 1", mode: "merge" })
  assert.equal((await other.call("m365_graph", { path: "/me" })).displayName, "Pat Lee", "the imported sign-in works")
  assert.equal((await other.rpc().state({})).cells.work["*"]["mail-send"], "deny")
  await h.rpc().updateAccount({ id: "work", preset: "full" })
  assert.equal((await h.rpc().state({})).accounts[0].signedIn, false, "changing access requires signing in again")
})
