import type { Plugin } from "@opencode/plugin"
import { readFile } from "node:fs/promises"
import { fileURLToPath } from "node:url"
import { installRedaction, PermissionBridge, pluginDataDir, SECURE_KIT_VERSION, sessionRoots, type PermissionHost } from "oc-secure-kit"
import { createAdmin } from "./admin.js"
import { defaultEffect } from "./model.js"
import { M365Rpc } from "./rpc.js"
import { M365Service, PLUGIN_ID } from "./service.js"
import { createTools, TOOL_NAMES } from "./tools.js"

export { PLUGIN_ID, SECURE_KIT_VERSION }

export interface M365Options {
  /** Absolute path; default `$XDG_DATA_HOME/oc-m365`. */
  readonly dataDirectory?: string
  /** Sign-in authority; default https://login.microsoftonline.com. For tests or national clouds. */
  readonly authority?: string
  /** Graph base; default https://graph.microsoft.com. For tests or national clouds. */
  readonly graphUrl?: string
}

export function parseOptions(raw: unknown): M365Options {
  const options = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>
  for (const key of ["dataDirectory", "authority", "graphUrl"]) if (options[key] !== undefined && typeof options[key] !== "string") throw new Error(`${PLUGIN_ID}: ${key} must be a string`)
  for (const key of ["authority", "graphUrl"]) if (options[key] !== undefined && !/^https?:\/\//.test(options[key] as string)) throw new Error(`${PLUGIN_ID}: ${key} must be an http(s) URL`)
  return Object.fromEntries(["dataDirectory", "authority", "graphUrl"].filter(key => options[key] !== undefined).map(key => [key, options[key]])) as M365Options
}

const SKILL = fileURLToPath(new URL("../skill/SKILL.md", import.meta.url))
const warn = (message: string, error?: unknown) =>
  console.warn(`[${PLUGIN_ID}] ${message}${error ? `: ${error instanceof Error ? error.message : String(error)}` : ""}`)

export function createPlugin(): Plugin.Plugin {
  return ({
    id: PLUGIN_ID,
    async setup(ctx) {
      const registrations: { dispose(): Promise<void> | void }[] = []
      const cleanup = async () => { await Promise.allSettled(registrations.splice(0).reverse().map(async item => item.dispose())) }
      try {
        const options = parseOptions(ctx.options)
        const service = await M365Service.open(pluginDataDir(PLUGIN_ID, options.dataDirectory), { ...(options.authority && { authority: options.authority }), ...(options.graphUrl && { graphUrl: options.graphUrl }) })
        registrations.push({ dispose: () => service.dispose() })
        const root = sessionRoots(async sessionID => (await ctx.session.get({ sessionID: sessionID as never })).parentID ?? undefined)
        const bridge = new PermissionBridge({ plugin: PLUGIN_ID, policy: service.policy, fallback: defaultEffect, root })
        registrations.push(await bridge.install(ctx.permission as unknown as PermissionHost))
        registrations.push(bridge.followReplies(ctx.event as never, error => warn("could not save Allow always", error)))
        service.attach(bridge)

        const tools = createTools(service, async sessionID => { const session = await ctx.session.get({ sessionID: sessionID as never }); return [session.location.directory, session.subpath].filter(Boolean).join("/") || undefined })
        registrations.push(await ctx.tool.transform(editor => {
          for (const tool of tools) {
            editor.add({
              name: tool.name, description: tool.description(), input: tool.input as never, options: { codemode: false },
              execute: async (input, context) => {
                try { return await tool.execute(input, context) }
                finally { await service.syncRedactor().catch(error => warn("redactor refresh failed", error)) }
              },
            })
          }
        }))
        registrations.push(await installRedaction(ctx.tool as never, service.redactor, new Set<string>(TOOL_NAMES), error => warn("redaction skipped", error)))

        try {
          const content = await readFile(SKILL, "utf8")
          const description = /^description:\s*(.+)$/m.exec(content)?.[1] ?? "Microsoft 365 know-how"
          registrations.push(await ctx.skill.transform(editor => {
            editor.add({ id: "m365" as never, name: "m365" as never, description, path: SKILL as never, content })
          }))
        } catch (error) { warn("skill not registered", error) }

        let reloading: Promise<void> | undefined
        registrations.push({ dispose: service.store.onChange(() => {
          reloading ??= Promise.resolve().then(() => ctx.tool.reload()).catch(error => warn("tool reload failed", error)).finally(() => { reloading = undefined })
        }) })

        const admin = createAdmin({
          service, root,
          agents: async () => ((await ctx.agent.list()).data ?? []).filter(agent => !agent.hidden).map(agent => ({ id: agent.id as string, name: agent.name, mode: agent.mode })),
        })
        const handlers = Object.fromEntries(Object.keys(M365Rpc.methods).map(name => [name, async (input: unknown) => (admin as Record<string, (input: never) => Promise<unknown>>)[name]!(input as never)]))
        registrations.push(await ctx.rpc.register(M365Rpc, handlers as never))
      } catch (error) {
        warn("disabled", error)
        await cleanup()
      }
      return cleanup
    },
  }) satisfies Plugin.Plugin
}

export default createPlugin()
