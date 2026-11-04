// Delete leftover e2e test events (identified by their description) on a given day.
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
const day = process.argv[2] ?? "2026-11-21";
const client = new Client({ name: "cleanup", version: "0" });
await client.connect(new StdioClientTransport({ command: "bun", args: ["run", new URL("../src/server.ts", import.meta.url).pathname] }));
const call = async (name: string, args: any) => {
  const r: any = await client.callTool({ name, arguments: args }, undefined, { timeout: 600_000 });
  if (r.isError) throw new Error(r.content[0].text);
  return JSON.parse(r.content[0].text);
};
for (let i = 0; i < 10; i++) {
  const ev = (await call("outlook_list_events", { start: day })).events.find((e: any) => e.bodyPreview?.includes("Created by the e2e test") || e.subject.startsWith("MCP e2e event") || e.subject === "MCP dbg event");
  if (!ev) break;
  console.log("deleting", ev.subject, ev.start, await call("outlook_delete_event", { ref: ev.ref }));
}
console.log("remaining:", (await call("outlook_list_events", { start: day })).events.map((e: any) => e.subject));
await client.close();
