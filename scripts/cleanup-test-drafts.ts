// Remove drafts created by the e2e tests (moved to Deleted Items).
// Usage: E2E_DRAFT_FOLDERS="you@example.com/Drafts,..." bun scripts/cleanup-test-drafts.ts [--apply]
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
const apply = process.argv.includes("--apply");
const folders = (process.env.E2E_DRAFT_FOLDERS ?? "").split(",").filter(Boolean);
if (!folders.length) {
  console.error("Set E2E_DRAFT_FOLDERS to a comma-separated list of Drafts folders.");
  process.exit(2);
}
const client = new Client({ name: "cleanup", version: "0" });
await client.connect(new StdioClientTransport({ command: "bun", args: ["run", new URL("../src/server.ts", import.meta.url).pathname] }));
const call = async (name: string, args: any) => {
  const r: any = await client.callTool({ name, arguments: args }, undefined, { timeout: 600_000 });
  if (r.isError) throw new Error(r.content[0].text);
  return JSON.parse(r.content[0].text);
};
// Only drafts created today whose subject or body carries the e2e markers.
const isTest = (m: any) => /^\d{1,2}:\d{2}$/.test(m.date) && (/^MCP (e2e|draft)/.test(m.subject) || /^MCP e2e/.test(m.preview));
for (const folder of folders) {
  for (let i = 0; i < 20; i++) {
    const list = (await call("outlook_list_messages", { folder, limit: 30 })).messages;
    if (i === 0) for (const m of list.filter((x: any) => /^\d{1,2}:\d{2}$/.test(x.date))) console.log(`${isTest(m) ? "TEST " : "keep "} ${folder} | ${m.date} | ${m.subject}`);
    const t = list.find(isTest);
    if (!t || !apply) break;
    await call("outlook_delete_message", { ref: t.ref });
    console.log("deleted:", t.subject, t.date);
  }
}
await client.close();
