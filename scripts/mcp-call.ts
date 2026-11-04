// Tiny MCP client for manual testing: bun scripts/mcp-call.ts <tool> '<json args>' [<tool> '<json>' ...]
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
const client = new Client({ name: "manual-test", version: "0" });
await client.connect(new StdioClientTransport({ command: "bun", args: ["run", new URL("../src/server.ts", import.meta.url).pathname] }));
const argv = process.argv.slice(2);
if (!argv.length) {
  const { tools } = await client.listTools();
  for (const t of tools) console.log(`${t.name}: ${t.description?.slice(0, 90)}`);
}
for (let i = 0; i < argv.length; i += 2) {
  const t0 = Date.now();
  const res: any = await client.callTool({ name: argv[i]!, arguments: JSON.parse(argv[i + 1] ?? "{}") }, undefined, { timeout: 600_000 });
  console.log(`=== ${argv[i]} (${Date.now() - t0} ms)${res.isError ? " ERROR" : ""}`);
  console.log(res.content?.[0]?.text);
}
await client.close();
