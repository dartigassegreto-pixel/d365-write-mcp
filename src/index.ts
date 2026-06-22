import "dotenv/config";
import express from "express";
import { randomUUID } from "node:crypto";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { dataverseRequest } from "./dynamics-client.js";

function buildServer() {
  const server = new Server(
    { name: "d365-write-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );
  registerHandlers(server);
  return server;
}

// ---- Tool definitions -------------------------------------------------

const TOOLS = [
  {
    name: "create_contact",
    description: "Create a new contact in Dynamics 365.",
    inputSchema: {
      type: "object",
      properties: {
        firstname: { type: "string" },
        lastname: { type: "string", description: "Required." },
        emailaddress1: { type: "string" },
        telephone1: { type: "string" },
        jobtitle: { type: "string" },
      },
      required: ["lastname"],
    },
  },
  {
    name: "update_contact",
    description: "Update fields on an existing contact, by contact ID.",
    inputSchema: {
      type: "object",
      properties: {
        contactid: { type: "string", description: "GUID of the contact to update." },
        firstname: { type: "string" },
        lastname: { type: "string" },
        emailaddress1: { type: "string" },
        telephone1: { type: "string" },
        jobtitle: { type: "string" },
      },
      required: ["contactid"],
    },
  },
  {
    name: "create_account",
    description: "Create a new account (company) in Dynamics 365.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Required." },
        telephone1: { type: "string" },
        websiteurl: { type: "string" },
        description: { type: "string" },
      },
      required: ["name"],
    },
  },
  {
    name: "update_account",
    description: "Update fields on an existing account, by account ID.",
    inputSchema: {
      type: "object",
      properties: {
        accountid: { type: "string", description: "GUID of the account to update." },
        name: { type: "string" },
        telephone1: { type: "string" },
        websiteurl: { type: "string" },
        description: { type: "string" },
      },
      required: ["accountid"],
    },
  },
  {
    name: "create_task",
    description:
      "Create a task in Dynamics 365, optionally linked to a contact, account, or opportunity (regardingobjectid).",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Required." },
        description: { type: "string" },
        scheduledend: { type: "string", description: "Due date, ISO 8601 e.g. 2026-06-22" },
        regarding_id: { type: "string", description: "GUID of the related record (contact/account/opportunity)." },
        regarding_type: {
          type: "string",
          enum: ["contact", "account", "opportunity"],
          description: "Entity type that regarding_id belongs to.",
        },
      },
      required: ["subject"],
    },
  },
  {
    name: "create_opportunity",
    description: "Create a new sales opportunity in Dynamics 365.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Required. Opportunity name/topic." },
        estimatedvalue: { type: "number" },
        description: { type: "string" },
        contact_id: { type: "string", description: "GUID of the primary contact." },
        account_id: { type: "string", description: "GUID of the parent account." },
      },
      required: ["name"],
    },
  },
  {
    name: "update_opportunity",
    description: "Update fields on an existing opportunity, by opportunity ID.",
    inputSchema: {
      type: "object",
      properties: {
        opportunityid: { type: "string", description: "GUID of the opportunity to update." },
        name: { type: "string" },
        estimatedvalue: { type: "number" },
        description: { type: "string" },
      },
      required: ["opportunityid"],
    },
  },
];

function registerHandlers(server: Server) {
server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

// ---- Tool execution -----------------------------------------------------

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  try {
    switch (name) {
      case "create_contact": {
        const result = await dataverseRequest("POST", "contacts", args);
        return ok(`Created contact: ${result.contactid}`);
      }

      case "update_contact": {
        const { contactid, ...fields } = args as any;
        await dataverseRequest("PATCH", `contacts(${contactid})`, fields);
        return ok(`Updated contact ${contactid}.`);
      }

      case "create_account": {
        const result = await dataverseRequest("POST", "accounts", args);
        return ok(`Created account: ${result.accountid}`);
      }

      case "update_account": {
        const { accountid, ...fields } = args as any;
        await dataverseRequest("PATCH", `accounts(${accountid})`, fields);
        return ok(`Updated account ${accountid}.`);
      }

      case "create_task": {
        const { regarding_id, regarding_type, ...fields } = args as any;
        const body: Record<string, unknown> = { ...fields };
        if (regarding_id && regarding_type) {
          const bind: Record<string, string> = {
            contact: "contacts",
            account: "accounts",
            opportunity: "opportunities",
          };
          body[`regardingobjectid_${regarding_type}@odata.bind`] = `/${bind[regarding_type]}(${regarding_id})`;
        }
        const result = await dataverseRequest("POST", "tasks", body);
        return ok(`Created task: ${result.activityid}`);
      }

      case "create_opportunity": {
        const { contact_id, account_id, ...fields } = args as any;
        const body: Record<string, unknown> = { ...fields };
        if (contact_id) body["parentcontactid@odata.bind"] = `/contacts(${contact_id})`;
        if (account_id) body["parentaccountid@odata.bind"] = `/accounts(${account_id})`;
        const result = await dataverseRequest("POST", "opportunities", body);
        return ok(`Created opportunity: ${result.opportunityid}`);
      }

      case "update_opportunity": {
        const { opportunityid, ...fields } = args as any;
        await dataverseRequest("PATCH", `opportunities(${opportunityid})`, fields);
        return ok(`Updated opportunity ${opportunityid}.`);
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (err: any) {
    return {
      content: [{ type: "text", text: `Error: ${err.message}` }],
      isError: true,
    };
  }
});
} // end registerHandlers

function ok(text: string) {
  return { content: [{ type: "text", text }] };
}

// ---- Start (HTTP, for remote hosting e.g. Render) ------------------------

const app = express();
app.use(express.json());

app.use((_req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, mcp-session-id");
  if (_req.method === "OPTIONS") { res.sendStatus(204); return; }
  next();
});

app.get("/", (_req, res) => {
  res.send("d365-write-mcp is running.");
});

// ---- Anthropic proxy (avoids browser CORS restrictions in artifacts) -----

app.post("/proxy", async (req, res) => {
  const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
  if (!ANTHROPIC_API_KEY) {
    res.status(500).json({ error: { message: "ANTHROPIC_API_KEY not set on server." } });
    return;
  }
  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": "mcp-client-2025-04-04",
      },
      body: JSON.stringify(req.body),
    });
    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err: any) {
    res.status(500).json({ error: { message: err.message } });
  }
});

app.post("/mcp", async (req, res) => {
  const server = buildServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: () => randomUUID(),
  });
  res.on("close", () => {
    transport.close();
    server.close();
  });
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
});

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
app.listen(PORT, () => {
  console.log(`d365-write-mcp listening on port ${PORT}`);
});
