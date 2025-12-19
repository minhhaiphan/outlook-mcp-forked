#!/usr/bin/env node
/**
 * Outlook MCP Server - Main entry point
 * 
 * A Model Context Protocol server that provides access to
 * Microsoft Outlook through the Microsoft Graph API.
 */

const http = require("http");
const url = require("url");

const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");

// const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
// const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const config = require('./config');

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools
  // Future modules: contactsTools, etc.
];

// Create server with tools capabilities
// const server = new Server(
//   { name: config.SERVER_NAME, version: config.SERVER_VERSION },
//   { 
//     capabilities: { 
//       tools: TOOLS.reduce((acc, tool) => {
//         acc[tool.name] = {};
//         return acc;
//       }, {})
//     } 
//   }
// );

// ---- MCP server (NEW: use McpServer and register tools) ----
const server = new McpServer({ name: config.SERVER_NAME, version: config.SERVER_VERSION });


// Handle all requests using fallback handler
server.fallbackRequestHandler = async (request) => {
  console.error("RAW REQUEST:", JSON.stringify(request, null, 2));
  try {
    const { method, params, id } = request;
    console.error(`=== FALLBACK HANDLER CALLED ===`);
    console.error(`REQUEST: ${method} [${id}]`);
    console.error(`Full request:`, JSON.stringify(request, null, 2));
    
    // Initialize handler
    if (method === "initialize") {
      console.error(`INITIALIZE REQUEST: ID [${id}]`);
      return {
        protocolVersion: "2024-11-05",
        capabilities: { 
          tools: TOOLS.reduce((acc, tool) => {
            acc[tool.name] = {};
            return acc;
          }, {})
        },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
      };
    }
    
    // Tools list handler
    if (method === "tools/list") {
      console.error(`TOOLS LIST REQUEST: ID [${id}]`);
      console.error(`TOOLS COUNT: ${TOOLS.length}`);
      console.error(`TOOLS NAMES: ${TOOLS.map(t => t.name).join(', ')}`);
      
      return {
        tools: TOOLS.map(tool => ({
          name: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema
        }))
      };
    }
    
    // Required empty responses for other capabilities
    if (method === "resources/list") return { resources: [] };
    if (method === "prompts/list") return { prompts: [] };
    console.log(method, "method")
    // Tool call handler
    if (method === "tools/call") {
      try {
        // const { name, arguments: args = {} } = params || {};
        
        // console.error(`TOOL CALL: ${name}`);
        // console.error(`TOOL ARGS:`, JSON.stringify(args, null, 2));
        
        // // Find the tool handler
        // const tool = TOOLS.find(t => t.name === name);
        
        // if (tool && tool.handler) {
        //   return await tool.handler(args);
        // }
        
        // // Tool not found
        // return {
        //   error: {
        //     code: -32601,
        //     message: `Tool not found: ${name}`
        //   }
        // };
        const p = params || {};
        const name = p.name;
        console.log(params, "params")
        // tolerate multiple client shapes
        const args =
          params?.arguments ??
          params?.input ??
          params?.params?.arguments ??
          params?.params?.input ??
          {};
      
        console.error("TOOLS/CALL name:", name);
        console.error("TOOLS/CALL args:", JSON.stringify(args, null, 2));
      
        const tool = TOOLS.find(t => t.name === name);
        if (!tool?.handler) {
          return { error: { code: -32601, message: `Tool not found: ${name}` } };
        }
        return await tool.handler(args);
      } catch (error) {
        console.error(`Error in tools/call:`, error);
        return {
          error: {
            code: -32603,
            message: `Error processing tool call: ${error.message}`
          }
        };
      }
    }
    
    // For any other method, return method not found
    return {
      error: {
        code: -32601,
        message: `Method not found: ${method}`
      }
    };
  } catch (error) {
    console.error(`Error in fallbackRequestHandler:`, error);
    return {
      error: {
        code: -32603,
        message: `Error processing request: ${error.message}`
      }
    };
  }
};


// Disable fallback handler to let server.tool() registrations handle everything
console.error('Using server.tool() registrations only, no fallback handler');

// // Make the script executable
// process.on('SIGTERM', () => {
//   console.error('SIGTERM received but staying alive');
// });

// // Start the server
// const transport = new StdioServerTransport();
// server.connect(transport)
//   .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
//   .catch(error => {
//     console.error(`Connection error: ${error.message}`);
//     process.exit(1);
//   });


// Register tools using server.tool() - this is the correct approach for the MCP SDK
for (const tool of TOOLS) {
  if (!tool?.name || !tool?.handler) continue;

  console.error(`Registering tool: ${tool.name}`);
  
  // The MCP SDK tool handler might use different signatures
  server.tool(tool.name, tool.description || "", tool.inputSchema || {}, async (...allArgs) => {
    console.error(`=== Tool Call: ${tool.name} ===`);
    console.error('All arguments received:');
    allArgs.forEach((arg, index) => {
      console.error(`Arg ${index}:`, JSON.stringify(arg, null, 2));
    });
    
    // Try to find the actual parameters in different locations
    let params = {};
    
    // Check each argument for tool parameters
    for (let i = 0; i < allArgs.length; i++) {
      const arg = allArgs[i];
      if (arg && typeof arg === 'object') {
        // Look for properties that look like tool parameters
        if (arg.subject || arg.body || arg.to) {
          console.error(`Found tool parameters in arg ${i}:`, arg);
          params = arg;
          break;
        }
        // Check if parameters are nested
        if (arg.arguments && typeof arg.arguments === 'object') {
          console.error(`Found nested parameters in arg ${i}.arguments:`, arg.arguments);
          params = arg.arguments;
          break;
        }
        if (arg.params && typeof arg.params === 'object') {
          console.error(`Found nested parameters in arg ${i}.params:`, arg.params);
          params = arg.params;
          break;
        }
      }
    }
    
    // TEMPORARY WORKAROUND: If no parameters found and this is create-draft, use test data
    // if (Object.keys(params).length === 0 && tool.name === 'create-draft') {
    //   console.error('No parameters found - using test data for create-draft');
    //   params = {
    //     subject: 'hello',
    //     body: 'this is body',
    //     to: 'hai@vggate.com'
    //   };
    // }
    
    console.error('Final params to pass to handler:', JSON.stringify(params, null, 2));
    
    return await tool.handler(params);
  });
}

console.error(`Registered ${TOOLS.length} tools: ${TOOLS.map(t => t.name).join(', ')}`);

// ---- HTTP transport (NEW) ----
const PORT = Number(process.env.PORT || process.env.MCP_HTTP_PORT || 3001);
const MCP_PATH = process.env.MCP_PATH || "/mcp";

// One transport instance is enough for Streamable HTTP
const transport = new StreamableHTTPServerTransport({
  // Optional: you can customize session ids, but default is usually fine
  // sessionIdFactory: () => crypto.randomUUID(),
});

// Connect MCP server to transport (required)
server.connect(transport).then(() => {
  console.error(`${config.SERVER_NAME} connected (HTTP Streamable)`);
}).catch((err) => {
  console.error("MCP connect error:", err);
  process.exit(1);
});

// Create HTTP server and forward requests to the MCP transport
const httpServer = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  // Basic health endpoint
  if (req.method === "GET" && (pathname === "/" || pathname === "/health")) {
    res.writeHead(200, { "content-type": "application/json" });
    res.end(JSON.stringify({ ok: true, name: config.SERVER_NAME, version: config.SERVER_VERSION }));
    return;
  }

  // MCP endpoint for n8n: POST/GET (depending on transport internals) to /mcp
  if (pathname === MCP_PATH) {
    try {
      await transport.handleRequest(req, res);
    } catch (e) {
      console.error("handleRequest error:", e);
      res.writeHead(500, { "content-type": "text/plain" });
      res.end("MCP transport error");
    }
    return;
  }

  res.writeHead(404, { "content-type": "text/plain" });
  res.end("Not Found");
});

httpServer.listen(PORT, () => {
  console.error(`HTTP server listening on :${PORT}`);
  console.error(`MCP endpoint: http://0.0.0.0:${PORT}${MCP_PATH}`);
});

// Keep process alive in Dokploy
process.on("SIGTERM", () => console.error("SIGTERM received"));
process.on("SIGINT", () => console.error("SIGINT received"));
