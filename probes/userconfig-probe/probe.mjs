#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  InitializeRequestSchema,
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { appendFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const LOG = join(homedir(), 'mcp-userconfig-probe.log');
writeFileSync(LOG, `=== probe boot ${new Date().toISOString()} ===\n`);

function log(label, obj) {
  appendFileSync(
    LOG,
    `[${new Date().toISOString()}] ${label}\n${JSON.stringify(obj, null, 2)}\n\n`
  );
}

const server = new Server(
  { name: 'userconfig-probe', version: '0.0.1' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(InitializeRequestSchema, async (request) => {
  log('INITIALIZE_REQUEST_FULL', request);
  log('INITIALIZE_PARAMS', request.params);
  log('INITIALIZE_PARAMS_META', request.params?._meta);
  log('INITIALIZE_PARAMS_CAPS', request.params?.capabilities);
  log('INITIALIZE_PARAMS_CAPS_EXP', request.params?.capabilities?.experimental);
  log('INITIALIZE_CLIENTINFO', request.params?.clientInfo);
  return {
    protocolVersion: '2024-11-05',
    capabilities: { tools: {} },
    serverInfo: { name: 'userconfig-probe', version: '0.0.1' },
  };
});

const originalNotification = server.notification?.bind(server);
server.notification = async (notification) => {
  log(`NOTIFICATION:${notification?.method ?? 'unknown'}`, notification);
  if (originalNotification) return originalNotification(notification);
};

server.fallbackNotificationHandler = async (notification) => {
  log(`FALLBACK_NOTIFICATION:${notification?.method ?? 'unknown'}`, notification);
};

server.setRequestHandler(ListToolsRequestSchema, async (request) => {
  log('LIST_TOOLS_REQUEST', request);
  return {
    tools: [
      {
        name: 'probe_dump',
        description:
          'Returns the full path to the probe log file. Call this to confirm the probe is reachable.',
        inputSchema: { type: 'object', properties: {}, required: [] },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  log(`CALL_TOOL:${request.params?.name}`, request);
  if (request.params?.name === 'probe_dump') {
    return {
      content: [
        {
          type: 'text',
          text: `Probe log file: ${LOG}\nThe file contains every initialize param and notification this server has received since boot.`,
        },
      ],
    };
  }
  return {
    content: [{ type: 'text', text: 'unknown tool' }],
    isError: true,
  };
});

const transport = new StdioServerTransport();
await server.connect(transport);
appendFileSync(LOG, `[${new Date().toISOString()}] connected to stdio\n\n`);
