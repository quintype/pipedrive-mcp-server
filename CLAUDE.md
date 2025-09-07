# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Model Context Protocol (MCP) server that integrates with the Pipedrive API v2. The server exposes Pipedrive CRM data (deals, persons, organizations, pipelines) to LLM applications like Claude Desktop through the MCP framework.

## Key Commands

- `npm run build` - Compile TypeScript to JavaScript in the build/ directory
- `npm start` - Run the compiled MCP server (requires build first)
- `npm run dev` - Run the server in development mode with tsx
- Setup requires a `.env` file with `PIPEDRIVE_API_TOKEN=your_token`

## Architecture

### Core Structure
- `src/index.ts` - Main MCP server implementation with all tools and prompts
- `src/types.d.ts` - TypeScript declarations for the Pipedrive SDK
- Built using `@modelcontextprotocol/sdk` for MCP protocol implementation
- Uses the `pipedrive` npm package for API integration

### API Integration Pattern
- Single ApiClient instance with API key authentication configured
- Separate API classes for each Pipedrive entity (DealsApi, PersonsApi, etc.)
- All tools return JSON-formatted responses wrapped in MCP content structure
- Consistent error handling pattern using `getErrorMessage()` utility

### MCP Tools Structure
The server exposes 13 tools with comprehensive parameter support:
- **Deals**: `get-deals`, `get-deal`, `search-deals` with filtering (owner_id, person_id, org_id, pipeline_id, stage_id, status), pagination (limit, cursor), sorting (sort_by, sort_direction), field selection (include_fields, custom_fields), and time-based filters (updated_since, updated_until)
- **Persons**: `get-persons`, `get-person`, `search-persons` with similar filtering and pagination options plus organization-based filtering
- **Organizations**: `get-organizations`, `get-organization`, `search-organizations` with filtering, pagination, and field selection capabilities
- **Pipelines**: `get-pipelines`, `get-pipeline` with sorting and pagination support
- **Special tools**: `get-stages`, `search-leads`, `search-all` for cross-entity operations

### MCP Prompts Structure
Pre-defined prompts for common operations like listing entities, analysis by groupings, and finding high-value deals.

## Development Notes

- Server runs on stdio transport for MCP communication
- Environment validation ensures PIPEDRIVE_API_TOKEN is present on startup
- TypeScript with ES2022/Node16 module resolution
- All async operations use proper try-catch with standardized error responses
- Custom types defined for Pipedrive SDK which lacks complete TypeScript support

## Claude Desktop Integration

The server is designed to be configured in `claude_desktop_config.json` with the built JavaScript file path and environment variables for the API token.