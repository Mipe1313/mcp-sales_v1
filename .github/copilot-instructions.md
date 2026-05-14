---
applyTo: "**"
---

# Dataverse MCP Server

This project is an MCP (Model Context Protocol) server written in **TypeScript** (Node.js ESM).

## Key references

- MCP TypeScript SDK v1.x: https://github.com/modelcontextprotocol/typescript-sdk/tree/v1.x
- MCP specification: https://modelcontextprotocol.io/specification/latest
- Dataverse Web API: https://learn.microsoft.com/power-apps/developer/data-platform/webapi/overview

## Project structure

| Path | Purpose |
|------|---------|
| `src/index.ts` | MCP server entry point; supports HTTP (Copilot Studio) and stdio (VS Code) |
| `src/auth.ts` | OAuth2 client_credentials token acquisition & cache |
| `src/dataverse.ts` | Dataverse Web API v9.2 client |
| `src/formUtils.ts` | Pure form-XML manipulation (add/remove/move fields, tabs, sections) |
| `src/tools/accountForms.ts` | Tool handler functions wired into the MCP server |
| `src/tools/solution.ts` | MCPautoSetup solution & Fellowmind publisher management |
| `.env.example` | Template for required environment variables |

## Transport modes

| Mode | How to start | Use case |
|------|-------------|----------|
| HTTP (default) | `npm start` | Copilot Studio remote connector |
| stdio | `npm run start:stdio` or `--stdio` flag | VS Code MCP debug |

## Environment variables

- `TENANT_ID` – Azure AD tenant ID
- `CLIENT_ID` – App registration client ID
- `CLIENT_SECRET` – App registration client secret
- `ENVIRONMENT_URL` – e.g. `https://orgname.crm4.dynamics.com`
- `PORT` – HTTP listen port (default 3000)

## Available MCP tools

1. `configure_environment` – set dynamic tenant / environment for the session
2. `list_account_forms` – list Main account forms
3. `get_form_xml` – fetch raw formXml
4. `get_attribute_metadata` – get field type / display name
5. `add_tab_to_form` – add a tab
6. `add_section_to_form` – add a section to a tab
7. `add_field_to_form` – add a field control to a section
8. `remove_field_from_form` – remove a field control
9. `set_field_properties` – set disabled / visible
10. `move_field_on_form` – relocate a field to another section
11. `update_form_xml_raw` – overwrite entire formXml (advanced)
12. `ensure_solution` – create MCPautoSetup / Fellowmind publisher if missing
13. `add_form_to_solution` – track a form in MCPautoSetup
14. `publish_customisations` – publish account entity changes
