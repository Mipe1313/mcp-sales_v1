import "dotenv/config";
import { randomUUID } from "crypto";
import { createRequire } from "module";
import express, { type Request, type Response } from "express";

const _require = createRequire(import.meta.url);
const { version: SERVER_VERSION } = _require("../package.json") as { version: string };

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";

import { DataverseClient } from "./dataverse.js";
import type { DynamicsConfig, FormCache, SessionConfig } from "./types.js";
import * as Forms from "./tools/accountForms.js";
import * as Solution from "./tools/solution.js";
import * as Schema from "./tools/schema.js";

// ----------------------------------------------------------------
// Tool definitions (for ListTools)
// ----------------------------------------------------------------

const TOOL_DEFINITIONS = [
  {
    name: "configure_environment",
    description:
      "Optionally override the Dynamics 365 / Dataverse tenant and environment for this session. " +
      "The server already has a default tenant and environment pre-configured — only call this tool " +
      "if the user explicitly wants to use a DIFFERENT tenant or environment URL. " +
      "If the user says 'default', 'use default', or does not mention a specific tenant, do NOT call this tool. " +
      "The server's app registration credentials are managed server-side and must not be passed here.",
    inputSchema: {
      type: "object",
      required: [],
      properties: {
        tenantId: { type: "string", description: "Customer's Azure AD tenant ID (GUID)" },
        environmentUrl: {
          type: "string",
          description: "Customer's Dataverse environment URL, e.g. https://orgname.crm4.dynamics.com",
        },
        targetEntity: {
          type: "string",
          description: "Logical name of the entity to operate on, e.g. 'account' (default: account)",
        },
        targetAppUniqueName: {
          type: "string",
          description: "Unique name of the model-driven app to scope operations to, e.g. 'msdyn_SalesHub'",
        },
      },
    },
  },
  {
    name: "test_connection",
    description:
      "Validate the current session's Dataverse connection. " +
      "Tests token acquisition, environment reachability, and API access (WhoAmI). " +
      "Call this after configure_environment to confirm the customer's setup is correct before making any changes.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "list_account_forms",
    description: "List all Main account forms in the Dataverse environment.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "get_form_xml",
    description: "Get the raw XML definition of a specific account form.",
    inputSchema: {
      type: "object",
      required: ["formId"],
      properties: {
        formId: { type: "string", description: "GUID of the system form" },
      },
    },
  },
  {
    name: "get_attribute_metadata",
    description: "Get attribute type and display name for a field on an entity.",
    inputSchema: {
      type: "object",
      required: ["entityName", "fieldName"],
      properties: {
        entityName: { type: "string", description: "Logical name of the entity, e.g. account" },
        fieldName: {
          type: "string",
          description: "Logical name of the field, e.g. telephone1",
        },
      },
    },
  },
  {
    name: "add_tab_to_form",
    description: "Add a new tab to an account form. Set autoCommit=true to push immediately, or false (default) to stage and call commit_form_changes later.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "tabLabel"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string", description: "Internal tab name (no spaces)" },
        tabLabel: { type: "string", description: "Display label for the tab" },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "add_section_to_form",
    description: "Add a new section to an existing tab on an account form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "sectionName", "sectionLabel"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string" },
        sectionName: { type: "string", description: "Internal section name (no spaces)" },
        sectionLabel: { type: "string" },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "list_tabs",
    description: "List all tabs on a form with their names, labels, IDs, and section counts.",
    inputSchema: {
      type: "object",
      required: ["formId"],
      properties: {
        formId: { type: "string", description: "GUID of the system form" },
      },
    },
  },
  {
    name: "rename_tab",
    description: "Rename the display label of an existing tab on a form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "newLabel"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string", description: "Internal name of the tab to rename" },
        newLabel: { type: "string", description: "New display label" },
        autoCommit: { type: "boolean" },
      },
    },
  },
  {
    name: "remove_tab_from_form",
    description: "Remove a tab (and all its sections and fields) from a form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string", description: "Internal name of the tab to remove" },
        autoCommit: { type: "boolean" },
      },
    },
  },
  {
    name: "list_sections",
    description: "List all sections in a tab on a form with their names, labels, and IDs.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string", description: "Internal name of the tab" },
      },
    },
  },
  {
    name: "rename_section",
    description: "Rename the display label of an existing section on a form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "sectionName", "newLabel"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string" },
        sectionName: { type: "string", description: "Internal name of the section to rename" },
        newLabel: { type: "string", description: "New display label" },
        autoCommit: { type: "boolean" },
      },
    },
  },
  {
    name: "remove_section_from_form",
    description: "Remove a section (and all its fields) from a tab on a form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "sectionName"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string" },
        sectionName: { type: "string", description: "Internal name of the section to remove" },
        autoCommit: { type: "boolean" },
      },
    },
  },
  {
    name: "move_section_on_form",
    description: "Move a section from its current tab to a different tab on the same form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "sectionName", "targetTabName"],
      properties: {
        formId: { type: "string" },
        sectionName: { type: "string", description: "Internal name of the section to move" },
        targetTabName: { type: "string", description: "Internal name of the destination tab" },
        autoCommit: { type: "boolean" },
      },
    },
  },
  {
    name: "add_field_to_form",
    description:
      "Add a field control to a specific tab and section on an account form. " +
      "The control classid is resolved automatically from field metadata if attributeType is omitted. " +
      "Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "tabName", "sectionName", "fieldName", "fieldLabel"],
      properties: {
        formId: { type: "string" },
        tabName: { type: "string" },
        sectionName: { type: "string" },
        fieldName: { type: "string", description: "Logical name of the field, e.g. telephone1" },
        fieldLabel: { type: "string", description: "Display label shown on the form" },
        attributeType: {
          type: "string",
          description:
            "Optional Dataverse attribute type (String, Lookup, Picklist, DateTime, …). " +
            "If omitted the type is fetched from entity metadata.",
        },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "remove_field_from_form",
    description: "Remove a field control from an account form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "fieldName"],
      properties: {
        formId: { type: "string" },
        fieldName: { type: "string" },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "set_field_properties",
    description: "Set visibility and/or disabled state of a field on an account form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "fieldName"],
      properties: {
        formId: { type: "string" },
        fieldName: { type: "string" },
        disabled: { type: "boolean", description: "true = read-only, false = editable" },
        visible: { type: "boolean", description: "true = visible, false = hidden" },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "move_field_on_form",
    description: "Move a field control to a different tab/section on the same form. Set autoCommit=true to push immediately, or false (default) to stage.",
    inputSchema: {
      type: "object",
      required: ["formId", "fieldName", "targetTabName", "targetSectionName"],
      properties: {
        formId: { type: "string" },
        fieldName: { type: "string" },
        targetTabName: { type: "string" },
        targetSectionName: { type: "string" },
        autoCommit: { type: "boolean", description: "Push to Dataverse immediately (true) or stage for commit_form_changes (false, default)" },
      },
    },
  },
  {
    name: "update_form_xml_raw",
    description: "Replace the entire formXml of an account form (advanced – use with care).",
    inputSchema: {
      type: "object",
      required: ["formId", "formXml"],
      properties: {
        formId: { type: "string" },
        formXml: { type: "string", description: "Complete Dataverse FormXml string" },
      },
    },
  },
  {
    name: "ensure_solution",
    description:
      "Ensure the MCPautoSetup solution (publisher: Fellowmind) exists; create it if missing.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "add_form_to_solution",
    description: "Add an account form to the MCPautoSetup solution.",
    inputSchema: {
      type: "object",
      required: ["formId"],
      properties: { formId: { type: "string" } },
    },
  },
  {
    name: "publish_customisations",
    description: "Publish all pending account-entity customisations to make them live.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "add_entity_to_solution",
    description: "Add a Dataverse table (entity) to the MCPautoSetup solution by its logical name.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
      },
    },
  },
  {
    name: "add_field_to_solution",
    description: "Add a specific field (attribute) on a table to the MCPautoSetup solution.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name of the field, e.g. telephone1" },
      },
    },
  },
  {
    name: "list_app_modules",
    description: "List all model-driven apps in the Dataverse environment.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "add_form_to_app",
    description:
      "Add a form to a model-driven app's form list so users can see it. " +
      "Uses TARGET_APP_UNIQUE_NAME from config if appUniqueName is omitted.",
    inputSchema: {
      type: "object",
      required: ["formId"],
      properties: {
        formId: { type: "string", description: "GUID of the form to add" },
        appUniqueName: {
          type: "string",
          description: "Unique name of the app, e.g. msdyn_SalesHub. Defaults to TARGET_APP_UNIQUE_NAME.",
        },
      },
    },
  },
  {
    name: "add_app_to_solution",
    description:
      "Add a model-driven app to the MCPautoSetup solution for ALM tracking. " +
      "Uses TARGET_APP_UNIQUE_NAME from config if appUniqueName is omitted.",
    inputSchema: {
      type: "object",
      properties: {
        appUniqueName: {
          type: "string",
          description: "Unique name of the app. Defaults to TARGET_APP_UNIQUE_NAME.",
        },
      },
    },
  },
  {
    name: "commit_form_changes",
    description:
      "Push all staged (autoCommit=false) changes for a form to Dataverse and add it to the solution. " +
      "Call this after staging multiple changes to the same form to apply them in a single operation.",
    inputSchema: {
      type: "object",
      properties: {
        formId: { type: "string", description: "Form ID (original or clone) to commit staged changes for" },
      },
      required: ["formId"],
    },
  },
  {
    name: "discard_form_changes",
    description: "Discard all staged (not yet committed) changes for a form without writing to Dataverse.",
    inputSchema: {
      type: "object",
      properties: {
        formId: { type: "string", description: "Form ID whose staged changes should be discarded" },
      },
      required: ["formId"],
    },
  },
  {
    name: "list_staged_changes",
    description: "List all forms that currently have staged (uncommitted) changes in this session.",
    inputSchema: { type: "object", properties: {} },
  },
  // ---- Schema / metadata tools ----
  {
    name: "create_column",
    description:
      "Create a new column (field) on a Dataverse table via the Metadata API. " +
      "Supported fieldTypes: String, Memo, Integer, Decimal, Double, Boolean, DateTime. " +
      "The field logical name should include the publisher prefix, e.g. fmk_myfield.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName", "displayName", "fieldType"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name for the new field including prefix, e.g. fmk_myfield" },
        displayName: { type: "string", description: "User-visible label for the field" },
        fieldType: {
          type: "string",
          enum: ["String", "Memo", "Integer", "Decimal", "Double", "Boolean", "DateTime"],
          description: "Dataverse field type",
        },
        maxLength: { type: "number", description: "Max length for String/Memo fields (default: 100 for String, 2000 for Memo)" },
        requiredLevel: { type: "string", enum: ["None", "Recommended", "Required"], description: "Required level (default: None)" },
        description: { type: "string", description: "Optional description for the field" },
        minValue: { type: "number", description: "Minimum value for Integer/Decimal/Double fields" },
        maxValue: { type: "number", description: "Maximum value for Integer/Decimal/Double fields" },
        dateTimeFormat: { type: "string", enum: ["DateOnly", "DateAndTime"], description: "DateTime format (default: DateOnly)" },
        trueLabel: { type: "string", description: "Label for true/yes option on Boolean fields (default: Yes)" },
        falseLabel: { type: "string", description: "Label for false/no option on Boolean fields (default: No)" },
      },
    },
  },
  {
    name: "create_choice_column",
    description:
      "Create a Choice (Picklist) or Multi-Select Choice column on a Dataverse table. " +
      "Use isMultiSelect=true for a multi-select field. " +
      "Provide choices as an array of {value, label} objects. Use values >= 100000000 for custom options.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName", "displayName", "choices"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name for the new field including prefix, e.g. fmk_status" },
        displayName: { type: "string", description: "User-visible label for the field" },
        choices: {
          type: "array",
          description: "Array of choice options",
          items: {
            type: "object",
            required: ["value", "label"],
            properties: {
              value: { type: "number", description: "Integer value for the option (use >= 100000000 for custom)" },
              label: { type: "string", description: "Display label for the option" },
            },
          },
        },
        defaultValue: { type: "number", description: "Default option value (must match one of the choice values)" },
        isMultiSelect: { type: "boolean", description: "true = Multi-Select Choice, false = single Choice (default: false)" },
        requiredLevel: { type: "string", enum: ["None", "Recommended", "Required"], description: "Required level (default: None)" },
        description: { type: "string", description: "Optional description for the field" },
      },
    },
  },
  {
    name: "create_lookup_column",
    description:
      "Create a Lookup field on a Dataverse table that references another table. " +
      "The user must specify which table to look up (referencedEntity) and which field from that table to display (displayFieldLogicalName). " +
      "Note: the standard Dataverse lookup control always shows the primary name field of the referenced table.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName", "displayName", "referencedEntity", "displayFieldLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table to add the lookup to, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name for the new lookup field including prefix, e.g. fmk_contactid" },
        displayName: { type: "string", description: "User-visible label for the lookup field" },
        referencedEntity: { type: "string", description: "Logical name of the table being looked up, e.g. contact" },
        displayFieldLogicalName: { type: "string", description: "Logical name of the field on the referenced table to display, e.g. fullname" },
        requiredLevel: { type: "string", enum: ["None", "Recommended", "Required"], description: "Required level (default: None)" },
        description: { type: "string", description: "Optional description for the field" },
        relationshipSchemaName: { type: "string", description: "Optional custom schema name for the relationship (auto-generated if omitted)" },
      },
    },
  },
  {
    name: "set_field_requirement",
    description:
      "Set the required level of an existing field on a Dataverse table. " +
      "This updates the field metadata directly (not via a business rule).",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName", "requirementLevel"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name of the field, e.g. telephone1" },
        requirementLevel: {
          type: "string",
          enum: ["None", "Recommended", "Required"],
          description: "None = optional, Recommended = soft required, Required = hard required",
        },
      },
    },
  },
  {
    name: "list_views",
    description: "List views (saved queries) for a Dataverse entity.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        queryType: {
          type: "number",
          description: "View type: 0=Public (default), 1=Advanced Find, 2=Associated, 4=Quick Find",
        },
      },
    },
  },
  {
    name: "add_field_to_view",
    description:
      "Add a column to an existing view (saved query) on a Dataverse entity. " +
      "Updates both the layoutxml and fetchxml, then publishes the entity.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "viewId", "fieldLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        viewId: { type: "string", description: "GUID of the view (savedqueryid) to update" },
        fieldLogicalName: { type: "string", description: "Logical name of the field to add, e.g. telephone1" },
        width: { type: "number", description: "Column width in pixels (default: 100)" },
      },
    },
  },
  {
    name: "list_business_rules",
    description: "List all business rules for a Dataverse entity.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
      },
    },
  },
  {
    name: "create_business_rule",
    description:
      "Create a new form Business Rule for a Dataverse entity. " +
      "Supports actions: SetRequired (value: required|recommended|none), " +
      "SetVisible (value: show|hide), SetValue (value: the value to set). " +
      "An optional single-field condition can be applied. " +
      "The rule is created in Draft state — activate it with activate_business_rule.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "ruleName", "actions"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        ruleName: { type: "string", description: "Display name for the business rule" },
        actions: {
          type: "array",
          description: "List of actions the rule should perform",
          items: {
            type: "object",
            required: ["type", "fieldLogicalName", "value"],
            properties: {
              type: {
                type: "string",
                enum: ["SetRequired", "SetVisible", "SetValue"],
                description: "Action type",
              },
              fieldLogicalName: { type: "string", description: "Target field logical name" },
              value: {
                type: "string",
                description: "For SetRequired: required|recommended|none. For SetVisible: show|hide. For SetValue: the value.",
              },
            },
          },
        },
        triggerOnCreate: { type: "boolean", description: "Trigger when a record is created (default: true)" },
        triggerOnUpdate: { type: "boolean", description: "Trigger when a record is updated (default: true)" },
        conditionField: { type: "string", description: "Optional: logical name of field to condition on" },
        conditionOperator: {
          type: "string",
          enum: ["Equal", "NotEqual", "Contains", "GreaterThan", "LessThan", "IsNull", "IsNotNull"],
          description: "Operator for the condition (default: Equal)",
        },
        conditionValue: { type: "string", description: "Value to compare in the condition" },
      },
    },
  },
  {
    name: "activate_business_rule",
    description: "Activate or deactivate a Dataverse business rule.",
    inputSchema: {
      type: "object",
      required: ["workflowId", "activate"],
      properties: {
        workflowId: { type: "string", description: "GUID of the business rule (workflow)" },
        activate: { type: "boolean", description: "true = activate, false = set to draft" },
      },
    },
  },
  {
    name: "update_column",
    description:
      "Modify properties of an existing column (field) on a Dataverse table. " +
      "Can update display name, description, required level, max length, min/max values.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name of the field to update, e.g. telephone1" },
        displayName: { type: "string", description: "New display label for the field" },
        description: { type: "string", description: "New description for the field" },
        requiredLevel: { type: "string", enum: ["None", "Recommended", "Required"], description: "New required level" },
        maxLength: { type: "number", description: "New max length (String/Memo fields only)" },
        minValue: { type: "number", description: "New minimum value (Integer/Decimal/Double fields only)" },
        maxValue: { type: "number", description: "New maximum value (Integer/Decimal/Double fields only)" },
      },
    },
  },
  {
    name: "add_choice_option",
    description: "Add a new option value to an existing Choice (Picklist) or Multi-Select Choice field.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName", "optionValue", "optionLabel"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name of the choice field" },
        optionValue: { type: "number", description: "Integer value for the new option (use >= 100000000 for custom)" },
        optionLabel: { type: "string", description: "Display label for the new option" },
      },
    },
  },
  {
    name: "get_choice_options",
    description: "List all current options on a Choice (Picklist) or Multi-Select Choice field.",
    inputSchema: {
      type: "object",
      required: ["entityLogicalName", "fieldLogicalName"],
      properties: {
        entityLogicalName: { type: "string", description: "Logical name of the table, e.g. account" },
        fieldLogicalName: { type: "string", description: "Logical name of the choice field" },
      },
    },
  },
];

// ----------------------------------------------------------------
// Server factory – one server per session
// ----------------------------------------------------------------

function createServer(sessionCfg: { current: SessionConfig }): Server {
  const server = new Server(
    { name: "dataverse-account-forms-mcp", version: "1.0.0" },
    {
      capabilities: { tools: {} },
      instructions: "Customise Dynamics 365 Account entity forms, columns, views, and business rules via Dataverse Web API. Supports creating and editing form layouts including tabs, sections, and fields.",
    },
  );

  // Per-session form XML staging cache
  const formCache: FormCache = new Map();

  // Build a DataverseClient — clientId/clientSecret are ALWAYS server-side only.
  // Customers provide tenantId + environmentUrl via configure_environment.
  function getClient(): DataverseClient {
    const cfg: DynamicsConfig = {
      tenantId: sessionCfg.current.tenantId ?? process.env["TENANT_ID"] ?? "",
      clientId: process.env["CLIENT_ID"] ?? "",
      clientSecret: process.env["CLIENT_SECRET"] ?? "",
      environmentUrl: sessionCfg.current.environmentUrl ?? process.env["ENVIRONMENT_URL"] ?? "",
      targetEntity: sessionCfg.current.targetEntity ?? process.env["TARGET_ENTITY"] ?? "account",
      targetAppUniqueName: sessionCfg.current.targetAppUniqueName ?? process.env["TARGET_APP_UNIQUE_NAME"] ?? undefined,
      solutionUniqueName: sessionCfg.current.solutionUniqueName ?? process.env["SOLUTION_UNIQUE_NAME"] ?? "MCPautoSetup",
    };

    if (!cfg.tenantId || !cfg.environmentUrl) {
      throw new McpError(
        ErrorCode.InvalidParams,
        "Dataverse environment not configured. " +
          "Call configure_environment with tenantId and environmentUrl, " +
          "or set TENANT_ID and ENVIRONMENT_URL in the server .env file.",
      );
    }
    if (!cfg.clientId || !cfg.clientSecret) {
      throw new McpError(
        ErrorCode.InternalError,
        "Server is not configured: CLIENT_ID and CLIENT_SECRET must be set in the server environment.",
      );
    }
    return new DataverseClient(cfg);
  }

  // List tools
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOL_DEFINITIONS }));

  // Call tool
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args = {} } = request.params;
    process.stderr.write(`[mcp] tool: ${name}\n`);

    try {
      let result: unknown;

      if (name === "configure_environment") {
        // Only allow tenantId, environmentUrl, targetEntity, targetAppUniqueName from callers.
        // clientId and clientSecret are server-side only and cannot be overridden.
        const a = args as Record<string, unknown>;
        sessionCfg.current = {
          ...sessionCfg.current,
          tenantId: a["tenantId"] as string | undefined,
          environmentUrl: a["environmentUrl"] as string | undefined,
          targetEntity: a["targetEntity"] as string | undefined,
          targetAppUniqueName: a["targetAppUniqueName"] as string | undefined,
        };
        result = { success: true, message: `Environment configured: tenant=${sessionCfg.current.tenantId}, url=${sessionCfg.current.environmentUrl}` };
      } else {
        const client = getClient();

        switch (name) {
          case "test_connection": {
            const r = await client.testConnection();
            if (!r.tokenOk) {
              result = {
                success: false,
                stage: "token",
                message: `❌ Token acquisition failed. Check tenantId, clientId, clientSecret.`,
                detail: r.tokenError,
              };
            } else if (!r.dataverseOk) {
              result = {
                success: false,
                stage: "dataverse",
                message: `✅ Token OK  |  ❌ Dataverse API call failed. Check environmentUrl and that the app has an Application User with a security role in PPAC.`,
                detail: r.dataverseError,
              };
            } else {
              result = {
                success: true,
                stage: "all",
                message: `✅ Token OK  |  ✅ Dataverse API OK  |  ✅ App user confirmed`,
                whoAmI: r.whoAmI,
              };
            }
            break;
          }

          case "list_account_forms":
            result = await Forms.listAccountForms(client);
            break;

          case "get_form_xml":
            result = await Forms.getFormXmlContent(client, guid(args, "formId"));
            break;

          case "get_attribute_metadata":
            result = await Forms.getAttributeMetadataTool(
              client,
              str(args, "entityName"),
              str(args, "fieldName"),
            );
            break;

          case "list_tabs":
            result = await Forms.listTabsOnFormTool(client, formCache, guid(args, "formId"));
            break;

          case "rename_tab":
            result = await Forms.renameTabTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "newLabel"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "remove_tab_from_form":
            result = await Forms.removeTabFromFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "list_sections":
            result = await Forms.listSectionsOnFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
            );
            break;

          case "rename_section":
            result = await Forms.renameSectionTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "sectionName"),
              str(args, "newLabel"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "remove_section_from_form":
            result = await Forms.removeSectionFromFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "sectionName"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "move_section_on_form":
            result = await Forms.moveSectionOnFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "sectionName"),
              str(args, "targetTabName"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "add_tab_to_form":
            result = await Forms.addTabToFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "tabLabel"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "add_section_to_form":
            result = await Forms.addSectionToFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "sectionName"),
              str(args, "sectionLabel"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "add_field_to_form":
            result = await Forms.addFieldToFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "tabName"),
              str(args, "sectionName"),
              str(args, "fieldName"),
              str(args, "fieldLabel"),
              boolDef(args, "autoCommit", false),
              (args as Record<string, unknown>)["attributeType"] as string | undefined,
            );
            break;

          case "remove_field_from_form":
            result = await Forms.removeFieldFromFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "fieldName"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "set_field_properties":
            result = await Forms.setFieldPropertiesTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "fieldName"),
              {
                disabled: bool(args, "disabled"),
                visible: bool(args, "visible"),
              },
              boolDef(args, "autoCommit", false),
            );
            break;

          case "move_field_on_form":
            result = await Forms.moveFieldOnFormTool(
              client,
              formCache,
              guid(args, "formId"),
              str(args, "fieldName"),
              str(args, "targetTabName"),
              str(args, "targetSectionName"),
              boolDef(args, "autoCommit", false),
            );
            break;

          case "update_form_xml_raw":
            result = await Forms.updateFormXmlRaw(
              client,
              guid(args, "formId"),
              str(args, "formXml"),
            );
            break;

          case "ensure_solution":
            result = await Forms.ensureSolutionTool(client);
            break;

          case "add_form_to_solution":
            result = await Forms.addFormToSolutionTool(client, guid(args, "formId"));
            break;

          case "publish_customisations":
            result = await Forms.publishCustomisations(client);
            break;

          case "add_entity_to_solution":
            result = await Solution.addEntityToSolutionTool(client, str(args, "entityLogicalName"));
            break;

          case "add_field_to_solution":
            result = await Solution.addFieldToSolutionTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
            );
            break;

          case "list_app_modules":
            result = await Solution.listAppModulesTool(client);
            break;

          case "add_form_to_app":
            result = await Solution.addFormToAppTool(
              client,
              guid(args, "formId"),
              (args as Record<string, unknown>)["appUniqueName"] as string | undefined,
            );
            break;

          case "add_app_to_solution":
            result = await Solution.addAppToSolutionTool(
              client,
              (args as Record<string, unknown>)["appUniqueName"] as string | undefined,
            );
            break;

          case "commit_form_changes":
            result = await Forms.commitFormChanges(client, formCache, guid(args, "formId"));
            break;

          case "discard_form_changes":
            result = await Forms.discardFormChanges(formCache, guid(args, "formId"));
            break;

          case "list_staged_changes":
            result = await Forms.listStagedChanges(formCache);
            break;

          // ---- Schema / metadata tools ----

          case "create_column":
            result = await Schema.createColumnTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              str(args, "displayName"),
              str(args, "fieldType"),
              {
                maxLength: (args as Record<string, unknown>)["maxLength"] as number | undefined,
                requiredLevel: (args as Record<string, unknown>)["requiredLevel"] as string | undefined,
                description: (args as Record<string, unknown>)["description"] as string | undefined,
                minValue: (args as Record<string, unknown>)["minValue"] as number | undefined,
                maxValue: (args as Record<string, unknown>)["maxValue"] as number | undefined,
                dateTimeFormat: (args as Record<string, unknown>)["dateTimeFormat"] as string | undefined,
                trueLabel: (args as Record<string, unknown>)["trueLabel"] as string | undefined,
                falseLabel: (args as Record<string, unknown>)["falseLabel"] as string | undefined,
              },
            );
            break;

          case "create_choice_column": {
            const rawChoices = (args as Record<string, unknown>)["choices"];
            const choices = Array.isArray(rawChoices)
              ? (rawChoices as Array<{ value: number; label: string }>)
              : [];
            result = await Schema.createChoiceColumnTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              str(args, "displayName"),
              choices,
              {
                defaultValue: (args as Record<string, unknown>)["defaultValue"] as number | undefined,
                isMultiSelect: bool(args, "isMultiSelect"),
                requiredLevel: (args as Record<string, unknown>)["requiredLevel"] as string | undefined,
                description: (args as Record<string, unknown>)["description"] as string | undefined,
              },
            );
            break;
          }

          case "create_lookup_column":
            result = await Schema.createLookupColumnTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              str(args, "displayName"),
              str(args, "referencedEntity"),
              str(args, "displayFieldLogicalName"),
              {
                requiredLevel: (args as Record<string, unknown>)["requiredLevel"] as string | undefined,
                description: (args as Record<string, unknown>)["description"] as string | undefined,
                relationshipSchemaName: (args as Record<string, unknown>)["relationshipSchemaName"] as string | undefined,
              },
            );
            break;

          case "set_field_requirement":
            result = await Schema.setFieldRequirementTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              str(args, "requirementLevel"),
            );
            break;

          case "list_views":
            result = await Schema.listViewsTool(
              client,
              str(args, "entityLogicalName"),
              (args as Record<string, unknown>)["queryType"] as number | undefined,
            );
            break;

          case "add_field_to_view":
            result = await Schema.addFieldToViewTool(
              client,
              str(args, "entityLogicalName"),
              guid(args, "viewId"),
              str(args, "fieldLogicalName"),
              (args as Record<string, unknown>)["width"] as number | undefined,
            );
            break;

          case "list_business_rules":
            result = await Schema.listBusinessRulesTool(client, str(args, "entityLogicalName"));
            break;

          case "create_business_rule": {
            const rawActions = (args as Record<string, unknown>)["actions"];
            const actions = Array.isArray(rawActions)
              ? (rawActions as Array<{ type: "SetRequired" | "SetVisible" | "SetValue"; fieldLogicalName: string; value: string }>)
              : [];
            result = await Schema.createBusinessRuleTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "ruleName"),
              actions,
              {
                triggerOnCreate: bool(args, "triggerOnCreate"),
                triggerOnUpdate: bool(args, "triggerOnUpdate"),
                conditionField: (args as Record<string, unknown>)["conditionField"] as string | undefined,
                conditionOperator: (args as Record<string, unknown>)["conditionOperator"] as string | undefined,
                conditionValue: (args as Record<string, unknown>)["conditionValue"] as string | undefined,
              },
            );
            break;
          }

          case "activate_business_rule":
            result = await Schema.activateBusinessRuleTool(
              client,
              guid(args, "workflowId"),
              boolDef(args, "activate", true),
            );
            break;

          case "update_column":
            result = await Schema.updateColumnTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              {
                displayName: (args as Record<string, unknown>)["displayName"] as string | undefined,
                description: (args as Record<string, unknown>)["description"] as string | undefined,
                requiredLevel: (args as Record<string, unknown>)["requiredLevel"] as string | undefined,
                maxLength: (args as Record<string, unknown>)["maxLength"] as number | undefined,
                minValue: (args as Record<string, unknown>)["minValue"] as number | undefined,
                maxValue: (args as Record<string, unknown>)["maxValue"] as number | undefined,
              },
            );
            break;

          case "add_choice_option":
            result = await Schema.addChoiceOptionTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
              (args as Record<string, unknown>)["optionValue"] as number,
              str(args, "optionLabel"),
            );
            break;

          case "get_choice_options":
            result = await Schema.getChoiceOptionsTool(
              client,
              str(args, "entityLogicalName"),
              str(args, "fieldLogicalName"),
            );
            break;

          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
        }
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (err) {
      if (err instanceof McpError) throw err;
      // Surface Dataverse API error details if available (axios error)
      const axiosBody = (err as any)?.response?.data;
      const axiosStatus = (err as any)?.response?.status;
      const baseMsg = err instanceof Error ? err.message : String(err);
      const detail = axiosBody
        ? ` | HTTP ${axiosStatus}: ${JSON.stringify(axiosBody)}`
        : "";
      const fullMsg = `${baseMsg}${detail}`;
      process.stderr.write(`[mcp] tool error: ${fullMsg}\n`);
      throw new McpError(ErrorCode.InternalError, fullMsg);
    }
  });

  return server;
}

// ----------------------------------------------------------------
// Argument helpers
// ----------------------------------------------------------------

function str(args: Record<string, unknown>, key: string): string {
  const v = args[key];
  if (typeof v !== "string" || !v) {
    throw new McpError(ErrorCode.InvalidParams, `Parameter '${key}' is required and must be a non-empty string`);
  }
  return v;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validates that a string param is a well-formed GUID. */
function guid(args: Record<string, unknown>, key: string): string {
  const v = str(args, key);
  if (!GUID_RE.test(v)) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `Parameter '${key}' must be a valid GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx), got: '${v}'`,
    );
  }
  return v;
}

function bool(
  args: Record<string, unknown>,
  key: string,
): boolean | undefined {
  const v = args[key];
  if (v === undefined || v === null) return undefined;
  if (typeof v === "boolean") return v;
  throw new McpError(ErrorCode.InvalidParams, `Parameter '${key}' must be a boolean`);
}

function boolDef(args: Record<string, unknown>, key: string, defaultVal: boolean): boolean {
  const v = args[key];
  if (v === undefined || v === null) return defaultVal;
  if (typeof v === "boolean") return v;
  if (v === "true") return true;
  if (v === "false") return false;
  return defaultVal;
}

// ----------------------------------------------------------------
// HTTP mode (Copilot Studio / remote)
// ----------------------------------------------------------------

async function startHttp(port: number): Promise<void> {
  const app = express();
  app.use(express.json());

  // Map sessionId → { transport, sessionConfig }
  const sessions = new Map<
    string,
    { transport: StreamableHTTPServerTransport; cfg: { current: SessionConfig } }
  >();

  app.post("/mcp", async (req: Request, res: Response) => {
    const existingId = req.headers["mcp-session-id"] as string | undefined;
    process.stderr.write(`[mcp] POST sessionId=${existingId ?? "none"} sessions=[${[...sessions.keys()].join(",")}]\n`);

    if (existingId && sessions.has(existingId)) {
      // Existing session – reuse transport
      await sessions.get(existingId)!.transport.handleRequest(req, res, req.body);
      return;
    }

    // New session
    const cfg: { current: SessionConfig } = { current: {} };
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (id) => {
        sessions.set(id, { transport, cfg });
      },
    });

    transport.onclose = () => {
      if (transport.sessionId) sessions.delete(transport.sessionId);
    };

    const server = createServer(cfg);
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  });

  const handleSessionRequest = async (req: Request, res: Response) => {
    const id = req.headers["mcp-session-id"] as string | undefined;
    if (!id || !sessions.has(id)) {
      res.status(400).json({ error: "Invalid or missing session ID" });
      return;
    }
    await sessions.get(id)!.transport.handleRequest(req, res);
  };

  app.get("/mcp", handleSessionRequest);
  app.delete("/mcp", handleSessionRequest);

  app.get("/health", (_req, res) => res.json({ status: "ok" }));

  app.listen(port, () => {
    process.stderr.write(`[dataverse-mcp] v${SERVER_VERSION} HTTP server listening on port ${port}\n`);
  });
}

// ----------------------------------------------------------------
// Stdio mode (VS Code local development)
// ----------------------------------------------------------------

async function startStdio(): Promise<void> {
  const cfg: { current: SessionConfig } = { current: {} };
  const server = createServer(cfg);
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write("[dataverse-mcp] stdio transport connected\n");
}

// ----------------------------------------------------------------
// Entry point
// ----------------------------------------------------------------

const useStdio =
  process.argv.includes("--stdio") ||
  process.env["TRANSPORT"]?.toLowerCase() === "stdio";

if (useStdio) {
  startStdio().catch((err) => {
    process.stderr.write(`[dataverse-mcp] Fatal: ${err}\n`);
    process.exit(1);
  });
} else {
  const port = parseInt(process.env["PORT"] ?? "3000", 10);
  startHttp(port).catch((err) => {
    process.stderr.write(`[dataverse-mcp] Fatal: ${err}\n`);
    process.exit(1);
  });
}
