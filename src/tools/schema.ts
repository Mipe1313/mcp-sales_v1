import { randomUUID } from "crypto";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import type { DataverseClient } from "../dataverse.js";
import { ensureSolution } from "./solution.js";

// ----------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
const xmlBuilder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", format: true });

function parseXml(xml: string): Record<string, unknown> {
  return xmlParser.parse(xml) as Record<string, unknown>;
}

function buildXml(obj: Record<string, unknown>): string {
  return xmlBuilder.build(obj) as string;
}

function toArray<T>(v: T | T[] | undefined): T[] {
  if (!v) return [];
  return Array.isArray(v) ? v : [v];
}

// ----------------------------------------------------------------
// create_column
// ----------------------------------------------------------------

export async function createColumnTool(
  client: DataverseClient,
  entityLogicalName: string,
  fieldLogicalName: string,
  displayName: string,
  fieldType: string,
  options: {
    maxLength?: number;
    requiredLevel?: string;
    description?: string;
    minValue?: number;
    maxValue?: number;
    dateTimeFormat?: string;
    trueLabel?: string;
    falseLabel?: string;
  },
): Promise<unknown> {
  await ensureSolution(client);
  const solutionUniqueName = client.solutionUniqueName;

  const validTypes = ["String", "Memo", "Integer", "Decimal", "Double", "Boolean", "DateTime"] as const;
  type ValidType = (typeof validTypes)[number];
  if (!validTypes.includes(fieldType as ValidType)) {
    throw new Error(`Invalid fieldType '${fieldType}'. Must be one of: ${validTypes.join(", ")}`);
  }

  const validLevels = ["None", "Recommended", "Required"] as const;
  type ValidLevel = (typeof validLevels)[number];
  const reqLevel = (options.requiredLevel as ValidLevel) ?? "None";
  if (!validLevels.includes(reqLevel)) {
    throw new Error(`Invalid requiredLevel '${reqLevel}'. Must be: None, Recommended, or Required`);
  }

  const metadataId = await client.createSimpleAttribute(
    entityLogicalName,
    fieldLogicalName,
    displayName,
    fieldType as ValidType,
    {
      maxLength: options.maxLength,
      requiredLevel: reqLevel,
      description: options.description,
      minValue: options.minValue,
      maxValue: options.maxValue,
      dateTimeFormat: options.dateTimeFormat as "DateOnly" | "DateAndTime" | undefined,
      trueLabel: options.trueLabel,
      falseLabel: options.falseLabel,
    },
    solutionUniqueName,
  );

  return {
    success: true,
    metadataId,
    logicalName: fieldLogicalName,
    entityLogicalName,
    fieldType,
    message: `Field '${fieldLogicalName}' (${fieldType}) created on '${entityLogicalName}'. MetadataId: ${metadataId}`,
  };
}

// ----------------------------------------------------------------
// create_choice_column
// ----------------------------------------------------------------

export async function createChoiceColumnTool(
  client: DataverseClient,
  entityLogicalName: string,
  fieldLogicalName: string,
  displayName: string,
  choices: Array<{ value: number; label: string }>,
  options: {
    defaultValue?: number;
    isMultiSelect?: boolean;
    requiredLevel?: string;
    description?: string;
  },
): Promise<unknown> {
  await ensureSolution(client);
  const solutionUniqueName = client.solutionUniqueName;

  if (!choices || choices.length === 0) {
    throw new Error("At least one choice option is required");
  }

  const validLevels = ["None", "Recommended", "Required"] as const;
  type ValidLevel = (typeof validLevels)[number];
  const reqLevel = (options.requiredLevel as ValidLevel) ?? "None";

  const metadataId = await client.createPicklistAttribute(
    entityLogicalName,
    fieldLogicalName,
    displayName,
    choices,
    {
      defaultValue: options.defaultValue,
      isMultiSelect: options.isMultiSelect ?? false,
      requiredLevel: reqLevel,
      description: options.description,
    },
    solutionUniqueName,
  );

  const typeName = options.isMultiSelect ? "MultiSelectPicklist" : "Picklist";
  return {
    success: true,
    metadataId,
    logicalName: fieldLogicalName,
    entityLogicalName,
    fieldType: typeName,
    choices,
    defaultValue: options.defaultValue,
    message: `${typeName} field '${fieldLogicalName}' created on '${entityLogicalName}' with ${choices.length} option(s). MetadataId: ${metadataId}`,
  };
}

// ----------------------------------------------------------------
// set_field_requirement
// ----------------------------------------------------------------

export async function setFieldRequirementTool(
  client: DataverseClient,
  entityLogicalName: string,
  fieldLogicalName: string,
  requirementLevel: string,
): Promise<unknown> {
  const validLevels = ["None", "Recommended", "Required"] as const;
  type ValidLevel = (typeof validLevels)[number];
  if (!validLevels.includes(requirementLevel as ValidLevel)) {
    throw new Error(`Invalid requirementLevel '${requirementLevel}'. Must be: None, Recommended, or Required`);
  }

  await client.updateAttributeRequirementLevel(
    entityLogicalName,
    fieldLogicalName,
    requirementLevel as ValidLevel,
  );

  return {
    success: true,
    entityLogicalName,
    fieldLogicalName,
    requirementLevel,
    message: `Field '${fieldLogicalName}' on '${entityLogicalName}' set to RequiredLevel='${requirementLevel}'`,
  };
}

// ----------------------------------------------------------------
// list_views
// ----------------------------------------------------------------

export async function listViewsTool(
  client: DataverseClient,
  entityLogicalName: string,
  queryType?: number,
): Promise<unknown> {
  const views = await client.listViews(entityLogicalName, queryType ?? 0);
  const typeLabels: Record<number, string> = {
    0: "Public View",
    1: "Advanced Find",
    2: "Associated",
    4: "Quick Find",
    64: "Lookup",
    128: "Workflow Participating",
  };
  return views.map((v) => ({
    viewId: v.savedqueryid,
    name: v.name,
    queryType: v.querytype,
    queryTypeLabel: typeLabels[v.querytype] ?? `Type ${v.querytype}`,
    statecode: v.statecode,
    description: v.description,
  }));
}

// ----------------------------------------------------------------
// add_field_to_view
// ----------------------------------------------------------------

export async function addFieldToViewTool(
  client: DataverseClient,
  entityLogicalName: string,
  viewId: string,
  fieldLogicalName: string,
  width = 100,
): Promise<unknown> {
  const view = await client.getSavedQueryById(viewId);
  if (!view.layoutxml) throw new Error("View has no layoutxml");
  if (!view.fetchxml) throw new Error("View has no fetchxml");

  // ---- Update layoutxml ----
  const layout = parseXml(view.layoutxml) as Record<string, unknown>;
  const grid = layout["grid"] as Record<string, unknown>;
  if (!grid) throw new Error("View layoutxml has no <grid> element");

  const row = grid["row"] as Record<string, unknown>;
  if (!row) throw new Error("View layoutxml has no <row> element");

  // Ensure cells is an array
  row["cell"] = toArray(row["cell"] as unknown);

  const cells = row["cell"] as Array<Record<string, unknown>>;

  // Check if field already present
  if (cells.some((c) => c["@_name"] === fieldLogicalName)) {
    return { success: false, message: `Field '${fieldLogicalName}' is already in the view layout` };
  }

  cells.push({ "@_name": fieldLogicalName, "@_width": String(width) });
  const updatedLayout = buildXml(layout);

  // ---- Update fetchxml ----
  const fetch = parseXml(view.fetchxml) as Record<string, unknown>;
  const fetchEntity = (fetch["fetch"] as Record<string, unknown>)?.["entity"] as Record<string, unknown>;
  if (!fetchEntity) throw new Error("fetchxml has no <entity> element");

  fetchEntity["attribute"] = toArray(fetchEntity["attribute"] as unknown);
  const attrs = fetchEntity["attribute"] as Array<Record<string, unknown>>;

  if (!attrs.some((a) => a["@_name"] === fieldLogicalName)) {
    attrs.push({ "@_name": fieldLogicalName });
  }
  const updatedFetch = buildXml(fetch);

  await client.updateSavedQueryLayoutXml(viewId, updatedLayout, updatedFetch);

  // Publish the entity to make the view change live
  await client.publishXml(`<importexportxml><entities><entity>${entityLogicalName}</entity></entities></importexportxml>`);

  return {
    success: true,
    viewId,
    viewName: view.name,
    fieldLogicalName,
    width,
    message: `Field '${fieldLogicalName}' added to view '${view.name}' and published`,
  };
}

// ----------------------------------------------------------------
// list_business_rules
// ----------------------------------------------------------------

export async function listBusinessRulesTool(
  client: DataverseClient,
  entityLogicalName: string,
): Promise<unknown> {
  const rules = await client.listBusinessRules(entityLogicalName);
  const stateLabels: Record<number, string> = { 0: "Draft", 1: "Activated", 2: "Suspended" };
  const scopeLabels: Record<number, string> = {
    1: "User",
    2: "BusinessUnit",
    4: "ParentChildBusinessUnit",
    8: "Organization",
  };
  return rules.map((r) => ({
    ruleId: r.workflowid,
    name: r.name,
    state: stateLabels[r.statecode] ?? `State ${r.statecode}`,
    scope: scopeLabels[r.scope] ?? `Scope ${r.scope}`,
    description: r.description,
  }));
}

// ----------------------------------------------------------------
// create_business_rule
// ----------------------------------------------------------------

/**
 * Generate clientdata JSON for a Dataverse form business rule.
 * Supports actions: SetRequired, SetVisible, SetValue.
 * Optionally a single "field equals value" condition can be applied to all actions.
 */
function buildBusinessRuleClientData(
  entityLogicalName: string,
  ruleName: string,
  triggerOnCreate: boolean,
  triggerOnUpdate: boolean,
  actions: Array<{
    type: "SetRequired" | "SetVisible" | "SetValue";
    fieldLogicalName: string;
    value: string;
  }>,
  condition?: {
    fieldLogicalName: string;
    operator: "Equal" | "NotEqual" | "Contains" | "GreaterThan" | "LessThan" | "IsNull" | "IsNotNull";
    value: string;
  },
): string {
  const conditionNodes = condition
    ? [
        {
          id: randomUUID(),
          type: "Condition",
          fieldName: condition.fieldLogicalName,
          operator: condition.operator,
          value: condition.value,
        },
      ]
    : [];

  const actionNodes = actions.map((a) => ({
    id: randomUUID(),
    type: a.type,
    fieldName: a.fieldLogicalName,
    value: a.value,
  }));

  const clientData = {
    schemaVersion: "2.0.0.0",
    businessRuleId: randomUUID(),
    displayName: ruleName,
    description: "",
    triggerOnCreate,
    triggerOnUpdate,
    scope: "Entity",
    entity: { entityName: entityLogicalName },
    conditionGroup: {
      type: "Group",
      logicalOperator: "AND",
      conditions: conditionNodes,
    },
    actions: actionNodes,
    elseActions: [],
  };

  return JSON.stringify(clientData);
}

export async function createBusinessRuleTool(
  client: DataverseClient,
  entityLogicalName: string,
  ruleName: string,
  actions: Array<{
    type: "SetRequired" | "SetVisible" | "SetValue";
    fieldLogicalName: string;
    value: string;
  }>,
  options: {
    triggerOnCreate?: boolean;
    triggerOnUpdate?: boolean;
    conditionField?: string;
    conditionOperator?: string;
    conditionValue?: string;
  } = {},
): Promise<unknown> {
  await ensureSolution(client);

  const triggerOnCreate = options.triggerOnCreate ?? true;
  const triggerOnUpdate = options.triggerOnUpdate ?? true;

  const condition =
    options.conditionField
      ? {
          fieldLogicalName: options.conditionField,
          operator: (options.conditionOperator ?? "Equal") as "Equal" | "NotEqual" | "Contains" | "GreaterThan" | "LessThan" | "IsNull" | "IsNotNull",
          value: options.conditionValue ?? "",
        }
      : undefined;

  const clientdata = buildBusinessRuleClientData(
    entityLogicalName,
    ruleName,
    triggerOnCreate,
    triggerOnUpdate,
    actions,
    condition,
  );

  const workflowId = await client.createWorkflowBusinessRule(
    entityLogicalName,
    ruleName,
    clientdata,
    client.solutionUniqueName,
  );

  return {
    success: true,
    workflowId,
    ruleName,
    entityLogicalName,
    triggerOnCreate,
    triggerOnUpdate,
    actions,
    condition,
    message: `Business rule '${ruleName}' created in Draft state. Activate it in the Power Apps maker portal or call activate_business_rule.`,
  };
}

// ----------------------------------------------------------------
// activate_business_rule
// ----------------------------------------------------------------

export async function activateBusinessRuleTool(
  client: DataverseClient,
  workflowId: string,
  activate: boolean,
): Promise<unknown> {
  await client.setBusinessRuleState(workflowId, activate ? 1 : 0);
  return {
    success: true,
    workflowId,
    state: activate ? "Activated" : "Draft",
    message: `Business rule ${workflowId} set to ${activate ? "Activated" : "Draft"}`,
  };
}
