import type { DataverseClient } from "../dataverse.js";

/** Fallback used only when client.solutionUniqueName is not available */
export const SOLUTION_UNIQUE_NAME = "MCPautoSetup";
export const PUBLISHER_UNIQUE_NAME = "fellowmind";
const PUBLISHER_FRIENDLY_NAME = "Fellowmind";
/** Publisher customisation prefix (max 8 chars, lower-case, no underscores at start) */
const PUBLISHER_PREFIX = "fmk";

/** SystemForm component type code in Dataverse solution components */
const COMPONENT_TYPE_SYSTEM_FORM = 24;
/** Entity (table) component type code */
const COMPONENT_TYPE_ENTITY = 1;
/** Attribute (field) component type code */
const COMPONENT_TYPE_ATTRIBUTE = 2;

/**
 * Ensure the "Fellowmind" publisher and "MCPautoSetup" solution exist.
 * Creates them if they are missing and returns the solution ID.
 */
export async function ensureSolution(client: DataverseClient): Promise<string> {
  const solutionName = client.solutionUniqueName;
  // ---- Publisher ----
  let publisher = await client.getPublisher(PUBLISHER_UNIQUE_NAME);
  let publisherId: string;

  if (!publisher) {
    publisherId = await client.createPublisher({
      uniquename: PUBLISHER_UNIQUE_NAME,
      friendlyname: PUBLISHER_FRIENDLY_NAME,
      customizationprefix: PUBLISHER_PREFIX,
    });
  } else {
    publisherId = publisher.publisherid;
  }

  // ---- Solution ----
  const existing = await client.getSolution(solutionName);
  let solutionId: string;

  if (existing) {
    solutionId = existing.solutionid;
  } else {
    solutionId = await client.createSolution({
      uniquename: solutionName,
      friendlyname: solutionName,
      "publisherid@odata.bind": `/publishers(${publisherId})`,
      version: "1.0.0.0",
    });
  }

  // Always ensure the target entity is in the solution (idempotent)
  const entityMetadataId = await client.getEntityMetadataId(client.targetEntity);
  if (entityMetadataId) {
    try {
      await client.addComponentToSolution(entityMetadataId, COMPONENT_TYPE_ENTITY, solutionName);
    } catch {
      // Non-fatal — already tracked
    }
  }

  // Ensure the target app is in the solution if configured (idempotent)
  const appUniqueName = client.targetAppUniqueName;
  if (appUniqueName) {
    const app = await client.getAppModuleByUniqueName(appUniqueName);
    if (app) {
      try {
        await client.addComponentToSolution(app.appmoduleid, 80, solutionName);
      } catch {
        // Non-fatal
      }
    }
  }

  return solutionId;
}

/** Add a SystemForm component to the MCPautoSetup solution. Ensures solution exists first.
 *  NOTE: Forms created via the Web API are not trackable via AddSolutionComponent —
 *  this call is best-effort and failures are logged + silently skipped. */
export async function addFormToSolution(
  client: DataverseClient,
  formId: string,
): Promise<void> {
  const solutionName = client.solutionUniqueName;
  await ensureSolution(client);

  try {
    await client.addComponentToSolution(formId, COMPONENT_TYPE_SYSTEM_FORM, solutionName);
  } catch (err: unknown) {
    const errData = (err as { response?: { data?: { error?: { code?: string; message?: string } } } })?.response?.data?.error;
    const code = errData?.code ?? "";
    if (code === "0x80048423") return; // already in solution — fine
    // Log and skip — form XML was already updated and published successfully
    process.stderr.write(`[mcp] warn: could not add form to solution (skipping): code=${code} msg=${errData?.message ?? String(err)}\n`);
  }
}

/** Add a table (entity) to the MCPautoSetup solution by logical name. */
export async function addEntityToSolutionTool(
  client: DataverseClient,
  entityLogicalName: string,
): Promise<{ success: boolean; message: string }> {
  const metadataId = await client.getEntityMetadataId(entityLogicalName);
  if (!metadataId) {
    throw new Error(`Entity '${entityLogicalName}' not found in this environment`);
  }
  await client.addComponentToSolution(metadataId, COMPONENT_TYPE_ENTITY, client.solutionUniqueName);
  return { success: true, message: `Table '${entityLogicalName}' added to '${client.solutionUniqueName}' solution` };
}

/** Add a field (attribute) to the MCPautoSetup solution by entity + field logical name. */
export async function addFieldToSolutionTool(
  client: DataverseClient,
  entityLogicalName: string,
  fieldLogicalName: string,
): Promise<{ success: boolean; message: string }> {
  const metadataId = await client.getAttributeMetadataId(entityLogicalName, fieldLogicalName);
  if (!metadataId) {
    throw new Error(`Field '${fieldLogicalName}' on entity '${entityLogicalName}' not found`);
  }
  await client.addComponentToSolution(metadataId, COMPONENT_TYPE_ATTRIBUTE, client.solutionUniqueName);
  return { success: true, message: `Field '${entityLogicalName}.${fieldLogicalName}' added to '${client.solutionUniqueName}' solution` };
}

/** List all model-driven apps in the environment. */
export async function listAppModulesTool(client: DataverseClient) {
  const apps = await client.listAppModules();
  return apps.map((a) => ({
    appModuleId: a.appmoduleid,
    uniqueName: a.uniquename,
    name: a.name,
    description: a.description ?? "",
  }));
}

/**
 * Add a form to a model-driven app and track the app in MCPautoSetup.
 * Resolves the app by unique name (uses TARGET_APP_UNIQUE_NAME if appUniqueName is omitted).
 */
export async function addFormToAppTool(
  client: DataverseClient,
  formId: string,
  appUniqueName?: string,
): Promise<{ success: boolean; message: string }> {
  const uniqueName = appUniqueName ?? client.targetAppUniqueName;
  if (!uniqueName) {
    return { success: true, message: "Skipped: no app unique name configured (set TARGET_APP_UNIQUE_NAME in .env to enable)" };
  }

  const app = await client.getAppModuleByUniqueName(uniqueName);
  if (!app) throw new Error(`App '${uniqueName}' not found in this environment`);

  await client.addFormToAppModule(app.appmoduleid, formId);
  return { success: true, message: `Form added to app '${app.name}' (${uniqueName})` };
}

/** Add a model-driven app to the MCPautoSetup solution by unique name. */
export async function addAppToSolutionTool(
  client: DataverseClient,
  appUniqueName?: string,
): Promise<{ success: boolean; message: string }> {
  const uniqueName = appUniqueName ?? client.targetAppUniqueName;
  if (!uniqueName) {
    return { success: true, message: "Skipped: no app unique name configured (set TARGET_APP_UNIQUE_NAME in .env to enable)" };
  }

  const app = await client.getAppModuleByUniqueName(uniqueName);
  if (!app) throw new Error(`App '${uniqueName}' not found in this environment`);

  await client.addComponentToSolution(app.appmoduleid, 80, client.solutionUniqueName);
  return { success: true, message: `App '${app.name}' added to '${client.solutionUniqueName}' solution` };
}
