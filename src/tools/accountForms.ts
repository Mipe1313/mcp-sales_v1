import type { DataverseClient } from "../dataverse.js";
import type { ClarificationResponse, FormCache } from "../types.js";
import {
  addFieldToForm,
  addSectionToForm,
  addTabToForm,
  findSectionCandidates,
  findTabCandidates,
  getClassIdForAttributeType,
  listSectionsOnForm,
  listTabsOnForm,
  moveFieldOnForm,
  moveSectionOnForm,
  removeFieldFromForm,
  removeSectionFromForm,
  removeTabFromForm,
  renameSection,
  renameTab,
  setFieldProperties,
} from "../formUtils.js";
import type { SectionCandidate, TabCandidate } from "../formUtils.js";
import { addFormToSolution, ensureSolution } from "./solution.js";

// ----------------------------------------------------------------
// Clarification helpers
// ----------------------------------------------------------------

function tabNotFound(tabName: string, candidates: TabCandidate[]): ClarificationResponse {
  return {
    needsClarification: true,
    field: "tabName",
    message: `Tab '${tabName}' was not found on this form. Please confirm the tab name from the candidates listed.`,
    candidates,
  };
}

function sectionNotFound(
  sectionName: string,
  tabName: string,
  candidates: SectionCandidate[],
): ClarificationResponse {
  return {
    needsClarification: true,
    field: "sectionName",
    message: `Section '${sectionName}' was not found in tab '${tabName}'. Please confirm the section name from the candidates listed.`,
    candidates,
  };
}

// ----------------------------------------------------------------
// Clone management
// ----------------------------------------------------------------
const CLONE_PREFIX = "[MCP] ";

/**
 * If `formId` points to an original (non-clone) form, find or create an
 * "[MCP] {name}" clone and return its formId.  If the form is already a
 * clone (name starts with CLONE_PREFIX) the same formId is returned unchanged.
 */
// ----------------------------------------------------------------
// Cache helpers
// ----------------------------------------------------------------

/** Get current XML for cloneId: from cache if staged, otherwise fetch from Dataverse. */
async function getXml(client: DataverseClient, cache: FormCache, cloneId: string): Promise<string> {
  const cached = cache.get(cloneId);
  if (cached) return cached.xml;
  const form = await client.getFormById(cloneId);
  if (!form.formxml) throw new Error("Form has no XML content");
  return form.formxml;
}

/** Stage updated XML in cache. If autoCommit, push to Dataverse immediately. */
async function stageOrCommit(
  client: DataverseClient,
  cache: FormCache,
  cloneId: string,
  cloneName: string,
  updatedXml: string,
  autoCommit: boolean,
): Promise<string> {
  cache.set(cloneId, { xml: updatedXml, cloneName });
  if (autoCommit) {
    await client.updateFormXml(cloneId, updatedXml);
    await addFormToSolution(client, cloneId);
    cache.delete(cloneId);
    return "committed";
  }
  return "staged";
}

async function ensureClone(client: DataverseClient, formId: string): Promise<{ cloneId: string; cloneName: string; wasCreated: boolean }> {
  const original = await client.getFormById(formId);
  if (!original.formxml) throw new Error("Form has no XML content");

  // Already a clone — use as-is
  if (original.name.startsWith(CLONE_PREFIX)) {
    return { cloneId: formId, cloneName: original.name, wasCreated: false };
  }

  const cloneName = `${CLONE_PREFIX}${original.name}`;

  // Check if clone already exists
  const existing = await client.findFormByName(cloneName);
  if (existing) {
    return { cloneId: existing.formid, cloneName, wasCreated: false };
  }

  // Create clone — register it in the solution at creation time via MSCRM.SolutionUniqueName header.
  // This avoids the "does not exist" error from AddSolutionComponent on a freshly-created form.
  const cloneInfo = await client.createForm(
    {
      name: cloneName,
      objecttypecode: client.targetEntity,
      type: 2,
      formxml: original.formxml,
      description: `MCP-managed clone of '${original.name}'`,
    },
    client.solutionUniqueName,
  );
  const cloneId = cloneInfo.formid;

  // Track clone in solution
  await ensureSolution(client);
  await addFormToSolution(client, cloneId);

  return { cloneId, cloneName, wasCreated: true };
}
async function resolveClassId(
  client: DataverseClient,
  fieldName: string,
  attributeType?: string,
): Promise<string> {
  if (attributeType) return getClassIdForAttributeType(attributeType);
  const meta = await client.getAttributeMetadata(client.targetEntity, fieldName);
  if (meta) return getClassIdForAttributeType(meta.AttributeType);
  return getClassIdForAttributeType("String");
}

// ----------------------------------------------------------------
// Tool handlers
// ----------------------------------------------------------------

export async function listAccountForms(client: DataverseClient) {
  const forms = await client.listAccountForms();
  return forms.map((f) => ({
    formId: f.formid,
    name: f.name,
    type: f.type,
    description: f.description ?? "",
  }));
}

export async function getFormXmlContent(client: DataverseClient, formId: string) {
  const form = await client.getFormById(formId);
  return { formId: form.formid, name: form.name, formXml: form.formxml ?? "" };
}

export async function addFieldToFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  sectionName: string,
  fieldName: string,
  fieldLabel: string,
  autoCommit: boolean,
  attributeType?: string,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const { match: secMatch, candidates: secCandidates } = findSectionCandidates(tabMatch, sectionName);
  if (!secMatch) return sectionNotFound(sectionName, resolvedTab, secCandidates);
  const resolvedSection = String(secMatch["@_name"] ?? sectionName);

  const classId = await resolveClassId(client, fieldName, attributeType);
  const updatedXml = addFieldToForm(xml, resolvedTab, resolvedSection, fieldName, classId, fieldLabel);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Field '${fieldName}' added to '${resolvedTab}/${resolvedSection}'${cloneNote} [${status}]` };
}

export async function removeFieldFromFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  fieldName: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);
  const updatedXml = removeFieldFromForm(xml, fieldName);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Field '${fieldName}' removed from form${cloneNote} [${status}]` };
}

export async function setFieldPropertiesTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  fieldName: string,
  props: { disabled?: boolean; visible?: boolean },
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);
  const updatedXml = setFieldProperties(xml, fieldName, props);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Field '${fieldName}' properties updated${cloneNote} [${status}]` };
}

export async function addTabToFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  tabLabel: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);
  const updatedXml = addTabToForm(xml, tabName, tabLabel);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Tab '${tabLabel}' added to form${cloneNote} [${status}]` };
}

export async function addSectionToFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  sectionName: string,
  sectionLabel: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const updatedXml = addSectionToForm(xml, resolvedTab, sectionName, sectionLabel);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Section '${sectionLabel}' added to tab '${resolvedTab}'${cloneNote} [${status}]` };
}

export async function moveFieldOnFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  fieldName: string,
  targetTabName: string,
  targetSectionName: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, targetTabName);
  if (!tabMatch) return tabNotFound(targetTabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? targetTabName);

  const { match: secMatch, candidates: secCandidates } = findSectionCandidates(tabMatch, targetSectionName);
  if (!secMatch) return sectionNotFound(targetSectionName, resolvedTab, secCandidates);
  const resolvedSection = String(secMatch["@_name"] ?? targetSectionName);

  const updatedXml = moveFieldOnForm(xml, fieldName, resolvedTab, resolvedSection);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Field '${fieldName}' moved to '${resolvedTab}/${resolvedSection}'${cloneNote} [${status}]` };
}

// ----------------------------------------------------------------
// Commit / discard staged changes
// ----------------------------------------------------------------

export async function commitFormChanges(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
) {
  // formId may be original or clone — resolve clone first
  const { cloneId, cloneName } = await ensureClone(client, formId);
  const entry = cache.get(cloneId);
  if (!entry) {
    return { success: true, formId: cloneId, message: `No staged changes for '${cloneName}' — nothing to commit` };
  }
  await client.updateFormXml(cloneId, entry.xml);
  await addFormToSolution(client, cloneId);
  cache.delete(cloneId);
  return { success: true, formId: cloneId, message: `Staged changes committed for '${cloneName}'` };
}

export async function discardFormChanges(
  cache: FormCache,
  formId: string,
) {
  const deleted = cache.delete(formId);
  // Also try deleting by cloneId in case formId was the original
  if (!deleted) {
    for (const [key] of cache) {
      if (key === formId) { cache.delete(key); break; }
    }
  }
  return { success: true, message: `Staged changes for form '${formId}' discarded` };
}

export async function listStagedChanges(cache: FormCache) {
  const entries = Array.from(cache.entries()).map(([cloneId, e]) => ({
    cloneId,
    cloneName: e.cloneName,
    xmlLength: e.xml.length,
  }));
  return { staged: entries, count: entries.length };
}

export async function updateFormXmlRaw(
  client: DataverseClient,
  formId: string,
  formXml: string,
) {
  await client.updateFormXml(formId, formXml);
  await addFormToSolution(client, formId);
  return { success: true, message: "Form XML updated successfully" };
}

export async function getAttributeMetadataTool(
  client: DataverseClient,
  entityName: string,
  fieldName: string,
) {
  const meta = await client.getAttributeMetadata(entityName, fieldName);
  if (!meta) throw new Error(`Attribute '${fieldName}' not found on entity '${entityName}'`);
  return {
    logicalName: meta.LogicalName,
    attributeType: meta.AttributeType,
    displayName: meta.DisplayName?.UserLocalizedLabel?.Label ?? fieldName,
  };
}

export async function ensureSolutionTool(client: DataverseClient) {
  const solutionId = await ensureSolution(client);
  return { success: true, solutionId, message: "Solution MCPautoSetup is ready" };
}

export async function addFormToSolutionTool(client: DataverseClient, formId: string) {
  await addFormToSolution(client, formId);
  return { success: true, message: `Form '${formId}' added to MCPautoSetup solution` };
}

// ----------------------------------------------------------------
// Tab / section CRUD tool handlers
// ----------------------------------------------------------------

export async function listTabsOnFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
) {
  const xml = cache.get(formId)?.xml ?? (await client.getFormById(formId)).formxml ?? "";
  return { tabs: listTabsOnForm(xml) };
}

export async function renameTabTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  newLabel: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const updatedXml = renameTab(xml, resolvedTab, newLabel);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Tab '${resolvedTab}' renamed to '${newLabel}'${cloneNote} [${status}]` };
}

export async function removeTabFromFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const updatedXml = removeTabFromForm(xml, resolvedTab);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Tab '${resolvedTab}' removed from form${cloneNote} [${status}]` };
}

export async function listSectionsOnFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
) {
  const xml = cache.get(formId)?.xml ?? (await client.getFormById(formId)).formxml ?? "";
  return { sections: listSectionsOnForm(xml, tabName) };
}

export async function renameSectionTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  sectionName: string,
  newLabel: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const { match: secMatch, candidates: secCandidates } = findSectionCandidates(tabMatch, sectionName);
  if (!secMatch) return sectionNotFound(sectionName, resolvedTab, secCandidates);
  const resolvedSection = String(secMatch["@_name"] ?? sectionName);

  const updatedXml = renameSection(xml, resolvedTab, resolvedSection, newLabel);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Section '${resolvedSection}' renamed to '${newLabel}'${cloneNote} [${status}]` };
}

export async function removeSectionFromFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  tabName: string,
  sectionName: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, tabName);
  if (!tabMatch) return tabNotFound(tabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? tabName);

  const { match: secMatch, candidates: secCandidates } = findSectionCandidates(tabMatch, sectionName);
  if (!secMatch) return sectionNotFound(sectionName, resolvedTab, secCandidates);
  const resolvedSection = String(secMatch["@_name"] ?? sectionName);

  const updatedXml = removeSectionFromForm(xml, resolvedTab, resolvedSection);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Section '${resolvedSection}' removed from tab '${resolvedTab}'${cloneNote} [${status}]` };
}

export async function moveSectionOnFormTool(
  client: DataverseClient,
  cache: FormCache,
  formId: string,
  sectionName: string,
  targetTabName: string,
  autoCommit: boolean,
) {
  const { cloneId, cloneName, wasCreated } = await ensureClone(client, formId);
  const xml = await getXml(client, cache, cloneId);

  const { match: tabMatch, candidates: tabCandidates } = findTabCandidates(xml, targetTabName);
  if (!tabMatch) return tabNotFound(targetTabName, tabCandidates);
  const resolvedTab = String(tabMatch["@_name"] ?? targetTabName);

  const updatedXml = moveSectionOnForm(xml, sectionName, resolvedTab);
  const status = await stageOrCommit(client, cache, cloneId, cloneName, updatedXml, autoCommit);
  const cloneNote = wasCreated ? ` (created clone '${cloneName}')` : ` (on clone '${cloneName}')`;
  return { success: true, formId: cloneId, staged: !autoCommit, message: `Section '${sectionName}' moved to tab '${resolvedTab}'${cloneNote} [${status}]` };
}

export async function publishCustomisations(client: DataverseClient) {
  const entity = client.targetEntity;
  const xml = `<importexportxml><entities><entity>${entity}</entity></entities></importexportxml>`;
  // Retry up to 5 times if another solution operation is running (0x80071151 / HTTP 429)
  for (let attempt = 1; attempt <= 5; attempt++) {
    try {
      await client.publishXml(xml);
      return { success: true, message: `${entity} entity customisations published` };
    } catch (err: unknown) {
      const code = (err as { response?: { data?: { error?: { code?: string } } } })?.response?.data?.error?.code ?? "";
      const status = (err as { response?: { status?: number } })?.response?.status;
      if ((code === "0x80071151" || status === 429) && attempt < 5) {
        const wait = attempt * 5000;
        process.stderr.write(`[mcp] publish blocked by concurrent operation, retrying in ${wait / 1000}s (${attempt}/5)...\n`);
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      throw err;
    }
  }
  return { success: true, message: `${entity} entity customisations published` };
}
