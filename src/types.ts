// ============================================================
// Shared TypeScript interfaces
// ============================================================

export interface DynamicsConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  /** e.g. https://orgname.crm4.dynamics.com  (no trailing slash) */
  environmentUrl: string;
  /** Logical name of the entity to operate on, e.g. 'account' (default: 'account') */
  targetEntity?: string;
  /** Unique name of the model-driven app to scope to, e.g. 'msdyn_SalesHub' */
  targetAppUniqueName?: string;
  /** Unique name of the solution to track changes in (default: MCPautoSetup) */
  solutionUniqueName?: string;
}

export interface TokenInfo {
  accessToken: string;
  /** Unix timestamp in ms when the token expires */
  expiresAt: number;
}

export interface FormInfo {
  formid: string;
  /** Stable cross-environment GUID used by AddSolutionComponent (ComponentType 24) */
  formidunique?: string;
  name: string;
  /** 2 = Main, 0 = Quick Create, 6 = Quick View */
  type: number;
  description?: string;
  formxml?: string;
}

export interface AttributeMetadata {
  LogicalName: string;
  AttributeType: string;
  DisplayName?: {
    UserLocalizedLabel?: { Label?: string };
  };
}

export interface SolutionInfo {
  solutionid: string;
  uniquename: string;
  friendlyname: string;
}

export interface PublisherInfo {
  publisherid: string;
  uniquename: string;
  friendlyname: string;
  customizationprefix: string;
}

export interface AppModuleInfo {
  appmoduleid: string;
  uniquename: string;
  name: string;
  description?: string;
}

/** Per-session connection details (overrides env-var defaults) */
export interface SessionConfig {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  environmentUrl?: string;
  targetEntity?: string;
  targetAppUniqueName?: string;
  solutionUniqueName?: string;
}

/** In-session staging cache: pending XML changes not yet pushed to Dataverse */
export interface FormCacheEntry {
  xml: string;
  cloneName: string;
}
/** Maps cloneId → staged XML. Shared across tool calls within one server instance. */
export type FormCache = Map<string, FormCacheEntry>;

export interface SavedQueryInfo {
  savedqueryid: string;
  name: string;
  querytype: number;
  returnedtypecode: number;
  layoutxml?: string;
  fetchxml?: string;
  statecode: number;
  description?: string;
}

export interface BusinessRuleInfo {
  workflowid: string;
  name: string;
  statecode: number;
  statuscode: number;
  scope: number;
  description?: string;
  primaryentity: string;
}

export interface AttributeFullMetadata {
  MetadataId: string;
  LogicalName: string;
  AttributeType: string;
  AttributeTypeName: { Value: string };
  RequiredLevel: { Value: string };
  DisplayName?: { UserLocalizedLabel?: { Label?: string } };
}
