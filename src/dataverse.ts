import axios from "axios";
import type {
  AppModuleInfo,
  AttributeMetadata,
  AttributeFullMetadata,
  BusinessRuleInfo,
  DynamicsConfig,
  FormInfo,
  PublisherInfo,
  SavedQueryInfo,
  SolutionInfo,
} from "./types.js";
import { getAccessToken } from "./auth.js";

/** Thin wrapper around the Dataverse Web API v9.2 */
export class DataverseClient {
  private readonly baseUrl: string;

  constructor(private readonly config: DynamicsConfig) {
    this.baseUrl = `${config.environmentUrl.replace(/\/+$/, "")}/api/data/v9.2`;
  }

  get targetEntity(): string {
    return this.config.targetEntity ?? "account";
  }

  get targetAppUniqueName(): string | undefined {
    return this.config.targetAppUniqueName || undefined;
  }

  get solutionUniqueName(): string {
    return this.config.solutionUniqueName || "MCPautoSetup";
  }

  // ------------------------------------------------------------------
  // Internals
  // ------------------------------------------------------------------

  private async headers(): Promise<Record<string, string>> {
    const token = await getAccessToken(this.config);
    return {
      Authorization: `Bearer ${token}`,
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
      "Content-Type": "application/json",
    };
  }

  /**
   * Retry an async operation when Dataverse returns the operation-lock 429
   * (code 0x80071151 – "another [Import] running").
   * Retries up to maxAttempts times with exponential backoff starting at delayMs.
   */
  private async withRetryOnLock<T>(
    fn: () => Promise<T>,
    maxAttempts = 5,
    delayMs = 4000,
  ): Promise<T> {
    let attempt = 0;
    while (true) {
      try {
        return await fn();
      } catch (err) {
        const code = (err as { response?: { data?: { error?: { code?: string } } } })
          ?.response?.data?.error?.code;
        attempt++;
        if (code === "0x80071151" && attempt < maxAttempts) {
          const wait = delayMs * Math.pow(2, attempt - 1);
          process.stderr.write(
            `[mcp] operation lock (0x80071151), retrying in ${wait}ms (attempt ${attempt}/${maxAttempts})\n`,
          );
          await new Promise((resolve) => setTimeout(resolve, wait));
        } else {
          throw err;
        }
      }
    }
  }

  /** Extract a Dataverse entity ID from the OData-EntityId response header. */
  private extractId(locationHeader: string | undefined): string {
    const match = locationHeader?.match(/\(([^)]+)\)/);
    if (!match) throw new Error("Could not extract entity ID from response headers");
    return match[1];
  }

  // ------------------------------------------------------------------
  // Forms (systemform)
  // ------------------------------------------------------------------

  /** List all Main (type=2) forms for the configured target entity. */
  async listAccountForms(): Promise<FormInfo[]> {
    const h = await this.headers();
    const entity = this.config.targetEntity ?? "account";
    const res = await axios.get<{ value: FormInfo[] }>(
      `${this.baseUrl}/systemforms` +
        `?$filter=objecttypecode eq '${entity}' and type eq 2` +
        `&$select=formid,formidunique,name,description,type`,
      { headers: h },
    );
    return res.data.value;
  }

  /** Fetch a single form including its formxml. */
  async getFormById(formId: string): Promise<FormInfo> {
    const h = await this.headers();
    const res = await axios.get<FormInfo>(
      `${this.baseUrl}/systemforms(${formId})?$select=formid,formidunique,name,description,type,formxml`,
      { headers: h },
    );
    return res.data;
  }

  /** Overwrite the formxml of an existing form. */
  async updateFormXml(formId: string, formXml: string): Promise<void> {
    const h = await this.headers();
    await axios.patch(
      `${this.baseUrl}/systemforms(${formId})`,
      { formxml: formXml },
      { headers: h },
    );
  }

  /** Find a Main form for the configured target entity by exact name, or null if not found. */
  async findFormByName(name: string): Promise<FormInfo | null> {
    const h = await this.headers();
    const entity = this.config.targetEntity ?? "account";
    const escaped = name.replace(/'/g, "''"); // OData single-quote escaping only — do NOT URL-encode
    const res = await axios.get<{ value: FormInfo[] }>(
      `${this.baseUrl}/systemforms` +
        `?$filter=objecttypecode eq '${entity}' and type eq 2 and name eq '${escaped}'` +
        `&$select=formid,formidunique,name,description,type`,
      { headers: h },
    );
    return res.data.value[0] ?? null;
  }

  /** Create a new systemform and return its FormInfo (includes formidunique).
   *  Pass solutionUniqueName to register the form in that solution at creation time
   *  via the MSCRM.SolutionUniqueName header — avoids a separate AddSolutionComponent call. */
  async createForm(
    data: {
      name: string;
      objecttypecode: string;
      type: number;
      formxml: string;
      description?: string;
    },
    solutionUniqueName?: string,
  ): Promise<import("./types.js").FormInfo> {
    const h = await this.headers();
    h["Prefer"] = "return=representation";
    if (solutionUniqueName) {
      h["MSCRM.SolutionUniqueName"] = solutionUniqueName;
    }
    const res = await axios.post(`${this.baseUrl}/systemforms`, data, { headers: h });
    // With return=representation Dataverse returns 201 + full entity body
    if (res.data?.formid) {
      return res.data as import("./types.js").FormInfo;
    }
    // Fallback: extract from OData-EntityId / Location header
    const formid = this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
    // Re-fetch to get formidunique
    return this.getFormById(formid);
  }

  // ------------------------------------------------------------------
  // Attribute metadata
  // ------------------------------------------------------------------

  /** Return the MetadataId (GUID) for an entity by its logical name, or null if not found. */
  async getEntityMetadataId(entityLogicalName: string): Promise<string | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ MetadataId: string }>(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=MetadataId`,
        { headers: h },
      );
      return res.data.MetadataId ?? null;
    } catch {
      return null;
    }
  }

  /** Return the MetadataId (GUID) for a field on an entity, or null if not found. */
  async getAttributeMetadataId(entityLogicalName: string, fieldLogicalName: string): Promise<string | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ value: Array<{ MetadataId: string }> }>(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
          `?$filter=LogicalName eq '${fieldLogicalName}'&$select=MetadataId`,
        { headers: h },
      );
      return res.data.value[0]?.MetadataId ?? null;
    } catch {
      return null;
    }
  }

  /** Return attribute metadata for a single field on an entity, or null if not found. */
  async getAttributeMetadata(
    entityName: string,
    fieldName: string,
  ): Promise<AttributeMetadata | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ value: AttributeMetadata[] }>(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityName}')/Attributes` +
          `?$filter=LogicalName eq '${fieldName}'` +
          `&$select=LogicalName,AttributeType,DisplayName`,
        { headers: h },
      );
      return res.data.value[0] ?? null;
    } catch {
      return null;
    }
  }

  // ------------------------------------------------------------------
  // Solutions & publishers
  // ------------------------------------------------------------------

  async getSolution(uniqueName: string): Promise<SolutionInfo | null> {
    const h = await this.headers();
    const res = await axios.get<{ value: SolutionInfo[] }>(
      `${this.baseUrl}/solutions` +
        `?$filter=uniquename eq '${uniqueName}'` +
        `&$select=solutionid,uniquename,friendlyname`,
      { headers: h },
    );
    return res.data.value[0] ?? null;
  }

  async getPublisher(uniqueName: string): Promise<PublisherInfo | null> {
    const h = await this.headers();
    const res = await axios.get<{ value: PublisherInfo[] }>(
      `${this.baseUrl}/publishers` +
        `?$filter=uniquename eq '${uniqueName}'` +
        `&$select=publisherid,uniquename,friendlyname,customizationprefix`,
      { headers: h },
    );
    return res.data.value[0] ?? null;
  }

  async createPublisher(data: {
    uniquename: string;
    friendlyname: string;
    customizationprefix: string;
  }): Promise<string> {
    const h = await this.headers();
    const res = await axios.post(`${this.baseUrl}/publishers`, data, { headers: h });
    return this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
  }

  async createSolution(data: {
    uniquename: string;
    friendlyname: string;
    "publisherid@odata.bind": string;
    version: string;
  }): Promise<string> {
    const h = await this.headers();
    const res = await axios.post(`${this.baseUrl}/solutions`, data, { headers: h });
    return this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
  }

  /**
   * Add a component to a solution.
   * componentType 24 = SystemForm, 80 = App Module
   */
  async addComponentToSolution(
    componentId: string,
    componentType: number,
    solutionUniqueName: string,
  ): Promise<void> {
    const h = await this.headers();
    // DoNotIncludeSubcomponents=true: add the component itself only, without pulling in
    // referenced sub-components (fields, libraries etc.) that may not be customizable.
    await axios.post(
      `${this.baseUrl}/AddSolutionComponent`,
      {
        ComponentId: componentId,
        ComponentType: componentType,
        SolutionUniqueName: solutionUniqueName,
        AddRequiredComponents: false,
        DoNotIncludeSubcomponents: true,
      },
      { headers: h },
    ).catch((err) => {
      const errData = (err as { response?: { data?: { error?: { code?: string; message?: string } } } })?.response?.data?.error;
      process.stderr.write(`[mcp] addComponentToSolution error: code=${errData?.code} msg=${errData?.message}\n`);
      throw err;
    });
  }

  /** Fetch iscustomizable and ismanaged flags for a systemform — used to diagnose AddSolutionComponent failures. */
  async getFormFlags(formId: string): Promise<{ iscustomizable: boolean; ismanaged: boolean } | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ iscustomizable: { Value: boolean } | boolean; ismanaged: boolean }>(
        `${this.baseUrl}/systemforms(${formId})?$select=iscustomizable,ismanaged`,
        { headers: h },
      );
      const raw = res.data.iscustomizable;
      const customizable = typeof raw === "boolean" ? raw : (raw as { Value: boolean }).Value;
      return { iscustomizable: customizable, ismanaged: res.data.ismanaged };
    } catch {
      return null;
    }
  }

  /** Query solutioncomponents to find which solutions already track a given form (componenttype 24). */
  async getFormSolutionComponents(formId: string): Promise<Array<{ solutionid: string; objectid: string }>> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ value: Array<{ solutionid: string; objectid: string }> }>(
        `${this.baseUrl}/solutioncomponents?$filter=objectid eq ${formId} and componenttype eq 24&$select=solutionid,objectid`,
        { headers: h },
      );
      return res.data.value;
    } catch {
      return [];
    }
  }

  /** Delete a systemform record. Used to remove orphaned [MCP] clones that have no solution tracking. */
  async deleteForm(formId: string): Promise<void> {
    const h = await this.headers();
    await axios.delete(`${this.baseUrl}/systemforms(${formId})`, { headers: h });
  }

  // ------------------------------------------------------------------
  // App modules (model-driven apps)
  // ------------------------------------------------------------------

  /** List all model-driven apps in the environment. */
  async listAppModules(): Promise<AppModuleInfo[]> {
    const h = await this.headers();
    const res = await axios.get<{ value: AppModuleInfo[] }>(
      `${this.baseUrl}/appmodules?$select=appmoduleid,uniquename,name,description`,
      { headers: h },
    );
    return res.data.value;
  }

  /** Get a single app module by unique name, or null if not found. */
  async getAppModuleByUniqueName(uniqueName: string): Promise<AppModuleInfo | null> {
    const h = await this.headers();
    const res = await axios.get<{ value: AppModuleInfo[] }>(
      `${this.baseUrl}/appmodules?$filter=uniquename eq '${uniqueName}'&$select=appmoduleid,uniquename,name,description`,
      { headers: h },
    );
    return res.data.value[0] ?? null;
  }

  /**
   * Add a form to a model-driven app's entity form list.
   * Uses the appmodulecomponents endpoint (PUT to add a component).
   */
  async addFormToAppModule(appModuleId: string, formId: string): Promise<void> {
    const h = await this.headers();
    // appmodulecomponents uses a different payload — component reference via @odata.id
    await axios.post(
      `${this.baseUrl}/appmodules(${appModuleId})/appmodulecomponents`,
      {
        "objectid@odata.bind": `/systemforms(${formId})`,
        componenttype: 24, // SystemForm
      },
      { headers: h },
    );
  }

  // ------------------------------------------------------------------
  // Publishing
  // ------------------------------------------------------------------

  /** Publish a targeted set of entities. */
  async publishXml(parameterXml: string): Promise<void> {
    const h = await this.headers();
    await axios.post(
      `${this.baseUrl}/PublishXml`,
      { ParameterXml: parameterXml },
      { headers: h },
    );
  }

  /** Publish all pending customisations. */
  async publishAllXml(): Promise<void> {
    const h = await this.headers();
    await axios.post(`${this.baseUrl}/PublishAllXml`, {}, { headers: h });
  }

  // ------------------------------------------------------------------
  // Attribute metadata – create / update
  // ------------------------------------------------------------------

  /** Derive PascalCase SchemaName from a logical name (e.g. fmk_myfield → fmk_MyField). */
  private toSchemaName(logicalName: string): string {
    const parts = logicalName.split("_");
    if (parts.length === 1) return logicalName.charAt(0).toUpperCase() + logicalName.slice(1);
    return parts[0] + "_" + parts.slice(1).map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join("_");
  }

  /** Build the standard OData Label structure. */
  private makeLabel(text: string, languageCode = 1033) {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.Label",
      LocalizedLabels: [
        {
          "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
          Label: text,
          LanguageCode: languageCode,
        },
      ],
      UserLocalizedLabel: {
        "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
        Label: text,
        LanguageCode: languageCode,
      },
    };
  }

  /** Build the standard RequiredLevel structure. */
  private makeRequiredLevel(level: "None" | "Recommended" | "Required") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.AttributeRequiredLevelManagedProperty",
      Value: level,
      CanBeChanged: true,
      ManagedPropertyLogicalName: "canmodifyrequirementlevelsettings",
    };
  }

  /**
   * Create a simple (non-choice) attribute on an entity via the Metadata API.
   * Supported fieldTypes: String, Memo, Integer, Decimal, Double, Boolean, DateTime.
   */
  async createSimpleAttribute(
    entityLogicalName: string,
    logicalName: string,
    displayName: string,
    fieldType: "String" | "Memo" | "Integer" | "Decimal" | "Double" | "Boolean" | "DateTime",
    options: {
      maxLength?: number;
      requiredLevel?: "None" | "Recommended" | "Required";
      description?: string;
      minValue?: number;
      maxValue?: number;
      dateTimeFormat?: "DateOnly" | "DateAndTime";
      trueLabel?: string;
      falseLabel?: string;
    } = {},
    solutionUniqueName?: string,
  ): Promise<string> {
    const h = await this.headers();
    if (solutionUniqueName) h["MSCRM.SolutionUniqueName"] = solutionUniqueName;

    const schemaName = this.toSchemaName(logicalName);
    const requiredLevel = this.makeRequiredLevel(options.requiredLevel ?? "None");
    const displayNameLabel = this.makeLabel(displayName);
    const descriptionLabel = options.description ? this.makeLabel(options.description) : this.makeLabel("");

    let body: Record<string, unknown>;

    switch (fieldType) {
      case "String":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
          AttributeType: "String",
          AttributeTypeName: { Value: "StringType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          MaxLength: options.maxLength ?? 100,
          FormatName: { Value: "Text" },
        };
        break;
      case "Memo":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
          AttributeType: "Memo",
          AttributeTypeName: { Value: "MemoType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          MaxLength: options.maxLength ?? 2000,
        };
        break;
      case "Integer":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
          AttributeType: "Integer",
          AttributeTypeName: { Value: "IntegerType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          Format: "None",
          MinValue: options.minValue ?? -2147483648,
          MaxValue: options.maxValue ?? 2147483647,
        };
        break;
      case "Decimal":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
          AttributeType: "Decimal",
          AttributeTypeName: { Value: "DecimalType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          Precision: 2,
          MinValue: options.minValue ?? -100000000000,
          MaxValue: options.maxValue ?? 100000000000,
        };
        break;
      case "Double":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.DoubleAttributeMetadata",
          AttributeType: "Double",
          AttributeTypeName: { Value: "DoubleType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          Precision: 5,
          MinValue: options.minValue ?? -100000000000,
          MaxValue: options.maxValue ?? 100000000000,
        };
        break;
      case "Boolean":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
          AttributeType: "Boolean",
          AttributeTypeName: { Value: "BooleanType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          OptionSet: {
            "@odata.type": "Microsoft.Dynamics.CRM.BooleanOptionSetMetadata",
            IsGlobal: false,
            OptionSetType: "Boolean",
            TrueOption: {
              Value: 1,
              Label: this.makeLabel(options.trueLabel ?? "Yes"),
            },
            FalseOption: {
              Value: 0,
              Label: this.makeLabel(options.falseLabel ?? "No"),
            },
          },
        };
        break;
      case "DateTime":
        body = {
          "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
          AttributeType: "DateTime",
          AttributeTypeName: { Value: "DateTimeType" },
          SchemaName: schemaName,
          LogicalName: logicalName,
          DisplayName: displayNameLabel,
          Description: descriptionLabel,
          RequiredLevel: requiredLevel,
          Format: options.dateTimeFormat ?? "DateOnly",
          DateTimeBehavior: { Value: "UserLocal" },
        };
        break;
    }

    const res = await this.withRetryOnLock(() =>
      axios.post(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes`,
        body,
        { headers: h },
      )
    );
    return this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
  }

  /**
   * Create a Picklist (single choice) or MultiSelectPicklist (multi choice) attribute.
   * choices: array of { value: number, label: string } — use values >= 100000000 for custom options.
   */
  async createPicklistAttribute(
    entityLogicalName: string,
    logicalName: string,
    displayName: string,
    choices: Array<{ value: number; label: string }>,
    options: {
      defaultValue?: number;
      isMultiSelect?: boolean;
      requiredLevel?: "None" | "Recommended" | "Required";
      description?: string;
    } = {},
    solutionUniqueName?: string,
  ): Promise<string> {
    const h = await this.headers();
    if (solutionUniqueName) h["MSCRM.SolutionUniqueName"] = solutionUniqueName;

    const schemaName = this.toSchemaName(logicalName);
    const isMulti = options.isMultiSelect ?? false;

    const body: Record<string, unknown> = {
      "@odata.type": isMulti
        ? "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata"
        : "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      AttributeType: isMulti ? "Virtual" : "Picklist",
      AttributeTypeName: { Value: isMulti ? "MultiSelectPicklistType" : "PicklistType" },
      SchemaName: schemaName,
      LogicalName: logicalName,
      DisplayName: this.makeLabel(displayName),
      Description: options.description ? this.makeLabel(options.description) : this.makeLabel(""),
      RequiredLevel: this.makeRequiredLevel(options.requiredLevel ?? "None"),
      OptionSet: {
        "@odata.type": "Microsoft.Dynamics.CRM.OptionSetMetadata",
        IsGlobal: false,
        OptionSetType: "Picklist",
        DisplayName: this.makeLabel(displayName + " Options"),
        Options: choices.map((c) => ({
          Value: c.value,
          Label: this.makeLabel(c.label),
        })),
      },
    };

    if (!isMulti && options.defaultValue !== undefined) {
      body["DefaultFormValue"] = options.defaultValue;
    }

    const res = await this.withRetryOnLock(() =>
      axios.post(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes`,
        body,
        { headers: h },
      )
    );
    return this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
  }

  /** Fetch extended attribute metadata including MetadataId and AttributeTypeName. */
  async getAttributeFullMetadata(
    entityLogicalName: string,
    fieldLogicalName: string,
  ): Promise<AttributeFullMetadata | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ value: AttributeFullMetadata[] }>(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes` +
          `?$filter=LogicalName eq '${fieldLogicalName}'` +
          `&$select=MetadataId,LogicalName,AttributeType,AttributeTypeName,RequiredLevel,DisplayName`,
        { headers: h },
      );
      return res.data.value[0] ?? null;
    } catch {
      return null;
    }
  }

  /** Map an AttributeTypeName.Value to the Dataverse OData type name for PATCH requests. */
  private attributeTypeNameToOdataType(typeName: string): string {
    const map: Record<string, string> = {
      StringType: "Microsoft.Dynamics.CRM.StringAttributeMetadata",
      MemoType: "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
      IntegerType: "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
      DecimalType: "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
      DoubleType: "Microsoft.Dynamics.CRM.DoubleAttributeMetadata",
      MoneyType: "Microsoft.Dynamics.CRM.MoneyAttributeMetadata",
      BooleanType: "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
      DateTimeType: "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
      PicklistType: "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      MultiSelectPicklistType: "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
      LookupType: "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
      StateType: "Microsoft.Dynamics.CRM.StateAttributeMetadata",
      StatusType: "Microsoft.Dynamics.CRM.StatusAttributeMetadata",
      BigIntType: "Microsoft.Dynamics.CRM.BigIntAttributeMetadata",
      UniqueidentifierType: "Microsoft.Dynamics.CRM.UniqueIdentifierAttributeMetadata",
    };
    return map[typeName] ?? "Microsoft.Dynamics.CRM.AttributeMetadata";
  }

  /** Update the RequiredLevel of an existing attribute. */
  async updateAttributeRequirementLevel(
    entityLogicalName: string,
    fieldLogicalName: string,
    requiredLevel: "None" | "Recommended" | "Required",
  ): Promise<void> {
    const h = await this.headers();
    h["MSCRM.MergeLabels"] = "true";

    const full = await this.getAttributeFullMetadata(entityLogicalName, fieldLogicalName);
    if (!full) throw new Error(`Field '${fieldLogicalName}' not found on '${entityLogicalName}'`);

    const odataType = this.attributeTypeNameToOdataType(full.AttributeTypeName.Value);

    await this.withRetryOnLock(() =>
      axios.put(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(${full.MetadataId})`,
        {
          "@odata.type": odataType,
          RequiredLevel: this.makeRequiredLevel(requiredLevel),
        },
        { headers: h },
      )
    );
  }

  // ------------------------------------------------------------------
  // Views (savedqueries)
  // ------------------------------------------------------------------

  /** Get the ObjectTypeCode for an entity (needed to query savedqueries). */
  async getEntityObjectTypeCode(entityLogicalName: string): Promise<number | null> {
    const h = await this.headers();
    try {
      const res = await axios.get<{ ObjectTypeCode: number }>(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=ObjectTypeCode`,
        { headers: h },
      );
      return res.data.ObjectTypeCode ?? null;
    } catch {
      return null;
    }
  }

  /**
   * List views (savedqueries) for an entity.
   * queryType: 0 = Public View (default), 1 = Advanced Find, 4 = Quick Find, 2 = Associated.
   */
  async listViews(entityLogicalName: string, queryType = 0): Promise<SavedQueryInfo[]> {
    const h = await this.headers();
    const typeCode = await this.getEntityObjectTypeCode(entityLogicalName);
    if (!typeCode) throw new Error(`Entity '${entityLogicalName}' not found`);

    const res = await axios.get<{ value: SavedQueryInfo[] }>(
      `${this.baseUrl}/savedqueries` +
        `?$filter=returnedtypecode eq ${typeCode} and querytype eq ${queryType} and statecode eq 0` +
        `&$select=savedqueryid,name,querytype,returnedtypecode,statecode,description`,
      { headers: h },
    );
    return res.data.value;
  }

  /** Get a single savedquery including its layoutxml and fetchxml. */
  async getSavedQueryById(viewId: string): Promise<SavedQueryInfo> {
    const h = await this.headers();
    const res = await axios.get<SavedQueryInfo>(
      `${this.baseUrl}/savedqueries(${viewId})?$select=savedqueryid,name,querytype,returnedtypecode,statecode,layoutxml,fetchxml`,
      { headers: h },
    );
    return res.data;
  }

  /** Update a savedquery's layoutxml (and optionally fetchxml). */
  async updateSavedQueryLayoutXml(
    viewId: string,
    layoutXml: string,
    fetchXml?: string,
  ): Promise<void> {
    const h = await this.headers();
    const body: Record<string, string> = { layoutxml: layoutXml };
    if (fetchXml) body["fetchxml"] = fetchXml;
    await axios.patch(`${this.baseUrl}/savedqueries(${viewId})`, body, { headers: h });
  }

  // ------------------------------------------------------------------
  // Business rules (workflows category=2)
  // ------------------------------------------------------------------

  /** List business rules for an entity. */
  async listBusinessRules(entityLogicalName: string): Promise<BusinessRuleInfo[]> {
    const h = await this.headers();
    const escaped = entityLogicalName.replace(/'/g, "''");
    const res = await axios.get<{ value: BusinessRuleInfo[] }>(
      `${this.baseUrl}/workflows` +
        `?$filter=category eq 2 and primaryentity eq '${escaped}'` +
        `&$select=workflowid,name,statecode,statuscode,scope,description,primaryentity`,
      { headers: h },
    );
    return res.data.value;
  }

  /**
   * Create a form Business Rule (category=2, scope=1) for an entity.
   * clientdata must be the JSON string representation of the rule definition.
   * Returns the new workflow GUID.
   */
  async createWorkflowBusinessRule(
    entityLogicalName: string,
    name: string,
    clientdata: string,
    solutionUniqueName?: string,
  ): Promise<string> {
    const h = await this.headers();
    if (solutionUniqueName) h["MSCRM.SolutionUniqueName"] = solutionUniqueName;

    const res = await axios.post(
      `${this.baseUrl}/workflows`,
      {
        name,
        category: 2,
        primaryentity: entityLogicalName,
        type: 1,
        scope: 1,
        statecode: 0,
        clientdata,
      },
      { headers: h },
    );
    return this.extractId(
      (res.headers["odata-entityid"] as string | undefined) ??
        (res.headers["location"] as string | undefined),
    );
  }

  /** Activate or deactivate a business rule. statecode: 0=Draft, 1=Activated */
  async setBusinessRuleState(workflowId: string, statecode: 0 | 1): Promise<void> {
    const h = await this.headers();
    const statuscode = statecode === 1 ? 2 : 1;
    await axios.patch(
      `${this.baseUrl}/workflows(${workflowId})`,
      { statecode, statuscode },
      { headers: h },
    );
  }

  /**
   * Update mutable properties on an existing attribute (display name, description,
   * required level, max length). Fetches full metadata first to determine the OData type.
   */
  async updateAttributeProperties(
    entityLogicalName: string,
    fieldLogicalName: string,
    updates: {
      displayName?: string;
      description?: string;
      requiredLevel?: "None" | "Recommended" | "Required";
      maxLength?: number;
      minValue?: number;
      maxValue?: number;
    },
    solutionUniqueName?: string,
  ): Promise<void> {
    const h = await this.headers();
    h["MSCRM.MergeLabels"] = "true";
    if (solutionUniqueName) h["MSCRM.SolutionUniqueName"] = solutionUniqueName;

    const full = await this.getAttributeFullMetadata(entityLogicalName, fieldLogicalName);
    if (!full) throw new Error(`Field '${fieldLogicalName}' not found on entity '${entityLogicalName}'`);

    const odataType = this.attributeTypeNameToOdataType(full.AttributeTypeName.Value);
    const body: Record<string, unknown> = { "@odata.type": odataType };

    if (updates.displayName !== undefined) {
      body["DisplayName"] = this.makeLabel(updates.displayName);
    }
    if (updates.description !== undefined) {
      body["Description"] = this.makeLabel(updates.description);
    }
    if (updates.requiredLevel !== undefined) {
      body["RequiredLevel"] = this.makeRequiredLevel(updates.requiredLevel);
    }
    if (updates.maxLength !== undefined) {
      body["MaxLength"] = updates.maxLength;
    }
    if (updates.minValue !== undefined) {
      body["MinValue"] = updates.minValue;
    }
    if (updates.maxValue !== undefined) {
      body["MaxValue"] = updates.maxValue;
    }

    await this.withRetryOnLock(() =>
      axios.put(
        `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(${full.MetadataId})`,
        body,
        { headers: h },
      )
    );
  }

  /**
   * Add a new option to an existing local Picklist or MultiSelectPicklist attribute.
   */
  async addOptionToPicklist(
    entityLogicalName: string,
    fieldLogicalName: string,
    optionValue: number,
    optionLabel: string,
    solutionUniqueName?: string,
  ): Promise<void> {
    const h = await this.headers();
    if (solutionUniqueName) h["MSCRM.SolutionUniqueName"] = solutionUniqueName;

    await this.withRetryOnLock(() =>
      axios.post(
        `${this.baseUrl}/InsertOptionValue`,
        {
          EntityLogicalName: entityLogicalName,
          AttributeLogicalName: fieldLogicalName,
          Value: optionValue,
          Label: this.makeLabel(optionLabel),
          MergeLabels: true,
        },
        { headers: h },
      )
    );
  }

  /**
   * List all options for a Picklist or MultiSelectPicklist attribute.
   */
  async getPicklistOptions(
    entityLogicalName: string,
    fieldLogicalName: string,
  ): Promise<Array<{ value: number; label: string }>> {
    const h = await this.headers();
    const res = await axios.get<{
      Options: Array<{ Value: number; Label: { UserLocalizedLabel: { Label: string } } }>;
    }>(
      `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${fieldLogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata/OptionSet`,
      { headers: h },
    );
    return (res.data.Options ?? []).map((o) => ({
      value: o.Value,
      label: o.Label?.UserLocalizedLabel?.Label ?? "",
    }));
  }
}

