import { XMLParser, XMLBuilder } from "fast-xml-parser";
import { v4 as uuidv4 } from "uuid";

// ---------------------------------------------------------------
// Dataverse attribute type → form control classid mapping
// ---------------------------------------------------------------
const CLASSID: Record<string, string> = {
  String: "{4273B735-4538-468D-8964-D0ED97F9CD56}",
  Memo: "{E0DECE4B-6FC8-4A8F-A065-082708572369}",
  Integer: "{C6D124CA-7EDA-4A60-AEA9-7FB8A3CF1DB4}",
  BigInt: "{C6D124CA-7EDA-4A60-AEA9-7FB8A3CF1DB4}",
  Decimal: "{C3EFE0C3-0EC6-42BE-8349-CBD9079C4B45}",
  Double: "{0D2C745A-E5A8-4C8F-BA63-C6D3BB604C15}",
  Money: "{533B9E00-756B-4312-95A0-DC888637AC78}",
  Boolean: "{67FAC785-CD58-4F9F-ABB3-4B7DDC6ED5ED}",
  DateTime: "{5B773807-9FB2-42DB-97C3-7A91EFF8ADFF}",
  Lookup: "{F3015350-44A2-4AA0-97B5-00166532B5E9}",
  Customer: "{270BD3DB-D9AF-4782-9025-509E298DEC0A}",
  Owner: "{270BD3DB-D9AF-4782-9025-509E298DEC0A}",
  Picklist: "{3EF39988-22BB-4F0B-BBBE-64B5A3748AEE}",
  State: "{5D68B988-0661-4DB2-BC3E-17598AD3BE6C}",
  Status: "{5D68B988-0661-4DB2-BC3E-17598AD3BE6C}",
  Virtual: "{4AA28AB7-9C13-4F57-A73D-AD894D048B5F}", // MultiSelectPicklist
  Uniqueidentifier: "{4273B735-4538-468D-8964-D0ED97F9CD56}",
  EntityName: "{4273B735-4538-468D-8964-D0ED97F9CD56}",
};

/** Return the control classid for a Dataverse attribute type (defaults to text). */
export function getClassIdForAttributeType(attributeType: string): string {
  return CLASSID[attributeType] ?? CLASSID["String"];
}

// ---------------------------------------------------------------
// XML parser/builder config
// ---------------------------------------------------------------
const ARRAY_TAGS = new Set(["tab", "column", "section", "row", "cell", "control", "label"]);

const PARSER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  isArray: (name: string) => ARRAY_TAGS.has(name),
  parseAttributeValue: false,
  trimValues: true,
};

const BUILDER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

// ---------------------------------------------------------------
// Parse / build helpers
// ---------------------------------------------------------------

export function parseFormXml(xml: string): Record<string, unknown> {
  return new XMLParser(PARSER_OPTS).parse(xml) as Record<string, unknown>;
}

export function buildFormXml(formObj: Record<string, unknown>): string {
  return new XMLBuilder(BUILDER_OPTS).build(formObj) as string;
}

/** Always return an array, even when fast-xml-parser collapsed it to a single object.
 *  Empty strings (produced by fast-xml-parser for empty XML elements) are filtered out. */
function toArray<T>(value: T | T[] | undefined | null): T[] {
  if (Array.isArray(value)) return value.filter((v) => v !== null && v !== undefined && v !== "") as T[];
  if (value === undefined || value === null || (value as unknown) === "") return [];
  return [value];
}

function newGuid(): string {
  return `{${uuidv4().toUpperCase()}}`;
}

// ---------------------------------------------------------------
// Form manipulation API
// ---------------------------------------------------------------

/** Add a new tab to the form. */
export function addTabToForm(
  formXml: string,
  tabName: string,
  tabLabel: string,
  languageCode = 1033,
): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  if (!root) throw new Error("Invalid form XML: missing <form> root");

  if (!root["tabs"]) root["tabs"] = { tab: [] };
  const tabs = root["tabs"] as Record<string, unknown>;
  tabs["tab"] = toArray(tabs["tab"]);

  (tabs["tab"] as unknown[]).push({
    "@_name": tabName,
    "@_id": newGuid(),
    "@_showlabel": "true",
    "@_locklevel": "0",
    "@_expanded": "true",
    labels: { label: [{ "@_description": tabLabel, "@_languagecode": String(languageCode) }] },
    columns: {
      column: [{ "@_width": "100%", sections: { section: [] } }],
    },
  });

  return buildFormXml(form);
}

/** Add a new section to an existing tab. */
export function addSectionToForm(
  formXml: string,
  tabName: string,
  sectionName: string,
  sectionLabel: string,
  languageCode = 1033,
): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  const tabs = toArray((root?.["tabs"] as Record<string, unknown>)?.["tab"]);
  const tab = tabs.find((t) => (t as Record<string, unknown>)["@_name"] === tabName);
  if (!tab) throw new Error(`Tab '${tabName}' not found`);

  const tabRec = tab as Record<string, unknown>;
  if (!tabRec["columns"]) tabRec["columns"] = { column: [] };
  const columns = toArray(
    ((tabRec["columns"] as Record<string, unknown>)["column"]),
  ) as Record<string, unknown>[];

  if (columns.length === 0) {
    const newCol: Record<string, unknown> = { "@_width": "100%", sections: { section: [] } };
    (tabRec["columns"] as Record<string, unknown>)["column"] = [newCol];
    columns.push(newCol);
  }

  const col = columns[0];
  if (!col["sections"]) col["sections"] = { section: [] };
  const sectionsWrapper = col["sections"] as Record<string, unknown>;
  sectionsWrapper["section"] = toArray(sectionsWrapper["section"]);

  (sectionsWrapper["section"] as unknown[]).push({
    "@_name": sectionName,
    "@_showlabel": "true",
    "@_locklevel": "0",
    "@_id": newGuid(),
    "@_IsUserDefined": "1",
    "@_layout": "varwidth",
    "@_showbar": "false",
    labels: {
      label: [{ "@_description": sectionLabel, "@_languagecode": String(languageCode) }],
    },
    rows: { row: [] },
  });

  return buildFormXml(form);
}

/** Add a field control to an existing tab/section. */
export function addFieldToForm(
  formXml: string,
  tabName: string,
  sectionName: string,
  fieldName: string,
  classId: string,
  fieldLabel: string,
  languageCode = 1033,
): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  const tabs = toArray((root?.["tabs"] as Record<string, unknown>)?.["tab"]) as Record<
    string,
    unknown
  >[];

  const tab = tabs.find((t) => t["@_name"] === tabName);
  if (!tab) throw new Error(`Tab '${tabName}' not found`);

  const columns = toArray(
    ((tab["columns"] as Record<string, unknown>)?.["column"]),
  ) as Record<string, unknown>[];

  let section: Record<string, unknown> | undefined;
  for (const col of columns) {
    const secs = toArray((col["sections"] as Record<string, unknown>)?.["section"]) as Record<
      string,
      unknown
    >[];
    section = secs.find((s) => s["@_name"] === sectionName);
    if (section) break;
  }
  if (!section) throw new Error(`Section '${sectionName}' not found in tab '${tabName}'`);

  if (!section["rows"]) section["rows"] = { row: [] };
  const rows = (section["rows"] as Record<string, unknown>)["row"] as unknown[];
  if (!Array.isArray(rows)) {
    (section["rows"] as Record<string, unknown>)["row"] =
      rows !== undefined && rows !== null ? [rows] : [];
  }

  const rowArr = (section["rows"] as Record<string, unknown>)["row"] as unknown[];
  rowArr.push({
    cell: [
      {
        "@_id": newGuid(),
        "@_showlabel": "true",
        "@_locklevel": "0",
        labels: {
          label: [{ "@_description": fieldLabel, "@_languagecode": String(languageCode) }],
        },
        control: [
          {
            "@_id": fieldName,
            "@_classid": classId,
            "@_datafieldname": fieldName,
            "@_disabled": "false",
            "@_uniqueid": newGuid(),
          },
        ],
      },
    ],
  });

  return buildFormXml(form);
}

/** Remove all controls for a field from the form. */
export function removeFieldFromForm(formXml: string, fieldName: string): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  const tabs = toArray((root?.["tabs"] as Record<string, unknown>)?.["tab"]) as Record<
    string,
    unknown
  >[];

  for (const tab of tabs) {
    for (const col of toArray(
      ((tab["columns"] as Record<string, unknown>)?.["column"]),
    ) as Record<string, unknown>[]) {
      for (const sec of toArray(
        ((col["sections"] as Record<string, unknown>)?.["section"]),
      ) as Record<string, unknown>[]) {
        const rows = toArray(
          ((sec["rows"] as Record<string, unknown>)?.["row"]),
        ) as Record<string, unknown>[];

        for (const row of rows) {
          const cells = toArray(row["cell"]) as Record<string, unknown>[];
          row["cell"] = cells.filter((cell) => {
            const controls = toArray(cell["control"]) as Record<string, unknown>[];
            return !controls.some(
              (c) => c["@_datafieldname"] === fieldName || c["@_id"] === fieldName,
            );
          });
        }

        // Remove rows that became empty
        (sec["rows"] as Record<string, unknown>)["row"] = rows.filter(
          (r) => toArray(r["cell"]).length > 0,
        );
      }
    }
  }

  return buildFormXml(form);
}

/** Set disabled / visible on an existing field control in the form. */
export function setFieldProperties(
  formXml: string,
  fieldName: string,
  props: { disabled?: boolean; visible?: boolean },
): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  const tabs = toArray((root?.["tabs"] as Record<string, unknown>)?.["tab"]) as Record<
    string,
    unknown
  >[];
  let found = false;

  for (const tab of tabs) {
    for (const col of toArray(
      ((tab["columns"] as Record<string, unknown>)?.["column"]),
    ) as Record<string, unknown>[]) {
      for (const sec of toArray(
        ((col["sections"] as Record<string, unknown>)?.["section"]),
      ) as Record<string, unknown>[]) {
        for (const row of toArray(
          ((sec["rows"] as Record<string, unknown>)?.["row"]),
        ) as Record<string, unknown>[]) {
          for (const cell of toArray(row["cell"]) as Record<string, unknown>[]) {
            const controls = toArray(cell["control"]) as Record<string, unknown>[];
            for (const ctrl of controls) {
              if (ctrl["@_datafieldname"] === fieldName || ctrl["@_id"] === fieldName) {
                if (props.disabled !== undefined) ctrl["@_disabled"] = String(props.disabled);
                if (props.visible !== undefined) cell["@_visible"] = String(props.visible);
                found = true;
              }
            }
          }
        }
      }
    }
  }

  if (!found) throw new Error(`Field '${fieldName}' not found in form`);
  return buildFormXml(form);
}

/** Move a field to a different tab/section. */
export function moveFieldOnForm(
  formXml: string,
  fieldName: string,
  targetTabName: string,
  targetSectionName: string,
): string {
  const form = parseFormXml(formXml);
  const root = form["form"] as Record<string, unknown>;
  const tabs = toArray((root?.["tabs"] as Record<string, unknown>)?.["tab"]) as Record<
    string,
    unknown
  >[];

  // 1. Extract the cell
  let extracted: Record<string, unknown> | undefined;

  outer: for (const tab of tabs) {
    for (const col of toArray(
      ((tab["columns"] as Record<string, unknown>)?.["column"]),
    ) as Record<string, unknown>[]) {
      for (const sec of toArray(
        ((col["sections"] as Record<string, unknown>)?.["section"]),
      ) as Record<string, unknown>[]) {
        const rows = toArray(
          ((sec["rows"] as Record<string, unknown>)?.["row"]),
        ) as Record<string, unknown>[];

        for (const row of rows) {
          const cells = toArray(row["cell"]) as Record<string, unknown>[];
          const idx = cells.findIndex((cell) => {
            const controls = toArray(cell["control"]) as Record<string, unknown>[];
            return controls.some(
              (c) => c["@_datafieldname"] === fieldName || c["@_id"] === fieldName,
            );
          });
          if (idx !== -1) {
            [extracted] = cells.splice(idx, 1);
            row["cell"] = cells;
            break outer;
          }
        }

        (sec["rows"] as Record<string, unknown>)["row"] = rows.filter(
          (r) => toArray(r["cell"]).length > 0,
        );
      }
    }
  }

  if (!extracted) throw new Error(`Field '${fieldName}' not found in form`);

  // 2. Add to target section
  const targetTab = tabs.find((t) => t["@_name"] === targetTabName);
  if (!targetTab) throw new Error(`Target tab '${targetTabName}' not found`);

  let targetSec: Record<string, unknown> | undefined;
  for (const col of toArray(
    ((targetTab["columns"] as Record<string, unknown>)?.["column"]),
  ) as Record<string, unknown>[]) {
    const secs = toArray((col["sections"] as Record<string, unknown>)?.["section"]) as Record<
      string,
      unknown
    >[];
    targetSec = secs.find((s) => s["@_name"] === targetSectionName);
    if (targetSec) break;
  }
  if (!targetSec) throw new Error(`Target section '${targetSectionName}' not found`);

  if (!targetSec["rows"]) targetSec["rows"] = { row: [] };
  const rowsWrapper = targetSec["rows"] as Record<string, unknown>;
  if (!Array.isArray(rowsWrapper["row"])) {
    rowsWrapper["row"] = rowsWrapper["row"] !== undefined ? [rowsWrapper["row"]] : [];
  }
  (rowsWrapper["row"] as unknown[]).push({ cell: [extracted] });

  return buildFormXml(form);
}
