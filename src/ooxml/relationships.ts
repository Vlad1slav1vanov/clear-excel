import { posix as packagePath } from "node:path";

import type JSZip from "jszip";

import type { ExcelSanitizationLimits, ReadBudget } from "../types.js";
import { readXmlPart, requireXml } from "./parts.js";
import { findElements, getNamespacedAttribute, removeElements, type XmlElement } from "./xml.js";

const PACKAGE_RELATIONSHIP_NAMESPACES = new Set([
  "",
  "http://schemas.openxmlformats.org/package/2006/relationships",
  "http://purl.oclc.org/ooxml/package/relationships",
]);

export interface Relationship {
  id: string;
  target: string;
  targetMode: string | undefined;
  type: string;
}

export type RelationshipGraph = Map<string, Relationship[]>;

export interface RemovedRelationships {
  xml: string;
  targets: Set<string>;
}

export function removeRelationships(
  xml: string,
  source: string,
  relationshipSuffixes: readonly string[],
): RemovedRelationships {
  const targets = new Set<string>();
  const removedIds = new Set<string>();
  for (const relationship of parseRelationships(xml, source)) {
    if (
      relationship.targetMode?.toLowerCase() !== "external" &&
      relationshipSuffixes.some((suffix) => relationship.type.endsWith(suffix))
    ) {
      removedIds.add(relationship.id);
      targets.add(relationship.target);
    }
  }
  const cleaned = removeElements(
    xml,
    (element) =>
      isRelationshipElement(element) &&
      removedIds.has(getNamespacedAttribute(element, "Id", [""]) ?? ""),
  );
  return { xml: cleaned, targets };
}

export function parseRelationships(xml: string, source: string): Relationship[] {
  const ids = new Set<string>();
  return findElements(xml, "Relationship")
    .filter(isRelationshipElement)
    .map((element) => {
      const id = getNamespacedAttribute(element, "Id", [""]);
      const target = getNamespacedAttribute(element, "Target", [""]);
      const type = getNamespacedAttribute(element, "Type", [""]);
      if (!id || !target || !type) {
        throw new Error(`Invalid relationship in ${source || "package root"}`);
      }
      if (ids.has(id)) {
        throw new Error(`Duplicate relationship ID in ${source || "package root"}: ${id}`);
      }
      ids.add(id);
      const targetMode = getNamespacedAttribute(element, "TargetMode", [""]);
      if (
        targetMode !== undefined &&
        !["internal", "external"].includes(targetMode.toLowerCase())
      ) {
        throw new Error(`Invalid relationship target mode: ${targetMode}`);
      }
      return {
        id,
        target: targetMode?.toLowerCase() === "external" ? target : resolveTarget(source, target),
        targetMode,
        type,
      };
    });
}

function isRelationshipElement(element: XmlElement): boolean {
  return (
    element.localName === "Relationship" &&
    PACKAGE_RELATIONSHIP_NAMESPACES.has(element.namespaceUri)
  );
}

export function resolveTarget(source: string, target: string): string {
  const resolved = target.startsWith("/")
    ? packagePath.normalize(target.slice(1))
    : packagePath.normalize(packagePath.join(packagePath.dirname(source), target));
  if (resolved === ".." || resolved.startsWith("../")) {
    throw new Error(`Relationship target escapes package: ${target}`);
  }
  return resolved;
}

export function relationshipFileForPart(partName: string): string {
  return packagePath.join(
    packagePath.dirname(partName),
    "_rels",
    `${packagePath.basename(partName)}.rels`,
  );
}

export function sourcePartForRelationshipFile(fileName: string): string | undefined {
  if (fileName === "_rels/.rels") {
    return "";
  }
  const match = fileName.match(/^(.*\/)_rels\/([^/]+)\.rels$/);
  return match ? `${match[1]}${match[2]}` : undefined;
}

export async function removeUnreachableParts(
  zip: JSZip,
  detachedRoots: Set<string>,
  readBudget: ReadBudget,
  limits: Readonly<ExcelSanitizationLimits>,
): Promise<void> {
  if (detachedRoots.size === 0) {
    return;
  }

  const graph = await buildRelationshipGraph(zip, readBudget, limits);
  const candidates = collectReachableParts(detachedRoots, graph);
  const packageReachable = collectReachableParts(new Set([""]), graph);
  packageReachable.delete("");
  const deletedParts = new Set(
    [...candidates].filter((partName) => !packageReachable.has(partName)),
  );
  for (const partName of deletedParts) {
    zip.remove(partName);
    zip.remove(relationshipFileForPart(partName));
  }

  const contentTypes = await requireXml(
    zip,
    "[Content_Types].xml",
    readBudget,
    limits.maxMetadataXmlBytes,
  );
  zip.file(
    "[Content_Types].xml",
    removeElements(contentTypes, (element) => {
      if (
        element.localName !== "Override" ||
        !["", "http://schemas.openxmlformats.org/package/2006/content-types"].includes(
          element.namespaceUri,
        )
      ) {
        return false;
      }
      const partName = getNamespacedAttribute(element, "PartName", [""])?.replace(/^\//, "");
      return partName !== undefined && deletedParts.has(partName);
    }),
  );
}

async function buildRelationshipGraph(
  zip: JSZip,
  readBudget: ReadBudget,
  limits: Readonly<ExcelSanitizationLimits>,
): Promise<RelationshipGraph> {
  const graph: RelationshipGraph = new Map();
  for (const fileName of Object.keys(zip.files)) {
    if (!fileName.endsWith(".rels")) {
      continue;
    }
    const source = sourcePartForRelationshipFile(fileName);
    const file = zip.file(fileName);
    if (source !== undefined && file) {
      graph.set(
        source,
        parseRelationships(
          // Reads intentionally stay sequential because they share a mutable byte budget.
          // oxlint-disable-next-line no-await-in-loop
          await readXmlPart(file, fileName, readBudget, limits.maxRelationshipsXmlBytes),
          source,
        ),
      );
    }
  }
  return graph;
}

export function collectReachableParts(
  initial: ReadonlySet<string>,
  graph: ReadonlyMap<string, readonly Relationship[]>,
): Set<string> {
  const reachable = new Set(initial);
  const pending = [...initial];
  while (pending.length > 0) {
    const source = pending.pop();
    if (source === undefined) {
      continue;
    }
    for (const relationship of graph.get(source) ?? []) {
      if (
        relationship.targetMode?.toLowerCase() !== "external" &&
        !reachable.has(relationship.target)
      ) {
        reachable.add(relationship.target);
        pending.push(relationship.target);
      }
    }
  }
  return reachable;
}
