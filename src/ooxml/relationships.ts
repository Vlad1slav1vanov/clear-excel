import { posix as packagePath } from "node:path";

import type JSZip from "jszip";

import type { ExcelSanitizationLimits, ReadBudget } from "../types.js";
import { readXmlPart, requireXml } from "./parts.js";
import { getAttribute } from "./xml.js";

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
  const cleaned = xml.replace(
    /<(?:\w+:)?Relationship\b[^>]*?(?:\/\s*>|>\s*<\/(?:\w+:)?Relationship\s*>)/gi,
    (tag) => {
      const type = getAttribute(tag, "Type") ?? "";
      const target = getAttribute(tag, "Target");
      const targetMode = getAttribute(tag, "TargetMode");
      if (
        !target ||
        targetMode?.toLowerCase() === "external" ||
        !relationshipSuffixes.some((suffix) => type.endsWith(suffix))
      ) {
        return tag;
      }
      targets.add(resolveTarget(source, target));
      return "";
    },
  );
  return { xml: cleaned, targets };
}

export function parseRelationships(xml: string, source: string): Relationship[] {
  return Array.from(
    xml.matchAll(/<(?:\w+:)?Relationship\b([^>]*?)(?:\/\s*>|>\s*<\/(?:\w+:)?Relationship\s*>)/gi),
  ).flatMap((match) => {
    const attributes = match[1] ?? "";
    const id = getAttribute(attributes, "Id");
    const target = getAttribute(attributes, "Target");
    const type = getAttribute(attributes, "Type");
    if (!id || !target || !type) {
      return [];
    }
    const targetMode = getAttribute(attributes, "TargetMode");
    return [
      {
        id,
        target: targetMode?.toLowerCase() === "external" ? target : resolveTarget(source, target),
        targetMode,
        type,
      },
    ];
  });
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
    contentTypes.replace(/<(?:\w+:)?Override\b[^>]*\/>/gi, (tag) => {
      const partName = getAttribute(tag, "PartName")?.replace(/^\//, "");
      return partName && deletedParts.has(partName) ? "" : tag;
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
