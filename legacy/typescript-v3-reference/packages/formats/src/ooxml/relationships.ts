import { ensureArray, parseXml } from "./xml.js";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function parseRelationships(xml: string | undefined): Relationship[] {
  if (!xml) return [];
  const parsed = parseXml<{ Relationships?: { Relationship?: unknown } }>(xml);
  return ensureArray(parsed.Relationships?.Relationship as Record<string, string> | Record<string, string>[] | undefined)
    .map((rel) => ({
      id: rel["@_Id"] ?? "",
      type: rel["@_Type"] ?? "",
      target: rel["@_Target"] ?? "",
      targetMode: rel["@_TargetMode"]
    }))
    .filter((rel) => rel.id && rel.target);
}

export function relationshipTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = `${baseDir}/${target}`.split("/");
  const resolved: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") resolved.pop();
    else resolved.push(part);
  }
  return resolved.join("/");
}

export function nextRelationshipId(xml: string): string {
  const ids = [...xml.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])).filter(Number.isFinite);
  return `rId${Math.max(0, ...ids) + 1}`;
}
