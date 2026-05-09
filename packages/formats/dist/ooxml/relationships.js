import { ensureArray, parseXml } from "./xml.js";
export function parseRelationships(xml) {
    if (!xml)
        return [];
    const parsed = parseXml(xml);
    return ensureArray(parsed.Relationships?.Relationship)
        .map((rel) => ({
        id: rel["@_Id"] ?? "",
        type: rel["@_Type"] ?? "",
        target: rel["@_Target"] ?? "",
        targetMode: rel["@_TargetMode"]
    }))
        .filter((rel) => rel.id && rel.target);
}
export function relationshipTarget(baseDir, target) {
    if (target.startsWith("/"))
        return target.slice(1);
    const parts = `${baseDir}/${target}`.split("/");
    const resolved = [];
    for (const part of parts) {
        if (!part || part === ".")
            continue;
        if (part === "..")
            resolved.pop();
        else
            resolved.push(part);
    }
    return resolved.join("/");
}
export function nextRelationshipId(xml) {
    const ids = [...xml.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])).filter(Number.isFinite);
    return `rId${Math.max(0, ...ids) + 1}`;
}
//# sourceMappingURL=relationships.js.map