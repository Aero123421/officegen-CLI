import { inspect } from "./inspect.js";
export async function diagnose(input, options = {}) {
    const inspected = isInspectResult(input) ? input : await inspect(input, { depth: "shallow" });
    const issues = [];
    const maxTextLength = options.maxTextLength ?? 220;
    for (const entry of inspected.objectMap) {
        if ((entry.text?.length ?? 0) > maxTextLength) {
            issues.push({
                code: "TEXT_OVERFLOW_RISK",
                severity: "warning",
                message: "Text object is long enough to risk overflow in approximate rendering.",
                stableObjectId: entry.stableObjectId,
                suggestedOps: [{ type: "setText", selector: { stableObjectId: entry.stableObjectId }, text: `${entry.text?.slice(0, maxTextLength - 1)}…` }]
            });
        }
    }
    if (Number(inspected.trusted.summary.macros ?? 0) > 0) {
        issues.push({
            code: "MACRO_PRESENT",
            severity: "warning",
            message: "Document contains a VBA project. Treat the file as untrusted and avoid executing embedded code."
        });
    }
    if (!inspected.objectMap.length && inspected.trusted.format !== "pdf") {
        issues.push({
            code: "NO_TEXT_OBJECTS",
            severity: "info",
            message: "No editable text objects were detected by the XML inspector."
        });
    }
    return {
        schema: "officegen.diagnose.result@1.2",
        issues,
        caveats: ["Diagnosis is based on approximate inspect/view data and does not execute external renderers."]
    };
}
export const diagnoseDocument = diagnose;
function isInspectResult(value) {
    return Boolean(value && typeof value === "object" && value.schema === "officegen.inspect.result@1.2");
}
//# sourceMappingURL=diagnose.js.map