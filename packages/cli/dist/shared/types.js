export const CLI_SPEC_VERSION = "1.2";
export const ENVELOPE_SCHEMA = "officegen.envelope@1.2";
export class CliFailure extends Error {
    payload;
    exitCode;
    constructor(payload, exitCode = 1) {
        super(payload.message);
        this.payload = payload;
        this.exitCode = exitCode;
    }
}
//# sourceMappingURL=types.js.map