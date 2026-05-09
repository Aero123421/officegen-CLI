import { Command } from "commander";
import type { RuntimeContext } from "../shared/types.js";
export declare function createProgram(context: RuntimeContext, stdout: (text: string) => void, _stderr: (text: string) => void, now: Date): Command;
export declare function writeNativeHelp(context: RuntimeContext, stdout: (text: string) => void): void;
export declare function writeCommandHelp(context: RuntimeContext, commandGroup: string, subcommand: string | undefined, stdout: (text: string) => void): void;
