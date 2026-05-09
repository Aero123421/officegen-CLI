export declare const VALUE_OPTIONS: Set<string>;
export declare function commandFromArgv(argv: string[]): string;
export declare function getTopCommand(argv: string[]): string | undefined;
export declare function positionalArgs(argv: string[], start: number): string[];
export declare function hasFlag(argv: string[], flag: string): boolean;
export declare function optionValue(argv: string[], name: string): string | undefined;
export declare function secondCommandToken(argv: string[]): string | undefined;
