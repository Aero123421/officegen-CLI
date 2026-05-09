export interface Relationship {
    id: string;
    type: string;
    target: string;
    targetMode?: string;
}
export declare function parseRelationships(xml: string | undefined): Relationship[];
export declare function relationshipTarget(baseDir: string, target: string): string;
export declare function nextRelationshipId(xml: string): string;
