/**
 * Type definitions for extracted COM API data
 */

export interface ComApi {
    metadata: {
        prog_id: string;
        version: string;
        generator: string;
    };
    classes: {
        [className: string]: ComClass;
    };
}

export interface ComClass {
    properties: {
        [propName: string]: {
            type: string;
            readonly: boolean;
        };
    };
    methods: {
        [methodName: string]: {
            parameters: any[];  // Simplified for MVP
        };
    };
}
