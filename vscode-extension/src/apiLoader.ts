import * as fs from 'fs';
import * as path from 'path';
import { ComApi } from './types';

/**
 * Load all COM API JSON files from data/apis directory
 */
export class ApiLoader {
    private apis: Map<string, ComApi> = new Map();

    constructor(private dataDir: string) {}

    load(): void {
        const apisDir = path.join(this.dataDir, 'apis');

        if (!fs.existsSync(apisDir)) {
            console.warn('APIs directory not found:', apisDir);
            return;
        }

        const files = fs.readdirSync(apisDir).filter(f => f.endsWith('.json'));

        for (const file of files) {
            try {
                const filePath = path.join(apisDir, file);
                const content = fs.readFileSync(filePath, 'utf-8');
                const api: ComApi = JSON.parse(content);

                this.apis.set(api.metadata.prog_id, api);
                console.log(`Loaded API: ${api.metadata.prog_id}`);
            } catch (error) {
                console.error(`Failed to load ${file}:`, error);
            }
        }
    }

    getAll(): ComApi[] {
        return Array.from(this.apis.values());
    }

    getByProgId(progId: string): ComApi | undefined {
        return this.apis.get(progId);
    }
}
