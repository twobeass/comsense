import * as vscode from 'vscode';
import { ComApi, ComClass } from './types';

/**
 * Provides IntelliSense completions for COM APIs
 */
export class ComCompletionProvider implements vscode.CompletionItemProvider {
    constructor(private apis: ComApi[]) {}

    provideCompletionItems(
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.CompletionItem[] {
        const lineText = document.lineAt(position).text;
        const textBeforeCursor = lineText.substring(0, position.character);

        // Check if we're after a dot (e.g., "DemoClass.")
        const dotMatch = textBeforeCursor.match(/([\w]+)\.$/);
        if (dotMatch) {
            const objectName = dotMatch[1];
            return this.getClassMembers(objectName);
        }

        // Otherwise, suggest class names
        return this.getAllClasses();
    }

    private getAllClasses(): vscode.CompletionItem[] {
        const items: vscode.CompletionItem[] = [];
        for (const api of this.apis) {
            for (const className of Object.keys(api.classes)) {
                const item = new vscode.CompletionItem(
                    className,
                    vscode.CompletionItemKind.Class
                );
                item.detail = api.metadata.prog_id;
                items.push(item);
            }
        }
        return items;
    }

    private getClassMembers(className: string): vscode.CompletionItem[] {
        const items: vscode.CompletionItem[] = [];
        for (const api of this.apis) {
            const comClass = api.classes[className];
            if (!comClass) continue;

            // Add properties
            for (const [propName, propInfo] of Object.entries(comClass.properties)) {
                const item = new vscode.CompletionItem(
                    propName,
                    vscode.CompletionItemKind.Property
                );
                item.detail = propInfo.type;
                if (propInfo.readonly) {
                    item.detail += ' (readonly)';
                }
                items.push(item);
            }

            // Add methods
            for (const methodName of Object.keys(comClass.methods)) {
                const item = new vscode.CompletionItem(
                    methodName,
                    vscode.CompletionItemKind.Method
                );
                item.insertText = new vscode.SnippetString(`${methodName}($0)`);
                items.push(item);
            }
        }
        return items;
    }
}
