import * as vscode from 'vscode';
import * as path from 'path';
import { ApiLoader } from './apiLoader';
import { ComCompletionProvider } from './completionProvider';

export function activate(context: vscode.ExtensionContext) {
    console.log('comsense extension activated');

    // Load COM APIs
    const dataDir = path.join(context.extensionPath, 'data');
    const loader = new ApiLoader(dataDir);
    loader.load();

    const apis = loader.getAll();
    console.log(`Loaded ${apis.length} COM APIs`);

    // Register completion provider for VBA/VB files
    const provider = new ComCompletionProvider(apis);

    const disposable = vscode.languages.registerCompletionItemProvider(
        { language: 'vb', scheme: 'file' },
        provider,
        '.'  // Trigger on dot
    );

    context.subscriptions.push(disposable);
}

export function deactivate() {}
