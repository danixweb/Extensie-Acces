import * as vscode from 'vscode';
import { CATEGORIES, Category, DatabaseRegistry, listingFor, OpenDatabase } from './model';

export type TreeNode = DbNode | CategoryNode | ObjectNode;

export interface DbNode {
  kind: 'db';
  db: OpenDatabase;
}

export interface CategoryNode {
  kind: 'category';
  db: OpenDatabase;
  category: Category;
}

export interface ObjectNode {
  kind: 'object';
  db: OpenDatabase;
  category: Category;
  name: string;
}

const CATEGORY_ICONS: Record<Category, string> = {
  Tables: 'table',
  Queries: 'search',
  Forms: 'browser',
  Reports: 'output',
  Macros: 'run-all',
  Modules: 'symbol-namespace',
};

const OBJECT_ICONS: Record<Category, string> = {
  Tables: 'table',
  Queries: 'database',
  Forms: 'window',
  Reports: 'file',
  Macros: 'zap',
  Modules: 'file-code',
};

function categoryLabel(category: Category): string {
  switch (category) {
    case 'Tables': return vscode.l10n.t('Tables');
    case 'Queries': return vscode.l10n.t('Queries');
    case 'Forms': return vscode.l10n.t('Forms');
    case 'Reports': return vscode.l10n.t('Reports');
    case 'Macros': return vscode.l10n.t('Macros');
    case 'Modules': return vscode.l10n.t('Modules');
  }
}

export class AccessTreeProvider implements vscode.TreeDataProvider<TreeNode> {
  private readonly changeEmitter = new vscode.EventEmitter<TreeNode | undefined>();
  readonly onDidChangeTreeData = this.changeEmitter.event;

  constructor(private readonly registry: DatabaseRegistry) {
    registry.onDidChange(() => this.changeEmitter.fire(undefined));
  }

  refresh(): void {
    this.changeEmitter.fire(undefined);
  }

  getTreeItem(node: TreeNode): vscode.TreeItem {
    switch (node.kind) {
      case 'db': {
        const item = new vscode.TreeItem(
          this.registry.labelFor(node.db),
          vscode.TreeItemCollapsibleState.Expanded,
        );
        item.id = node.db.key;
        item.contextValue = 'database';
        item.description = node.db.dbPath;
        item.tooltip = node.db.dbPath;
        item.iconPath = new vscode.ThemeIcon('database');
        return item;
      }
      case 'category': {
        const count = listingFor(node.db.listing, node.category).length;
        const item = new vscode.TreeItem(
          categoryLabel(node.category),
          count > 0
            ? vscode.TreeItemCollapsibleState.Collapsed
            : vscode.TreeItemCollapsibleState.None,
        );
        item.id = `${node.db.key}/${node.category}`;
        item.contextValue = 'category';
        item.description = String(count);
        item.iconPath = new vscode.ThemeIcon(CATEGORY_ICONS[node.category]);
        return item;
      }
      case 'object': {
        const item = new vscode.TreeItem(node.name, vscode.TreeItemCollapsibleState.None);
        item.id = `${node.db.key}/${node.category}/${node.name}`;
        item.contextValue = `object:${node.category}`;
        item.iconPath = new vscode.ThemeIcon(OBJECT_ICONS[node.category]);
        item.command = {
          command: 'accessExplorer.openObject',
          title: vscode.l10n.t('Open Object'),
          arguments: [node.db.key, node.category, node.name],
        };
        return item;
      }
    }
  }

  getChildren(node?: TreeNode): TreeNode[] {
    if (!node) {
      return this.registry.all.map((db) => ({ kind: 'db', db }) satisfies DbNode);
    }
    if (node.kind === 'db') {
      return CATEGORIES.map(
        (category) => ({ kind: 'category', db: node.db, category }) satisfies CategoryNode,
      );
    }
    if (node.kind === 'category') {
      return listingFor(node.db.listing, node.category).map(
        (name) =>
          ({ kind: 'object', db: node.db, category: node.category, name }) satisfies ObjectNode,
      );
    }
    return [];
  }
}
