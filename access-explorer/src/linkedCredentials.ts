import * as path from 'node:path';
import * as vscode from 'vscode';
import { describeError } from './errors';
import { OpenDatabase } from './model';

const SECRET_PREFIX = 'accessExplorer.linkedCredentials:';

interface LinkedCredentials {
  uid: string;
  pwd: string;
}

/**
 * Databases with ODBC-linked tables (e.g. to SQL Server) show a native modal login dialog the
 * first time linked data is touched — invisible to Access COM automation, so it would hang the
 * bridge. This prompts once for the connection's username/password, applies them to every
 * ODBC-linked TableDef (persisted into the .accdb itself — the standard Access "saved link"
 * mechanism), and remembers them in VS Code's encrypted Secret Storage, keyed per database file,
 * so future opens of the same file connect silently instead of prompting again.
 */
export class LinkedCredentialsManager {
  constructor(private readonly secrets: vscode.SecretStorage) {}

  private keyFor(db: OpenDatabase): string {
    return SECRET_PREFIX + db.key;
  }

  async forget(db: OpenDatabase): Promise<void> {
    await this.secrets.delete(this.keyFor(db));
  }

  /** No-op if the database has no ODBC-linked tables. */
  async ensureConnected(db: OpenDatabase): Promise<void> {
    if (!db.listing.hasLinkedTables) {
      return;
    }
    const key = this.keyFor(db);
    const label = path.basename(db.dbPath);
    const saved = await this.secrets.get(key);
    if (saved) {
      try {
        const { uid, pwd } = JSON.parse(saved) as LinkedCredentials;
        await db.bridge.relinkCredentials(uid, pwd);
        return;
      } catch (err) {
        void vscode.window.showWarningMessage(
          vscode.l10n.t(
            'The saved data source credentials for {0} no longer work: {1}. Enter them again.',
            label,
            describeError(err),
          ),
        );
        await this.secrets.delete(key);
      }
    }
    await this.promptAndConnect(db);
  }

  private async promptAndConnect(db: OpenDatabase): Promise<void> {
    const label = path.basename(db.dbPath);
    const uid = await vscode.window.showInputBox({
      ignoreFocusOut: true,
      title: vscode.l10n.t('Data source username — {0}', label),
      prompt: vscode.l10n.t(
        'This database has linked tables that need a username and password to connect. They are remembered securely for next time.',
      ),
    });
    if (!uid) {
      return;
    }
    const pwd = await vscode.window.showInputBox({
      password: true,
      ignoreFocusOut: true,
      title: vscode.l10n.t('Data source password — {0}', label),
    });
    if (pwd === undefined) {
      return;
    }
    try {
      await db.bridge.relinkCredentials(uid, pwd);
      await this.secrets.store(this.keyFor(db), JSON.stringify({ uid, pwd }));
      void vscode.window.showInformationMessage(vscode.l10n.t('Linked tables connected for {0}.', label));
    } catch (err) {
      void vscode.window.showErrorMessage(
        vscode.l10n.t('Could not connect the linked tables in {0}: {1}', label, describeError(err)),
      );
    }
  }
}
