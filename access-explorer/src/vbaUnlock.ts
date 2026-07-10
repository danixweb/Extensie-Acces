import * as path from 'node:path';
import * as vscode from 'vscode';
import { OpenDatabase } from './model';

const SECRET_PREFIX = 'accessExplorer.vbaPassword:';

/**
 * There is no public COM API to supply a VBA project password programmatically. This drives a
 * semi-automatic flow instead: Access + the VBA editor are shown briefly, the (known or
 * previously-saved) password is put on the clipboard, and the user pastes it into Access's own
 * password dialog. Far more reliable across Access versions/locales than simulating keystrokes.
 */
export class VbaUnlockManager {
  constructor(private readonly secrets: vscode.SecretStorage) {}

  private keyFor(db: OpenDatabase): string {
    return SECRET_PREFIX + db.key;
  }

  async forget(db: OpenDatabase): Promise<void> {
    await this.secrets.delete(this.keyFor(db));
  }

  /** Prompts for (or reuses) the VBA project password and drives the unlock flow. */
  async unlock(db: OpenDatabase): Promise<boolean> {
    const key = this.keyFor(db);
    const saved = await this.secrets.get(key);
    const label = path.basename(db.dbPath);

    if (saved) {
      const unlockAction = vscode.l10n.t('Unlock');
      const forgetAction = vscode.l10n.t('Forget saved password');
      const choice = await vscode.window.showInformationMessage(
        vscode.l10n.t(
          'The VBA project in {0} is password-protected. A saved password is available.',
          label,
        ),
        unlockAction,
        forgetAction,
      );
      if (choice === forgetAction) {
        await this.secrets.delete(key);
        return this.unlock(db);
      }
      if (choice !== unlockAction) {
        return false;
      }
      return this.runUnlock(db, saved);
    }

    const entered = await vscode.window.showInputBox({
      password: true,
      ignoreFocusOut: true,
      title: vscode.l10n.t('VBA project password — {0}', label),
      prompt: vscode.l10n.t(
        "The VBA project is password-protected. The password is copied to the clipboard so you can paste it into Access's own dialog — it is not sent anywhere else.",
      ),
    });
    if (!entered) {
      return false;
    }
    await this.secrets.store(key, entered);
    return this.runUnlock(db, entered);
  }

  private async runUnlock(db: OpenDatabase, password: string): Promise<boolean> {
    await vscode.env.clipboard.writeText(password);
    const cfg = vscode.workspace.getConfiguration('accessExplorer');
    const timeoutSeconds = cfg.get<number>('vbaUnlockTimeoutSeconds', 120);

    const { unlocked } = await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: vscode.l10n.t(
          'Access is opening. In its window: press Ctrl+R (Project Explorer), expand this database\'s VBA project, then double-click any module inside it — that is what makes Access show the password box. Paste the password (Ctrl+V) and press OK.',
        ),
      },
      () => db.bridge.unlockVba(timeoutSeconds),
    );

    if (unlocked) {
      db.vbaLocked = false;
      void vscode.window.showInformationMessage(vscode.l10n.t('VBA project unlocked for this session.'));
    } else {
      void vscode.window.showWarningMessage(
        vscode.l10n.t('VBA project unlock timed out. Run "Unlock VBA Project" again when you are ready.'),
      );
    }
    return unlocked;
  }
}
