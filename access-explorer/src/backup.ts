import * as fs from 'node:fs/promises';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { OpenDatabase } from './model';

const BACKUP_DIR_NAME = '.accdb-backups';

/**
 * Creates a timestamped copy of the .accdb before every write (the copy itself is
 * done by the bridge, which already holds a shared handle) and prunes old copies.
 * If the backup fails the write MUST be aborted — callers rely on the throw.
 */
export class BackupManager {
  private readonly lastBackupAt = new Map<string, number>();

  backupDirFor(db: OpenDatabase): string {
    return path.join(path.dirname(db.dbPath), BACKUP_DIR_NAME);
  }

  async beforeWrite(db: OpenDatabase): Promise<void> {
    const cfg = vscode.workspace.getConfiguration('accessExplorer');
    if (!cfg.get<boolean>('backupBeforeWrite', true)) {
      return;
    }
    const debounceMs = cfg.get<number>('backupDebounceSeconds', 30) * 1000;
    const last = this.lastBackupAt.get(db.key) ?? 0;
    if (debounceMs > 0 && Date.now() - last < debounceMs) {
      return;
    }
    const base = path.basename(db.dbPath, path.extname(db.dbPath));
    const stamp = new Date()
      .toISOString()
      .replace(/[-:]/g, '')
      .replace('T', '-')
      .slice(0, 15); // yyyyMMdd-HHmmss
    const target = path.join(this.backupDirFor(db), `${base}.${stamp}.accdb`);
    await db.bridge.backup(target);
    this.lastBackupAt.set(db.key, Date.now());
    await this.prune(db, cfg.get<number>('backupCount', 10));
  }

  private async prune(db: OpenDatabase, keep: number): Promise<void> {
    try {
      const dir = this.backupDirFor(db);
      const base = path.basename(db.dbPath, path.extname(db.dbPath));
      const entries = (await fs.readdir(dir))
        .filter((f) => f.startsWith(`${base}.`) && f.toLowerCase().endsWith('.accdb'))
        .sort(); // timestamp format sorts chronologically
      for (const stale of entries.slice(0, Math.max(0, entries.length - keep))) {
        await fs.rm(path.join(dir, stale), { force: true });
      }
    } catch {
      // Pruning failures must never block a save.
    }
  }
}
