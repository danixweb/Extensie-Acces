import * as vscode from 'vscode';

export type BridgeErrorCode =
  | 'ACCESS_NOT_INSTALLED'
  | 'DB_LOCKED'
  | 'NO_PERMISSION'
  | 'DB_NOT_FOUND'
  | 'OBJECT_NOT_FOUND'
  | 'SQL_SYNTAX'
  | 'VBA_COMPILE_ERROR'
  | 'MACRO_SECURITY'
  | 'VBE_TRUST_REQUIRED'
  | 'LINKED_AUTH_FAILED'
  | 'BRIDGE_TIMEOUT'
  | 'BRIDGE_CRASHED'
  | 'CONFLICT'
  | 'COM_ERROR';

/** Error thrown by AccessBridge; `code` is stable and maps to a localized message. */
export class BridgeError extends Error {
  constructor(
    readonly code: string,
    detail?: string,
    readonly comNumber?: number,
  ) {
    super(detail ?? code);
    this.name = 'BridgeError';
  }
}

/** Maps a bridge error to a user-facing localized message. */
export function describeError(err: unknown, opts?: { vbaLocked?: boolean }): string {
  const detail = err instanceof Error ? err.message : String(err);
  const code = err instanceof BridgeError ? err.code : 'COM_ERROR';
  switch (code as BridgeErrorCode) {
    case 'ACCESS_NOT_INSTALLED':
      return vscode.l10n.t('Microsoft Access is not installed (COM class not registered).');
    case 'DB_LOCKED':
      return vscode.l10n.t(
        'The database is locked by another process (possibly open exclusively in Access). Close it and retry. Details: {0}',
        detail,
      );
    case 'NO_PERMISSION':
      return vscode.l10n.t('You do not have permission to access this database. Details: {0}', detail);
    case 'DB_NOT_FOUND':
      return vscode.l10n.t('Database file not found or not a valid Access database. Details: {0}', detail);
    case 'OBJECT_NOT_FOUND':
      return vscode.l10n.t('The object no longer exists in the database. Refresh the tree. Details: {0}', detail);
    case 'SQL_SYNTAX':
      return vscode.l10n.t('Access rejected the SQL statement: {0}', detail);
    case 'VBA_COMPILE_ERROR':
      return vscode.l10n.t('The VBA project does not compile: {0}', detail);
    case 'VBE_TRUST_REQUIRED':
      if (opts?.vbaLocked) {
        return vscode.l10n.t(
          'The VBA project is password-protected. Run "Access: Unlock VBA Project" (right-click the database) and unlock it, then Refresh.',
        );
      }
      return vscode.l10n.t(
        'This Access version needs the VBA project object model to read class modules. In Access enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > "Trust access to the VBA project object model", then Refresh.',
      );
    case 'LINKED_AUTH_FAILED':
      return vscode.l10n.t(
        'The username/password were rejected for the linked data source. Details: {0}',
        detail,
      );
    case 'MACRO_SECURITY':
      return vscode.l10n.t(
        'Access blocked the operation (Trust Center / macro security). Add the folder to trusted locations. Details: {0}',
        detail,
      );
    case 'BRIDGE_TIMEOUT':
      return vscode.l10n.t(
        'The Access operation timed out — Access may be showing a hidden dialog. The connection was reset; reopen the database. Details: {0}',
        detail,
      );
    case 'BRIDGE_CRASHED':
      return vscode.l10n.t('The connection to Access was lost. Reopen the database. Details: {0}', detail);
    case 'CONFLICT':
      return vscode.l10n.t(
        'The object changed outside VS Code since it was opened. Run Refresh, review the new content, then re-apply your edit.',
      );
    default:
      return vscode.l10n.t('Access COM error: {0}', detail);
  }
}
