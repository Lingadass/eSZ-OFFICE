import Dexie, {type Table} from 'dexie';
import type {
  DocMetaRecord,
  DocVersionRecord,
  OfficeFileBlobRecord,
  OfficeFileRecord,
  OfficeFolderRecord,
  RecoverySnapshot,
  SettingEntry,
} from '../types';

export class ESZOfficeDatabase extends Dexie {
  files!: Table<OfficeFileRecord, string>;
  fileBlobs!: Table<OfficeFileBlobRecord, string>;
  folders!: Table<OfficeFolderRecord, string>;
  recoverySnapshots!: Table<RecoverySnapshot, string>;
  settings!: Table<SettingEntry, string>;
  docVersions!: Table<DocVersionRecord, string>;
  docMeta!: Table<DocMetaRecord, string>;

  constructor() {
    super('eSZOfficeDB');
    this.version(1).stores({
      files: 'id, name, moduleType, folderId, updatedAt, lastOpenedAt, isFavorite, isTrashed',
      fileBlobs: 'fileId, updatedAt',
      folders: 'id, name, parentId, updatedAt, isTrashed',
      recoverySnapshots: 'fileId, savedAt, dirty',
      settings: 'key',
    });
    this.version(2).stores({
      files: 'id, name, moduleType, folderId, updatedAt, lastOpenedAt, isFavorite, isTrashed',
      fileBlobs: 'fileId, updatedAt',
      folders: 'id, name, parentId, updatedAt, isTrashed',
      recoverySnapshots: 'fileId, savedAt, dirty',
      settings: 'key',
      docVersions: 'id, fileId, savedAt',
      docMeta: 'fileId',
    });
  }
}

export const db = new ESZOfficeDatabase();
