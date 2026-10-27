export interface IdentitySet {
  user?: { displayName?: string; email?: string };
  application?: { displayName?: string };
}

/** The subset of driveItem properties this Worker uses. */
export interface DriveItem {
  id: string;
  name?: string;
  size?: number;
  webUrl?: string;
  eTag?: string;
  cTag?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: IdentitySet;
  parentReference?: { driveId?: string; id?: string; path?: string };
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  package?: { type?: string };
  root?: Record<string, unknown>;
  deleted?: { state?: string };
  "@microsoft.graph.downloadUrl"?: string;
}

export interface DeltaPage {
  value: DriveItem[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

export interface Drive {
  id: string;
  name: string;
  webUrl: string;
  driveType?: string;
}

export interface Site {
  id: string;
  webUrl: string;
  displayName?: string;
}
