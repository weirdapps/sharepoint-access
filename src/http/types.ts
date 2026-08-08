// src/http/types.ts
//
// Shapes returned by classic SPO REST with odata=nometadata. Fields are
// optional because SharePoint omits rather than nulls, and `Length` really
// does arrive as a string.

export interface SpFolder {
  Name?: string;
  ServerRelativeUrl?: string;
  TimeLastModified?: string;
  ItemCount?: number;
}

export interface SpFile {
  Name?: string;
  ServerRelativeUrl?: string;
  /** Byte length, delivered as a decimal string. */
  Length?: string | number;
  TimeLastModified?: string;
}

export interface SpFolderListing {
  Name?: string;
  ServerRelativeUrl?: string;
  Exists?: boolean;
  ItemCount?: number;
  Folders?: SpFolder[];
  Files?: SpFile[];
}

export interface SpList {
  Title?: string;
  Id?: string;
  RootFolder?: { ServerRelativeUrl?: string };
}

export interface SpListsResponse {
  value?: SpList[];
}

export interface SpSearchCell {
  Key?: string;
  Value?: string | null;
}

export interface SpSearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      TotalRows?: number;
      Table?: { Rows?: Array<{ Cells?: SpSearchCell[] }> };
    };
  };
}

/** Normalised entry shared by the ls output. */
export interface Entry {
  name: string;
  serverRelativeUrl: string;
  size?: number;
  modified?: string;
}
