export type SourceFileType = 'pdf' | 'docx' | 'pptx';

export type OutputFormat = 'auto' | 'pdf' | 'docx' | 'pptx';

export type PageSizeStandard = 'original' | 'letter' | 'a4';

export type SlideAspectRatio = '16:9' | '4:3' | '16:10';

export interface StructuredBlock {
  type: 'h1' | 'h2' | 'h3' | 'p' | 'li' | 'table-row';
  text: string;
  cells?: string[];
  isHeaderRow?: boolean;
}

export interface SlideContent {
  slideNumber: number;
  title: string;
  blocks: StructuredBlock[];
  bgColorHex?: string;
  textColorHex?: string;
}

export interface QueuedDocument {
  id: string;
  file: File;
  name: string;
  extension: SourceFileType;
  sizeBytes: number;
  addedAt: string;
  status: 'parsing' | 'ready' | 'error';
  errorMessage?: string;
  // Inspection metadata
  pageCount: number;
  wordCount: number;
  dimensionsLabel: string;
  // Per-document merge settings
  included: boolean;
  pageRangeInput: string; // e.g., "All" or "1-3, 5"
  rotationDegrees: 0 | 90 | 180 | 270;
  // Extracted content for preview & cross-format conversion
  arrayBuffer: ArrayBuffer;
  previewHtml: string;
  rawText: string;
  blocks: StructuredBlock[];
  slides?: SlideContent[];
  docxHeaders?: string[];
  docxFooters?: string[];
}

export interface MergeConfiguration {
  outputFilename: string;
  outputFormat: OutputFormat;
  pageSize: PageSizeStandard;
  slideRatio: SlideAspectRatio;
  insertPageBreakBetweenDocs: boolean;
  stampUnifiedPageNumbers: boolean;
  autoDownloadOnMerge: boolean;
}

export interface RejectedFileError {
  id: string;
  fileName: string;
  detectedType: string;
  sizeBytes: number;
  reason: string;
}

export interface UploadProgressState {
  phase: 'idle' | 'validating' | 'uploading' | 'complete' | 'error';
  percent: number;
  processedFiles: number;
  totalFiles: number;
  currentFileName: string;
  stepDetail: string;
}

export interface MergeProgressState {
  phase: 'idle' | 'merging' | 'complete' | 'error';
  percent: number;
  processedDocs: number;
  totalDocs: number;
  currentDocName: string;
  stepDetail: string;
}

export interface MergedArtifact {
  id: string;
  filename: string;
  format: 'pdf' | 'docx' | 'pptx';
  blob: Blob;
  blobUrl: string;
  sizeBytes: number;
  totalPages: number;
  totalWords: number;
  sourceCount: number;
  mergedAt: string;
  durationMs: number;
  manifest: Array<{
    docId: string;
    name: string;
    extension: SourceFileType;
    pagesIncluded: number;
    startPage: number;
    endPage: number;
  }>;
}
