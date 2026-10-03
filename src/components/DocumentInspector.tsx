import React, { useState } from 'react';
import {
  FileText,
  Eye,
  FileOutput,
} from 'lucide-react';
import {
  QueuedDocument,
  MergedArtifact,
  MergeConfiguration,
} from '../types/document';
import { PdfCanvasViewer } from './PdfCanvasViewer';

interface DocumentInspectorProps {
  mergedArtifact: MergedArtifact | null;
  inspectedDoc: QueuedDocument | null;
  config: MergeConfiguration;
  onDownloadAlternateFormat: (format: 'pdf' | 'docx' | 'pptx') => void;
  onSelectDocToInspect: (id: string) => void;
  queuedDocs: QueuedDocument[];
}

export const DocumentInspector: React.FC<DocumentInspectorProps> = ({
  mergedArtifact,
  inspectedDoc,
  config,
  onDownloadAlternateFormat,
  onSelectDocToInspect,
  queuedDocs,
}) => {
  const [viewMode, setViewMode] = useState<'merged' | 'source'>('merged');

  return (
    <div className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl flex flex-col overflow-hidden text-[#00CEC9]">
      {/* Inspector Header */}
      <div className="px-6 py-4 border-b border-[#00CEC9]/25 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-[#1E272E] border border-[#0984E3] flex items-center justify-center text-[#00CEC9]">
            <Eye className="w-4 h-4" />
          </div>
          <div>
            <h2 className="text-base font-bold text-[#00CEC9]">
              Live Document Inspector
            </h2>
            <p className="text-xs text-[#00CEC9]/80">
              Preview compiled output or inspect individual uploaded files
            </p>
          </div>
        </div>

        {/* Animated 2-Tab Segmented View Toggle (Merged Output | Source File) */}
        <div className="relative grid grid-cols-2 items-center p-1 bg-[#1E272E] border border-[#0984E3] rounded-lg w-60">
          <div
            aria-hidden="true"
            className={`absolute top-1 bottom-1 left-1 w-[calc(50%-4px)] bg-[#00CEC9] border border-[#0984E3] rounded-md transition-transform duration-300 ease-out ${
              viewMode === 'source' ? 'translate-x-full' : 'translate-x-0'
            }`}
          />
          <button
            type="button"
            onClick={() => setViewMode('merged')}
            className={`relative z-10 px-3 py-1.5 text-xs font-bold rounded-md transition-colors cursor-pointer whitespace-nowrap text-center ${
              viewMode === 'merged' ? 'text-[#1E272E]' : 'text-[#00CEC9]'
            }`}
          >
            Merged Output
          </button>
          <button
            type="button"
            onClick={() => setViewMode('source')}
            className={`relative z-10 px-3 py-1.5 text-xs font-bold rounded-md transition-colors cursor-pointer whitespace-nowrap text-center ${
              viewMode === 'source' ? 'text-[#1E272E]' : 'text-[#00CEC9]'
            }`}
          >
            Source File
          </button>
        </div>
      </div>

      {/* Inspector Viewport */}
      <div className="p-6 space-y-5">
        {viewMode === 'merged' ? (
          mergedArtifact ? (
            <div className="space-y-5">
              {/* Quick Alternate Format Export Bar */}
              <div className="bg-[#1E272E] border border-[#0984E3] rounded-lg p-3.5 flex flex-wrap items-center justify-between gap-3 text-[#00CEC9]">
                <div className="text-xs">
                  <span className="font-bold text-[#00CEC9]">
                    {mergedArtifact.filename}
                  </span>
                  <span className="text-[#0984E3] mx-2">·</span>
                  <span className="text-[#00CEC9]/85 font-mono-tabular font-semibold">
                    {mergedArtifact.sourceCount} files merged into {mergedArtifact.totalPages}{' '}
                    {mergedArtifact.format === 'pptx'
                      ? mergedArtifact.totalPages === 1
                        ? 'slide'
                        : 'slides'
                      : mergedArtifact.totalPages === 1
                      ? 'page'
                      : 'pages'}
                  </span>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => onDownloadAlternateFormat('pdf')}
                    className="px-2.5 py-1 text-xs font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded transition-colors cursor-pointer"
                  >
                    Export .PDF
                  </button>
                  <button
                    type="button"
                    onClick={() => onDownloadAlternateFormat('docx')}
                    className="px-2.5 py-1 text-xs font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded transition-colors cursor-pointer"
                  >
                    Export .DOCX
                  </button>
                  <button
                    type="button"
                    onClick={() => onDownloadAlternateFormat('pptx')}
                    className="px-2.5 py-1 text-xs font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded transition-colors cursor-pointer"
                  >
                    Export .PPTX
                  </button>
                </div>
              </div>

              {/* Canvas PDF Preview or OpenXML Manifest Preview */}
              {mergedArtifact.format === 'pdf' ? (
                <PdfCanvasViewer
                  blob={mergedArtifact.blob}
                  blobUrl={mergedArtifact.blobUrl}
                  filename={mergedArtifact.filename}
                />
              ) : (
                <div className="border border-[#0984E3] rounded-lg p-5 bg-[#1E272E] space-y-3 text-[#00CEC9]">
                  <div className="flex items-center justify-between border-b border-[#00CEC9]/25 pb-3">
                    <div>
                      <h3 className="text-sm font-bold text-[#00CEC9]">
                        OpenXML {mergedArtifact.format.toUpperCase()} Package Ready
                      </h3>
                      <p className="text-xs text-[#00CEC9]/85">
                        {mergedArtifact.format === 'docx'
                          ? 'Native Word tables, styles, headers, and footers are compiled into this .docx archive.'
                          : `Slide backgrounds, tables, and ${config.slideRatio} aspect ratio geometry are compiled into this .pptx archive.`}
                      </p>
                    </div>
                  </div>

                  <div className="divide-y divide-[#00CEC9]/20 text-xs">
                    {mergedArtifact.manifest.map((entry, idx) => (
                      <div key={entry.docId} className="py-2 flex items-center justify-between">
                        <span className="font-bold text-[#00CEC9]">
                          {String(idx + 1).padStart(2, '0')}. {entry.name}
                        </span>
                        <span className="font-mono-tabular font-semibold text-[#0984E3]">
                          {mergedArtifact.format === 'pptx' ? 'Slides' : 'Pages'} {entry.startPage}–
                          {entry.endPage}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ) : (
            <div className="py-12 text-center text-xs font-medium text-[#00CEC9]/80">
              <FileOutput className="w-8 h-8 text-[#00CEC9] mx-auto mb-2" />
              Upload PDF, DOCX, or PPTX files on the Merge tab to inspect the merged output here.
            </div>
          )
        ) : queuedDocs.length > 0 && inspectedDoc ? (
          <div className="space-y-4">
            {/* Source File Selector */}
            <div className="flex flex-wrap items-center gap-2">
              {queuedDocs.map((doc, idx) => (
                <button
                  key={doc.id}
                  type="button"
                  onClick={() => onSelectDocToInspect(doc.id)}
                  className={`px-3 py-1.5 text-xs font-bold rounded-lg border transition-colors cursor-pointer ${
                    inspectedDoc.id === doc.id
                      ? 'bg-[#00CEC9] text-[#1E272E] border-[#00CEC9]'
                      : 'bg-[#1E272E] text-[#00CEC9] border-[#0984E3] hover:bg-[#0984E3]/30'
                  }`}
                >
                  {String(idx + 1).padStart(2, '0')}. {doc.name}
                </button>
              ))}
            </div>

            {inspectedDoc.extension === 'pdf' ? (
              <PdfCanvasViewer
                arrayBuffer={inspectedDoc.arrayBuffer}
                filename={inspectedDoc.name}
              />
            ) : (
              <div className="border border-[#0984E3] rounded-lg p-5 bg-[#1E272E] space-y-4 max-h-[440px] overflow-y-auto text-[#00CEC9]">
                {inspectedDoc.docxHeaders && inspectedDoc.docxHeaders.length > 0 && (
                  <div className="text-xs font-mono-tabular text-[#00CEC9] border-b border-[#0984E3] pb-2 bg-[#1E272E] p-2.5 rounded">
                    <strong>Header:</strong> {inspectedDoc.docxHeaders.join(' | ')}
                  </div>
                )}

                {inspectedDoc.extension === 'pptx' &&
                inspectedDoc.slides &&
                inspectedDoc.slides.length > 0 ? (
                  <div className="space-y-3">
                    {inspectedDoc.slides.map((s) => (
                      <div
                        key={s.slideNumber}
                        className="p-3.5 rounded-lg border border-[#0984E3] bg-[#1E272E] text-[#00CEC9] space-y-1.5"
                      >
                        <div className="flex items-center justify-between gap-2">
                          <div className="text-xs font-mono-tabular font-bold text-[#00CEC9]">
                            Slide {s.slideNumber}: {s.title}
                          </div>
                          {s.bgColorHex && (
                            <span className="text-[11px] font-mono-tabular text-[#0984E3]">
                              BG #{s.bgColorHex}
                            </span>
                          )}
                        </div>
                        {s.blocks.map((b, bIdx) =>
                          b.type === 'table-row' && b.cells ? (
                            <div
                              key={bIdx}
                              className="grid gap-2 text-xs font-mono-tabular bg-[#1E272E] text-[#00CEC9] p-1.5 rounded border border-[#0984E3]"
                              style={{
                                gridTemplateColumns: `repeat(${Math.max(1, b.cells.length)}, minmax(0, 1fr))`,
                              }}
                            >
                              {b.cells.map((cell, cIdx) => (
                                <span
                                  key={cIdx}
                                  className={`truncate ${b.isHeaderRow ? 'font-bold' : ''}`}
                                >
                                  {cell}
                                </span>
                              ))}
                            </div>
                          ) : (
                            <p key={bIdx} className="text-xs text-[#00CEC9]/90">
                              {b.text}
                            </p>
                          )
                        )}
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="space-y-2">
                    {inspectedDoc.blocks.map((block, i) =>
                      block.type === 'table-row' && block.cells ? (
                        <div
                          key={i}
                          className={`grid gap-2 text-xs font-mono-tabular p-2 rounded border border-[#0984E3] ${
                            block.isHeaderRow
                              ? 'bg-[#1E272E] font-bold text-[#00CEC9]'
                              : 'bg-[#1E272E] text-[#00CEC9]/90'
                          }`}
                          style={{
                            gridTemplateColumns: `repeat(${Math.max(1, block.cells.length)}, minmax(0, 1fr))`,
                          }}
                        >
                          {block.cells.map((cell, cIdx) => (
                            <span key={cIdx} className="break-words">
                              {cell}
                            </span>
                          ))}
                        </div>
                      ) : (
                        <p
                          key={i}
                          className={
                            block.type === 'h1' || block.type === 'h2'
                              ? 'text-sm font-bold text-[#00CEC9]'
                              : 'text-xs text-[#00CEC9]/90'
                          }
                        >
                          {block.text}
                        </p>
                      )
                    )}
                  </div>
                )}

                {inspectedDoc.docxFooters && inspectedDoc.docxFooters.length > 0 && (
                  <div className="text-xs font-mono-tabular text-[#00CEC9] border-t border-[#0984E3] pt-2 bg-[#1E272E] p-2.5 rounded">
                    <strong>Footer:</strong> {inspectedDoc.docxFooters.join(' | ')}
                  </div>
                )}
              </div>
            )}
          </div>
        ) : (
          <div className="py-12 text-center text-xs font-medium text-[#00CEC9]/80">
            <FileText className="w-8 h-8 text-[#00CEC9] mx-auto mb-2" />
            No source files uploaded yet.
          </div>
        )}
      </div>
    </div>
  );
};
