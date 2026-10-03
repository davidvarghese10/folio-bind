import React, { useEffect, useRef, useState } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {
  ZoomIn,
  ZoomOut,
  ChevronLeft,
  ChevronRight,
  Download,
  RefreshCw,
  AlertCircle,
} from 'lucide-react';

if (typeof window !== 'undefined') {
  try {
    pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
      pdfWorkerUrl,
      document.baseURI || window.location.href
    ).toString();
  } catch {
    pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  }
}

interface PdfCanvasViewerProps {
  blob?: Blob;
  arrayBuffer?: ArrayBuffer;
  blobUrl?: string;
  filename: string;
}

export const PdfCanvasViewer: React.FC<PdfCanvasViewerProps> = ({
  blob,
  arrayBuffer,
  blobUrl,
  filename,
}) => {
  const [numPages, setNumPages] = useState<number>(0);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [scale, setScale] = useState<number>(1.15);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const pdfDocRef = useRef<pdfjsLib.PDFDocumentProxy | null>(null);
  const pageCanvasesRef = useRef<Array<HTMLCanvasElement | null>>([]);
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);

  // Load PDF document into pdfjs
  useEffect(() => {
    let cancelled = false;
    let loadedDoc: pdfjsLib.PDFDocumentProxy | null = null;
    let currentLoadingTask: { promise: Promise<pdfjsLib.PDFDocumentProxy>; destroy?: () => unknown } | null = null;

    const safelyDestroyDoc = (docInstance: pdfjsLib.PDFDocumentProxy | null) => {
      if (!docInstance) return;
      try {
        const maybeDestroy = (docInstance as unknown as { destroy?: () => unknown }).destroy;
        if (typeof maybeDestroy === 'function') {
          Promise.resolve(maybeDestroy.call(docInstance)).catch(() => {});
        } else {
          const maybeCleanup = (docInstance as unknown as { cleanup?: () => unknown }).cleanup;
          if (typeof maybeCleanup === 'function') {
            Promise.resolve(maybeCleanup.call(docInstance)).catch(() => {});
          }
        }
      } catch {
        // Ignore cleanup errors
      }
    };

    const loadPdf = async () => {
      setIsLoading(true);
      setErrorMsg(null);
      try {
        let rawBytes: Uint8Array | null = null;
        if (blob) {
          const buf = await blob.arrayBuffer();
          rawBytes = new Uint8Array(buf.slice(0));
        } else if (arrayBuffer) {
          rawBytes = new Uint8Array(arrayBuffer.slice(0));
        } else if (blobUrl) {
          const res = await fetch(blobUrl);
          const buf = await res.arrayBuffer();
          rawBytes = new Uint8Array(buf);
        }

        if (!rawBytes || rawBytes.byteLength === 0) {
          throw new Error('No PDF binary data available to render.');
        }

        const loadingTask = pdfjsLib.getDocument({ data: rawBytes });
        currentLoadingTask = loadingTask;
        const doc = await loadingTask.promise;
        if (cancelled) {
          safelyDestroyDoc(doc);
          return;
        }

        loadedDoc = doc;
        pdfDocRef.current = doc;
        setNumPages(doc.numPages);
        setCurrentPage(1);
      } catch (err) {
        if (!cancelled) {
          setErrorMsg(
            err instanceof Error ? err.message : 'Unable to render PDF preview.'
          );
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    };

    loadPdf();

    return () => {
      cancelled = true;
      safelyDestroyDoc(loadedDoc);
      try {
        if (currentLoadingTask && typeof currentLoadingTask.destroy === 'function') {
          Promise.resolve(currentLoadingTask.destroy()).catch(() => {});
        }
      } catch {
        // Ignore loadingTask cleanup errors
      }
      pdfDocRef.current = null;
    };
  }, [blob, arrayBuffer, blobUrl]);

  // Render all pages onto their respective canvases whenever numPages or scale changes
  useEffect(() => {
    const doc = pdfDocRef.current;
    if (!doc || numPages === 0 || isLoading) return;

    let cancelled = false;
    const renderTasks: Array<{ cancel: () => void }> = [];

    const renderAllPages = async () => {
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      for (let pageNum = 1; pageNum <= numPages; pageNum++) {
        if (cancelled) break;
        const canvas = pageCanvasesRef.current[pageNum - 1];
        if (!canvas) continue;

        try {
          const page = await doc.getPage(pageNum);
          if (cancelled) break;

          const viewport = page.getViewport({ scale });
          const outputViewport = page.getViewport({ scale: scale * dpr });

          canvas.width = Math.floor(outputViewport.width);
          canvas.height = Math.floor(outputViewport.height);
          canvas.style.width = `${Math.floor(viewport.width)}px`;
          canvas.style.height = `${Math.floor(viewport.height)}px`;

          const ctx = canvas.getContext('2d');
          if (!ctx) continue;

          const renderTask = page.render({
            canvasContext: ctx,
            canvas,
            viewport: outputViewport,
          });
          renderTasks.push(renderTask);
          await renderTask.promise;
        } catch {
          // Ignore cancelled render task errors
        }
      }
    };

    renderAllPages();

    return () => {
      cancelled = true;
      renderTasks.forEach((t) => {
        try {
          t.cancel();
        } catch {
          // ignore
        }
      });
    };
  }, [numPages, scale, isLoading]);

  const jumpToPage = (targetPage: number) => {
    const clamped = Math.max(1, Math.min(numPages, targetPage));
    setCurrentPage(clamped);
    const targetCanvas = pageCanvasesRef.current[clamped - 1];
    if (targetCanvas && scrollContainerRef.current) {
      targetCanvas.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  const handleScroll = () => {
    const container = scrollContainerRef.current;
    if (!container || numPages === 0) return;
    const containerRect = container.getBoundingClientRect();
    const midY = containerRect.top + containerRect.height / 3;

    for (let i = 0; i < numPages; i++) {
      const canvas = pageCanvasesRef.current[i];
      if (!canvas) continue;
      const rect = canvas.getBoundingClientRect();
      if (rect.top <= midY && rect.bottom >= containerRect.top) {
        if (currentPage !== i + 1) {
          setCurrentPage(i + 1);
        }
      }
    }
  };

  return (
    <div className="border-2 border-[#00CEC9] rounded-lg overflow-hidden bg-[#1E272E]">
      {/* Viewer Control Bar */}
      <div className="px-4 py-2.5 bg-[#1E272E] text-[#00CEC9] text-xs font-bold flex flex-wrap items-center justify-between gap-3 border-b border-[#0984E3]">
        <div className="flex items-center gap-3 min-w-0">
          <span className="font-mono-tabular truncate max-w-[220px] sm:max-w-xs">
            {filename}
          </span>
          {numPages > 0 && (
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                disabled={currentPage <= 1}
                onClick={() => jumpToPage(currentPage - 1)}
                className="p-1 bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded disabled:opacity-40 cursor-pointer disabled:cursor-not-allowed transition-colors"
                title="Previous page"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
              <span className="font-mono-tabular text-[#00CEC9] px-1">
                Page {currentPage} / {numPages}
              </span>
              <button
                type="button"
                disabled={currentPage >= numPages}
                onClick={() => jumpToPage(currentPage + 1)}
                className="p-1 bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded disabled:opacity-40 cursor-pointer disabled:cursor-not-allowed transition-colors"
                title="Next page"
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          {/* Zoom Controls */}
          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => setScale((s) => Math.max(0.6, Number((s - 0.15).toFixed(2))))}
              className="p-1 bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded cursor-pointer transition-colors"
              title="Zoom out"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="font-mono-tabular w-12 text-center text-[#00CEC9]">
              {Math.round(scale * 100)}%
            </span>
            <button
              type="button"
              onClick={() => setScale((s) => Math.min(2.2, Number((s + 0.15).toFixed(2))))}
              className="p-1 bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded cursor-pointer transition-colors"
              title="Zoom in"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          {blobUrl && (
            <a
              href={blobUrl}
              download={filename}
              className="px-2.5 py-1 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border border-[#00CEC9] rounded inline-flex items-center gap-1 font-bold transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Save PDF</span>
            </a>
          )}
        </div>
      </div>

      {/* Canvas Pages Viewport */}
      <div
        ref={scrollContainerRef}
        onScroll={handleScroll}
        className="w-full h-[480px] overflow-auto bg-[#1E272E] p-4 flex flex-col items-center gap-5"
      >
        {isLoading ? (
          <div className="my-auto flex flex-col items-center gap-2 text-xs font-bold text-[#00CEC9]">
            <RefreshCw className="w-6 h-6 animate-spin text-[#00CEC9]" />
            <span>Rendering PDF pages...</span>
          </div>
        ) : errorMsg ? (
          <div className="my-auto flex flex-col items-center gap-2 text-xs font-bold text-[#00CEC9] text-center max-w-md">
            <AlertCircle className="w-6 h-6 text-[#00CEC9]" />
            <span>{errorMsg}</span>
          </div>
        ) : (
          Array.from({ length: numPages }, (_, idx) => (
            <div
              key={idx + 1}
              className="relative border-2 border-[#0984E3] rounded shadow-lg bg-[#1E272E] shrink-0"
            >
              <canvas
                ref={(el) => {
                  pageCanvasesRef.current[idx] = el;
                }}
                className="block max-w-full h-auto"
              />
              <div className="px-2.5 py-1 bg-[#1E272E] border-t border-[#0984E3] text-[11px] font-mono-tabular font-bold text-[#00CEC9] flex items-center justify-between">
                <span>Page {idx + 1}</span>
                <span>{numPages} total</span>
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
};
