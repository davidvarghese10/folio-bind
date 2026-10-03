/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  Upload,
  ArrowUp,
  ArrowDown,
  Trash2,
  Download,
  FolderPlus,
  GripVertical,
  AlertCircle,
  CheckCircle2,
  RefreshCw,
  X,
  ShieldAlert,
  SlidersHorizontal,
  Layers,
} from 'lucide-react';
import {
  QueuedDocument,
  MergeConfiguration,
  MergedArtifact,
  RejectedFileError,
  UploadProgressState,
  MergeProgressState,
} from './types/document';
import {
  validateUploadedFiles,
  inspectUploadedFile,
  mergeDocuments,
  formatBytes,
} from './utils/documentEngine';
import { DocumentInspector } from './components/DocumentInspector';
import { MergeRulesPanel } from './components/MergeRulesPanel';

const DEFAULT_CONFIG: MergeConfiguration = {
  outputFilename: 'Merged_Document',
  outputFormat: 'auto',
  pageSize: 'letter',
  slideRatio: '16:9',
  insertPageBreakBetweenDocs: true,
  stampUnifiedPageNumbers: false,
  autoDownloadOnMerge: false,
};

const INITIAL_UPLOAD_PROGRESS: UploadProgressState = {
  phase: 'idle',
  percent: 0,
  processedFiles: 0,
  totalFiles: 0,
  currentFileName: '',
  stepDetail: 'Waiting for PDF, DOCX, or PPTX files...',
};

const INITIAL_MERGE_PROGRESS: MergeProgressState = {
  phase: 'idle',
  percent: 0,
  processedDocs: 0,
  totalDocs: 0,
  currentDocName: '',
  stepDetail: 'Upload files to merge automatically...',
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'merge' | 'advanced'>('merge');
  const [queuedDocs, setQueuedDocs] = useState<QueuedDocument[]>([]);
  const [config, setConfig] = useState<MergeConfiguration>(DEFAULT_CONFIG);
  const [mergedArtifact, setMergedArtifact] = useState<MergedArtifact | null>(null);
  const [sessionHistory, setSessionHistory] = useState<MergedArtifact[]>([]);
  const [isMerging, setIsMerging] = useState<boolean>(false);
  const [isIngesting, setIsIngesting] = useState<boolean>(false);
  const [isDragOver, setIsDragOver] = useState<boolean>(false);
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [inspectedDocId, setInspectedDocId] = useState<string | null>(null);
  const [bannerNotice, setBannerNotice] = useState<string | null>(null);
  const [rejectedFiles, setRejectedFiles] = useState<RejectedFileError[]>([]);

  // Real-time progress state for Stage 1 (File Upload & Validation)
  const [uploadProgress, setUploadProgress] =
    useState<UploadProgressState>(INITIAL_UPLOAD_PROGRESS);

  // Real-time progress state for Stage 2 (Document Merging)
  const [mergeProgress, setMergeProgress] =
    useState<MergeProgressState>(INITIAL_MERGE_PROGRESS);

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const pendingAutoDownloadRef = useRef<boolean>(false);
  const cursorRef = useRef<HTMLDivElement | null>(null);
  const cursorRingRef = useRef<HTMLDivElement | null>(null);
  const [cursorVisible, setCursorVisible] = useState<boolean>(false);

  // Track pointer position and button hover state for constant dark blue hollow circle cursor
  useEffect(() => {
    const handlePointerMove = (e: MouseEvent) => {
      if (!cursorVisible) {
        setCursorVisible(true);
      }
      if (cursorRef.current) {
        cursorRef.current.style.transform = `translate3d(${e.clientX}px, ${e.clientY}px, 0)`;
      }
      if (cursorRingRef.current) {
        const target = e.target as HTMLElement | null;
        const isButton = Boolean(target?.closest('button, a'));
        cursorRingRef.current.classList.toggle('cursor-enlarged', isButton);
      }
    };

    const handleMouseLeave = () => {
      setCursorVisible(false);
    };

    window.addEventListener('mousemove', handlePointerMove, { passive: true });
    document.addEventListener('mouseleave', handleMouseLeave);
    return () => {
      window.removeEventListener('mousemove', handlePointerMove);
      document.removeEventListener('mouseleave', handleMouseLeave);
    };
  }, [cursorVisible]);

  const triggerBrowserDownload = useCallback((artifact: MergedArtifact) => {
    const link = document.createElement('a');
    link.href = artifact.blobUrl;
    link.download = artifact.filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, []);

  // Process and validate newly selected or dropped files with real-time upload progress
  const handleFilesAdded = useCallback(
    async (fileList: FileList | File[]) => {
      const rawArray = Array.from(fileList);
      if (rawArray.length === 0) return;

      setBannerNotice(null);
      setUploadProgress({
        phase: 'validating',
        percent: 5,
        processedFiles: 0,
        totalFiles: rawArray.length,
        currentFileName: rawArray[0].name,
        stepDetail: `Validating ${rawArray.length} selected ${
          rawArray.length === 1 ? 'file' : 'files'
        } (PDF, DOCX & PPTX only)...`,
      });

      const { validFiles, rejectedFiles: newlyRejected } = validateUploadedFiles(rawArray);

      if (newlyRejected.length > 0) {
        setRejectedFiles((prev) => [...newlyRejected, ...prev].slice(0, 6));
      } else {
        setRejectedFiles([]);
      }

      if (validFiles.length === 0) {
        setUploadProgress({
          phase: 'error',
          percent: 100,
          processedFiles: 0,
          totalFiles: rawArray.length,
          currentFileName: newlyRejected[0]?.fileName || '',
          stepDetail:
            'Upload blocked: Unsupported file format. Only .pdf, .docx, and .pptx files are accepted.',
        });
        return;
      }

      setIsIngesting(true);

      try {
        const inspectedDocs: QueuedDocument[] = [];

        for (let i = 0; i < validFiles.length; i++) {
          const file = validFiles[i];
          const doc = await inspectUploadedFile(file, (withinFilePercent, detail) => {
            const overallPercent = Math.min(
              99,
              Math.round(((i + withinFilePercent / 100) / validFiles.length) * 100)
            );
            setUploadProgress({
              phase: 'uploading',
              percent: overallPercent,
              processedFiles: i,
              totalFiles: validFiles.length,
              currentFileName: file.name,
              stepDetail: detail,
            });
          });
          inspectedDocs.push(doc);
        }

        if (config.autoDownloadOnMerge) {
          pendingAutoDownloadRef.current = true;
        }

        setQueuedDocs((prev) => {
          const next = [...prev, ...inspectedDocs];
          if (!inspectedDocId && inspectedDocs.length > 0) {
            setInspectedDocId(inspectedDocs[0].id);
          }
          setUploadProgress({
            phase: 'complete',
            percent: 100,
            processedFiles: next.length,
            totalFiles: next.length,
            currentFileName: validFiles[validFiles.length - 1].name,
            stepDetail: `Uploaded & validated ${next.length} ${
              next.length === 1 ? 'file' : 'files'
            }`,
          });
          return next;
        });
      } finally {
        setIsIngesting(false);
      }
    },
    [config.autoDownloadOnMerge, inspectedDocId]
  );

  // Automatic real-time merge whenever queuedDocs or config changes
  useEffect(() => {
    const activeDocs = queuedDocs.filter((d) => d.included && d.status === 'ready');
    if (activeDocs.length === 0) {
      setMergedArtifact(null);
      setMergeProgress(INITIAL_MERGE_PROGRESS);
      return;
    }

    let cancelled = false;
    setIsMerging(true);
    setMergeProgress({
      phase: 'merging',
      percent: 5,
      processedDocs: 0,
      totalDocs: activeDocs.length,
      currentDocName: activeDocs[0].name,
      stepDetail: `Merging ${activeDocs.length} ${
        activeDocs.length === 1 ? 'document' : 'documents'
      }...`,
    });

    const timer = setTimeout(async () => {
      try {
        const artifact = await mergeDocuments(queuedDocs, config, (progressUpdate) => {
          if (cancelled) return;
          setMergeProgress({
            phase: progressUpdate.percent >= 100 ? 'complete' : 'merging',
            ...progressUpdate,
          });
        });
        if (cancelled) return;

        setMergedArtifact(artifact);
        const unitLabel =
          artifact.format === 'pptx'
            ? artifact.totalPages === 1
              ? 'slide'
              : 'slides'
            : artifact.totalPages === 1
            ? 'page'
            : 'pages';

        setMergeProgress({
          phase: 'complete',
          percent: 100,
          processedDocs: activeDocs.length,
          totalDocs: activeDocs.length,
          currentDocName: artifact.filename,
          stepDetail: `Ready: ${artifact.filename} (${artifact.totalPages} ${unitLabel} · ${formatBytes(
            artifact.sizeBytes
          )})`,
        });

        setSessionHistory((prev) => {
          const filtered = prev.filter(
            (item) =>
              !(
                item.filename === artifact.filename &&
                item.sizeBytes === artifact.sizeBytes &&
                item.totalPages === artifact.totalPages
              )
          );
          return [artifact, ...filtered].slice(0, 8);
        });

        if (pendingAutoDownloadRef.current && config.autoDownloadOnMerge) {
          pendingAutoDownloadRef.current = false;
          triggerBrowserDownload(artifact);
        }
      } catch (err) {
        if (!cancelled) {
          const message =
            err instanceof Error ? err.message : 'Failed to compile merged document.';
          setBannerNotice(message);
          setMergeProgress((prev) => ({
            ...prev,
            phase: 'error',
            stepDetail: message,
          }));
        }
      } finally {
        if (!cancelled) {
          setIsMerging(false);
        }
      }
    }, 80);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [queuedDocs, config, triggerBrowserDownload]);

  const handleDownloadAlternateFormat = async (targetFormat: 'pdf' | 'docx' | 'pptx') => {
    try {
      setIsMerging(true);
      const altArtifact = await mergeDocuments(
        queuedDocs,
        {
          ...config,
          outputFormat: targetFormat,
        },
        (progressUpdate) => {
          setMergeProgress({
            phase: progressUpdate.percent >= 100 ? 'complete' : 'merging',
            ...progressUpdate,
          });
        }
      );
      triggerBrowserDownload(altArtifact);
      setSessionHistory((prev) => [altArtifact, ...prev].slice(0, 8));
    } finally {
      setIsMerging(false);
    }
  };

  // Queue Manipulation Handlers
  const moveDoc = (index: number, direction: -1 | 1) => {
    setQueuedDocs((prev) => {
      const targetIndex = index + direction;
      if (targetIndex < 0 || targetIndex >= prev.length) return prev;
      const copy = [...prev];
      const [moved] = copy.splice(index, 1);
      copy.splice(targetIndex, 0, moved);
      return copy;
    });
  };

  const removeDoc = (id: string) => {
    setQueuedDocs((prev) => {
      const next = prev.filter((d) => d.id !== id);
      if (next.length === 0) {
        setUploadProgress(INITIAL_UPLOAD_PROGRESS);
        setMergeProgress(INITIAL_MERGE_PROGRESS);
        setMergedArtifact(null);
      } else {
        setUploadProgress((curr) => ({
          ...curr,
          phase: 'complete',
          percent: 100,
          processedFiles: next.length,
          totalFiles: next.length,
          currentFileName: next[next.length - 1].name,
          stepDetail: `${next.length} ${next.length === 1 ? 'file' : 'files'} in queue`,
        }));
      }
      return next;
    });
    if (inspectedDocId === id) {
      setInspectedDocId(null);
    }
  };

  const clearAllDocs = () => {
    setQueuedDocs([]);
    setMergedArtifact(null);
    setInspectedDocId(null);
    setRejectedFiles([]);
    setUploadProgress(INITIAL_UPLOAD_PROGRESS);
    setMergeProgress(INITIAL_MERGE_PROGRESS);
  };

  const clearHistory = () => {
    setSessionHistory([]);
  };

  const toggleIncludeDoc = (id: string) => {
    setQueuedDocs((prev) =>
      prev.map((d) => (d.id === id ? { ...d, included: !d.included } : d))
    );
  };

  const updatePageRange = (id: string, range: string) => {
    setQueuedDocs((prev) =>
      prev.map((d) => (d.id === id ? { ...d, pageRangeInput: range } : d))
    );
  };

  const cycleRotation = (id: string) => {
    const rotations: Array<0 | 90 | 180 | 270> = [0, 90, 180, 270];
    setQueuedDocs((prev) =>
      prev.map((d) => {
        if (d.id !== id) return d;
        const nextIdx = (rotations.indexOf(d.rotationDegrees) + 1) % rotations.length;
        return { ...d, rotationDegrees: rotations[nextIdx] };
      })
    );
  };

  // Drag and drop row reordering
  const handleRowDragStart = (index: number) => {
    setDraggedIndex(index);
  };

  const handleRowDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;
    setQueuedDocs((prev) => {
      const copy = [...prev];
      const [draggedItem] = copy.splice(draggedIndex, 1);
      copy.splice(index, 0, draggedItem);
      return copy;
    });
    setDraggedIndex(index);
  };

  const handleRowDragEnd = () => {
    setDraggedIndex(null);
  };

  const inspectedDoc = queuedDocs.find((d) => d.id === inspectedDocId) || queuedDocs[0] || null;

  return (
    <div className="min-h-screen flex flex-col bg-[#1E272E] text-[#00CEC9]">
      {/* Constant Dark Blue Hollow Circle Cursor */}
      <div
        ref={cursorRef}
        aria-hidden="true"
        className={`fixed top-0 left-0 w-0 h-0 z-[9999] pointer-events-none hidden md:block ${
          cursorVisible ? 'opacity-100' : 'opacity-0'
        }`}
      >
        <div ref={cursorRingRef} className="custom-hollow-cursor-ring" />
      </div>

      {/* Hidden File Input shared across tabs */}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept=".pdf,.docx,.pptx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        onChange={(e) => {
          if (e.target.files && e.target.files.length > 0) {
            handleFilesAdded(e.target.files);
            e.target.value = '';
          }
        }}
        className="hidden"
      />

      {/* Top Navigation Bar — FolioBind anchored at Top-Left Corner | Animated Sliding Tab Switcher */}
      <header className="sticky top-0 z-30 w-full bg-[#1E272E] border-b-2 border-[#0984E3] px-5 py-3 text-[#00CEC9]">
        <div className="w-full flex items-center justify-between gap-4">
          {/* Top-Left Corner: Brand Wordmark */}
          <button
            type="button"
            onClick={() => setActiveTab('merge')}
            className="font-display text-2xl tracking-tight text-[#00CEC9] hover:text-[#0984E3] transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            FolioBind
          </button>

          {/* Animated 2-Tab Switcher (Dark background in top-right bar with sliding indicator) */}
          <nav
            aria-label="Workspace views"
            className="relative grid grid-cols-2 items-center p-1 bg-[#1E272E] border border-[#0984E3] rounded-lg w-56"
          >
            {/* Sliding Indicator Box */}
            <div
              aria-hidden="true"
              className={`absolute top-1 bottom-1 left-1 w-[calc(50%-4px)] bg-[#00CEC9] border border-[#0984E3] rounded-md transition-transform duration-300 ease-out ${
                activeTab === 'advanced' ? 'translate-x-full' : 'translate-x-0'
              }`}
            />

            <button
              type="button"
              onClick={() => setActiveTab('merge')}
              className={`relative z-10 px-3 py-1.5 text-xs font-bold rounded-md transition-colors flex items-center justify-center gap-1.5 cursor-pointer whitespace-nowrap ${
                activeTab === 'merge' ? 'text-[#1E272E]' : 'text-[#00CEC9]'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>Merge</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('advanced')}
              className={`relative z-10 px-3 py-1.5 text-xs font-bold rounded-md transition-colors flex items-center justify-center gap-1.5 cursor-pointer whitespace-nowrap ${
                activeTab === 'advanced' ? 'text-[#1E272E]' : 'text-[#00CEC9]'
              }`}
            >
              <SlidersHorizontal className="w-3.5 h-3.5" />
              <span>Advanced</span>
            </button>
          </nav>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-4xl w-full mx-auto px-6 py-8 space-y-6">
        {activeTab === 'merge' ? (
          <>
            {/* Focused Upload & Merge Header */}
            <section className="text-center space-y-2">
              <h1 className="font-display text-3xl sm:text-4xl text-[#00CEC9] tracking-tight">
                Upload PDF, DOCX, or PPTX files to merge automatically
              </h1>
              <p className="text-sm text-[#00CEC9]/85 font-semibold max-w-xl mx-auto">
                Select or drop two or more <span className="font-mono-tabular font-bold text-[#0984E3]">.pdf</span>
                , <span className="font-mono-tabular font-bold text-[#0984E3]">.docx</span>, or{' '}
                <span className="font-mono-tabular font-bold text-[#0984E3]">.pptx</span> files below.
              </p>
            </section>

            {/* Drag-and-Drop Multi-File Upload Zone */}
            <section>
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragOver(true);
                }}
                onDragLeave={() => setIsDragOver(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsDragOver(false);
                  if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                    handleFilesAdded(e.dataTransfer.files);
                  }
                }}
                onClick={() => fileInputRef.current?.click()}
                className={`rounded-2xl border-2 border-dashed border-[#00CEC9] p-8 flex flex-col items-center justify-center text-center transition-colors cursor-pointer text-[#00CEC9] ${
                  isDragOver
                    ? 'bg-[#00CEC9] text-[#1E272E] border-[#1E272E]'
                    : 'bg-[#1E272E] hover:border-[#00CEC9]'
                }`}
              >
                <div className="w-14 h-14 rounded-2xl bg-[#00CEC9] border-2 border-[#0984E3] flex items-center justify-center mb-3 text-[#1E272E]">
                  <Upload className="w-6 h-6" />
                </div>

                <p className="text-base font-bold text-[#00CEC9]">
                  {isIngesting
                    ? `Uploading & validating files (${uploadProgress.percent}%)...`
                    : 'Drop PDF, DOCX, or PPTX files here, or click to select'}
                </p>
                <p className="text-xs font-medium text-[#00CEC9]/85 mt-1">
                  Only <span className="font-mono-tabular font-bold text-[#0984E3]">.pdf</span>,{' '}
                  <span className="font-mono-tabular font-bold text-[#0984E3]">.docx</span>, and{' '}
                  <span className="font-mono-tabular font-bold text-[#0984E3]">.pptx</span> formats are accepted
                </p>

                <div className="mt-5" onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    onClick={() => fileInputRef.current?.click()}
                    className="px-5 py-2.5 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border-2 border-[#00CEC9] text-xs font-bold rounded-lg transition-colors flex items-center gap-2 cursor-pointer"
                  >
                    <FolderPlus className="w-4 h-4" />
                    <span>Choose PDF / DOCX / PPTX Files</span>
                  </button>
                </div>
              </div>
            </section>

            {/* Input Validation Error Alert (when an unsupported file is uploaded) */}
            {rejectedFiles.length > 0 && (
              <section
                role="alert"
                aria-live="assertive"
                className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl p-5 text-[#00CEC9]"
              >
                <div className="flex items-start justify-between gap-4">
                  <div className="flex items-start gap-3">
                    <ShieldAlert className="w-5 h-5 text-[#00CEC9] shrink-0 mt-0.5" />
                    <div className="space-y-1">
                      <h2 className="text-sm font-bold text-[#00CEC9]">
                        Unsupported File Format Rejected ({rejectedFiles.length}{' '}
                        {rejectedFiles.length === 1 ? 'file' : 'files'})
                      </h2>
                      <p className="text-xs font-medium text-[#00CEC9]/90">
                        Only <strong>PDF (.pdf)</strong>, <strong>Word (.docx)</strong>, and{' '}
                        <strong>PowerPoint (.pptx)</strong> files are allowed:
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => setRejectedFiles([])}
                    className="p-1.5 bg-[#0984E3] hover:bg-[#00CEC9] text-[#1E272E] border border-[#00CEC9] rounded-lg cursor-pointer shrink-0 transition-colors"
                    aria-label="Dismiss unsupported file error message"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                <div className="mt-3 pt-3 border-t border-[#00CEC9]/25 divide-y divide-[#00CEC9]/20">
                  {rejectedFiles.map((err) => (
                    <div
                      key={err.id}
                      className="py-2 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs"
                    >
                      <div className="font-mono-tabular font-bold text-[#00CEC9]">
                        {err.fileName} ({err.detectedType})
                      </div>
                      <div className="font-medium text-[#00CEC9]/85">{err.reason}</div>
                    </div>
                  ))}
                </div>
              </section>
            )}

            {/* Real-Time Visual Progress Indicator (Upload Stage & Merge Stage) */}
            <section
              aria-label="Real-time upload and document merge progress"
              className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl p-5 grid grid-cols-1 md:grid-cols-2 gap-6 text-[#00CEC9]"
            >
              {/* Stage 1: Upload Progress */}
              <div className="space-y-2">
                <div className="flex items-center justify-between gap-2 text-xs">
                  <div className="flex items-center gap-2 font-bold text-[#00CEC9]">
                    {uploadProgress.phase === 'validating' ||
                    uploadProgress.phase === 'uploading' ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#00CEC9] shrink-0" />
                    ) : uploadProgress.phase === 'error' ? (
                      <AlertCircle className="w-3.5 h-3.5 text-[#00CEC9] shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-3.5 h-3.5 text-[#00CEC9] shrink-0" />
                    )}
                    <span>1. File Upload Progress</span>
                  </div>
                  <span className="font-mono-tabular font-bold text-[#00CEC9]">
                    {uploadProgress.percent}% ({uploadProgress.processedFiles}/
                    {uploadProgress.totalFiles})
                  </span>
                </div>

                <div
                  className="w-full h-3.5 bg-[#0984E3]/30 border border-[#00CEC9] rounded-full overflow-hidden p-0.5"
                  role="progressbar"
                  aria-valuenow={uploadProgress.percent}
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-label="File upload progress"
                >
                  <div
                    className="h-full bg-[#00CEC9] transition-all duration-150 rounded-full"
                    style={{ width: `${Math.max(0, Math.min(100, uploadProgress.percent))}%` }}
                  />
                </div>

                <div className="text-xs font-medium text-[#00CEC9]/85 truncate">
                  {uploadProgress.stepDetail}
                </div>
              </div>

              {/* Stage 2: Document Merging Progress */}
              <div className="space-y-2 md:border-l md:border-[#00CEC9]/25 md:pl-6">
                <div className="flex items-center justify-between gap-2 text-xs">
                  <div className="flex items-center gap-2 font-bold text-[#00CEC9]">
                    {isMerging || mergeProgress.phase === 'merging' ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#0984E3] shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-3.5 h-3.5 text-[#0984E3] shrink-0" />
                    )}
                    <span>2. Document Merge Progress</span>
                  </div>
                  <span className="font-mono-tabular font-bold text-[#00CEC9]">
                    {mergeProgress.percent}% ({mergeProgress.processedDocs}/
                    {mergeProgress.totalDocs})
                  </span>
                </div>

                <div
                  className="w-full h-3.5 bg-[#0984E3]/30 border border-[#0984E3] rounded-full overflow-hidden p-0.5"
                  role="progressbar"
                  aria-valuenow={mergeProgress.percent}
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-label="Document merging progress"
                >
                  <div
                    className="h-full bg-[#0984E3] transition-all duration-150 rounded-full"
                    style={{ width: `${Math.max(0, Math.min(100, mergeProgress.percent))}%` }}
                  />
                </div>

                <div className="text-xs font-medium text-[#00CEC9]/85 truncate">
                  {mergeProgress.stepDetail}
                </div>
              </div>
            </section>

            {/* Simple Merged Output & File Queue Card */}
            <section className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl p-6 space-y-5 text-[#00CEC9]">
              <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-[#00CEC9]/25">
                <div>
                  <h2 className="text-base font-bold text-[#00CEC9]">
                    Uploaded Files ({queuedDocs.length})
                  </h2>
                  <p className="text-xs text-[#00CEC9]/80 mt-0.5">
                    {queuedDocs.length === 0
                      ? 'No files uploaded yet'
                      : 'Drag rows or use arrows to change merge order'}
                  </p>
                </div>

                {queuedDocs.length > 0 && (
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => fileInputRef.current?.click()}
                      className="px-3 py-1.5 text-xs font-bold bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border border-[#00CEC9] rounded-lg transition-colors cursor-pointer"
                    >
                      + Add More
                    </button>
                    <button
                      type="button"
                      onClick={clearAllDocs}
                      className="px-3 py-1.5 text-xs font-bold bg-[#0984E3] hover:bg-[#00CEC9] text-[#1E272E] border border-[#00CEC9] rounded-lg transition-colors cursor-pointer"
                    >
                      Clear All
                    </button>
                  </div>
                )}
              </div>

              {queuedDocs.length > 0 && (
                <div className="divide-y divide-[#00CEC9]/20">
                  {queuedDocs.map((doc, index) => (
                    <div
                      key={doc.id}
                      draggable
                      onDragStart={() => handleRowDragStart(index)}
                      onDragOver={(e) => handleRowDragOver(e, index)}
                      onDragEnd={handleRowDragEnd}
                      className={`py-3 px-2 flex items-center justify-between gap-3 rounded transition-colors ${
                        !doc.included ? 'opacity-50' : 'hover:bg-[#0984E3]/25'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0 flex-1">
                        <div
                          className="text-[#00CEC9] cursor-grab active:cursor-grabbing"
                          title="Drag to reorder"
                        >
                          <GripVertical className="w-4 h-4" />
                        </div>

                        <input
                          type="checkbox"
                          checked={doc.included}
                          onChange={() => toggleIncludeDoc(doc.id)}
                          aria-label={`Include ${doc.name}`}
                          className="h-4 w-4 rounded accent-[#00CEC9] cursor-pointer shrink-0"
                        />

                        <span className="font-mono-tabular text-xs font-bold text-[#0984E3] w-6 shrink-0">
                          {String(index + 1).padStart(2, '0')}
                        </span>

                        <div className="min-w-0 flex-1">
                          <div className="text-sm font-bold truncate text-[#00CEC9]">
                            {doc.name}
                          </div>
                          <div className="flex items-center gap-2 text-xs text-[#00CEC9]/85 font-mono-tabular font-medium">
                            <span className="uppercase font-bold text-[#0984E3]">{doc.extension}</span>
                            <span>·</span>
                            <span>
                              {doc.pageCount}{' '}
                              {doc.extension === 'pptx'
                                ? doc.pageCount === 1
                                  ? 'slide'
                                  : 'slides'
                                : doc.pageCount === 1
                                ? 'page'
                                : 'pages'}
                            </span>
                            <span>·</span>
                            <span>{formatBytes(doc.sizeBytes)}</span>
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0">
                        <button
                          type="button"
                          disabled={index === 0}
                          onClick={() => moveDoc(index, -1)}
                          title="Move up"
                          className="p-1.5 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border border-[#00CEC9] disabled:opacity-35 rounded transition-colors cursor-pointer disabled:cursor-not-allowed"
                        >
                          <ArrowUp className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          disabled={index === queuedDocs.length - 1}
                          onClick={() => moveDoc(index, 1)}
                          title="Move down"
                          className="p-1.5 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border border-[#00CEC9] disabled:opacity-35 rounded transition-colors cursor-pointer disabled:cursor-not-allowed"
                        >
                          <ArrowDown className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => removeDoc(doc.id)}
                          title="Remove file"
                          className="p-1.5 bg-[#0984E3] hover:bg-[#00CEC9] text-[#1E272E] border border-[#00CEC9] rounded transition-colors cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* Primary Merged Download CTA */}
              <div className="pt-2 flex flex-col sm:flex-row items-center gap-3">
                <button
                  type="button"
                  disabled={!mergedArtifact || isMerging}
                  onClick={() => mergedArtifact && triggerBrowserDownload(mergedArtifact)}
                  className="w-full py-3.5 px-5 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border-2 border-[#00CEC9] disabled:opacity-45 font-bold text-sm rounded-xl transition-colors flex items-center justify-center gap-2 cursor-pointer disabled:cursor-not-allowed"
                >
                  <Download className="w-4 h-4 shrink-0" />
                  <span>
                    {mergedArtifact
                      ? `Download Merged File: ${mergedArtifact.filename} (${
                          mergedArtifact.totalPages
                        } ${
                          mergedArtifact.format === 'pptx'
                            ? mergedArtifact.totalPages === 1
                              ? 'slide'
                              : 'slides'
                            : mergedArtifact.totalPages === 1
                            ? 'page'
                            : 'pages'
                        } · ${formatBytes(mergedArtifact.sizeBytes)})`
                      : 'Upload PDF, DOCX, or PPTX Files Above to Merge Automatically'}
                  </span>
                </button>
              </div>
            </section>
          </>
        ) : (
          /* ADVANCED TAB: Formatting Rules, Per-File Page/Slide Range, Live Document Inspector & Session History */
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h1 className="font-display text-3xl text-[#00CEC9] tracking-tight">
                  Advanced Merge Options &amp; Inspector
                </h1>
                <p className="text-xs text-[#00CEC9]/80 font-semibold mt-0.5">
                  Configure output format, slide aspect ratios, page geometry, and inspect merged content
                </p>
              </div>
              <button
                type="button"
                onClick={() => setActiveTab('merge')}
                className="px-3.5 py-2 text-xs font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border-2 border-[#00CEC9] rounded-lg transition-colors cursor-pointer"
              >
                Back to Merge
              </button>
            </div>

            {bannerNotice && (
              <div className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl px-4 py-3 flex items-center justify-between gap-4 text-xs text-[#00CEC9]">
                <span className="font-bold">{bannerNotice}</span>
                <button
                  type="button"
                  onClick={() => setBannerNotice(null)}
                  className="font-bold text-[#00CEC9] underline cursor-pointer"
                >
                  Dismiss
                </button>
              </div>
            )}

            {/* Merge Formatting & Output Rules */}
            <MergeRulesPanel
              config={config}
              onChangeConfig={setConfig}
              onResetDefaults={() => setConfig(DEFAULT_CONFIG)}
              queuedDocs={queuedDocs}
              onUpdatePageRange={updatePageRange}
              onCycleRotation={cycleRotation}
            />

            {/* Live Document Inspector */}
            <DocumentInspector
              mergedArtifact={mergedArtifact}
              inspectedDoc={inspectedDoc}
              config={config}
              onDownloadAlternateFormat={handleDownloadAlternateFormat}
              onSelectDocToInspect={(id) => setInspectedDocId(id)}
              queuedDocs={queuedDocs}
            />

            {/* Session Merge History */}
            <section className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl p-6 text-[#00CEC9]">
              <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-[#00CEC9]/25">
                <div>
                  <h2 className="text-base font-bold text-[#00CEC9]">
                    Session Merge History
                  </h2>
                  <p className="text-xs text-[#00CEC9]/80 mt-0.5">
                    Previously compiled documents from this browser session
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  {sessionHistory.length > 0 && (
                    <span className="font-mono-tabular text-xs font-bold text-[#00CEC9]">
                      {sessionHistory.length}{' '}
                      {sessionHistory.length === 1 ? 'snapshot' : 'snapshots'}
                    </span>
                  )}
                  <button
                    type="button"
                    disabled={sessionHistory.length === 0}
                    onClick={clearHistory}
                    className="px-3 py-1.5 text-xs font-bold bg-[#0984E3] hover:bg-[#00CEC9] text-[#1E272E] border border-[#00CEC9] disabled:opacity-40 rounded-lg inline-flex items-center gap-1.5 transition-colors cursor-pointer disabled:cursor-not-allowed"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Clear History</span>
                  </button>
                </div>
              </div>

              {sessionHistory.length === 0 ? (
                <div className="py-8 text-center text-xs font-medium text-[#00CEC9]/80">
                  Merged document snapshots will appear here once you upload files.
                </div>
              ) : (
                <div className="overflow-x-auto pt-2">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-[#00CEC9]/25 text-xs text-[#00CEC9]">
                        <th className="py-3 pr-4 font-bold">Filename</th>
                        <th className="py-3 px-4 font-bold">Format</th>
                        <th className="py-3 px-4 font-bold text-right">Files</th>
                        <th className="py-3 px-4 font-bold text-right">Pages / Slides</th>
                        <th className="py-3 px-4 font-bold text-right">Size</th>
                        <th className="py-3 pl-4 font-bold text-right">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#00CEC9]/20 text-xs">
                      {sessionHistory.map((item) => (
                        <tr key={item.id} className="hover:bg-[#0984E3]/25 transition-colors">
                          <td className="py-3 pr-4 font-bold text-[#00CEC9]">{item.filename}</td>
                          <td className="py-3 px-4 font-mono-tabular uppercase font-bold text-[#0984E3]">
                            {item.format}
                          </td>
                          <td className="py-3 px-4 font-mono-tabular font-semibold text-right">
                            {item.sourceCount}
                          </td>
                          <td className="py-3 px-4 font-mono-tabular font-semibold text-right">
                            {item.totalPages}
                          </td>
                          <td className="py-3 px-4 font-mono-tabular font-semibold text-right">
                            {formatBytes(item.sizeBytes)}
                          </td>
                          <td className="py-3 pl-4 text-right">
                            <button
                              type="button"
                              onClick={() => triggerBrowserDownload(item)}
                              className="px-3 py-1 bg-[#00CEC9] hover:bg-[#0984E3] text-[#1E272E] border border-[#00CEC9] font-bold rounded-md inline-flex items-center gap-1.5 transition-colors cursor-pointer"
                            >
                              <Download className="w-3 h-3" />
                              <span>Download</span>
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </div>
        )}
      </main>
    </div>
  );
}
