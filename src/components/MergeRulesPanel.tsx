import React from 'react';
import { Settings2, RotateCcw, RotateCw } from 'lucide-react';
import {
  MergeConfiguration,
  OutputFormat,
  PageSizeStandard,
  SlideAspectRatio,
  QueuedDocument,
} from '../types/document';

interface MergeRulesPanelProps {
  config: MergeConfiguration;
  onChangeConfig: (next: MergeConfiguration) => void;
  onResetDefaults: () => void;
  queuedDocs: QueuedDocument[];
  onUpdatePageRange: (id: string, range: string) => void;
  onCycleRotation: (id: string) => void;
}

export const MergeRulesPanel: React.FC<MergeRulesPanelProps> = ({
  config,
  onChangeConfig,
  onResetDefaults,
  queuedDocs,
  onUpdatePageRange,
  onCycleRotation,
}) => {
  const updateField = <K extends keyof MergeConfiguration>(
    key: K,
    value: MergeConfiguration[K]
  ) => {
    onChangeConfig({
      ...config,
      [key]: value,
    });
  };

  const formatOptions: { id: OutputFormat; label: string }[] = [
    { id: 'auto', label: 'Auto' },
    { id: 'pdf', label: 'PDF' },
    { id: 'docx', label: 'DOCX' },
    { id: 'pptx', label: 'PPTX' },
  ];
  const activeFormatIdx = Math.max(
    0,
    formatOptions.findIndex((f) => f.id === config.outputFormat)
  );
  const formatTranslateClass =
    activeFormatIdx === 0
      ? 'translate-x-0'
      : activeFormatIdx === 1
      ? 'translate-x-full'
      : activeFormatIdx === 2
      ? 'translate-x-[200%]'
      : 'translate-x-[300%]';

  const pageSizeOptions: { id: PageSizeStandard; label: string }[] = [
    { id: 'letter', label: 'Letter' },
    { id: 'a4', label: 'A4' },
    { id: 'original', label: 'Native' },
  ];
  const activePageSizeIdx = Math.max(
    0,
    pageSizeOptions.findIndex((s) => s.id === config.pageSize)
  );
  const pageSizeTranslateClass =
    activePageSizeIdx === 0
      ? 'translate-x-0'
      : activePageSizeIdx === 1
      ? 'translate-x-full'
      : 'translate-x-[200%]';

  const slideRatioOptions: { id: SlideAspectRatio; label: string }[] = [
    { id: '16:9', label: '16:9' },
    { id: '4:3', label: '4:3' },
    { id: '16:10', label: '16:10' },
  ];
  const activeSlideRatioIdx = Math.max(
    0,
    slideRatioOptions.findIndex((r) => r.id === config.slideRatio)
  );
  const slideRatioTranslateClass =
    activeSlideRatioIdx === 0
      ? 'translate-x-0'
      : activeSlideRatioIdx === 1
      ? 'translate-x-full'
      : 'translate-x-[200%]';

  return (
    <div className="bg-[#1E272E] border-2 border-[#00CEC9] rounded-xl p-6 space-y-6 text-[#00CEC9]">
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-[#00CEC9]/25">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-[#1E272E] border border-[#0984E3] flex items-center justify-center text-[#00CEC9]">
            <Settings2 className="w-4 h-4" />
          </div>
          <div>
            <h2 className="text-base font-bold text-[#00CEC9]">
              Output Format &amp; Merge Rules
            </h2>
            <p className="text-xs text-[#00CEC9]/80">
              Changes automatically trigger a fresh background merge
            </p>
          </div>
        </div>

        <button
          type="button"
          onClick={onResetDefaults}
          className="px-3 py-1.5 text-xs font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded-lg transition-colors inline-flex items-center gap-1.5 cursor-pointer"
          title="Reset rules to default"
        >
          <RotateCcw className="w-3.5 h-3.5" />
          <span>Reset Defaults</span>
        </button>
      </div>

      {/* Main settings grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
        {/* Output Filename */}
        <div>
          <label
            htmlFor="output-filename-input"
            className="block text-xs font-bold text-[#00CEC9] mb-1.5"
          >
            Merged Output Filename
          </label>
          <input
            id="output-filename-input"
            type="text"
            value={config.outputFilename}
            onChange={(e) => updateField('outputFilename', e.target.value)}
            placeholder="Merged_Document"
            className="w-full px-3 py-2 text-sm font-mono-tabular font-semibold bg-[#1E272E] border border-[#0984E3] rounded-lg text-[#00CEC9] placeholder:text-[#00CEC9]/50 focus:outline-none focus:ring-2 focus:ring-[#00CEC9]"
          />
        </div>

        {/* Target Output Format (Animated Sliding Switcher) */}
        <div>
          <span className="block text-xs font-bold text-[#00CEC9] mb-1.5">
            Output Format
          </span>
          <div className="relative grid grid-cols-4 items-center p-1 bg-[#1E272E] border border-[#0984E3] rounded-lg">
            <div
              aria-hidden="true"
              className={`absolute top-1 bottom-1 left-1 w-[calc(25%-2px)] bg-[#00CEC9] border border-[#0984E3] rounded-md transition-transform duration-300 ease-out ${formatTranslateClass}`}
            />
            {formatOptions.map((fmt) => {
              const active = config.outputFormat === fmt.id;
              return (
                <button
                  key={fmt.id}
                  type="button"
                  onClick={() => updateField('outputFormat', fmt.id)}
                  className={`relative z-10 py-1.5 px-2 text-xs font-bold rounded-md transition-colors cursor-pointer text-center ${
                    active ? 'text-[#1E272E]' : 'text-[#00CEC9]'
                  }`}
                >
                  {fmt.label}
                </button>
              );
            })}
          </div>
        </div>

        {/* Page Size (PDF & DOCX) (Animated Sliding Switcher) */}
        <div>
          <span className="block text-xs font-bold text-[#00CEC9] mb-1.5">
            Page Size (PDF / DOCX)
          </span>
          <div className="relative grid grid-cols-3 items-center p-1 bg-[#1E272E] border border-[#0984E3] rounded-lg">
            <div
              aria-hidden="true"
              className={`absolute top-1 bottom-1 left-1 w-[calc((100%-8px)/3)] bg-[#00CEC9] border border-[#0984E3] rounded-md transition-transform duration-300 ease-out ${pageSizeTranslateClass}`}
            />
            {pageSizeOptions.map((sz) => {
              const active = config.pageSize === sz.id;
              return (
                <button
                  key={sz.id}
                  type="button"
                  onClick={() => updateField('pageSize', sz.id)}
                  className={`relative z-10 py-1.5 px-2 text-xs font-bold rounded-md transition-colors cursor-pointer text-center ${
                    active ? 'text-[#1E272E]' : 'text-[#00CEC9]'
                  }`}
                >
                  {sz.label}
                </button>
              );
            })}
          </div>
        </div>

        {/* Slide Aspect Ratio (PPTX & Slides) (Animated Sliding Switcher) */}
        <div>
          <span className="block text-xs font-bold text-[#00CEC9] mb-1.5">
            Slide Aspect Ratio (PPTX)
          </span>
          <div className="relative grid grid-cols-3 items-center p-1 bg-[#1E272E] border border-[#0984E3] rounded-lg">
            <div
              aria-hidden="true"
              className={`absolute top-1 bottom-1 left-1 w-[calc((100%-8px)/3)] bg-[#00CEC9] border border-[#0984E3] rounded-md transition-transform duration-300 ease-out ${slideRatioTranslateClass}`}
            />
            {slideRatioOptions.map((ratio) => {
              const active = config.slideRatio === ratio.id;
              return (
                <button
                  key={ratio.id}
                  type="button"
                  onClick={() => updateField('slideRatio', ratio.id)}
                  className={`relative z-10 py-1.5 px-2 text-xs font-bold rounded-md transition-colors cursor-pointer text-center ${
                    active ? 'text-[#1E272E]' : 'text-[#00CEC9]'
                  }`}
                >
                  {ratio.label}
                </button>
              );
            })}
          </div>
        </div>
      </div>

      {/* Checkboxes */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-3 border-t border-[#00CEC9]/25">
        <label className="flex items-start gap-2.5 cursor-pointer select-none">
          <input
            type="checkbox"
            checked={config.insertPageBreakBetweenDocs}
            onChange={(e) => updateField('insertPageBreakBetweenDocs', e.target.checked)}
            className="mt-0.5 h-4 w-4 rounded accent-[#00CEC9]"
          />
          <span className="text-xs">
            <span className="font-bold text-[#00CEC9] block">
              Section / Page Breaks
            </span>
            <span className="text-[#00CEC9]/80">
              Start each merged file on a new page and preserve its headers &amp; footers
            </span>
          </span>
        </label>

        <label className="flex items-start gap-2.5 cursor-pointer select-none">
          <input
            type="checkbox"
            checked={config.stampUnifiedPageNumbers}
            onChange={(e) => updateField('stampUnifiedPageNumbers', e.target.checked)}
            className="mt-0.5 h-4 w-4 rounded accent-[#00CEC9]"
          />
          <span className="text-xs">
            <span className="font-bold text-[#00CEC9] block">
              Continuous Page Numbering
            </span>
            <span className="text-[#00CEC9]/80">
              Stamp unified page numbers across the merged output
            </span>
          </span>
        </label>

        <label className="flex items-start gap-2.5 cursor-pointer select-none">
          <input
            type="checkbox"
            checked={config.autoDownloadOnMerge}
            onChange={(e) => updateField('autoDownloadOnMerge', e.target.checked)}
            className="mt-0.5 h-4 w-4 rounded accent-[#00CEC9]"
          />
          <span className="text-xs">
            <span className="font-bold text-[#00CEC9] block">
              Auto-Download on Upload
            </span>
            <span className="text-[#00CEC9]/80">
              Trigger browser file download automatically after upload completes
            </span>
          </span>
        </label>
      </div>

      {/* Per-File Page / Slide Range & Rotation Controls */}
      {queuedDocs.length > 0 && (
        <div className="pt-4 border-t border-[#00CEC9]/25 space-y-3">
          <div>
            <h3 className="text-xs font-bold text-[#00CEC9]">
              Per-Document Page / Slide Range &amp; Rotation
            </h3>
            <p className="text-xs text-[#00CEC9]/80">
              Specify page or slide subsets (e.g., <span className="font-mono-tabular font-bold text-[#00CEC9]">All</span> or <span className="font-mono-tabular font-bold text-[#00CEC9]">1-3, 5</span>) and page rotation for each file
            </p>
          </div>

          <div className="divide-y divide-[#00CEC9]/20 border border-[#0984E3] rounded-lg bg-[#1E272E] text-[#00CEC9]">
            {queuedDocs.map((doc, idx) => (
              <div
                key={doc.id}
                className="p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs"
              >
                <div className="flex items-center gap-2.5 min-w-0">
                  <span className="font-mono-tabular font-bold text-[#0984E3] w-6">
                    {String(idx + 1).padStart(2, '0')}
                  </span>
                  <span className="font-bold text-[#00CEC9] truncate">{doc.name}</span>
                  <span className="text-[#00CEC9]/80 font-mono-tabular font-semibold shrink-0">
                    ({doc.pageCount}{' '}
                    {doc.extension === 'pptx'
                      ? doc.pageCount === 1
                        ? 'slide'
                        : 'slides'
                      : doc.pageCount === 1
                      ? 'page'
                      : 'pages'}
                    )
                  </span>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <label className="text-[#00CEC9] font-bold">Range:</label>
                  <input
                    type="text"
                    value={doc.pageRangeInput}
                    onChange={(e) => onUpdatePageRange(doc.id, e.target.value)}
                    placeholder="All or 1-3"
                    aria-label={`Page or slide range for ${doc.name}`}
                    className="w-28 px-2.5 py-1 text-xs font-mono-tabular font-bold bg-[#1E272E] border border-[#0984E3] rounded text-[#00CEC9]"
                  />

                  <button
                    type="button"
                    onClick={() => onCycleRotation(doc.id)}
                    className="px-2.5 py-1 text-xs font-mono-tabular font-bold bg-[#1E272E] hover:bg-[#00CEC9] text-[#00CEC9] hover:text-[#1E272E] border border-[#0984E3] rounded inline-flex items-center gap-1 transition-colors cursor-pointer"
                    title="Rotate pages 90°"
                  >
                    <RotateCw className="w-3 h-3" />
                    <span>{doc.rotationDegrees}°</span>
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
