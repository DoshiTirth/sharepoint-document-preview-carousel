/**
 * The panel that opens when someone clicks "Summarize AI". Handles its own
 * loading/progress/error states; the parent component just tells it which
 * file to summarize (or that nothing/an unsupported item is selected) and
 * renders it above or below the preview box.
 */
import * as React from 'react';
import { SummarizeFileProgress, summarizeFile, SummarizationUnavailableError } from '../summarizeFile';
import { SummaryMarkdown } from './SummaryMarkdown';

export interface ISummaryTarget {
  fileName: string;
  fileId: string;
  fileVersion: string;
  fetchFileBytes: () => Promise<ArrayBuffer>;
}

export interface ISummaryPanelProps {
  /** undefined when nothing summarizable is selected (no selection, a
   *  Power BI report, an unsupported file type) - the panel shows a clear
   *  fallback message rather than crashing or showing stale content. */
  target: ISummaryTarget | undefined;
  pdfWorkerSrc?: string;
  onClose: () => void;
}

interface IPanelState {
  status: 'loading' | 'ready' | 'error';
  summary?: string;
  errorMessage?: string;
  progress?: SummarizeFileProgress;
}

function progressLabel(progress: SummarizeFileProgress | undefined): string {
  if (!progress) return 'Getting started\u2026';
  switch (progress.stage) {
    case 'cache-hit':
      return 'Loading saved summary\u2026';
    case 'checking-webgpu':
      return 'Checking device compatibility\u2026';
    case 'parsing':
      return 'Reading the document\u2026';
    case 'loading-model': {
      const pct = progress.modelLoadProgress !== undefined ? Math.round(progress.modelLoadProgress * 100) : undefined;
      return pct !== undefined ? `Loading summarization model\u2026 ${pct}%` : 'Loading summarization model\u2026';
    }
    case 'summarizing-chunk':
      return progress.chunkCount > 1
        ? `Summarizing section ${progress.chunkIndex} of ${progress.chunkCount}\u2026`
        : 'Summarizing\u2026';
    case 'combining':
      return 'Combining into final summary\u2026';
    default:
      return 'Working\u2026';
  }
}

const styles: { [key: string]: React.CSSProperties } = {
  panel: {
    background: '#ffffff',
    border: '1px solid #d1d1d1',
    borderRadius: 8,
    padding: '14px 16px',
    marginTop: 12,
    marginBottom: 12,
    position: 'relative',
  },
  closeButton: {
    position: 'absolute',
    top: 10,
    right: 10,
    background: 'transparent',
    border: 'none',
    fontSize: 16,
    cursor: 'pointer',
    color: '#605e5c',
    lineHeight: 1,
    padding: 4,
  },
  title: {
    fontSize: 14,
    fontWeight: 600,
    marginBottom: 8,
    paddingRight: 24,
  },
  statusRow: {
    fontSize: 13,
    color: '#605e5c',
  },
  errorText: {
    fontSize: 13,
    color: '#a80000',
  },
};

export class SummaryPanel extends React.Component<ISummaryPanelProps, IPanelState> {
  // Guards against two real issues with async work inside a React class
  // component:
  // 1. Unmount safety: if the panel is closed (unmounted) while a
  //    summarization is still in flight, the promise chain would otherwise
  //    call setState() on an unmounted component - React warns about this,
  //    and until the promise settles, the component instance and everything
  //    its closures reference stay reachable (can't be garbage collected).
  // 2. Staleness: if the selected file changes while the panel is still
  //    open (target changes twice before the first summarization finishes),
  //    without this guard the OLDER request could resolve after the newer
  //    one and overwrite fresh state with stale results - and both
  //    inference runs would keep consuming CPU/GPU concurrently for no
  //    reason, since only the latest result is ever wanted.
  // A monotonically increasing token solves both: each run captures its own
  // token, and only applies its result if that token is still the current
  // one by the time it resolves.
  private latestRequestToken = 0;

  public constructor(props: ISummaryPanelProps) {
    super(props);
    this.state = { status: 'loading' };
  }

  public componentDidMount(): void {
    this.runSummarization();
  }

  public componentWillUnmount(): void {
    // Invalidates any in-flight request's token so its eventual
    // then()/catch() becomes a no-op instead of calling setState().
    this.latestRequestToken += 1;
  }

  public componentDidUpdate(previousProps: ISummaryPanelProps): void {
    const previousTarget = previousProps.target;
    const currentTarget = this.props.target;
    const targetChanged =
      previousTarget?.fileId !== currentTarget?.fileId ||
      previousTarget?.fileVersion !== currentTarget?.fileVersion;

    if (targetChanged) {
      this.setState({ status: 'loading', summary: undefined, errorMessage: undefined, progress: undefined });
      this.runSummarization();
    }
  }

  private runSummarization(): void {
    const { target, pdfWorkerSrc } = this.props;
    const requestToken = ++this.latestRequestToken;
    const isStillCurrent = (): boolean => requestToken === this.latestRequestToken;

    if (!target) {
      this.setState({
        status: 'error',
        errorMessage:
          'Select a PDF, Word, Excel, or PowerPoint document from the carousel to summarize it. ' +
          'Power BI reports and folders can\u2019t be summarized.',
      });
      return;
    }

    summarizeFile({
      fileName: target.fileName,
      fileId: target.fileId,
      fileVersion: target.fileVersion,
      fetchFileBytes: target.fetchFileBytes,
      pdfWorkerSrc,
      onProgress: (progress) => {
        if (isStillCurrent()) this.setState({ progress });
      },
    })
      .then((result) => {
        if (isStillCurrent()) this.setState({ status: 'ready', summary: result.summary });
      })
      .catch((error: unknown) => {
        if (!isStillCurrent()) return;
        const message =
          error instanceof SummarizationUnavailableError
            ? error.message
            : 'Something went wrong while summarizing this document. Please try again.';
        this.setState({ status: 'error', errorMessage: message });
      });
  }

  public render(): React.ReactElement<ISummaryPanelProps> {
    const { onClose } = this.props;
    const { status, summary, errorMessage, progress } = this.state;

    return (
      <div style={styles.panel}>
        <button style={styles.closeButton} onClick={onClose} aria-label="Close summary">
          &#10005;
        </button>
        <div style={styles.title}>Summarize AI</div>

        {status === 'loading' && <div style={styles.statusRow}>{progressLabel(progress)}</div>}
        {status === 'error' && <div style={styles.errorText}>{errorMessage}</div>}
        {status === 'ready' && summary && <SummaryMarkdown markdown={summary} />}
      </div>
    );
  }
}
