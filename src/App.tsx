import { useCallback, useEffect, useRef, useState } from 'react';
import WebViewer, { type WebViewerInstance } from '@pdftron/webviewer';
import './App.css';

type DocumentStatus = 'ready' | 'loading' | 'loaded' | 'error';

const ACCEPTED_FILES = '.pdf,.doc,.docx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const App = () => {
  const viewerElement = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<WebViewerInstance | null>(null);
  const instancePromiseRef = useRef<Promise<WebViewerInstance> | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [status, setStatus] = useState<DocumentStatus>('ready');
  const [fileName, setFileName] = useState('No CV open');
  const [pageCount, setPageCount] = useState(0);
  const [message, setMessage] = useState('Open a CV to get started. PDF is recommended for the most reliable editing.');
  const [hasDocument, setHasDocument] = useState(false);
  const [viewerReady, setViewerReady] = useState(false);

  useEffect(() => {
    if (!viewerElement.current) return;
    let cancelled = false;
    const viewerPromise = WebViewer({
      path: '/lib/webviewer',
      licenseKey: import.meta.env.VITE_DEMO_KEY,
      fullAPI: true,
      preloadWorker: 'contentEdit',
      enableOfficeEditing: true,
    }, viewerElement.current);
    instancePromiseRef.current = viewerPromise;
    void viewerPromise.then((instance) => {
      if (cancelled) { instance.UI.dispose(); return; }
      instanceRef.current = instance;
      instance.UI.enableElements(['contentEditButton']);
      setViewerReady(true);
      instance.Core.documentViewer.addEventListener('documentLoaded', () => {
        const document = instance.Core.documentViewer.getDocument();
        setPageCount(document.getPageCount());
        setHasDocument(true);
        setStatus('loaded');
        setMessage('Your CV is open. Use the editor toolbar above the document, then download a PDF copy.');
      });
      instance.Core.documentViewer.addEventListener('documentUnloaded', () => {
        setHasDocument(false);
        setPageCount(0);
        setStatus('ready');
      });
    }).catch(() => {
      if (!cancelled) { setStatus('error'); setMessage('The editor could not start. Check the WebViewer files and Apryse license configuration.'); }
    });
    return () => { cancelled = true; instanceRef.current?.UI.dispose(); instanceRef.current = null; };
  }, []);

  const openFile = useCallback(async (file?: File) => {
    if (!file) return;
    const extension = file.name.split('.').pop()?.toLowerCase();
    const isPdf = extension === 'pdf' || file.type === 'application/pdf';
    const isDocx = extension === 'docx' || extension === 'doc';
    if (!isPdf && !isDocx) {
      setStatus('error'); setMessage('That file type is not supported here. Choose a PDF or DOCX CV.'); return;
    }
    setStatus('loading'); setHasDocument(false); setFileName(file.name);
    setMessage('Preparing the editor… your file stays on this device.');
    try {
      const instance = instanceRef.current ?? await instancePromiseRef.current;
      if (!instance) throw new Error('Editor did not initialize.');
      setMessage(isDocx ? 'Opening your Word file in Apryse Office Editor…' : 'Opening your PDF CV…');
      await instance.UI.loadDocument(file, { filename: file.name, enableOfficeEditing: isDocx });
    } catch {
      setStatus('error');
      setMessage('This file did not open. Try a standard PDF, or check whether your Apryse license includes Office Editor.');
    }
  }, []);

  const startTextEdit = useCallback(() => {
    const instance = instanceRef.current;
    if (!instance || !hasDocument) return;
    instance.UI.setToolMode(instance.Core.Tools.ToolNames.CONTENT_EDIT);
    setMessage('Text edit mode is on. If Shift moves a capital to the end, choose Edit PDF Content > Text or paste the corrected word. Click Done editing to leave this mode.');
  }, [hasDocument]);

  const finishTextEdit = useCallback(() => {
    const instance = instanceRef.current;
    if (!instance || !hasDocument) return;
    instance.UI.setToolMode(instance.Core.Tools.ToolNames.PAN);
    setMessage('Text editing is finished. You can continue reviewing or download your edited PDF.');
  }, [hasDocument]);

  const exportFile = useCallback(async () => {
    const instance = instanceRef.current;
    if (!instance || !hasDocument || status !== 'loaded') return;
    try {
      const loadedDocument = instance.Core.documentViewer.getDocument();
      const data = await loadedDocument.getFileData({ includeAnnotations: true, downloadType: 'pdf' });
      const url = URL.createObjectURL(new Blob([data], { type: 'application/pdf' }));
      const link = document.createElement('a');
      const baseName = fileName.replace(/\.(pdf|docx?)$/i, '') || 'cv';
      link.href = url;
      link.download = `${baseName}-edited.pdf`;
      document.body.append(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage('Downloaded an edited PDF copy. Your original file has not been changed.');
    } catch {
      setMessage('PDF download failed. Check whether this document format is supported by your Apryse license.');
    }
  }, [fileName, hasDocument, status]);

  return <main className="studio-shell">
    <header className="titlebar"><span className="titlebar-mark">Y</span><span>Yllen CV Studio</span><span className="titlebar-caption">DOCUMENT WORKSPACE</span><span className="titlebar-buttons" aria-hidden="true"><i>_</i><i>□</i><i>×</i></span></header>
    <div className="workspace">
      <aside className="sidebar">
        <div className="side-heading">FILE</div>
        <button className="menu-button" onClick={() => fileInputRef.current?.click()}><span>📂</span> Open CV…</button>
        <button className="menu-button" onClick={() => void exportFile()} disabled={!hasDocument || status !== 'loaded'}><span>💾</span> Download edited PDF</button>
        <div className="side-heading tools-heading">DOCUMENT</div>
        <div className="side-info"><b>Current file</b><span title={fileName}>{fileName}</span></div>
        <div className="side-info"><b>Pages</b><span>{hasDocument ? pageCount : '—'}</span></div>
        <div className="side-spacer" />
        <div className="privacy-box"><b>Private workspace</b><p>Files open in your browser. This app does not upload them to a server.</p></div>
        <div className="side-foot">YLLEN CV STUDIO · APRYSE</div>
      </aside>
      <section className="editor-area">
        <div className="toolbar-heading"><div><small>CV WORKSPACE</small><h1>{hasDocument ? fileName : 'Edit your CV'}</h1></div><div className="heading-actions"><button className="secondary-button" onClick={startTextEdit} disabled={!hasDocument || status !== 'loaded'}>✎ Edit text</button><button className="secondary-button done-button" onClick={finishTextEdit} disabled={!hasDocument || status !== 'loaded'}>✓ Done editing</button><button className="primary-button" onClick={() => fileInputRef.current?.click()}><span>＋</span> {hasDocument ? 'Open another CV' : 'Open your CV'}</button></div></div>
        <div className="status-bar"><span className={`state-lamp ${status}`} />{status === 'loading' ? 'Opening document…' : status === 'error' ? 'Could not open document' : hasDocument ? `${pageCount} page${pageCount === 1 ? '' : 's'} · ${message}` : message}</div>
        <div className="editor-frame">
          {!hasDocument && <div className="empty-state"><div className="empty-icon">CV</div><h2>{viewerReady ? 'Start with your CV' : 'Starting the editor…'}</h2><p>{viewerReady ? 'Choose a PDF or Word document from your device. To change your existing name, use Edit text and double-click the name on the page.' : 'The PDF editor is loading. You can choose your file now; it will open as soon as the editor is ready.'}</p><button className="primary-button" onClick={() => fileInputRef.current?.click()} disabled={!viewerReady}><span>📄</span> Choose a file from this computer</button><small>PDF recommended · DOCX editing requires Office Editor in your Apryse license</small></div>}
          <div className="viewer-canvas" ref={viewerElement} aria-label="Apryse document editor" />
        </div>
        <div className="bottom-bar"><span>{status === 'loading' ? 'Opening…' : status === 'error' ? 'Attention needed' : hasDocument ? 'Document open' : 'Ready for a CV'}</span><span>Local session · Original CV remains unchanged</span></div>
      </section>
    </div>
    <input ref={fileInputRef} className="file-picker" type="file" accept={ACCEPTED_FILES} onChange={(event) => { void openFile(event.target.files?.[0]); event.currentTarget.value = ''; }} />
  </main>;
};

export default App;
