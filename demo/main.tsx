import { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { XlsxEditor } from '@owload/xlsx-editor';
import type { XlsxEditorHandle } from '@owload/xlsx-editor';

declare global {
  interface Window {
    __saved?: Uint8Array;
    __dirty?: boolean;
    __handle?: XlsxEditorHandle | null;
  }
}

function Demo() {
  const [doc, setDoc] = useState<{ key: number; data: Uint8Array | null; name: string }>({ key: 0, data: null, name: 'new.xlsx' });
  const [log, setLog] = useState('');
  const handle = useRef<XlsxEditorHandle | null>(null);
  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: 6, font: '12px system-ui', display: 'flex', gap: 8, alignItems: 'center' }}>
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            f.arrayBuffer().then((buf) => setDoc((d) => ({ key: d.key + 1, data: new Uint8Array(buf), name: f.name })));
          }}
        />
        <button onClick={() => window.__handle?.save()}>save via ref</button>
        <span>{log}</span>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <XlsxEditor
          key={doc.key}
          data={doc.data}
          fileName={doc.name}
          ref={(h) => {
            handle.current = h;
            window.__handle = h;
          }}
          onSave={async (bytes) => {
            await new Promise((r) => setTimeout(r, 200));
            window.__saved = bytes;
            setLog(`saved ${bytes.length} bytes`);
          }}
          onDirtyChange={(d) => {
            window.__dirty = d;
          }}
          onError={(e) => setLog('error: ' + e.message)}
        />
      </div>
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Demo />);
