import { useState } from 'react';

interface Props {
  replace: boolean;
  message: string;
  onNext(q: string, matchCase: boolean): void;
  onReplace(q: string, to: string, matchCase: boolean): void;
  onReplaceAll(q: string, to: string, matchCase: boolean): void;
  onClose(): void;
}

export function FindBar({ replace, message, onNext, onReplace, onReplaceAll, onClose }: Props) {
  const [q, setQ] = useState('');
  const [to, setTo] = useState('');
  const [cs, setCs] = useState(false);
  return (
    <div className="xe-findbar" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <input autoFocus placeholder="Find" value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && onNext(q, cs)} />
      {replace && <input placeholder="Replace with" value={to} onChange={(e) => setTo(e.target.value)} />}
      <label>
        <input type="checkbox" checked={cs} onChange={(e) => setCs(e.target.checked)} /> Match case
      </label>
      <button onClick={() => onNext(q, cs)}>Find next</button>
      {replace && <button onClick={() => onReplace(q, to, cs)}>Replace</button>}
      {replace && <button onClick={() => onReplaceAll(q, to, cs)}>Replace all</button>}
      <span className="xe-msg">{message}</span>
      <button className="xe-x" onClick={onClose}>✕</button>
    </div>
  );
}
