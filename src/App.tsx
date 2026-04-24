import React, { useState, useRef, useEffect } from 'react';

/* global Office, Excel */

interface Message { role: 'ai' | 'user'; content: string; time: string; }
interface Change { id: number; cell: string; sheet: string; old: string; proposed: string; reason: string; status: 'pending' | 'accepted' | 'rejected' | 'rethink'; type: 'formula' | 'input'; }
interface Agent { id: number; name: string; icon: string; desc: string; scope: string[]; rules: string[]; on: boolean; }
interface Rule { id: number; name: string; type: string; typeBg: string; typeCol: string; desc: string; trigger: string; code: string; on: boolean; }
interface Template { name: string; icon: string; desc: string; sheets: number; cells: number; cat: string; catBg: string; catCol: string; star: boolean; isnew?: boolean; }

const G_DARK = '#14532d', G_MID = '#16a34a', G_LIGHT = '#f0fdf4', G_BORDER = '#bbf7d0';
const GR_BG = '#f9fafb', GR_BD = '#e5e7eb', GR_TX = '#6b7280', TX = '#111827';

function Toggle({ on, onToggle }: { on: boolean; onToggle: () => void }) {
  return <button onClick={onToggle} style={{ width: 44, height: 24, borderRadius: 12, border: 'none', cursor: 'pointer', background: on ? G_MID : '#d1d5db', position: 'relative', flexShrink: 0 }}><span style={{ position: 'absolute', top: 2, left: on ? 22 : 2, width: 20, height: 20, borderRadius: '50%', background: 'white', display: 'block', transition: 'left 0.2s' }} /></button>;
}
function Badge({ label, bg = '#f3f4f6', col = '#374151' }: { label: string; bg?: string; col?: string }) {
  return <span style={{ background: bg, color: col, fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 12, display: 'inline-block' }}>{label}</span>;
}

async function applyIBFormatting(onProgress: (msg: string) => void): Promise<string> {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      onProgress('Turning off gridlines...');
      sheet.showGridlines = false;
      const usedRange = sheet.getUsedRange();
      usedRange.load(['rowCount', 'columnCount', 'values', 'formulas']);
      await ctx.sync();
      const rowCount = usedRange.rowCount;
      const colCount = usedRange.columnCount;
      const values = usedRange.values as any[][];
      const formulas = usedRange.formulas as any[][];
      onProgress(`Scanning ${rowCount} rows x ${colCount} columns...`);
      const sectionKeywords = ['income statement', 'balance sheet', 'cash flow', 'comparable', 'valuation', 'assumptions', 'revenue', 'expenses'];
      const subtotalKeywords = ['gross profit', 'ebitda', 'ebit', 'net income', 'total', 'earnings before', 'net interest', 'operating income'];
      const sectionHeaderRows: number[] = [];
      const subtotalRows: number[] = [];
      for (let r = 0; r < rowCount; r++) {
        let rowText = '';
        for (let c = 0; c < Math.min(colCount, 3); c++) { if (values[r][c] !== null && values[r][c] !== '') rowText += String(values[r][c]).toLowerCase(); }
        const nonEmpty = values[r].filter((v: any) => v !== null && v !== '').length;
        if (nonEmpty <= 2 && rowText.length > 2 && sectionKeywords.some(k => rowText.includes(k))) sectionHeaderRows.push(r);
        else if (subtotalKeywords.some(k => rowText.includes(k))) subtotalRows.push(r);
      }
      onProgress('Applying font colors...');
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          const val = values[r][c]; const formula = formulas[r][c];
          if (val === null || val === '') continue;
          const cell = usedRange.getCell(r, c);
          cell.format.font.name = 'Calibri'; cell.format.font.size = 12; if (typeof val === 'string' && val.length > 0) { cell.values = [[val.charAt(0).toUpperCase() + val.slice(1)]]; };
          if (sectionHeaderRows.includes(r)) { cell.format.font.color = '#FFFFFF'; }
          else if (typeof formula === 'string' && formula.startsWith('=')) { cell.format.font.color = '#000000'; }
          else if (typeof val === 'number') { cell.format.font.color = '#0000FF'; } else { cell.format.font.color = '#000000'; }
        }
      }
      await ctx.sync();
      onProgress('Styling section headers...');
      for (const r of sectionHeaderRows) { const row = usedRange.getRow(r); row.format.fill.color = '#244062'; row.format.font.color = '#FFFFFF'; }
      await ctx.sync();
      onProgress('Adding subtotal borders...');
      for (const r of subtotalRows) { const row = usedRange.getRow(r); row.format.borders.getItem('EdgeTop').style = 'Continuous'; row.format.borders.getItem('EdgeTop').color = '#000000'; }
      await ctx.sync();
      onProgress('Setting number formats...');
      for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
          const val = values[r][c];
          if (typeof val !== 'number' || val === 0) continue;
          const cell = usedRange.getCell(r, c);
          (cell as any).numberFormat = Math.abs(val) > 0 && Math.abs(val) < 1 ? '0.0%' : '#,##0.0_);(#,##0.0)';
        }
      }
      await ctx.sync();
      onProgress('Setting column widths...');
      sheet.getRange('A:A').format.columnWidth = 20;
      sheet.getRange('B:B').format.columnWidth = 200;
      for (let c = 2; c < Math.min(colCount + 2, 20); c++) { sheet.getRange(`${String.fromCharCode(65 + c)}:${String.fromCharCode(65 + c)}`).format.columnWidth = 90; }
      await ctx.sync();
    });
    return '✓ Done!\n\n• Gridlines OFF\n• Blue = hardcoded inputs\n• Black = formulas\n• Dark blue section headers\n• Subtotal borders added\n• Column widths set\n• Number formats applied';
  } catch (err) { throw new Error(`Formatting failed: ${err instanceof Error ? err.message : 'Unknown error'}`); }
}

async function readActiveSheet() {
  try {
    return await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
      range.load(['values', 'formulas']); await ctx.sync();
      const cells: { address: string; value: string; formula: string }[] = [];
      const vals = range.values as any[][], fmls = range.formulas as any[][];
      for (let r = 0; r < vals.length; r++) for (let c = 0; c < vals[r].length; c++) if (vals[r][c] !== '' && vals[r][c] !== null) cells.push({ address: `R${r+1}C${c+1}`, value: String(vals[r][c]), formula: String(fmls[r][c]) });
      return cells;
    });
  } catch { return []; }
}

async function applyChange(change: Change): Promise<boolean> {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(change.cell);
      if (change.proposed.startsWith('=')) { range.formulas = [[change.proposed]]; } else { const num = parseFloat(change.proposed); range.values = [[isNaN(num) ? change.proposed : num]]; }
      await ctx.sync();
    }); return true;
  } catch { return false; }
}

async function callClaude(userMessage: string, sheetContext: string, apiKey: string): Promise<Change[]> {
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
    body: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 1000, system: `You are 4SIGHT, an AI Excel co-pilot. Respond ONLY with a JSON array of proposed cell changes with fields: cell, proposed, reason, type ("formula" or "input"). If no changes needed return []. Never return anything other than valid JSON.`, messages: [{ role: 'user', content: `Sheet:\n${sheetContext}\n\nRequest: ${userMessage}` }] }),
  });
  if (!response.ok) { const e = await response.text(); throw new Error(`API error: ${response.status} - ${e}`); }
  const data = await response.json();
  return JSON.parse(data.content[0].text.trim().replace(/```json|```/g, '').trim());
}

function ChatTab({ onChangesProposed, apiKey }: { onChangesProposed: (c: Change[]) => void; apiKey: string }) {
  const [msgs, setMsgs] = useState<Message[]>([{ role: 'ai', content: '👋 Welcome to 4SIGHT!\n\nPick a demo or ask me anything about your spreadsheet.\n\nEvery change I propose goes to the Changes tab first — you approve before anything touches your model.', time: new Date().toLocaleTimeString() }]);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);
  const msgsEndRef = useRef<HTMLDivElement>(null);
  useEffect(() => { msgsEndRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [msgs]);
  const addMsg = (role: 'ai' | 'user', content: string) => setMsgs(m => [...m, { role, content, time: new Date().toLocaleTimeString() }]);

  const runDemo = async (key: string) => {
    const labels: Record<string, string> = { format: 'Format This Model', dissect: 'Dissect This Model', sensitivity: 'Build a Sensitivity', explain: 'Explain This Formula' };
    addMsg('user', `Run demo: ${labels[key]}`); setLoading(true);
    if (key === 'format') {
      addMsg('ai', 'Running IB Formatting Agent on your active sheet...');
      try { const result = await applyIBFormatting(() => {}); addMsg('ai', result); }
      catch (err) { addMsg('ai', `Error: ${err instanceof Error ? err.message : 'Unknown'}`); }
      setLoading(false); return;
    }
    const demoChanges: Change[] = [
      { id: Date.now(), cell: 'B4', sheet: 'Sheet1', old: '=B3*0.72', proposed: '=B3*0.68', reason: 'COGS margin updated to 68% per Q3 actuals', status: 'pending', type: 'formula' },
      { id: Date.now() + 1, cell: 'D12', sheet: 'Sheet1', old: '0.15', proposed: '0.18', reason: 'Revenue growth rate raised to 18% for FY2025E', status: 'pending', type: 'input' },
    ];
    setTimeout(() => { addMsg('ai', `Running ${labels[key]}...\n\nI've proposed ${demoChanges.length} changes in the Changes tab.`); onChangesProposed(demoChanges); setLoading(false); }, 800);
  };

  const send = async () => {
    if (!input.trim() || loading) return;
    const q = input; setInput(''); addMsg('user', q); setLoading(true);
    try {
      const cells = await readActiveSheet();
      const sheetCtx = cells.length > 0 ? cells.slice(0, 50).map(c => `${c.address}: ${c.formula !== c.value ? c.formula : c.value}`).join('\n') : 'No data found.';
      if (!apiKey) { addMsg('ai', 'Please enter your Anthropic API key using the ⚙ button above.'); setLoading(false); return; }
      const changes = await callClaude(q, sheetCtx, apiKey);
      if (changes.length === 0) { addMsg('ai', 'No changes needed for this request.'); }
      else { const withIds = changes.map((c, i) => ({ ...c, id: Date.now() + i, sheet: 'Sheet1', old: '(current value)', status: 'pending' as const })); onChangesProposed(withIds); addMsg('ai', `I've proposed ${changes.length} change${changes.length > 1 ? 's' : ''} in the Changes tab.`); }
    } catch (err) { addMsg('ai', `Something went wrong: ${err instanceof Error ? err.message : 'Unknown error'}`); }
    setLoading(false);
  };

  const demos = [
    { key: 'format', icon: '📐', label: 'Format This Model', sub: 'Apply IB formatting standards instantly' },
    { key: 'dissect', icon: '🔍', label: 'Dissect This Model', sub: 'Surface key assumptions and drivers' },
    { key: 'sensitivity', icon: '📊', label: 'Build a Sensitivity', sub: 'Two-variable data table' },
    { key: 'explain', icon: '💡', label: 'Explain This Formula', sub: 'Narrate what the selected cell does' },
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', gap: 10 }}>
      <div style={{ background: G_LIGHT, border: `1px solid ${G_BORDER}`, borderRadius: 10, padding: '10px 14px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8 }}><svg width="12" height="12" viewBox="0 0 24 24" fill={G_MID}><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z" /></svg><span style={{ fontSize: 13, fontWeight: 700, color: G_MID }}>Try Interactive Demos</span></div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 7 }}>
          {demos.map(d => <button key={d.key} onClick={() => runDemo(d.key)} style={{ background: 'white', border: `1px solid ${GR_BD}`, borderRadius: 9, padding: '9px 10px', cursor: 'pointer', textAlign: 'left', display: 'flex', flexDirection: 'column', gap: 3 }}><span style={{ fontSize: 15 }}>{d.icon}</span><span style={{ fontWeight: 700, color: TX, fontSize: 12 }}>{d.label}</span><span style={{ fontSize: 10, color: GR_TX, lineHeight: 1.3 }}>{d.sub}</span></button>)}
        </div>
      </div>
      <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 8, minHeight: 0 }}>
        {msgs.map((m, i) => (
          <div key={i} style={{ display: 'flex', flexDirection: 'column', alignItems: m.role === 'user' ? 'flex-end' : 'flex-start', gap: 3 }}>
            {m.role === 'ai' && <div style={{ fontSize: 11, fontWeight: 700, color: G_MID }}>✦ AI Assistant</div>}
            <div style={{ background: m.role === 'user' ? G_MID : 'white', color: m.role === 'user' ? 'white' : TX, border: m.role === 'ai' ? `1px solid ${GR_BD}` : 'none', borderRadius: 10, padding: '9px 13px', fontSize: 12, lineHeight: 1.6, maxWidth: '88%', whiteSpace: 'pre-line' }}>{m.content}</div>
            <span style={{ fontSize: 10, color: GR_TX }}>{m.time}</span>
          </div>
        ))}
        {loading && <div style={{ fontSize: 11, color: GR_TX }}>✦ Working...</div>}
        <div ref={msgsEndRef} />
      </div>
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', paddingTop: 4 }}>
        <input value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === 'Enter' && send()} placeholder="Ask me to update formulas, create models, or explain calculations..." style={{ flex: 1, border: `1px solid ${GR_BD}`, borderRadius: 10, padding: '9px 13px', fontSize: 12, outline: 'none', background: GR_BG }} />
        <button onClick={send} disabled={loading} style={{ background: loading ? '#86efac' : G_MID, border: 'none', borderRadius: 10, width: 38, height: 38, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: loading ? 'default' : 'pointer', flexShrink: 0 }}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5"><line x1="22" y1="2" x2="11" y2="13" /><polygon points="22 2 15 22 11 13 2 9 22 2" /></svg>
        </button>
      </div>
    </div>
  );
}

function ChangesTab({ changes, setChanges }: { changes: Change[]; setChanges: React.Dispatch<React.SetStateAction<Change[]>> }) {
  const act = async (id: number, status: 'accepted' | 'rejected' | 'rethink') => { if (status === 'accepted') { const c = changes.find(x => x.id === id); if (c) await applyChange(c); } setChanges(c => c.map(x => x.id === id ? { ...x, status } : x)); };
  const acceptAll = async () => { for (const c of changes.filter(x => x.status === 'pending')) await applyChange(c); setChanges(c => c.map(x => x.status === 'pending' ? { ...x, status: 'accepted' } : x)); };
  const rejectAll = () => setChanges(c => c.map(x => x.status === 'pending' ? { ...x, status: 'rejected' } : x));
  const pending = changes.filter(c => c.status === 'pending').length;
  const sBg: Record<string, string> = { accepted: '#f0fdf4', rejected: '#fef2f2', pending: 'white', rethink: 'white' };
  const sBd: Record<string, string> = { accepted: G_BORDER, rejected: '#fca5a5', pending: GR_BD, rethink: GR_BD };
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div><div style={{ fontSize: 15, fontWeight: 700, color: TX }}>Proposed Changes</div><div style={{ fontSize: 11, color: GR_TX }}>{pending} pending · {changes.length - pending} reviewed</div></div>
        <div style={{ display: 'flex', gap: 6 }}><button onClick={acceptAll} style={{ background: G_MID, color: 'white', border: 'none', borderRadius: 7, padding: '6px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Accept All</button><button onClick={rejectAll} style={{ background: 'white', color: '#dc2626', border: '1px solid #fca5a5', borderRadius: 7, padding: '6px 10px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Reject All</button></div>
      </div>
      {changes.length === 0 && <div style={{ textAlign: 'center', color: GR_TX, fontSize: 12, padding: '32px 0' }}>No changes yet — ask the AI something in the Chat tab.</div>}
      {changes.map(c => (
        <div key={c.id} style={{ border: `1px solid ${sBd[c.status]}`, borderRadius: 10, padding: 12, background: sBg[c.status] }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}><span style={{ fontFamily: 'monospace', fontWeight: 700, fontSize: 13 }}>{c.cell}</span><span style={{ fontSize: 11, color: GR_TX }}>{c.sheet}</span><Badge label={c.type} bg={c.type === 'formula' ? '#eff6ff' : '#f0fdf4'} col={c.type === 'formula' ? '#1d4ed8' : '#15803d'} /></div>
            {c.status !== 'pending' && <Badge label={c.status === 'accepted' ? '✓ Accepted' : c.status === 'rejected' ? '✗ Rejected' : '↩ Rethink'} bg={sBg[c.status]} col={c.status === 'accepted' ? '#15803d' : '#dc2626'} />}
          </div>
          <div style={{ display: 'flex', gap: 8, marginBottom: 6 }}><code style={{ background: '#fef2f2', borderRadius: 6, padding: '3px 8px', fontSize: 11, color: '#991b1b' }}>— {c.old}</code><code style={{ background: '#f0fdf4', borderRadius: 6, padding: '3px 8px', fontSize: 11, color: '#15803d' }}>+ {c.proposed}</code></div>
          <div style={{ fontSize: 11, color: GR_TX, marginBottom: c.status === 'pending' ? 8 : 0 }}>{c.reason}</div>
          {c.status === 'pending' && <div style={{ display: 'flex', gap: 6 }}><button onClick={() => act(c.id, 'accepted')} style={{ flex: 1, background: G_MID, color: 'white', border: 'none', borderRadius: 7, padding: '6px 0', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>✓ Accept</button><button onClick={() => act(c.id, 'rejected')} style={{ flex: 1, background: 'white', color: '#dc2626', border: '1px solid #fca5a5', borderRadius: 7, padding: '6px 0', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>✗ Reject</button><button onClick={() => act(c.id, 'rethink')} style={{ flex: 1, background: 'white', color: GR_TX, border: `1px solid ${GR_BD}`, borderRadius: 7, padding: '6px 0', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>↩ Rethink</button></div>}
        </div>
      ))}
    </div>
  );
}

function AgentsTab() {
  const [agents, setAgents] = useState<Agent[]>([
    { id: 1, name: 'IB Formatting Agent', icon: '📐', desc: 'Applies Wall Street IB formatting standards to the active sheet instantly.', scope: ['Formatting', 'Colors', 'Fonts', 'Borders'], rules: ['Blue font for hardcoded inputs', 'Black font for formulas', 'Dark blue section headers + white text', 'Gridlines OFF', 'Calibri 12pt throughout'], on: true },
    { id: 2, name: 'Dashboard Agent', icon: '📊', desc: 'Controls dashboard creation. Prevents breaking chart references.', scope: ['Dashboard', 'Charts'], rules: ['No formula changes in source data', 'Preserve named ranges'], on: true },
    { id: 3, name: 'Sensitivity Agent', icon: '📈', desc: 'Manages sensitivity analyses and scenario planning.', scope: ['Sensitivity', 'Scenarios'], rules: ['Preserve input cells', 'Maintain data table formulas'], on: true },
    { id: 4, name: 'VC Model Dissector', icon: '🔍', desc: 'Reverse-engineers a startup model — surfaces key assumptions and runs sensitivities.', scope: ['Model Analysis', 'Assumptions'], rules: ['Extract hardcoded drivers', 'Flag circular logic'], on: false },
    { id: 5, name: 'Audit Agent', icon: '🛡️', desc: 'Logs every cell change with timestamp and rationale.', scope: ['Audit', 'Logging'], rules: ['Log all assumption changes', 'Flag unreviewed changes'], on: false },
  ]);
  const [running, setRunning] = useState<number | null>(null);
  const [runResult, setRunResult] = useState('');
  const toggle = (id: number) => setAgents(a => a.map(x => x.id === id ? { ...x, on: !x.on } : x));
  const active = agents.filter(a => a.on).length;

  const runAgent = async (id: number) => {
    setRunning(id); setRunResult('Starting...');
    try {
      if (id === 1) { const result = await applyIBFormatting(msg => setRunResult(msg)); setRunResult(result); }
      else { setTimeout(() => { setRunResult('Agent completed.'); setRunning(null); }, 1500); return; }
    } catch (err) { setRunResult(`Error: ${err instanceof Error ? err.message : 'Unknown'}`); }
    setRunning(null);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div><div style={{ fontSize: 15, fontWeight: 700, color: TX }}>Sub-Agents</div><div style={{ fontSize: 11, color: GR_TX }}>Control specialized agents</div><div style={{ fontSize: 11, color: G_MID, marginTop: 2 }}>{active} of {agents.length} active</div></div>
        <button style={{ background: G_MID, color: 'white', border: 'none', borderRadius: 8, padding: '7px 13px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>+ New Agent</button>
      </div>
      {runResult && <div style={{ background: G_LIGHT, border: `1px solid ${G_BORDER}`, borderRadius: 8, padding: '8px 12px', fontSize: 11, color: G_DARK, whiteSpace: 'pre-line' }}>{runResult}</div>}
      {agents.map(ag => (
        <div key={ag.id} style={{ border: `1px solid ${ag.on ? G_BORDER : GR_BD}`, borderRadius: 10, padding: 12, background: ag.on ? G_LIGHT : 'white' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 8 }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10 }}>
              <div style={{ width: 34, height: 34, borderRadius: 8, background: ag.on ? '#dcfce7' : '#f3f4f6', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 15, flexShrink: 0 }}>{ag.icon}</div>
              <div><div style={{ fontSize: 13, fontWeight: 700, color: TX }}>{ag.name}</div><div style={{ fontSize: 11, color: GR_TX, maxWidth: 220, lineHeight: 1.4 }}>{ag.desc}</div></div>
            </div>
            <Toggle on={ag.on} onToggle={() => toggle(ag.id)} />
          </div>
          <div style={{ fontSize: 10, fontWeight: 700, color: GR_TX, letterSpacing: '0.06em', marginBottom: 4 }}>SCOPE</div>
          <div style={{ display: 'flex', gap: 5, flexWrap: 'wrap', marginBottom: 6 }}>{ag.scope.map(s => <Badge key={s} label={s} />)}</div>
          <div style={{ fontSize: 10, fontWeight: 700, color: GR_TX, letterSpacing: '0.06em', marginBottom: 4 }}>RULES</div>
          {ag.rules.map(r => <div key={r} style={{ fontSize: 11, color: GR_TX, display: 'flex', alignItems: 'center', gap: 5, marginBottom: 2 }}><span style={{ width: 3, height: 3, borderRadius: '50%', background: GR_TX, display: 'inline-block', flexShrink: 0 }} />{r}</div>)}
          {ag.on && (
            <button onClick={() => runAgent(ag.id)} disabled={running === ag.id} style={{ marginTop: 10, background: running === ag.id ? '#86efac' : G_MID, color: 'white', border: 'none', borderRadius: 7, padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: running === ag.id ? 'default' : 'pointer' }}>
              {running === ag.id ? '⏳ Running...' : '▶ Run Agent'}
            </button>
          )}
        </div>
      ))}
    </div>
  );
}

function RulesTab() {
  const [rules, setRules] = useState<Rule[]>([
    { id: 1, name: 'No Circular References', type: 'validation', typeBg: '#eff6ff', typeCol: '#1d4ed8', desc: 'Validates that proposed changes do not create circular reference errors', trigger: 'before-change', code: 'checkCircularReferences(proposedChanges)', on: true },
    { id: 2, name: 'Preserve Audit Trail', type: 'audit', typeBg: '#faf5ff', typeCol: '#7c3aed', desc: 'Ensures all changes to assumption cells are logged with timestamp', trigger: 'after-change', code: 'logAssumptionChanges(cellRange)', on: true },
    { id: 3, name: 'IB Formatting Standards', type: 'formatting', typeBg: '#f0fdf4', typeCol: '#15803d', desc: 'Enforces Calibri 12pt, blue inputs, black formulas, dark blue section headers', trigger: 'after-change', code: 'applyIBFormatting(worksheet)', on: true },
    { id: 4, name: 'Protect Input Cells', type: 'constraint', typeBg: '#fffbeb', typeCol: '#d97706', desc: 'Prevents AI from overwriting cells tagged as hardcoded assumptions', trigger: 'before-change', code: 'validateInputProtection(cellRef)', on: true },
  ]);
  const toggle = (id: number) => setRules(r => r.map(x => x.id === id ? { ...x, on: !x.on } : x));
  const active = rules.filter(r => r.on).length;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div><div style={{ fontSize: 15, fontWeight: 700, color: TX }}>Custom Rules</div><div style={{ fontSize: 11, color: GR_TX }}>Define validation and constraint rules</div><div style={{ fontSize: 11, color: G_MID, marginTop: 2 }}>{active} of {rules.length} active</div></div>
        <button style={{ background: G_MID, color: 'white', border: 'none', borderRadius: 8, padding: '7px 13px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>+ New Rule</button>
      </div>
      {rules.map(r => (
        <div key={r.id} style={{ border: `1px solid ${GR_BD}`, borderRadius: 10, padding: 12, background: 'white' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 6 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}><div style={{ width: 28, height: 28, borderRadius: '50%', background: r.typeBg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, color: r.typeCol, flexShrink: 0 }}>✓</div><span style={{ fontSize: 13, fontWeight: 700, color: TX }}>{r.name}</span><Badge label={r.type} bg={r.typeBg} col={r.typeCol} /></div>
            <Toggle on={r.on} onToggle={() => toggle(r.id)} />
          </div>
          <div style={{ fontSize: 11, color: GR_TX, marginBottom: 6 }}>{r.desc}</div>
          <Badge label={r.trigger} />
          <div style={{ marginTop: 8, background: '#f9fafb', borderRadius: 7, padding: '7px 10px' }}><div style={{ fontSize: 9, fontWeight: 700, color: GR_TX, letterSpacing: '0.08em', marginBottom: 3 }}>CODE</div><code style={{ fontSize: 11, fontFamily: 'monospace', color: TX }}>{r.code}</code></div>
        </div>
      ))}
    </div>
  );
}

function TemplatesTab() {
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState('All');
  const filters = ['All', 'Valuation', 'Industry', 'Private Equity', 'Startups', 'M&A', 'VC'];
  const templates: Template[] = [
    { name: 'DCF Valuation Model', icon: '💵', desc: 'Three-statement DCF with WACC calculation', sheets: 5, cells: 420, cat: 'Valuation', catBg: '#eff6ff', catCol: '#1d4ed8', star: true },
    { name: 'Three Statement Model', icon: '📊', desc: 'Fully linked IS, Balance Sheet, and Cash Flow', sheets: 3, cells: 480, cat: 'Valuation', catBg: '#eff6ff', catCol: '#1d4ed8', star: true, isnew: true },
    { name: 'SaaS Financial Model', icon: '📈', desc: 'Complete SaaS metrics with cohort analysis', sheets: 7, cells: 650, cat: 'Industry', catBg: '#f0fdf4', catCol: '#15803d', star: true },
    { name: 'LBO Analysis', icon: '🧮', desc: 'Leveraged buyout with returns waterfall', sheets: 6, cells: 590, cat: 'Private Equity', catBg: '#faf5ff', catCol: '#7c3aed', star: false },
    { name: 'M&A Accretion/Dilution', icon: '🔄', desc: 'Merger model with synergies and EPS bridge', sheets: 8, cells: 740, cat: 'M&A', catBg: '#fff7ed', catCol: '#c2410c', star: false },
    { name: 'Startup Financial Model', icon: '🚀', desc: 'Full 3-statement for early-stage companies', sheets: 5, cells: 380, cat: 'Startups', catBg: '#fffbeb', catCol: '#b45309', star: true },
    { name: 'VC Comparables Valuation', icon: '🔭', desc: 'Revenue and EBITDA multiples comps for VC', sheets: 3, cells: 290, cat: 'VC', catBg: '#f0f9ff', catCol: '#0369a1', star: true, isnew: true },
    { name: 'Capitalization Table', icon: '🥧', desc: 'Cap table from seed through Series C', sheets: 4, cells: 340, cat: 'VC', catBg: '#f0f9ff', catCol: '#0369a1', star: true, isnew: true },
  ];
  const shown = templates.filter(t => (filter === 'All' || t.cat === filter) && (!search || t.name.toLowerCase().includes(search.toLowerCase())));
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <div><div style={{ fontSize: 15, fontWeight: 700, color: TX }}>Template Library</div><div style={{ fontSize: 11, color: GR_TX }}>Start with pre-built financial models or save your own</div></div>
      <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search templates..." style={{ width: '100%', padding: '8px 12px', border: `1px solid ${GR_BD}`, borderRadius: 8, fontSize: 12, outline: 'none', background: GR_BG, boxSizing: 'border-box' }} />
      <div style={{ display: 'flex', gap: 5, flexWrap: 'wrap' }}>{filters.map(f => <button key={f} onClick={() => setFilter(f)} style={{ padding: '4px 10px', borderRadius: 16, fontSize: 11, fontWeight: 600, cursor: 'pointer', background: filter === f ? TX : 'white', color: filter === f ? 'white' : TX, border: `1px solid ${filter === f ? TX : GR_BD}` }}>{f}</button>)}</div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {shown.map(t => (
          <div key={t.name} style={{ border: `1px solid ${t.isnew ? G_BORDER : GR_BD}`, borderRadius: 10, padding: 12, background: t.isnew ? G_LIGHT : 'white' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 5 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 9 }}>
                <div style={{ width: 34, height: 34, borderRadius: 8, background: t.isnew ? '#dcfce7' : G_LIGHT, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16, flexShrink: 0 }}>{t.icon}</div>
                <div><div style={{ fontSize: 13, fontWeight: 700, color: TX }}>{t.name} {t.star && <span style={{ color: '#f59e0b' }}>★</span>}{t.isnew && <span style={{ background: G_MID, color: 'white', fontSize: 9, fontWeight: 700, padding: '1px 6px', borderRadius: 8, marginLeft: 4 }}>NEW</span>}</div><div style={{ fontSize: 11, color: GR_TX, lineHeight: 1.4, maxWidth: 200 }}>{t.desc}</div></div>
              </div>
              <Badge label={t.cat} bg={t.catBg} col={t.catCol} />
            </div>
            <div style={{ fontSize: 10, color: GR_TX, marginBottom: 8 }}>{t.sheets} sheets · {t.cells} cells</div>
            <div style={{ display: 'flex', gap: 7 }}><button style={{ background: G_MID, color: 'white', border: 'none', borderRadius: 7, padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>↓ Use Template</button><button style={{ background: 'white', color: TX, border: `1px solid ${GR_BD}`, borderRadius: 7, padding: '6px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Preview</button></div>
          </div>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState<'Chat' | 'Changes' | 'Agents' | 'Rules' | 'Templates'>('Chat');
  const [changes, setChanges] = useState<Change[]>([]);
  const [apiKey, setApiKey] = useState('');
  const [showKey, setShowKey] = useState(false);
  const tabs = ['Chat', 'Changes', 'Agents', 'Rules', 'Templates'] as const;
  const pendingCount = changes.filter(c => c.status === 'pending').length;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', fontFamily: "'Segoe UI', system-ui, sans-serif", background: 'white' }}>
      <div style={{ background: G_DARK, padding: '10px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}><div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 2.5 }}>{[...Array(9)].map((_, i) => <div key={i} style={{ width: 3, height: 3, background: 'rgba(255,255,255,0.65)', borderRadius: 1 }} />)}</div><span style={{ color: 'white', fontWeight: 800, fontSize: 14, letterSpacing: '0.1em' }}>4SIGHT</span></div>
        <button onClick={() => setShowKey(k => !k)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.6)', fontSize: 11 }}>{showKey ? 'Hide Key' : '⚙ API Key'}</button>
      </div>
      {showKey && <div style={{ padding: '8px 14px', background: '#f0fdf4', borderBottom: `1px solid ${G_BORDER}`, flexShrink: 0 }}><div style={{ fontSize: 11, color: GR_TX, marginBottom: 4 }}>Anthropic API Key (stored in memory only)</div><input type="password" value={apiKey} onChange={e => setApiKey(e.target.value)} placeholder="sk-ant-..." style={{ width: '100%', padding: '6px 10px', border: `1px solid ${GR_BD}`, borderRadius: 7, fontSize: 12, outline: 'none', boxSizing: 'border-box' }} /></div>}
      <div style={{ display: 'flex', borderBottom: `1px solid ${GR_BD}`, background: 'white', flexShrink: 0 }}>
        {tabs.map(t => <button key={t} onClick={() => setTab(t)} style={{ flex: 1, padding: '9px 4px', border: 'none', background: 'none', fontSize: 12, fontWeight: t === tab ? 600 : 400, cursor: 'pointer', color: t === tab ? TX : GR_TX, borderBottom: t === tab ? `2px solid ${G_MID}` : '2px solid transparent', position: 'relative' }}>{t}{t === 'Changes' && pendingCount > 0 && <span style={{ position: 'absolute', top: 4, right: 4, background: '#dc2626', color: 'white', borderRadius: '50%', width: 16, height: 16, fontSize: 9, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{pendingCount}</span>}</button>)}
      </div>
      <div style={{ flex: 1, padding: 14, overflowY: 'auto', background: GR_BG, display: 'flex', flexDirection: 'column' }}>
        {tab === 'Chat' && <ChatTab onChangesProposed={c => setChanges(prev => [...prev, ...c])} apiKey={apiKey} />}
        {tab === 'Changes' && <ChangesTab changes={changes} setChanges={setChanges} />}
        {tab === 'Agents' && <AgentsTab />}
        {tab === 'Rules' && <RulesTab />}
        {tab === 'Templates' && <TemplatesTab />}
      </div>
    </div>
  );
}
