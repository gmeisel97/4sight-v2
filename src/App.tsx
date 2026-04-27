import React, { useState, useRef, useEffect } from 'react';

/* eslint-disable */
/* global Office, Excel */

// ── Types ────────────────────────────────────────────────────────────────────
interface Message { role: 'ai' | 'user'; content: string; time: string; }
interface Change { id: number; cell: string; sheet: string; old: string; proposed: string; reason: string; status: 'pending'|'accepted'|'rejected'|'rethink'; type: 'formula'|'input'; }
interface Rule { id: number; name: string; type: string; typeBg: string; typeCol: string; desc: string; trigger: string; code: string; on: boolean; }
interface Template { name: string; icon: string; desc: string; sheets: number; cells: number; cat: string; catBg: string; catCol: string; star: boolean; isnew?: boolean; }
interface DissectResult { drivers: { cell: string; label: string; value: string; controls: string; whyMatters: string; }[]; risks: { title: string; detail: string; impact: string; }[]; sensitivity: { input1: { cell: string; label: string; value: string; range: string; why: string; }; input2: { cell: string; label: string; value: string; range: string; why: string; }; }; }
interface CapTableInvestor { name: string; common: number; options: number; seed1: number; seed2: number; safe: number; seriesA: number; seriesB: number; invested: number; liquidationPref: number; seniority: number; }
interface ExitScenario { name: string; probability: number; exitRevenue: number; exitMultiple: number; exitDate: string; exitEV: number; }
interface AgentDef { id: number; name: string; icon: string; desc: string; scope: string[]; rules: string[]; on: boolean; }

// ── Colors ────────────────────────────────────────────────────────────────────
const G_DARK='#14532d',G_MID='#16a34a',G_LIGHT='#f0fdf4',G_BORDER='#bbf7d0';
const GR_BG='#f9fafb',GR_BD='#e5e7eb',GR_TX='#6b7280',TX='#111827';
const fmtDollar=(n:number)=>n>=1e6?`$${(n/1e6).toFixed(1)}M`:n>=1e3?`$${(n/1e3).toFixed(0)}K`:`$${n.toFixed(0)}`;
const fmtMoic=(n:number)=>`${n.toFixed(2)}x`;

// ── UI Helpers ────────────────────────────────────────────────────────────────
function Toggle({on,onToggle}:{on:boolean;onToggle:()=>void}){
  return <button onClick={onToggle} style={{width:44,height:24,borderRadius:12,border:'none',cursor:'pointer',background:on?G_MID:'#d1d5db',position:'relative',flexShrink:0}}><span style={{position:'absolute',top:2,left:on?22:2,width:20,height:20,borderRadius:'50%',background:'white',display:'block',transition:'left 0.2s'}}/></button>;
}
function Badge({label,bg='#f3f4f6',col='#374151'}:{label:string;bg?:string;col?:string}){
  return <span style={{background:bg,color:col,fontSize:11,fontWeight:600,padding:'2px 8px',borderRadius:12,display:'inline-block'}}>{label}</span>;
}
function SHdr({color,label}:{color:string;label:string}){
  return <div style={{marginBottom:8}}><span style={{background:color,color:'white',fontSize:10,fontWeight:700,padding:'2px 8px',borderRadius:4}}>{label}</span></div>;
}
function Input({label,value,onChange,type='text',placeholder=''}:{label:string;value:string;onChange:(v:string)=>void;type?:string;placeholder?:string}){
  return(
    <div style={{marginBottom:8}}>
      <div style={{fontSize:11,fontWeight:600,color:GR_TX,marginBottom:3}}>{label}</div>
      <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{width:'100%',padding:'6px 10px',border:`1px solid ${GR_BD}`,borderRadius:7,fontSize:12,outline:'none',boxSizing:'border-box' as any}}/>
    </div>
  );
}
function Btn({label,onClick,color=G_MID,small=false,disabled=false}:{label:string;onClick:()=>void;color?:string;small?:boolean;disabled?:boolean}){
  return <button onClick={onClick} disabled={disabled} style={{background:disabled?'#d1d5db':color,color:'white',border:'none',borderRadius:7,padding:small?'5px 10px':'7px 14px',fontSize:small?11:12,fontWeight:600,cursor:disabled?'default':'pointer'}}>{label}</button>;
}

// ── Excel Helpers ─────────────────────────────────────────────────────────────
async function readAllSheets(){
  try{
    return await Excel.run(async(ctx)=>{
      const sheets=ctx.workbook.worksheets;
      sheets.load('items/name');await ctx.sync();
      const allCells:{address:string;value:string;formula:string;sheet:string}[]=[];
      for(const sheet of sheets.items){
        try{
          const range=sheet.getUsedRange();
          range.load(['values','formulas','address']);await ctx.sync();
          const vals=range.values as any[][],fmls=range.formulas as any[][];
          for(let r=0;r<Math.min(vals.length,80);r++)for(let c=0;c<Math.min(vals[r].length,20);c++){
            if(vals[r][c]!==''&&vals[r][c]!==null)
              allCells.push({address:`${String.fromCharCode(65+c)}${r+1}`,value:String(vals[r][c]),formula:String(fmls[r][c]),sheet:sheet.name});
          }
        }catch{}
      }
      return allCells;
    });
  }catch{return[];}
}

async function readActiveSheet(){
  try{
    return await Excel.run(async(ctx)=>{
      const range=ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
      range.load(['values','formulas','address']);await ctx.sync();
      const cells:{address:string;value:string;formula:string}[]=[];
      const vals=range.values as any[][],fmls=range.formulas as any[][];
      const base=range.address.split('!')[0];
      for(let r=0;r<vals.length;r++)for(let c=0;c<vals[r].length;c++){
        if(vals[r][c]!==''&&vals[r][c]!==null)
          cells.push({address:`${base}!${String.fromCharCode(65+c)}${r+1}`,value:String(vals[r][c]),formula:String(fmls[r][c])});
      }
      return cells;
    });
  }catch{return[];}
}

async function applyChange(change:Change):Promise<boolean>{
  try{
    await Excel.run(async(ctx)=>{
      const range=ctx.workbook.worksheets.getActiveWorksheet().getRange(change.cell);
      if(change.proposed.startsWith('=')){range.formulas=[[change.proposed]];}
      else{const num=parseFloat(change.proposed);range.values=[[isNaN(num)?change.proposed:num]];}
      await ctx.sync();
    });return true;
  }catch{return false;}
}

async function callClaude(userMessage:string,sheetContext:string,apiKey:string):Promise<Change[]>{
  const r=await fetch('/api/claude',{method:'POST',
    headers:{'Content-Type':'application/json','anthropic-version':'2023-06-01'},
    body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:1000,
      system:`You are 4SIGHT, an AI Excel co-pilot. Respond ONLY with a JSON array of proposed cell changes with fields: cell, proposed, reason, type ("formula" or "input"). If no changes needed return []. Never return anything other than valid JSON.`,
      messages:[{role:'user',content:`Sheet:\n${sheetContext}\n\nRequest: ${userMessage}`}]})});
  if(!r.ok){const e=await r.text();throw new Error(`API error: ${r.status} - ${e}`);}
  const d=await r.json();
  return JSON.parse(d.content[0].text.trim().replace(/```json|```/g,'').trim());
}

async function claudeJSON(prompt:string,system:string,apiKey:string):Promise<any>{
  const r=await fetch('/api/claude',{method:'POST',
    headers:{'Content-Type':'application/json','anthropic-version':'2023-06-01'},
    body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:2000,system,messages:[{role:'user',content:prompt}]})});
  if(!r.ok){const e=await r.text();throw new Error(`API error: ${r.status} - ${e}`);}
  const d=await r.json();
  const raw=d.content[0].text.trim();
  const match=raw.match(/\{[\s\S]*\}|\[[\s\S]*\]/);
  if(!match)throw new Error('No JSON in response');
  return JSON.parse(match[0]);
}

// ── IB Formatting ─────────────────────────────────────────────────────────────
async function applyIBFormatting(onProgress:(msg:string)=>void):Promise<string>{
  try{
    await Excel.run(async(ctx)=>{
      const sheet=ctx.workbook.worksheets.getActiveWorksheet();
      onProgress('Turning off gridlines...');
      sheet.showGridlines=false;
      const usedRange=sheet.getUsedRange();
      usedRange.load(['rowCount','columnCount','values','formulas']);await ctx.sync();
      const rowCount=usedRange.rowCount,colCount=usedRange.columnCount;
      const values=usedRange.values as any[][],formulas=usedRange.formulas as any[][];
      onProgress(`Scanning ${rowCount} rows x ${colCount} columns...`);
      const sectionKw=['income statement','balance sheet','cash flow','comparable','valuation','assumptions','revenue','expenses'];
      const subtotalKw=['gross profit','ebitda','ebit','net income','total','earnings before','net interest','operating income'];
      const sectionRows:number[]=[],subtotalRows:number[]=[];
      for(let r=0;r<rowCount;r++){
        let rowText='';
        for(let c=0;c<Math.min(colCount,3);c++){if(values[r][c]!==null&&values[r][c]!=='')rowText+=String(values[r][c]).toLowerCase();}
        const nonEmpty=values[r].filter((v:any)=>v!==null&&v!=='').length;
        if(nonEmpty<=2&&rowText.length>2&&sectionKw.some(k=>rowText.includes(k)))sectionRows.push(r);
        else if(subtotalKw.some(k=>rowText.includes(k)))subtotalRows.push(r);
      }
      onProgress('Applying font colors...');
      for(let r=0;r<rowCount;r++){
        for(let c=0;c<colCount;c++){
          const val=values[r][c],formula=formulas[r][c];
          if(val===null||val==='')continue;
          const cell=usedRange.getCell(r,c);
          cell.format.font.name='Calibri';cell.format.font.size=12;
          if(sectionRows.includes(r)){cell.format.font.color='#FFFFFF';}
          else if(typeof formula==='string'&&formula.startsWith('=')){cell.format.font.color='#000000';}
          else if(typeof val==='number'){cell.format.font.color='#0000FF';}
          else{cell.format.font.color='#000000';}
        }
      }
      await ctx.sync();
      for(const r of sectionRows){const row=usedRange.getRow(r);row.format.fill.color='#244062';row.format.font.color='#FFFFFF';}
      await ctx.sync();
      for(const r of subtotalRows){const row=usedRange.getRow(r);row.format.borders.getItem('EdgeTop').style='Continuous';row.format.borders.getItem('EdgeTop').color='#000000';}
      await ctx.sync();
      for(let r=0;r<rowCount;r++)for(let c=0;c<colCount;c++){
        const val=values[r][c];
        if(typeof val!=='number'||val===0)continue;
        (usedRange.getCell(r,c) as any).numberFormat=Math.abs(val)<1?'0.0%':'#,##0.0_);(#,##0.0)';
      }
      await ctx.sync();
      sheet.getRange('A:A').format.columnWidth=20;sheet.getRange('B:B').format.columnWidth=200;
      for(let c=2;c<Math.min(colCount+2,20);c++)sheet.getRange(`${String.fromCharCode(65+c)}:${String.fromCharCode(65+c)}`).format.columnWidth=90;
      await ctx.sync();
    });
    return '✓ Done!\n\n• Gridlines OFF\n• Blue = hardcoded numbers\n• Black = formulas & text\n• Dark blue section headers\n• Subtotal borders added\n• Column widths set\n• Number formats applied';
  }catch(err){throw new Error(`Formatting failed: ${err instanceof Error?err.message:'Unknown'}`);}
}

// ── VC Model Dissector ────────────────────────────────────────────────────────
async function dissectModel(apiKey:string,onProgress:(msg:string)=>void):Promise<DissectResult>{
  onProgress('Reading spreadsheet...');
  const result=await Excel.run(async(ctx)=>{
    const sheet=ctx.workbook.worksheets.getActiveWorksheet();
    const usedRange=sheet.getUsedRange();
    usedRange.load(['rowCount','columnCount','values','formulas','address']);await ctx.sync();
    const rows=usedRange.rowCount,cols=usedRange.columnCount;
    const vals=usedRange.values as any[][],fmls=usedRange.formulas as any[][];
    const baseAddr=usedRange.address.split('!')[0];
    const cells:string[]=[];
    for(let r=0;r<Math.min(rows,80);r++){
      for(let c=0;c<Math.min(cols,15);c++){
        const val=vals[r][c],formula=fmls[r][c];
        if(val===null||val==='')continue;
        const col=String.fromCharCode(65+c);
        const addr=`${baseAddr}!${col}${r+1}`;
        const isFormula=typeof formula==='string'&&formula.startsWith('=');
        let label='';
        for(let lc=c-1;lc>=0;lc--){if(vals[r][lc]&&typeof vals[r][lc]==='string'){label=String(vals[r][lc]);break;}}
        if(!label&&r>0)for(let lc=c;lc>=0;lc--){if(vals[r-1][lc]&&typeof vals[r-1][lc]==='string'){label=String(vals[r-1][lc]);break;}}
        cells.push(`${addr} [${isFormula?'FORMULA':'INPUT'}] label:"${label}" value:${val}${isFormula?` formula:${formula}`:''}`);
      }
    }
    return cells.join('\n');
  });
  onProgress('Analyzing model with AI...');
  const sys=`You are 4SIGHT, an AI co-pilot for investment bankers and VCs analyzing financial models. Return a JSON object: {"drivers":[{"cell":"B5","label":"short name","value":"formatted value","controls":"what this drives","whyMatters":"quantified impact"}],"risks":[{"title":"short name","detail":"specific issue","impact":"financial impact"}],"sensitivity":{"input1":{"cell":"B3","label":"name","value":"current","range":"low to high","why":"reason"},"input2":{"cell":"B4","label":"name","value":"current","range":"low to high","why":"reason"}}} Rules: 3-5 hardcoded INPUT drivers with actual values. 2-4 risks. 2 sensitivity inputs. Return ONLY valid JSON.`;
  return await claudeJSON(`Analyze this financial model:\n\n${result}`,sys,apiKey) as DissectResult;
}

// ── Cap Table Agent ───────────────────────────────────────────────────────────
interface CapTableState { mode:'detect'|'existing'|'scratch'; existingData:CapTableInvestor[]; scratchInvestors:{name:string;shares:string;invested:string;round:string}[]; nextRound:{investment:string;preMoney:string;optionPool:string;liquidationPref:string;participating:boolean;roundName:string}; result:{headers:string[];rows:(string|number)[][];totalShares:number}|null; changes:Change[]; status:string; }

async function detectCapTable(apiKey:string):Promise<{hasCapTable:boolean;investors:CapTableInvestor[]}>{
  const cells=await readActiveSheet();
  if(cells.length===0)return{hasCapTable:false,investors:[]};
  const context=cells.slice(0,100).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
  const system=`Detect a cap table in this spreadsheet. Return JSON: {"hasCapTable":true/false,"investors":[{"name":"investor name","common":0,"options":0,"seed1":0,"seed2":0,"safe":0,"seriesA":0,"seriesB":0,"invested":0,"liquidationPref":1,"seniority":1}]} seniority: 1=most senior. If no cap table, return {"hasCapTable":false,"investors":[]}.`;
  try{return await claudeJSON(context,system,apiKey);}catch{return{hasCapTable:false,investors:[]};}
}

function CapTableAgent({apiKey,onChangesProposed}:{apiKey:string;onChangesProposed:(c:Change[])=>void}){
  const [state,setState]=useState<CapTableState>({mode:'detect',existingData:[],scratchInvestors:[],nextRound:{investment:'',preMoney:'',optionPool:'0.15',liquidationPref:'1',participating:false,roundName:'Series A'},result:null,changes:[],status:''});
  const [detecting,setDetecting]=useState(false);

  const detect=async()=>{
    if(!apiKey){setState(s=>({...s,status:'Please enter your API key first.'}));return;}
    setDetecting(true);setState(s=>({...s,status:'Scanning spreadsheet for cap table...'}));
    try{
      const res=await detectCapTable(apiKey);
      if(res.hasCapTable&&res.investors.length>0){setState(s=>({...s,mode:'existing',existingData:res.investors,status:`Found ${res.investors.length} investor classes.`}));}
      else{setState(s=>({...s,mode:'scratch',status:'No cap table found. Build one from scratch below.'}));}
    }catch(err){setState(s=>({...s,status:`Error: ${err instanceof Error?err.message:'Unknown'}`}));}
    setDetecting(false);
  };

  const computeExisting=()=>{
    const nr=state.nextRound;
    const investment=parseFloat(nr.investment)||0,preMoney=parseFloat(nr.preMoney)||0;
    const optionPool=parseFloat(nr.optionPool)||0.15;
    if(!investment||!preMoney){setState(s=>({...s,status:'Please enter investment and pre-money valuation.'}));return;}
    const totalShares=state.existingData.reduce((s,i)=>s+i.common+i.options+i.seed1+i.seed2+i.safe+i.seriesA+i.seriesB,0)||10000000;
    const pps=preMoney/totalShares;
    const newShares=Math.round(investment/pps);
    const optionShares=Math.round(totalShares*optionPool);
    const postShares=totalShares+newShares+optionShares;
    const headers=['Investor','Pre Shares','Pre %','Post Shares','Post %','Invested'];
    const rows=state.existingData.map(inv=>{
      const pre=inv.common+inv.options+inv.seed1+inv.seed2+inv.safe+inv.seriesA+inv.seriesB;
      return[inv.name,pre,(pre/totalShares*100).toFixed(1)+'%',pre,(pre/postShares*100).toFixed(1)+'%',inv.invested?fmtDollar(inv.invested):'-'];
    });
    rows.push([`New ${nr.roundName}`,0,'0.0%',newShares,(newShares/postShares*100).toFixed(1)+'%',fmtDollar(investment)]);
    rows.push(['New Options',0,'0.0%',optionShares,(optionShares/postShares*100).toFixed(1)+'%','-']);
    const changes:Change[]=[];
    const baseRow=50;
    [{cell:`B${baseRow}`,val:`${nr.roundName} Round Summary`},{cell:`D${baseRow+1}`,val:String(preMoney)},{cell:`D${baseRow+2}`,val:String(investment)},{cell:`D${baseRow+3}`,val:String(pps.toFixed(4))},{cell:`D${baseRow+4}`,val:String(newShares)},{cell:`D${baseRow+5}`,val:String(optionShares)},{cell:`D${baseRow+6}`,val:String(postShares)},{cell:`D${baseRow+7}`,val:String(preMoney+investment)}].forEach((x,i)=>
      changes.push({id:Date.now()+i,cell:x.cell,sheet:'Sheet1',old:'',proposed:x.val,reason:'Round calculation',status:'pending',type:'input'}));
    setState(s=>({...s,result:{headers,rows,totalShares:postShares},changes,status:'Preview below — click "Write to Spreadsheet" to push to Changes tab.'}));
  };

  const writeScratch=()=>{
    const changes:Change[]=[];
    let row=4;
    ['Investor','Common','Options','Seed-1','Seed-2','SAFE','Series A','Series B','Total','%','$ Invested'].forEach((h,i)=>
      changes.push({id:Date.now()+i,cell:`${String.fromCharCode(66+i)}${row}`,sheet:'Sheet1',old:'',proposed:h,reason:'Header',status:'pending',type:'input'}));
    row++;
    state.scratchInvestors.forEach((inv,idx)=>{
      const shares=parseFloat(inv.shares)||0;
      const cols=[inv.name,inv.round==='Common'?shares:0,inv.round==='Options'?shares:0,inv.round==='Seed-1'?shares:0,inv.round==='Seed-2'?shares:0,inv.round==='SAFE'?shares:0,inv.round==='Series A'?shares:0,inv.round==='Series B'?shares:0,shares,'TBD',inv.invested||'0'];
      cols.forEach((v,ci)=>changes.push({id:Date.now()+idx*20+ci+100,cell:`${String.fromCharCode(66+ci)}${row+idx}`,sheet:'Sheet1',old:'',proposed:String(v),reason:`${inv.name}`,status:'pending',type:'input'}));
    });
    onChangesProposed(changes);setState(s=>({...s,status:'Cap table sent to Changes tab.'}));
  };

  const s=state;
  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      {s.status&&<div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:8,padding:'8px 12px',fontSize:11,color:G_DARK}}>{s.status}</div>}
      {s.mode==='detect'&&(
        <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
          <div style={{fontSize:13,fontWeight:700,color:TX,marginBottom:6}}>Detect or Build Cap Table</div>
          <div style={{fontSize:11,color:GR_TX,marginBottom:12}}>Scan your active sheet for an existing cap table, or build one from scratch.</div>
          <div style={{display:'flex',gap:8}}>
            <Btn label={detecting?'Scanning...':'Scan Active Sheet'} onClick={detect} disabled={detecting}/>
            <Btn label='Build From Scratch' onClick={()=>setState(s=>({...s,mode:'scratch',status:'Add your investors below.'}))} color='#6b7280'/>
          </div>
        </div>
      )}
      {s.mode==='existing'&&(
        <>
          <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
            <SHdr color='#244062' label='EXISTING INVESTORS'/>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead><tr>{['Investor','Shares','Invested'].map(h=><th key={h} style={{textAlign:'left',padding:'4px 8px',background:'#f3f4f6',fontWeight:700}}>{h}</th>)}</tr></thead>
              <tbody>{s.existingData.map((inv,i)=>{const shares=inv.common+inv.options+inv.seed1+inv.seed2+inv.safe+inv.seriesA+inv.seriesB;return<tr key={i} style={{borderBottom:`1px solid ${GR_BD}`}}><td style={{padding:'4px 8px'}}>{inv.name}</td><td style={{padding:'4px 8px',color:GR_TX}}>{shares.toLocaleString()}</td><td style={{padding:'4px 8px',color:'#0000FF'}}>{inv.invested?fmtDollar(inv.invested):'-'}</td></tr>;})}</tbody>
            </table>
          </div>
          <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
            <SHdr color='#244062' label='MODEL NEXT ROUND'/>
            <Input label='Round Name' value={s.nextRound.roundName} onChange={v=>setState(ss=>({...ss,nextRound:{...ss.nextRound,roundName:v}}))}/>
            <Input label='Investment Amount ($)' value={s.nextRound.investment} onChange={v=>setState(ss=>({...ss,nextRound:{...ss.nextRound,investment:v}}))} type='number'/>
            <Input label='Pre-Money Valuation ($)' value={s.nextRound.preMoney} onChange={v=>setState(ss=>({...ss,nextRound:{...ss.nextRound,preMoney:v}}))} type='number'/>
            <Input label='Target Option Pool (%)' value={s.nextRound.optionPool} onChange={v=>setState(ss=>({...ss,nextRound:{...ss.nextRound,optionPool:v}}))} type='number'/>
            <Input label='Liquidation Pref (x)' value={s.nextRound.liquidationPref} onChange={v=>setState(ss=>({...ss,nextRound:{...ss.nextRound,liquidationPref:v}}))} type='number'/>
            <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:10}}>
              <Toggle on={s.nextRound.participating} onToggle={()=>setState(ss=>({...ss,nextRound:{...ss.nextRound,participating:!ss.nextRound.participating}}))}/>
              <span style={{fontSize:12,color:TX}}>Participating Preferred</span>
            </div>
            <Btn label='Preview New Round' onClick={computeExisting}/>
          </div>
          {s.result&&(
            <div style={{background:'white',border:`1px solid ${G_BORDER}`,borderRadius:10,padding:14}}>
              <SHdr color='#14532d' label='PREVIEW'/>
              <div style={{overflowX:'auto',marginBottom:12}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                  <thead><tr>{s.result.headers.map(h=><th key={h} style={{textAlign:'left',padding:'4px 6px',background:'#244062',color:'white',fontWeight:700}}>{h}</th>)}</tr></thead>
                  <tbody>{s.result.rows.map((row,i)=><tr key={i} style={{borderBottom:`1px solid ${GR_BD}`,background:i%2===0?'white':'#f9fafb'}}>{row.map((cell,j)=><td key={j} style={{padding:'4px 6px',color:j===0?TX:GR_TX}}>{cell}</td>)}</tr>)}</tbody>
                </table>
              </div>
              <Btn label='Write to Spreadsheet' onClick={()=>{onChangesProposed(s.changes);setState(ss=>({...ss,status:'Sent to Changes tab for approval.'}));}}/>
            </div>
          )}
        </>
      )}
      {s.mode==='scratch'&&(
        <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
          <SHdr color='#244062' label='BUILD CAP TABLE'/>
          {s.scratchInvestors.map((inv,i)=>(
            <div key={i} style={{background:'#f9fafb',borderRadius:8,padding:10,marginBottom:8}}>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                <Input label='Investor Name' value={inv.name} onChange={v=>setState(ss=>{const arr=[...ss.scratchInvestors];arr[i]={...arr[i],name:v};return{...ss,scratchInvestors:arr};})}/>
                <div style={{marginBottom:8}}>
                  <div style={{fontSize:11,fontWeight:600,color:GR_TX,marginBottom:3}}>Round</div>
                  <select value={inv.round} onChange={e=>setState(ss=>{const arr=[...ss.scratchInvestors];arr[i]={...arr[i],round:e.target.value};return{...ss,scratchInvestors:arr};})} style={{width:'100%',padding:'6px 10px',border:`1px solid ${GR_BD}`,borderRadius:7,fontSize:12,outline:'none'}}>
                    {['Common','Options','Seed-1','Seed-2','SAFE','Series A','Series B'].map(r=><option key={r}>{r}</option>)}
                  </select>
                </div>
                <Input label='Shares' value={inv.shares} onChange={v=>setState(ss=>{const arr=[...ss.scratchInvestors];arr[i]={...arr[i],shares:v};return{...ss,scratchInvestors:arr};})} type='number'/>
                <Input label='$ Invested' value={inv.invested} onChange={v=>setState(ss=>{const arr=[...ss.scratchInvestors];arr[i]={...arr[i],invested:v};return{...ss,scratchInvestors:arr};})} type='number'/>
              </div>
            </div>
          ))}
          <div style={{display:'flex',gap:8,marginBottom:12}}>
            <Btn label='+ Add Investor' onClick={()=>setState(s=>({...s,scratchInvestors:[...s.scratchInvestors,{name:'',shares:'',invested:'',round:'Common'}]}))} small/>
            <Btn label='Reset' onClick={()=>setState(s=>({...s,scratchInvestors:[]}))} color='#6b7280' small/>
          </div>
          {s.scratchInvestors.length>0&&<Btn label='Write Cap Table to Spreadsheet' onClick={writeScratch}/>}
        </div>
      )}
    </div>
  );
}

// ── Exit Analysis Agent ───────────────────────────────────────────────────────
interface ExitWaterfallScenario { scenario:string; rows:{investor:string;liqPref:number;seniority:number;prefProceeds:number;commonProceeds:number;total:number;moic:number;irr:number}[]; weightedMoic:number; }
function calcIRR(invested:number,proceeds:number,exitDateStr:string,entryDateStr:string='2022-01-01'):number{
  try{
    const entry=new Date(entryDateStr),exit=new Date(exitDateStr);
    const years=(exit.getTime()-entry.getTime())/(365.25*24*60*60*1000);
    if(years<=0||invested<=0)return 0;
    return Math.pow(proceeds/invested,1/years)-1;
  }catch{return 0;}
}

async function runExitAnalysis(apiKey:string,cells:{address:string;value:string;formula:string;sheet?:string}[],scenarios:ExitScenario[]):Promise<ExitWaterfallScenario[]>{
  const context=cells.slice(0,200).map(c=>`[${c.sheet||'Sheet'}] ${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
  const system=`Extract investor data from this cap table spreadsheet. Return JSON: {"investors":[{"name":"investor name","totalShares":number,"invested":number,"liquidationPref":number,"seniority":number}],"totalShares":number} seniority: 1=most senior. liquidationPref = total liquidation preference dollars. Return ONLY valid JSON.`;
  const capData=await claudeJSON(context,system,apiKey);
  const investors=capData.investors||[];
  const totalShares=capData.totalShares||investors.reduce((s:number,i:any)=>s+i.totalShares,0)||1;
  return scenarios.map(scenario=>{
    let remaining=scenario.exitEV;
    const proceeds:Record<string,{pref:number;common:number}>={};
    investors.forEach((i:any)=>{proceeds[i.name]={pref:0,common:0};});
    for(const seniority of[1,2,3,4]){
      const group=investors.filter((i:any)=>i.seniority===seniority);
      if(!group.length)continue;
      const totalPref=group.reduce((s:number,i:any)=>s+i.liquidationPref,0);
      if(remaining>=totalPref){group.forEach((i:any)=>{proceeds[i.name].pref=i.liquidationPref;});remaining-=totalPref;}
      else{group.forEach((i:any)=>{proceeds[i.name].pref=remaining*(i.liquidationPref/totalPref);});remaining=0;break;}
    }
    if(remaining>0)investors.forEach((i:any)=>{proceeds[i.name].common=remaining*(i.totalShares/totalShares);});
    const rows=investors.map((i:any)=>{const total=proceeds[i.name].pref+proceeds[i.name].common;const moic=i.invested>0?total/i.invested:0;const irr=i.invested>0?calcIRR(i.invested,total,scenario.exitDate):0;return{investor:i.name,liqPref:i.liquidationPref,seniority:i.seniority,prefProceeds:proceeds[i.name].pref,commonProceeds:proceeds[i.name].common,total,moic,irr};});
    const eipRow=rows.find(r=>r.investor.toLowerCase().includes('eip'));
    return{scenario:scenario.name,rows,weightedMoic:eipRow?eipRow.moic:rows.reduce((s,r)=>s+r.moic,0)/Math.max(rows.length,1)};
  });
}

function ExitAnalysisAgent({apiKey,onChangesProposed}:{apiKey:string;onChangesProposed:(c:Change[])=>void}){
  const [scenarios,setScenarios]=useState<ExitScenario[]>([
    {name:'Downside',probability:0.2,exitRevenue:0,exitMultiple:0,exitDate:'2026-12-31',exitEV:20000000},
    {name:'Base',probability:0.5,exitRevenue:0,exitMultiple:0,exitDate:'2028-12-31',exitEV:200000000},
    {name:'Upside',probability:0.3,exitRevenue:0,exitMultiple:0,exitDate:'2030-12-31',exitEV:500000000},
  ]);
  const [entryDate,setEntryDate]=useState('2022-01-01');
  const [scanning,setScanning]=useState(false);
  const [scanNote,setScanNote]=useState('');
  const [waterfall,setWaterfall]=useState<ExitWaterfallScenario[]>([]);
  const [status,setStatus]=useState('');
  const [running,setRunning]=useState(false);

  const scanSheets=async()=>{
    if(!apiKey){setScanNote('Enter API key first.');return;}
    setScanning(true);setScanNote('Scanning all sheets...');
    try{
      const cells=await readAllSheets();
      const context=cells.slice(0,150).map(c=>`[${c.sheet}] ${c.address}: ${c.value}`).join('\n');
      const found=await claudeJSON(context,`Scan this workbook for exit-related data. Return JSON: {"exitRevenue":number_or_0,"exitMultiple":number_or_0,"exitDate":"YYYY-MM-DD or empty","entryDate":"YYYY-MM-DD or empty","sheetsFound":["list of relevant sheet names"],"notes":"brief summary of what was found"} Return 0 if not found. Return ONLY valid JSON.`,apiKey);
      let note=`Scanned all sheets. `;
      if(found.sheetsFound?.length)note+=`Found relevant data in: ${found.sheetsFound.join(', ')}. `;
      if(found.notes)note+=found.notes;
      setScanNote(note);
      if(found.entryDate)setEntryDate(found.entryDate);
      if(found.exitRevenue||found.exitMultiple||found.exitDate){
        setScenarios(s=>s.map(sc=>({...sc,
          exitRevenue:found.exitRevenue&&sc.exitRevenue===0?found.exitRevenue:sc.exitRevenue,
          exitMultiple:found.exitMultiple&&sc.exitMultiple===0?found.exitMultiple:sc.exitMultiple,
          exitDate:found.exitDate&&!sc.exitDate?found.exitDate:sc.exitDate,
        })));
      }
    }catch(err){setScanNote(`Scan error: ${err instanceof Error?err.message:'Unknown'}`);}
    setScanning(false);
  };

  const run=async()=>{
    if(!apiKey){setStatus('Please enter your API key first.');return;}
    setRunning(true);setStatus('Reading cap table across all sheets...');
    try{
      const cells=await readAllSheets();
      const resolvedScenarios=scenarios.map(sc=>({...sc,exitEV:sc.exitRevenue&&sc.exitMultiple?sc.exitRevenue*sc.exitMultiple:sc.exitEV}));
      setStatus('Computing exit waterfalls...');
      setWaterfall(await runExitAnalysis(apiKey,cells,resolvedScenarios));
      setStatus('Preview below. Click "Write to Spreadsheet" to push the waterfall table.');
    }catch(err){setStatus(`Error: ${err instanceof Error?err.message:'Unknown'}`);}
    setRunning(false);
  };

  const writeToSheet=()=>{
    const changes:Change[]=[];let row=70;
    changes.push({id:Date.now(),cell:`B${row}`,sheet:'Sheet1',old:'',proposed:'Exit Waterfall Analysis',reason:'Header',status:'pending',type:'input'});row+=2;
    waterfall.forEach((sc,si)=>{
      changes.push({id:Date.now()+si*100,cell:`B${row}`,sheet:'Sheet1',old:'',proposed:`Scenario: ${sc.scenario}`,reason:'Scenario',status:'pending',type:'input'});row++;
      ['Investor','Liq. Pref','Pref $','Common $','Total','MOIC'].forEach((h,hi)=>changes.push({id:Date.now()+si*100+hi+10,cell:`${String.fromCharCode(66+hi)}${row}`,sheet:'Sheet1',old:'',proposed:h,reason:'Header',status:'pending',type:'input'}));row++;
      sc.rows.forEach((r,ri)=>[r.investor,fmtDollar(r.liqPref),fmtDollar(r.prefProceeds),fmtDollar(r.commonProceeds),fmtDollar(r.total),fmtMoic(r.moic)].forEach((v,vi)=>changes.push({id:Date.now()+si*100+ri*10+vi+50,cell:`${String.fromCharCode(66+vi)}${row+ri}`,sheet:'Sheet1',old:'',proposed:String(v),reason:`${sc.scenario} waterfall`,status:'pending',type:'input'})));
      row+=sc.rows.length+2;
    });
    onChangesProposed(changes);setStatus('Waterfall sent to Changes tab for approval.');
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      {status&&<div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:8,padding:'8px 12px',fontSize:11,color:G_DARK}}>{status}</div>}
      <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
        <SHdr color='#244062' label='EXIT SCENARIOS'/>
        {scenarios.map((sc,i)=>(
          <div key={i} style={{background:'#f9fafb',borderRadius:8,padding:10,marginBottom:8}}>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:6}}>
              <Input label='Name' value={sc.name} onChange={v=>setScenarios(s=>{const a=[...s];a[i]={...a[i],name:v};return a;})}/>
              <Input label='Exit EV ($)' value={String(sc.exitEV)} onChange={v=>setScenarios(s=>{const a=[...s];a[i]={...a[i],exitEV:parseFloat(v)||0};return a;})} type='number'/>
              <Input label='Probability' value={String(sc.probability)} onChange={v=>setScenarios(s=>{const a=[...s];a[i]={...a[i],probability:parseFloat(v)||0};return a;})} type='number'/>
            </div>
          </div>
        ))}
        <Btn label='+ Add Scenario' onClick={()=>setScenarios(s=>[...s,{name:'Scenario',probability:0.1,exitRevenue:0,exitMultiple:0,exitDate:'',exitEV:100000000}])} small/>
      </div>
      <Btn label={running?'Running...':'Run Exit Analysis'} onClick={run} disabled={running}/>
      {waterfall.length>0&&(
        <>
          {waterfall.map((sc,si)=>(
            <div key={si} style={{background:'white',border:`1px solid ${G_BORDER}`,borderRadius:10,padding:14}}>
              <SHdr color='#244062' label={`${sc.scenario.toUpperCase()} SCENARIO`}/>
              <div style={{overflowX:'auto',marginBottom:8}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                  <thead><tr>{['Investor','Liq. Pref','Pref $','Common $','Total','MoM','IRR'].map(h=><th key={h} style={{textAlign:'left',padding:'4px 6px',background:'#244062',color:'white',fontWeight:700}}>{h}</th>)}</tr></thead>
                  <tbody>{sc.rows.map((r,ri)=><tr key={ri} style={{borderBottom:`1px solid ${GR_BD}`,background:ri%2===0?'white':'#f9fafb'}}>
                    <td style={{padding:'4px 6px',color:TX,fontWeight:600}}>{r.investor}</td>
                    <td style={{padding:'4px 6px',color:GR_TX}}>{fmtDollar(r.liqPref)}</td>
                    <td style={{padding:'4px 6px',color:GR_TX}}>{fmtDollar(r.prefProceeds)}</td>
                    <td style={{padding:'4px 6px',color:GR_TX}}>{fmtDollar(r.commonProceeds)}</td>
                    <td style={{padding:'4px 6px',color:TX,fontWeight:700}}>{fmtDollar(r.total)}</td>
                    <td style={{padding:'4px 6px',color:r.moic>=2?G_MID:r.moic<1?'#dc2626':TX,fontWeight:700}}>{fmtMoic(r.moic)}</td>
                    <td style={{padding:'4px 6px',color:GR_TX}}>{r.irr?`${(r.irr*100).toFixed(1)}%`:'-'}</td>
                  </tr>)}</tbody>
                </table>
              </div>
              <div style={{fontSize:11,color:GR_TX}}>Weighted MoIC: <span style={{fontWeight:700,color:TX}}>{fmtMoic(sc.weightedMoic)}</span></div>
            </div>
          ))}
          <Btn label='Write Waterfall to Spreadsheet' onClick={writeToSheet}/>
        </>
      )}
    </div>
  );
}

// ── DCF Valuation Agent ───────────────────────────────────────────────────────
function DCFAgent({apiKey}:{apiKey:string}){
  const [wacc,setWacc]=useState('12');
  const [termGrowth,setTermGrowth]=useState('3');
  const [result,setResult]=useState<{ev:number;tv:number;evLow:number;evHigh:number;rows:{year:number;fcf:number;pv:number}[]}|null>(null);
  const [status,setStatus]=useState('');
  const [running,setRunning]=useState(false);

  const run=async()=>{
    if(!apiKey){setStatus('Please enter your API key first.');return;}
    setRunning(true);setStatus('Reading cash flow projections...');
    try{
      const cells=await readActiveSheet();
      const context=cells.slice(0,100).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
      const data=await claudeJSON(context,`Extract free cash flow projections. Return JSON: {"fcfs":[{"year":1,"fcf":number},...], "netDebt":number} Use EBITDA-taxes-capex-WC if FCF not explicit. Return ONLY valid JSON.`,apiKey);
      const fcfs:number[]=(data.fcfs||[]).map((f:any)=>f.fcf);
      const netDebt=data.netDebt||0;
      const w=parseFloat(wacc)/100,g=parseFloat(termGrowth)/100;
      const rows=fcfs.map((fcf,i)=>({year:i+1,fcf,pv:fcf/Math.pow(1+w,i+1)}));
      const pvSum=rows.reduce((s,r)=>s+r.pv,0);
      const lastFcf=fcfs[fcfs.length-1]||0;
      const tv=lastFcf*(1+g)/(w-g);
      const pvTv=tv/Math.pow(1+w,fcfs.length);
      const ev=pvSum+pvTv-netDebt;
      setResult({ev,tv,evLow:ev*0.85,evHigh:ev*1.15,rows});
      setStatus('DCF complete.');
    }catch(err){setStatus(`Error: ${err instanceof Error?err.message:'Unknown'}`);}
    setRunning(false);
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      {status&&<div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:8,padding:'8px 12px',fontSize:11,color:G_DARK}}>{status}</div>}
      <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
        <SHdr color='#244062' label='DCF ASSUMPTIONS'/>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
          <Input label='WACC (%)' value={wacc} onChange={setWacc} type='number'/>
          <Input label='Terminal Growth (%)' value={termGrowth} onChange={setTermGrowth} type='number'/>
        </div>
        <div style={{fontSize:11,color:GR_TX,marginBottom:10}}>4SIGHT reads FCF projections from your active sheet.</div>
        <Btn label={running?'Running DCF...':'Run DCF Analysis'} onClick={run} disabled={running}/>
      </div>
      {result&&(
        <div style={{background:'white',border:`1px solid ${G_BORDER}`,borderRadius:10,padding:14}}>
          <SHdr color='#14532d' label='DCF RESULTS'/>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:12}}>
            {[['Enterprise Value',fmtDollar(result.ev)],['Terminal Value',fmtDollar(result.tv)],['EV Low (-15%)',fmtDollar(result.evLow)],['EV High (+15%)',fmtDollar(result.evHigh)]].map(([l,v])=>(
              <div key={l} style={{background:'#f9fafb',borderRadius:8,padding:'10px 12px'}}>
                <div style={{fontSize:10,color:GR_TX,marginBottom:3}}>{l}</div>
                <div style={{fontSize:16,fontWeight:800,color:TX}}>{v}</div>
              </div>
            ))}
          </div>
          <SHdr color='#6b7280' label='FCF BY YEAR'/>
          <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
            <thead><tr>{['Year','FCF','PV of FCF'].map(h=><th key={h} style={{textAlign:'left',padding:'4px 6px',background:'#f3f4f6',fontWeight:700}}>{h}</th>)}</tr></thead>
            <tbody>{result.rows.map(r=><tr key={r.year} style={{borderBottom:`1px solid ${GR_BD}`}}>
              <td style={{padding:'4px 6px'}}>{`Year ${r.year}`}</td>
              <td style={{padding:'4px 6px',color:'#0000FF'}}>{fmtDollar(r.fcf)}</td>
              <td style={{padding:'4px 6px',color:GR_TX}}>{fmtDollar(r.pv)}</td>
            </tr>)}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ── VC Valuation Agent ────────────────────────────────────────────────────────
function VCValuationAgent({apiKey}:{apiKey:string}){
  const [investment,setInvestment]=useState('');
  const [ownership,setOwnership]=useState('');
  const [targetMoic,setTargetMoic]=useState('10');
  const [exitMultiple,setExitMultiple]=useState('');
  const [exitYear,setExitYear]=useState('5');
  const [autoDetectYear,setAutoDetectYear]=useState(false);
  const [result,setResult]=useState<{requiredExitEV:number;requiredRevenue:number;projectedRevenue:number;impliedMoic:number;verdict:string;marketShareNote:string;detectedYear:number}|null>(null);
  const [status,setStatus]=useState('');
  const [running,setRunning]=useState(false);

  const run=async()=>{
    if(!apiKey){setStatus('Please enter your API key first.');return;}
    if(!investment||!ownership||!exitMultiple){setStatus('Please fill in investment, ownership, and exit multiple.');return;}
    setRunning(true);setStatus('Reading revenue projections...');
    try{
      const cells=await readActiveSheet();
      const context=cells.slice(0,100).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
      const data=await claudeJSON(context,`Extract revenue projections. Return JSON: {"revenueByYear":[{"year":1,"revenue":number},...], "currentRevenue":number, "tam":number, "suggestedExitYear":number} suggestedExitYear = year revenue growth decelerates most, or last forecast year. Return ONLY valid JSON.`,apiKey);
      const inv=parseFloat(investment),own=parseFloat(ownership)/100,moic=parseFloat(targetMoic),mult=parseFloat(exitMultiple);
      const revenueByYear:any[]=data.revenueByYear||[];
      const detectedYear=autoDetectYear?(data.suggestedExitYear||revenueByYear.length):parseInt(exitYear);
      const projectedRevenue=revenueByYear[detectedYear-1]?.revenue||revenueByYear[revenueByYear.length-1]?.revenue||0;
      const requiredExitEV=(inv*moic)/own;
      const requiredRevenue=requiredExitEV/mult;
      const impliedMoic=projectedRevenue>0?(projectedRevenue*mult*own)/inv:0;
      const tam=data.tam||0;
      const impliedMktShare=tam>0?requiredRevenue/tam:null;
      let verdict='';
      if(impliedMoic>=moic*0.9){verdict='✓ The model supports the target return. Projected revenue is sufficient at this multiple.';}
      else if(impliedMoic>=moic*0.6){verdict='⚠ The model partially supports the return target. Revenue projections would need to improve or the multiple expand.';}
      else{verdict='✗ The model does not support the target return. A significant gap exists between projected and required revenue.';}
      const marketShareNote=impliedMktShare?`Required revenue implies ${(impliedMktShare*100).toFixed(1)}% market share of the identified TAM.`:'Could not calculate implied market share — TAM not found in model.';
      setResult({requiredExitEV,requiredRevenue,projectedRevenue,impliedMoic,verdict,marketShareNote,detectedYear});
      setStatus('Analysis complete.');
    }catch(err){setStatus(`Error: ${err instanceof Error?err.message:'Unknown'}`);}
    setRunning(false);
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      {status&&<div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:8,padding:'8px 12px',fontSize:11,color:G_DARK}}>{status}</div>}
      <div style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:14}}>
        <SHdr color='#244062' label='DEAL TERMS'/>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
          <Input label='Investment ($)' value={investment} onChange={setInvestment} type='number' placeholder='e.g. 12000000'/>
          <Input label='Ownership at Entry (%)' value={ownership} onChange={setOwnership} type='number' placeholder='e.g. 16'/>
          <Input label='Target MOIC' value={targetMoic} onChange={setTargetMoic} type='number' placeholder='e.g. 10'/>
          <Input label='EV/Revenue Exit Multiple' value={exitMultiple} onChange={setExitMultiple} type='number' placeholder='e.g. 3'/>
        </div>
        <div style={{marginBottom:10}}>
          <div style={{fontSize:11,fontWeight:600,color:GR_TX,marginBottom:3}}>Expected Exit Year</div>
          <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
            <Toggle on={autoDetectYear} onToggle={()=>setAutoDetectYear(v=>!v)}/>
            <span style={{fontSize:12,color:TX}}>{autoDetectYear?'Auto-detect from model':'Manual input'}</span>
          </div>
          {!autoDetectYear&&<input type='number' value={exitYear} onChange={e=>setExitYear(e.target.value)} placeholder='e.g. 5' style={{width:'100%',padding:'6px 10px',border:`1px solid ${GR_BD}`,borderRadius:7,fontSize:12,outline:'none',boxSizing:'border-box' as any}}/>}
        </div>
        <div style={{fontSize:11,color:GR_TX,marginBottom:10}}>4SIGHT reads revenue projections from your model to run the back-solve.</div>
        <Btn label={running?'Running...':'Run VC Valuation'} onClick={run} disabled={running}/>
      </div>
      {result&&(
        <div style={{background:'white',border:`1px solid ${G_BORDER}`,borderRadius:10,padding:14}}>
          <SHdr color='#14532d' label='VC BACK-SOLVE'/>
          {autoDetectYear&&<div style={{fontSize:11,color:GR_TX,background:'#f9fafb',borderRadius:8,padding:'8px 10px',marginBottom:10}}>📅 Auto-detected exit: <strong>Year {result.detectedYear}</strong></div>}
          <div style={{display:'flex',flexDirection:'column',gap:8,marginBottom:12}}>
            {([['Required Exit EV',fmtDollar(result.requiredExitEV),`Needed to return ${targetMoic}x`],['Required Revenue at Exit',fmtDollar(result.requiredRevenue),`At ${exitMultiple}x EV/Revenue`],[`Projected Revenue (Yr ${result.detectedYear})`,fmtDollar(result.projectedRevenue),'From model projections'],['Implied MOIC from Model',fmtMoic(result.impliedMoic),'Based on projected revenue & ownership']] as [string,string,string][]).map(([label,value,sub])=>(
              <div key={label} style={{background:'#f9fafb',borderRadius:8,padding:'10px 12px',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <div><div style={{fontSize:11,color:GR_TX}}>{label}</div><div style={{fontSize:10,color:GR_TX,marginTop:1}}>{sub}</div></div>
                <div style={{fontSize:16,fontWeight:800,color:TX,flexShrink:0,marginLeft:10}}>{value}</div>
              </div>
            ))}
          </div>
          <div style={{background:result.verdict.startsWith('✓')?'#f0fdf4':result.verdict.startsWith('⚠')?'#fffbeb':'#fef2f2',border:`1px solid ${result.verdict.startsWith('✓')?G_BORDER:result.verdict.startsWith('⚠')?'#fde68a':'#fca5a5'}`,borderRadius:8,padding:'10px 12px',marginBottom:8}}>
            <div style={{fontSize:12,fontWeight:600,color:result.verdict.startsWith('✓')?G_DARK:result.verdict.startsWith('⚠')?'#92400e':'#991b1b'}}>{result.verdict}</div>
          </div>
          <div style={{fontSize:11,color:GR_TX,background:'#f9fafb',borderRadius:8,padding:'8px 10px'}}>{result.marketShareNote}</div>
        </div>
      )}
    </div>
  );
}

// ── Sensitivity Builder ───────────────────────────────────────────────────────
async function buildSensitivity(apiKey:string,onProgress:(msg:string)=>void):Promise<string>{
  onProgress('Reading spreadsheet...');
  const cells=await readActiveSheet();
  if(cells.length===0)throw new Error('No data found.');
  onProgress('Identifying key assumptions...');
  const ctx=cells.slice(0,100).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
  const parsed=await claudeJSON(ctx,`Find the 2 most important hardcoded input cells for sensitivity analysis. Return ONLY JSON: {"input1":{"cell":"B3","label":"Revenue Growth","currentValue":0.18,"low":0.10,"mid":0.18,"high":0.26},"input2":{"cell":"B4","label":"Gross Margin","currentValue":0.68,"low":0.60,"mid":0.68,"high":0.76},"output":{"cell":"B10","label":"Net Income"}}`,apiKey);
  onProgress('Writing sensitivity table...');
  await Excel.run(async(context)=>{
    const sheet=context.workbook.worksheets.getActiveWorksheet();
    const usedRange=sheet.getUsedRange();usedRange.load(['rowCount']);await context.sync();
    const startRow=usedRange.rowCount+3;
    const hdr=sheet.getRangeByIndexes(startRow,1,1,4);
    hdr.values=[[`Sensitivity: ${parsed.input1.label} vs ${parsed.input2.label}`,'','','']];
    hdr.format.fill.color='#244062';hdr.format.font.color='#FFFFFF';hdr.format.font.name='Calibri';
    const colHdr=sheet.getRangeByIndexes(startRow+1,2,1,3);
    colHdr.values=[[parsed.input2.low,parsed.input2.mid,parsed.input2.high]];colHdr.format.font.bold=true;
    sheet.getRangeByIndexes(startRow+1,1,1,1).values=[[parsed.input2.label]];
    const i1Vals=[parsed.input1.low,parsed.input1.mid,parsed.input1.high];
    for(let i=0;i<3;i++){
      sheet.getRangeByIndexes(startRow+2+i,1,1,1).values=[[i1Vals[i]]];
      sheet.getRangeByIndexes(startRow+2+i,1,1,1).format.font.bold=true;
      for(let j=0;j<3;j++){
        const cell=sheet.getRangeByIndexes(startRow+2+i,2+j,1,1);cell.values=[['—']];
        if(i===1&&j===1)cell.format.fill.color='#DCE6F1';
      }
    }
    sheet.getRangeByIndexes(startRow+2,0,1,1).values=[[parsed.input1.label]];
    await context.sync();
  });
  return `✓ Sensitivity table built.\n\nRow axis: ${parsed.input1.label}\nColumn axis: ${parsed.input2.label}\nOutput: ${parsed.output.label}\n\nBlue cell = base case.`;
}

// ── Dissect Output Renderer ───────────────────────────────────────────────────
function DissectOutput({result}:{result:DissectResult}){
  return(
    <div style={{display:'flex',flexDirection:'column',gap:14}}>
      <div>
        <SHdr color='#244062' label='KEY DRIVERS'/>
        {result.drivers.map((d,i)=>(
          <div key={i} style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:10,padding:'10px 12px',marginBottom:8}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:6}}>
              <div style={{display:'flex',alignItems:'center',gap:7}}>
                <span style={{fontFamily:'monospace',fontSize:11,fontWeight:700,color:'#244062',background:'#eff6ff',padding:'2px 6px',borderRadius:4}}>{d.cell}</span>
                <span style={{fontSize:12,fontWeight:700,color:TX}}>{d.label}</span>
              </div>
              <span style={{fontSize:13,fontWeight:800,color:'#0000FF',flexShrink:0,marginLeft:8}}>{d.value}</span>
            </div>
            <div style={{fontSize:11,color:GR_TX,marginBottom:4}}><span style={{fontWeight:600,color:TX}}>Controls: </span>{d.controls}</div>
            <div style={{fontSize:11,color:GR_TX}}><span style={{fontWeight:600,color:TX}}>Why it matters: </span>{d.whyMatters}</div>
          </div>
        ))}
      </div>
      <div>
        <SHdr color='#7f1d1d' label='RISKS & FLAGS'/>
        {result.risks.map((r,i)=>(
          <div key={i} style={{background:'#fef2f2',border:'1px solid #fca5a5',borderRadius:10,padding:'10px 12px',marginBottom:8}}>
            <div style={{fontSize:12,fontWeight:700,color:'#991b1b',marginBottom:4}}>🚩 {r.title}</div>
            <div style={{fontSize:11,color:'#7f1d1d',marginBottom:4}}>{r.detail}</div>
            {r.impact&&<div style={{fontSize:11,color:'#991b1b',fontWeight:600}}>Impact: {r.impact}</div>}
          </div>
        ))}
      </div>
      <div>
        <SHdr color='#14532d' label='SENSITIVITY RECOMMENDATION'/>
        {[result.sensitivity.input1,result.sensitivity.input2].map((s,i)=>(
          <div key={i} style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:10,padding:'10px 12px',marginBottom:8}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
              <div style={{display:'flex',alignItems:'center',gap:7}}>
                <span style={{background:'#244062',color:'white',fontSize:10,fontWeight:700,padding:'2px 6px',borderRadius:4,fontFamily:'monospace'}}>{s.cell}</span>
                <span style={{fontSize:12,fontWeight:700,color:TX}}>{s.label}</span>
              </div>
              <div style={{textAlign:'right'}}>
                <div style={{fontSize:10,color:GR_TX}}>Current</div>
                <div style={{fontSize:13,fontWeight:800,color:'#0000FF'}}>{s.value}</div>
              </div>
            </div>
            <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
              <span style={{fontSize:10,fontWeight:700,color:GR_TX}}>STRESS RANGE</span>
              <span style={{fontSize:12,fontWeight:700,color:'#14532d',background:'#dcfce7',padding:'2px 10px',borderRadius:20}}>{s.range}</span>
            </div>
            <div style={{fontSize:11,color:GR_TX}}>{s.why}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Chat Tab ──────────────────────────────────────────────────────────────────
function ChatTab({onChangesProposed,apiKey}:{onChangesProposed:(c:Change[])=>void;apiKey:string}){
  const [msgs,setMsgs]=useState<Message[]>([{role:'ai',content:'👋 Welcome to 4SIGHT!\n\nPick a demo or ask me anything about your spreadsheet.\n\nEvery change I propose goes to the Changes tab first — you approve before anything touches your model.',time:new Date().toLocaleTimeString()}]);
  const [input,setInput]=useState('');
  const [loading,setLoading]=useState(false);
  const msgsEndRef=useRef<HTMLDivElement>(null);
  useEffect(()=>{msgsEndRef.current?.scrollIntoView({behavior:'smooth'});},[msgs]);
  const addMsg=(role:'ai'|'user',content:string)=>setMsgs(m=>[...m,{role,content,time:new Date().toLocaleTimeString()}]);

  const runDemo=async(key:string)=>{
    const labels:Record<string,string>={format:'Format This Model',dissect:'Dissect This Model',sensitivity:'Build a Sensitivity',explain:'Explain This Formula'};
    addMsg('user',`Run demo: ${labels[key]}`);setLoading(true);
    if(key==='format'){
      addMsg('ai','Running IB Formatting Agent...');
      try{addMsg('ai',await applyIBFormatting(()=>{}));}catch(err){addMsg('ai',`Error: ${err instanceof Error?err.message:'Unknown'}`);}
      setLoading(false);return;
    }
    if(key==='dissect'){
      if(!apiKey){addMsg('ai','Please enter your API key using the ⚙ button above.');setLoading(false);return;}
      addMsg('ai','Running VC Model Dissector — check the Agents tab for the full analysis.');
      setLoading(false);return;
    }
    if(key==='sensitivity'){
      if(!apiKey){addMsg('ai','Please enter your API key using the ⚙ button above.');setLoading(false);return;}
      addMsg('ai','Building sensitivity table...');
      try{addMsg('ai',await buildSensitivity(apiKey,()=>{}));}catch(err){addMsg('ai',`Error: ${err instanceof Error?err.message:'Unknown'}`);}
      setLoading(false);return;
    }
    if(key==='explain'){
      if(!apiKey){addMsg('ai','Please enter your API key using the ⚙ button above.');setLoading(false);return;}
      try{
        const cells=await readActiveSheet();
        const ctx=cells.slice(0,50).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n');
        const r=await fetch('/api/claude',{method:'POST',headers:{'Content-Type':'application/json','anthropic-version':'2023-06-01'},body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:300,system:'You are 4SIGHT. Find the most interesting formula in this spreadsheet and explain it in plain English in under 80 words.',messages:[{role:'user',content:ctx}]})});
        const d=await r.json();addMsg('ai',d.content[0].text);
      }catch(err){addMsg('ai',`Error: ${err instanceof Error?err.message:'Unknown'}`);}
      setLoading(false);return;
    }
  };

  const send=async()=>{
    if(!input.trim()||loading)return;
    const q=input;setInput('');addMsg('user',q);setLoading(true);
    try{
      const cells=await readActiveSheet();
      const ctx=cells.length>0?cells.slice(0,50).map(c=>`${c.address}: value=${c.value}${c.formula!==c.value?` formula=${c.formula}`:''}`).join('\n'):'No data found.';
      if(!apiKey){addMsg('ai','Please enter your API key using the ⚙ button above.');setLoading(false);return;}
      const changes=await callClaude(q,ctx,apiKey);
      if(changes.length===0){addMsg('ai','No changes needed for this request.');}
      else{const withIds=changes.map((c,i)=>({...c,id:Date.now()+i,sheet:'Sheet1',old:'(current value)',status:'pending' as const}));onChangesProposed(withIds);addMsg('ai',`I've proposed ${changes.length} change${changes.length>1?'s':''} in the Changes tab.`);}
    }catch(err){addMsg('ai',`Something went wrong: ${err instanceof Error?err.message:'Unknown error'}`);}
    setLoading(false);
  };

  const demos=[{key:'format',icon:'📐',label:'Format This Model',sub:'Apply IB formatting standards instantly'},{key:'dissect',icon:'🔍',label:'Dissect This Model',sub:'Surface key assumptions and drivers'},{key:'sensitivity',icon:'📊',label:'Build a Sensitivity',sub:'Two-variable data table'},{key:'explain',icon:'💡',label:'Explain This Formula',sub:'Narrate what the selected cell does'}];

  return(
    <div style={{display:'flex',flexDirection:'column',height:'100%',gap:10}}>
      <div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:10,padding:'10px 14px'}}>
        <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:8}}><svg width="12" height="12" viewBox="0 0 24 24" fill={G_MID}><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg><span style={{fontSize:13,fontWeight:700,color:G_MID}}>Try Interactive Demos</span></div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:7}}>
          {demos.map(d=><button key={d.key} onClick={()=>runDemo(d.key)} style={{background:'white',border:`1px solid ${GR_BD}`,borderRadius:9,padding:'9px 10px',cursor:'pointer',textAlign:'left',display:'flex',flexDirection:'column',gap:3}}><span style={{fontSize:15}}>{d.icon}</span><span style={{fontWeight:700,color:TX,fontSize:12}}>{d.label}</span><span style={{fontSize:10,color:GR_TX,lineHeight:1.3}}>{d.sub}</span></button>)}
        </div>
      </div>
      <div style={{flex:1,overflowY:'auto',display:'flex',flexDirection:'column',gap:8,minHeight:0}}>
        {msgs.map((m,i)=>(
          <div key={i} style={{display:'flex',flexDirection:'column',alignItems:m.role==='user'?'flex-end':'flex-start',gap:3}}>
            {m.role==='ai'&&<div style={{fontSize:11,fontWeight:700,color:G_MID}}>✦ AI Assistant</div>}
            <div style={{background:m.role==='user'?G_MID:'white',color:m.role==='user'?'white':TX,border:m.role==='ai'?`1px solid ${GR_BD}`:'none',borderRadius:10,padding:'9px 13px',fontSize:12,lineHeight:1.6,maxWidth:'88%',whiteSpace:'pre-line'}}>{m.content}</div>
            <span style={{fontSize:10,color:GR_TX}}>{m.time}</span>
          </div>
        ))}
        {loading&&<div style={{fontSize:11,color:GR_TX}}>✦ Working...</div>}
        <div ref={msgsEndRef}/>
      </div>
      <div style={{display:'flex',gap:8,alignItems:'center',paddingTop:4}}>
        <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&send()} placeholder="Ask me to update formulas, create models, or explain calculations..." style={{flex:1,border:`1px solid ${GR_BD}`,borderRadius:10,padding:'9px 13px',fontSize:12,outline:'none',background:GR_BG}}/>
        <button onClick={send} disabled={loading} style={{background:loading?'#86efac':G_MID,border:'none',borderRadius:10,width:38,height:38,display:'flex',alignItems:'center',justifyContent:'center',cursor:loading?'default':'pointer',flexShrink:0}}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        </button>
      </div>
    </div>
  );
}

// ── Changes Tab ───────────────────────────────────────────────────────────────
function ChangesTab({changes,setChanges}:{changes:Change[];setChanges:React.Dispatch<React.SetStateAction<Change[]>>}){
  const act=async(id:number,status:'accepted'|'rejected'|'rethink')=>{if(status==='accepted'){const c=changes.find(x=>x.id===id);if(c)await applyChange(c);}setChanges(c=>c.map(x=>x.id===id?{...x,status}:x));};
  const acceptAll=async()=>{for(const c of changes.filter(x=>x.status==='pending'))await applyChange(c);setChanges(c=>c.map(x=>x.status==='pending'?{...x,status:'accepted'}:x));};
  const rejectAll=()=>setChanges(c=>c.map(x=>x.status==='pending'?{...x,status:'rejected'}:x));
  const pending=changes.filter(c=>c.status==='pending').length;
  const sBg:Record<string,string>={accepted:'#f0fdf4',rejected:'#fef2f2',pending:'white',rethink:'white'};
  const sBd:Record<string,string>={accepted:G_BORDER,rejected:'#fca5a5',pending:GR_BD,rethink:GR_BD};
  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div><div style={{fontSize:15,fontWeight:700,color:TX}}>Proposed Changes</div><div style={{fontSize:11,color:GR_TX}}>{pending} pending · {changes.length-pending} reviewed</div></div>
        <div style={{display:'flex',gap:6}}><button onClick={acceptAll} style={{background:G_MID,color:'white',border:'none',borderRadius:7,padding:'6px 12px',fontSize:12,fontWeight:600,cursor:'pointer'}}>Accept All</button><button onClick={rejectAll} style={{background:'white',color:'#dc2626',border:'1px solid #fca5a5',borderRadius:7,padding:'6px 10px',fontSize:12,fontWeight:600,cursor:'pointer'}}>Reject All</button></div>
      </div>
      {changes.length===0&&<div style={{textAlign:'center',color:GR_TX,fontSize:12,padding:'32px 0'}}>No changes yet — ask the AI something in the Chat tab.</div>}
      {changes.map(c=>(
        <div key={c.id} style={{border:`1px solid ${sBd[c.status]}`,borderRadius:10,padding:12,background:sBg[c.status]}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
            <div style={{display:'flex',alignItems:'center',gap:7}}><span style={{fontFamily:'monospace',fontWeight:700,fontSize:13}}>{c.cell}</span><span style={{fontSize:11,color:GR_TX}}>{c.sheet}</span><Badge label={c.type} bg={c.type==='formula'?'#eff6ff':'#f0fdf4'} col={c.type==='formula'?'#1d4ed8':'#15803d'}/></div>
            {c.status!=='pending'&&<Badge label={c.status==='accepted'?'✓ Accepted':c.status==='rejected'?'✗ Rejected':'↩ Rethink'} bg={sBg[c.status]} col={c.status==='accepted'?'#15803d':'#dc2626'}/>}
          </div>
          <div style={{display:'flex',gap:8,marginBottom:6}}><code style={{background:'#fef2f2',borderRadius:6,padding:'3px 8px',fontSize:11,color:'#991b1b'}}>— {c.old}</code><code style={{background:'#f0fdf4',borderRadius:6,padding:'3px 8px',fontSize:11,color:'#15803d'}}>+ {c.proposed}</code></div>
          <div style={{fontSize:11,color:GR_TX,marginBottom:c.status==='pending'?8:0}}>{c.reason}</div>
          {c.status==='pending'&&<div style={{display:'flex',gap:6}}><button onClick={()=>act(c.id,'accepted')} style={{flex:1,background:G_MID,color:'white',border:'none',borderRadius:7,padding:'6px 0',fontSize:12,fontWeight:600,cursor:'pointer'}}>✓ Accept</button><button onClick={()=>act(c.id,'rejected')} style={{flex:1,background:'white',color:'#dc2626',border:'1px solid #fca5a5',borderRadius:7,padding:'6px 0',fontSize:12,fontWeight:600,cursor:'pointer'}}>✗ Reject</button><button onClick={()=>act(c.id,'rethink')} style={{flex:1,background:'white',color:GR_TX,border:`1px solid ${GR_BD}`,borderRadius:7,padding:'6px 0',fontSize:12,fontWeight:600,cursor:'pointer'}}>↩ Rethink</button></div>}
        </div>
      ))}
    </div>
  );
}

// ── Agents Tab ────────────────────────────────────────────────────────────────
function AgentsTab({apiKey,onChangesProposed}:{apiKey:string;onChangesProposed:(c:Change[])=>void}){
  const [agents,setAgents]=useState<AgentDef[]>([
    {id:1,name:'IB Formatting Agent',icon:'📐',desc:'Applies Wall Street IB formatting standards to the active sheet instantly.',scope:['Formatting','Colors','Fonts','Borders'],rules:['Blue font for hardcoded numbers','Black font for formulas & text','Dark blue (#244062) section headers','Gridlines OFF','Calibri 12pt throughout'],on:true},
    {id:2,name:'VC Model Dissector',icon:'🔍',desc:'Reverse-engineers a financial model — identifies key drivers, flags risks, and recommends sensitivities.',scope:['Model Analysis','Assumptions','Risks'],rules:['Identify top 3-5 hardcoded drivers with actual values','Flag structural issues & aggressive assumptions','Recommend 2-variable sensitivity with specific ranges'],on:true},
    {id:3,name:'Cap Table Agent',icon:'📋',desc:'Detects an existing cap table and models the next round, or builds one from scratch.',scope:['Cap Table','Dilution','Rounds'],rules:['Handles common, options, seed, series A/B/C','Models convertible notes & SAFEs','Calculates dilution and option pool refresh'],on:true},
    {id:4,name:'Exit Analysis Agent',icon:'🚪',desc:'Runs a multi-scenario exit waterfall — pays liquidation preferences in seniority order and distributes remainder to common.',scope:['Exit','Waterfall','MOIC'],rules:['Seniority-ordered liquidation pref payouts','Conversion threshold check per investor','Weighted average MOIC across scenarios'],on:true},
    {id:5,name:'DCF Valuation Agent',icon:'💹',desc:'Reads FCF projections from your model and computes enterprise value with terminal value.',scope:['DCF','Valuation','WACC'],rules:['Reads FCF from active sheet','WACC and terminal growth user-inputted','Outputs EV range ±15%'],on:true},
    {id:6,name:'VC Valuation Agent',icon:'🎯',desc:'Forward revenue multiple back-solve: given your target MOIC, computes required exit EV and revenue, then checks against model projections.',scope:['VC Valuation','Revenue Multiple','Back-Solve'],rules:['User sets target MOIC and EV/Rev multiple','Back-solves required exit EV and revenue','Auto-detect or manual exit year'],on:true},
    {id:7,name:'Sensitivity Agent',icon:'📊',desc:'Builds a two-variable sensitivity table on the most impactful assumptions.',scope:['Sensitivity','Scenarios'],rules:['AI identifies key inputs','Writes table below existing data','Highlights base case'],on:true},
  ]);
  const [activeAgentId,setActiveAgentId]=useState<number|null>(null);
  const [running,setRunning]=useState<number|null>(null);
  const [runStatus,setRunStatus]=useState('');
  const [dissectResult,setDissectResult]=useState<DissectResult|null>(null);
  const toggle=(id:number)=>{setAgents(a=>a.map(x=>x.id===id?{...x,on:!x.on}:x));if(activeAgentId===id){setActiveAgentId(null);setDissectResult(null);setRunStatus('');}};
  const active=agents.filter(a=>a.on).length;

  const runSimpleAgent=async(id:number)=>{
    setRunning(id);setRunStatus('Starting...');setDissectResult(null);
    try{
      if(id===1){setRunStatus(await applyIBFormatting(msg=>setRunStatus(msg)));}
      else if(id===2){
        if(!apiKey){setRunStatus('Please enter your API key using the ⚙ button at the top.');setRunning(null);return;}
        setRunStatus('Reading spreadsheet...');
        setDissectResult(await dissectModel(apiKey,msg=>setRunStatus(msg)));setRunStatus('');
      }else if(id===7){
        if(!apiKey){setRunStatus('Please enter your API key using the ⚙ button at the top.');setRunning(null);return;}
        setRunStatus(await buildSensitivity(apiKey,msg=>setRunStatus(msg)));
      }
    }catch(err){setRunStatus(`Error: ${err instanceof Error?err.message:'Unknown'}`);}
    setRunning(null);
  };

  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div><div style={{fontSize:15,fontWeight:700,color:TX}}>Sub-Agents</div><div style={{fontSize:11,color:GR_TX}}>Control specialized agents</div><div style={{fontSize:11,color:G_MID,marginTop:2}}>{active} of {agents.length} active</div></div>
        <button style={{background:G_MID,color:'white',border:'none',borderRadius:8,padding:'7px 13px',fontSize:12,fontWeight:700,cursor:'pointer'}}>+ New Agent</button>
      </div>
      {runStatus&&<div style={{background:G_LIGHT,border:`1px solid ${G_BORDER}`,borderRadius:8,padding:'10px 12px',fontSize:11,color:G_DARK,whiteSpace:'pre-line',lineHeight:1.6}}>{runStatus}</div>}
      {dissectResult&&<DissectOutput result={dissectResult}/>}
      {agents.map(ag=>(
        <div key={ag.id} style={{border:`1px solid ${ag.on?G_BORDER:GR_BD}`,borderRadius:10,padding:12,background:ag.on?G_LIGHT:'white'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8}}>
            <div style={{display:'flex',alignItems:'flex-start',gap:10}}>
              <div style={{width:34,height:34,borderRadius:8,background:ag.on?'#dcfce7':'#f3f4f6',display:'flex',alignItems:'center',justifyContent:'center',fontSize:15,flexShrink:0}}>{ag.icon}</div>
              <div><div style={{fontSize:13,fontWeight:700,color:TX}}>{ag.name}</div><div style={{fontSize:11,color:GR_TX,maxWidth:220,lineHeight:1.4}}>{ag.desc}</div></div>
            </div>
            <Toggle on={ag.on} onToggle={()=>toggle(ag.id)}/>
          </div>
          <div style={{fontSize:10,fontWeight:700,color:GR_TX,letterSpacing:'0.06em',marginBottom:4}}>SCOPE</div>
          <div style={{display:'flex',gap:5,flexWrap:'wrap',marginBottom:6}}>{ag.scope.map(s=><Badge key={s} label={s}/>)}</div>
          <div style={{fontSize:10,fontWeight:700,color:GR_TX,letterSpacing:'0.06em',marginBottom:4}}>RULES</div>
          {ag.rules.map(r=><div key={r} style={{fontSize:11,color:GR_TX,display:'flex',alignItems:'center',gap:5,marginBottom:2}}><span style={{width:3,height:3,borderRadius:'50%',background:GR_TX,display:'inline-block',flexShrink:0}}/>{r}</div>)}
          {ag.on&&(
            <div style={{marginTop:10}}>
              {[1,2,7].includes(ag.id)&&(
                <button onClick={()=>runSimpleAgent(ag.id)} disabled={running===ag.id} style={{background:running===ag.id?'#86efac':G_MID,color:'white',border:'none',borderRadius:7,padding:'6px 14px',fontSize:12,fontWeight:600,cursor:running===ag.id?'default':'pointer'}}>
                  {running===ag.id?'⏳ Running...':'▶ Run Agent'}
                </button>
              )}
              {[3,4,5,6].includes(ag.id)&&(
                <button onClick={()=>setActiveAgentId(activeAgentId===ag.id?null:ag.id)} style={{background:activeAgentId===ag.id?'#6b7280':G_MID,color:'white',border:'none',borderRadius:7,padding:'6px 14px',fontSize:12,fontWeight:600,cursor:'pointer'}}>
                  {activeAgentId===ag.id?'▼ Close':'▶ Open Agent'}
                </button>
              )}
            </div>
          )}
          {activeAgentId===ag.id&&ag.on&&(
            <div style={{marginTop:12,borderTop:`1px solid ${GR_BD}`,paddingTop:12}}>
              {ag.id===3&&<CapTableAgent apiKey={apiKey} onChangesProposed={onChangesProposed}/>}
              {ag.id===4&&<ExitAnalysisAgent apiKey={apiKey} onChangesProposed={onChangesProposed}/>}
              {ag.id===5&&<DCFAgent apiKey={apiKey}/>}
              {ag.id===6&&<VCValuationAgent apiKey={apiKey}/>}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ── Rules Tab ─────────────────────────────────────────────────────────────────
function RulesTab(){
  const [rules,setRules]=useState<Rule[]>([
    {id:1,name:'No Circular References',type:'validation',typeBg:'#eff6ff',typeCol:'#1d4ed8',desc:'Validates that proposed changes do not create circular reference errors',trigger:'before-change',code:'checkCircularReferences(proposedChanges)',on:true},
    {id:2,name:'Preserve Audit Trail',type:'audit',typeBg:'#faf5ff',typeCol:'#7c3aed',desc:'Ensures all changes to assumption cells are logged with timestamp',trigger:'after-change',code:'logAssumptionChanges(cellRange)',on:true},
    {id:3,name:'IB Formatting Standards',type:'formatting',typeBg:'#f0fdf4',typeCol:'#15803d',desc:'Enforces Calibri 12pt, blue inputs, black formulas, dark blue section headers',trigger:'after-change',code:'applyIBFormatting(worksheet)',on:true},
    {id:4,name:'Protect Input Cells',type:'constraint',typeBg:'#fffbeb',typeCol:'#d97706',desc:'Prevents AI from overwriting cells tagged as hardcoded assumptions',trigger:'before-change',code:'validateInputProtection(cellRef)',on:true},
  ]);
  const toggle=(id:number)=>setRules(r=>r.map(x=>x.id===id?{...x,on:!x.on}:x));
  const active=rules.filter(r=>r.on).length;
  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div><div style={{fontSize:15,fontWeight:700,color:TX}}>Custom Rules</div><div style={{fontSize:11,color:GR_TX}}>Define validation and constraint rules</div><div style={{fontSize:11,color:G_MID,marginTop:2}}>{active} of {rules.length} active</div></div>
        <button style={{background:G_MID,color:'white',border:'none',borderRadius:8,padding:'7px 13px',fontSize:12,fontWeight:700,cursor:'pointer'}}>+ New Rule</button>
      </div>
      {rules.map(r=>(
        <div key={r.id} style={{border:`1px solid ${GR_BD}`,borderRadius:10,padding:12,background:'white'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:6}}>
            <div style={{display:'flex',alignItems:'center',gap:7}}><div style={{width:28,height:28,borderRadius:'50%',background:r.typeBg,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,color:r.typeCol,flexShrink:0}}>✓</div><span style={{fontSize:13,fontWeight:700,color:TX}}>{r.name}</span><Badge label={r.type} bg={r.typeBg} col={r.typeCol}/></div>
            <Toggle on={r.on} onToggle={()=>toggle(r.id)}/>
          </div>
          <div style={{fontSize:11,color:GR_TX,marginBottom:6}}>{r.desc}</div>
          <Badge label={r.trigger}/>
          <div style={{marginTop:8,background:'#f9fafb',borderRadius:7,padding:'7px 10px'}}><div style={{fontSize:9,fontWeight:700,color:GR_TX,letterSpacing:'0.08em',marginBottom:3}}>CODE</div><code style={{fontSize:11,fontFamily:'monospace',color:TX}}>{r.code}</code></div>
        </div>
      ))}
    </div>
  );
}

// ── Templates Tab ─────────────────────────────────────────────────────────────
function TemplatesTab(){
  const [search,setSearch]=useState('');
  const [filter,setFilter]=useState('All');
  const filters=['All','Valuation','Industry','Private Equity','Startups','M&A','VC'];
  const templates:Template[]=[
    {name:'DCF Valuation Model',icon:'💵',desc:'Three-statement DCF with WACC calculation',sheets:5,cells:420,cat:'Valuation',catBg:'#eff6ff',catCol:'#1d4ed8',star:true},
    {name:'Three Statement Model',icon:'📊',desc:'Fully linked IS, Balance Sheet, and Cash Flow',sheets:3,cells:480,cat:'Valuation',catBg:'#eff6ff',catCol:'#1d4ed8',star:true,isnew:true},
    {name:'SaaS Financial Model',icon:'📈',desc:'Complete SaaS metrics with cohort analysis',sheets:7,cells:650,cat:'Industry',catBg:'#f0fdf4',catCol:'#15803d',star:true},
    {name:'LBO Analysis',icon:'🧮',desc:'Leveraged buyout with returns waterfall',sheets:6,cells:590,cat:'Private Equity',catBg:'#faf5ff',catCol:'#7c3aed',star:false},
    {name:'M&A Accretion/Dilution',icon:'🔄',desc:'Merger model with synergies and EPS bridge',sheets:8,cells:740,cat:'M&A',catBg:'#fff7ed',catCol:'#c2410c',star:false},
    {name:'Startup Financial Model',icon:'🚀',desc:'Full 3-statement for early-stage companies',sheets:5,cells:380,cat:'Startups',catBg:'#fffbeb',catCol:'#b45309',star:true},
    {name:'VC Comparables Valuation',icon:'🔭',desc:'Revenue and EBITDA multiples comps for VC',sheets:3,cells:290,cat:'VC',catBg:'#f0f9ff',catCol:'#0369a1',star:true,isnew:true},
    {name:'Capitalization Table',icon:'🥧',desc:'Cap table from seed through Series C',sheets:4,cells:340,cat:'VC',catBg:'#f0f9ff',catCol:'#0369a1',star:true,isnew:true},
  ];
  const shown=templates.filter(t=>(filter==='All'||t.cat===filter)&&(!search||t.name.toLowerCase().includes(search.toLowerCase())));
  return(
    <div style={{display:'flex',flexDirection:'column',gap:10}}>
      <div><div style={{fontSize:15,fontWeight:700,color:TX}}>Template Library</div><div style={{fontSize:11,color:GR_TX}}>Start with pre-built financial models or save your own</div></div>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search templates..." style={{width:'100%',padding:'8px 12px',border:`1px solid ${GR_BD}`,borderRadius:8,fontSize:12,outline:'none',background:GR_BG,boxSizing:'border-box' as any}}/>
      <div style={{display:'flex',gap:5,flexWrap:'wrap'}}>{filters.map(f=><button key={f} onClick={()=>setFilter(f)} style={{padding:'4px 10px',borderRadius:16,fontSize:11,fontWeight:600,cursor:'pointer',background:filter===f?TX:'white',color:filter===f?'white':TX,border:`1px solid ${filter===f?TX:GR_BD}`}}>{f}</button>)}</div>
      <div style={{display:'flex',flexDirection:'column',gap:8}}>
        {shown.map(t=>(
          <div key={t.name} style={{border:`1px solid ${t.isnew?G_BORDER:GR_BD}`,borderRadius:10,padding:12,background:t.isnew?G_LIGHT:'white'}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:5}}>
              <div style={{display:'flex',alignItems:'center',gap:9}}>
                <div style={{width:34,height:34,borderRadius:8,background:t.isnew?'#dcfce7':G_LIGHT,display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,flexShrink:0}}>{t.icon}</div>
                <div><div style={{fontSize:13,fontWeight:700,color:TX}}>{t.name} {t.star&&<span style={{color:'#f59e0b'}}>★</span>}{t.isnew&&<span style={{background:G_MID,color:'white',fontSize:9,fontWeight:700,padding:'1px 6px',borderRadius:8,marginLeft:4}}>NEW</span>}</div><div style={{fontSize:11,color:GR_TX,lineHeight:1.4,maxWidth:200}}>{t.desc}</div></div>
              </div>
              <Badge label={t.cat} bg={t.catBg} col={t.catCol}/>
            </div>
            <div style={{fontSize:10,color:GR_TX,marginBottom:8}}>{t.sheets} sheets · {t.cells} cells</div>
            <div style={{display:'flex',gap:7}}><button style={{background:G_MID,color:'white',border:'none',borderRadius:7,padding:'6px 14px',fontSize:12,fontWeight:600,cursor:'pointer'}}>↓ Use Template</button><button style={{background:'white',color:TX,border:`1px solid ${GR_BD}`,borderRadius:7,padding:'6px 14px',fontSize:12,fontWeight:600,cursor:'pointer'}}>Preview</button></div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Root App ──────────────────────────────────────────────────────────────────
export default function App(){
  const [tab,setTab]=useState<'Chat'|'Changes'|'Agents'|'Rules'|'Templates'>('Chat');
  const [changes,setChanges]=useState<Change[]>([]);
  const [apiKey]=useState('proxy');
  const tabs=['Chat','Changes','Agents','Rules','Templates'] as const;
  const pendingCount=changes.filter(c=>c.status==='pending').length;
  return(
    <div style={{display:'flex',flexDirection:'column',height:'100vh',fontFamily:"'Segoe UI', system-ui, sans-serif",background:'white'}}>
      <div style={{background:G_DARK,padding:'10px 16px',display:'flex',justifyContent:'space-between',alignItems:'center',flexShrink:0}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}><div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:2.5}}>{[...Array(9)].map((_,i)=><div key={i} style={{width:3,height:3,background:'rgba(255,255,255,0.65)',borderRadius:1}}/>)}</div><span style={{color:'white',fontWeight:800,fontSize:14,letterSpacing:'0.1em'}}>4SIGHT</span></div>
      
      </div>
     
      <div style={{display:'flex',borderBottom:`1px solid ${GR_BD}`,background:'white',flexShrink:0}}>
        {tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:'9px 4px',border:'none',background:'none',fontSize:12,fontWeight:t===tab?600:400,cursor:'pointer',color:t===tab?TX:GR_TX,borderBottom:t===tab?`2px solid ${G_MID}`:'2px solid transparent',position:'relative'}}>{t}{t==='Changes'&&pendingCount>0&&<span style={{position:'absolute',top:4,right:4,background:'#dc2626',color:'white',borderRadius:'50%',width:16,height:16,fontSize:9,fontWeight:700,display:'flex',alignItems:'center',justifyContent:'center'}}>{pendingCount}</span>}</button>)}
      </div>
      <div style={{flex:1,padding:14,overflowY:'auto',background:GR_BG,display:'flex',flexDirection:'column'}}>
        {tab==='Chat'&&<ChatTab onChangesProposed={c=>setChanges(prev=>[...prev,...c])} apiKey={apiKey}/>}
        {tab==='Changes'&&<ChangesTab changes={changes} setChanges={setChanges}/>}
        {tab==='Agents'&&<AgentsTab apiKey={apiKey} onChangesProposed={c=>setChanges(prev=>[...prev,...c])}/>}
        {tab==='Rules'&&<RulesTab/>}
        {tab==='Templates'&&<TemplatesTab/>}
      </div>
    </div>
  );
}
