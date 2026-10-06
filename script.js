const KEY='essayWorkshop 2_v1',$=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];let sourceCount=2;
const esc=s=>String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function sourceHTML(i){return `<div class="source"><h3>Source ${i}</h3><label>Author / Organization<input data-save id="src${i}author"></label><label>Article / Page Title<input data-save id="src${i}title"></label><label>Website / Container<input data-save id="src${i}site"></label><label>Publication Date<input data-save id="src${i}date"></label><label>URL<input data-save id="src${i}url" type="url"></label><label>MLA Works Cited Entry<textarea data-save id="src${i}mla" placeholder="Paste or type the complete MLA citation here."></textarea></label></div>`}
function data(){let d={sourceCount};$$('[data-save]').forEach(e=>d[e.id]=e.type==='checkbox'?e.checked:e.value);return d}function save(){localStorage.setItem(KEY,JSON.stringify(data()));progress()}function load(){try{return JSON.parse(localStorage.getItem(KEY)||'{}')}catch{return {}}}function restore(d){$$('[data-save]').forEach(e=>{if(d[e.id]!==undefined)e.type==='checkbox'?e.checked=d[e.id]:e.value=d[e.id]})}function bind(){$$('[data-save]').forEach(e=>{e.oninput=save;e.onchange=save})}function renderSources(n){$('#sources').innerHTML='';for(let i=1;i<=n;i++)$('#sources').insertAdjacentHTML('beforeend',sourceHTML(i));sourceCount=n;bind()}function progress(){let a=$$('[data-save]'),done=a.filter(e=>e.type==='checkbox'?e.checked:String(e.value).trim()).length,p=Math.round(done/a.length*100);$('#bar').style.width=p+'%';$('#pct').textContent=p+'% complete'}function v(id){return $('#'+id)?.value.trim()||''}
function join(ids){return ids.map(v).filter(Boolean).join(' ')}
function build(which){if(which==='intro')$('#introDraft').value=join(['hook','context','importance','introThesis']);else if(which==='con')$('#conDraft').value=join(['conClaim','conSynth','conMatter','conFinal']);else {let n=which.slice(-1);$('#'+which+'draft').value=join([`bp${n}topic`,`bp${n}evidence`,`bp${n}explain`,`bp${n}compare`,`bp${n}transition`])}save()}
function assemble(){let paras=['introDraft','bp1draft','bp2draft','bp3draft','conDraft'].map(v).filter(Boolean);$('#fullEssay').value=paras.join('\n\n');save()}
function xmlText(s){return esc(s).replace(/\n/g,'</w:t><w:br/><w:t xml:space="preserve">')}
function para(text,opts={}){let pPr='';if(opts.center)pPr+='<w:jc w:val="center"/>';if(opts.indent)pPr+='<w:ind w:firstLine="720"/>';if(opts.hanging)pPr+='<w:ind w:left="720" w:hanging="720"/>';if(opts.pageBreak)pPr+='<w:pageBreakBefore/>';return `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto" w:before="0" w:after="0"/>${pPr}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${xmlText(text)}</w:t></w:r></w:p>`}
async function docx(){if(!v('fullEssay')) assemble();let d=data(),zip=new JSZip(),last=(v('student').trim().split(/\s+/).pop()||'Student'),essay=v('fullEssay');let paras=essay.split(/\n\s*\n/).map(x=>x.trim()).filter(Boolean);let body=para(v('student'))+para(v('teacher'))+para(v('course'))+para(v('date'))+para(v('title')||'Compare and Contrast Essay',{center:true})+paras.map(x=>para(x,{indent:true})).join('');let cites=[];for(let i=1;i<=sourceCount;i++){let c=v(`src${i}mla`);if(c)cites.push(c)}if(cites.length){body+=para('Works Cited',{center:true,pageBreak:true})+cites.map(c=>para(c,{hanging:true})).join('')}
let document=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`;
let header=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>${esc(last)} </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:hdr>`;
zip.file('[Content_Types].xml',`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`);zip.folder('_rels').file('.rels',`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);let word=zip.folder('word');word.file('document.xml',document);word.file('header1.xml',header);word.folder('_rels').file('document.xml.rels',`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`);let blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=(v('title')||'Compare_Contrast_Essay').replace(/[^a-z0-9]+/gi,'_')+'.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
let d=load();renderSources(d.sourceCount||2);restore(d);bind();progress();$$('[data-put]').forEach(b=>b.onclick=()=>{let e=$('#'+b.dataset.put);e.value=b.dataset.text;e.focus();save()});$$('[data-copy]').forEach(b=>b.onclick=()=>{$('#'+b.dataset.to).value=$('#'+b.dataset.copy).value;save()});$$('[data-build]').forEach(b=>b.onclick=()=>build(b.dataset.build));$('#addSource').onclick=()=>{let old=data();renderSources(++sourceCount);restore(old);save()};$('#assemble').onclick=assemble;$('#docx').onclick=docx;$('#save').onclick=()=>{save();$('#save').textContent='✓ Saved';setTimeout(()=>$('#save').textContent='Save Progress',1000)};$('#reset').onclick=()=>{if(confirm('Reset all Workshop 2 responses?')){localStorage.removeItem(KEY);location.reload()}};

// Portable progress-file workflow: students can save and resume on any device.
function downloadProgress(){
  save();
  const payload={app:'Socratic Seminar Workshop 2 - Compare & Contrast Essay',version:2,savedAt:new Date().toISOString(),responses:data()};
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  const student=(v('student')||'Student').replace(/[^a-z0-9]+/gi,'_');
  a.download=student+'_Essay_Workshop_Progress.json';
  document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  $('#status').textContent='✓ Progress file downloaded. Save it somewhere you can access from ANY lab computer next class.';
}
function importProgress(file){
  const reader=new FileReader();
  reader.onload=()=>{try{
    const parsed=JSON.parse(reader.result),incoming=parsed.responses||parsed;
    if(!incoming || typeof incoming!=='object') throw new Error('Invalid progress file');
    renderSources(Number(incoming.sourceCount)||2);restore(incoming);bind();save();progress();
    $('#status').textContent='✓ Previous progress loaded. Everything saved in that file is ready to continue.';
    window.scrollTo({top:0,behavior:'smooth'});
  }catch(err){$('#status').textContent='That file could not be loaded. Please choose a Workshop 2 progress file.';}};
  reader.readAsText(file);
}
$('#downloadProgress').onclick=downloadProgress;
$('#uploadProgress').onclick=()=>$('#progressFile').click();
$('#progressFile').onchange=e=>{const f=e.target.files&&e.target.files[0];if(f)importProgress(f);e.target.value='';};

// Computer-lab duplicate controls at the top of the activity.
const topUpload=document.querySelector('#uploadProgressTop');
const topDownload=document.querySelector('#downloadProgressTop');
if(topUpload) topUpload.onclick=()=>document.querySelector('#progressFile').click();
if(topDownload) topDownload.onclick=downloadProgress;
const labObserver=new MutationObserver(()=>{const s=document.querySelector('#status'),t=document.querySelector('#statusTop');if(s&&t)t.textContent=s.textContent;});
if(document.querySelector('#status')) labObserver.observe(document.querySelector('#status'),{childList:true,subtree:true,characterData:true});


// --- Socratic Seminar -> Workshop 2 import ---
function firstBlank(ids, value){
  if(!value) return;
  for(const id of ids){
    const el=document.getElementById(id);
    if(el && !String(el.value||"").trim()){ el.value=value; el.dispatchEvent(new Event("input",{bubbles:true})); return; }
  }
}
function joinParts(parts){return parts.filter(x=>x&&String(x).trim()).join("\n\n")}
function importSeminarData(d){
  if(!d || typeof d!=="object") throw new Error("This file does not contain readable seminar notes.");
  const sources=[];
  for(let i=1;i<=8;i++){
    const title=d[`s${i}title`], author=d[`s${i}author`], link=d[`s${i}link`],
          summary=d[`s${i}summary`], evidence=d[`s${i}evidence`], connect=d[`s${i}connect`];
    if(title||summary||evidence||connect) sources.push({title,author,link,summary,evidence,connect});
  }
  const online=d.online||"", person=d.person||"", both=d.both||"", challenge=d.challenge||"";
  const position=d.position||d.meaning||"", strongest=d.strongEvidence||d.suggests||"",
        comparison=d.strongComparison||joinParts([online,person,both]);
  // Populate only blank essay fields so existing essay work is protected.
  firstBlank(["context"], joinParts([d.meaning, online, person]));
  firstBlank(["importance"], d.matters||d.useWhen);
  firstBlank(["thesis"], position);
  firstBlank(["bp1point"], both ? "Similarity between online and in-person communication" : "");
  firstBlank(["bp1evidence"], sources[0]?.evidence||strongest);
  firstBlank(["bp1explain"], sources[0]?.connect||d.matters);
  firstBlank(["bp2point"], online ? "Online/technology-enhanced communication" : "");
  firstBlank(["bp2evidence"], sources[1]?.evidence||online);
  firstBlank(["bp2explain"], sources[1]?.connect||d.useWhen);
  firstBlank(["bp3point"], person ? "In-person communication" : "");
  firstBlank(["bp3evidence"], sources[2]?.evidence||person);
  firstBlank(["bp3explain"], sources[2]?.connect||challenge);
  // Try common source fields in essay builder.
  sources.forEach((s,idx)=>{
    const i=idx+1;
    firstBlank([`source${i}title`,`src${i}title`,`wc${i}title`],s.title);
    firstBlank([`source${i}author`,`src${i}author`,`wc${i}author`],s.author);
    firstBlank([`source${i}url`,`source${i}link`,`src${i}url`],s.link);
  });
  const q=[d.q1,d.q2,d.q3].filter(Boolean);
  const html=[
    position && `<p><b>Current thinking:</b> ${escapeHtml(position)}</p>`,
    online && `<p><b>Online/technology-enhanced:</b> ${escapeHtml(online)}</p>`,
    person && `<p><b>In-person:</b> ${escapeHtml(person)}</p>`,
    both && `<p><b>Similarities:</b> ${escapeHtml(both)}</p>`,
    challenge && `<p><b>Perspective to consider:</b> ${escapeHtml(challenge)}</p>`,
    strongest && `<p><b>Strong evidence:</b> ${escapeHtml(strongest)}</p>`,
    comparison && `<p><b>Strong comparison:</b> ${escapeHtml(comparison)}</p>`,
    q.length && `<p><b>Seminar questions:</b><br>${q.map(escapeHtml).join("<br>")}</p>`,
    sources.length && `<p><b>Imported sources:</b><br>${sources.map(s=>escapeHtml([s.title,s.author,s.link].filter(Boolean).join(" — "))).join("<br>")}</p>`
  ].filter(Boolean).join("");
  const card=document.getElementById("importedSeminarCard");
  const notes=document.getElementById("importedSeminarNotes");
  if(card&&notes){notes.innerHTML=html||"Seminar file imported.";card.style.display="block";}
  if(typeof save==="function") save();
  const st=document.getElementById("seminarStatus");
  if(st) st.textContent="✓ Socratic Seminar notes imported. Blank essay-planning fields have been prefilled; your existing essay work was not overwritten.";
}
function escapeHtml(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]));}
const semBtn=document.getElementById("uploadSeminar"), semFile=document.getElementById("seminarFile");
if(semBtn&&semFile){
  semBtn.addEventListener("click",()=>semFile.click());
  semFile.addEventListener("change",async()=>{
    const f=semFile.files&&semFile.files[0]; if(!f)return;
    try{importSeminarData(JSON.parse(await f.text()));}
    catch(err){document.getElementById("seminarStatus").textContent="Could not import this file. Please choose the JSON progress/notes file downloaded from the Socratic Seminar activity.";}
    semFile.value="";
  });
}
const essayResume=document.getElementById("uploadEssayProgress");
if(essayResume){
  essayResume.addEventListener("click",()=>{
    const old=document.getElementById("uploadProgressTop")||document.getElementById("uploadProgress");
    if(old) old.click();
  });
}
