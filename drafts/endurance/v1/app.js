'use strict';
const $ = id => document.getElementById(id);
const pad = n => String(n).padStart(2,'0');
const escapeHtml = s => String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const paragraphs = items => items.map(p=>`<p>${escapeHtml(p)}</p>`).join('');
const arrow = '<svg class="icon chapter-arrow" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 12h16m-6-6 6 6-6 6"/></svg>';
const iconPaths = {
 fuel:'<path d="M12 3c6 5 8 8 8 12a8 8 0 0 1-16 0c0-4 2-7 8-12Z"/><path d="M12 11v8m0-4 4-3"/>',
 atp:'<circle cx="12" cy="12" r="3"/><circle cx="12" cy="12" r="9"/><path d="M12 1v4m0 14v4M1 12h4m14 0h4"/>',
 work:'<path d="M4 17 9 12l4 3 7-9M14 6h6v6"/><path d="M4 22h16"/>',
 signal:'<path d="M2 12h4l3-8 6 16 3-8h4"/>',
 link:'<path d="m10 14 4-4m-6 6-1 1a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 10a4 4 0 0 0 6 0l4-4a4 4 0 0 0-6-6l-1 1" transform="translate(0 -1) scale(.95)"/>',
 cycle:'<path d="M20 9a8 8 0 0 0-14-3L3 9m0-6v6h6m-5 6a8 8 0 0 0 14 3l3-3m0 6v-6h-6"/>'
};
let currentChapter=0,currentTopic=0,currentStep=1,focusMode=false;
let lastSearchTrigger=null;
const mq=window.matchMedia('(max-width: 760px)');
const reduced=window.matchMedia('(prefers-reduced-motion: reduce)');
// Store only a local reading preference. Failure (e.g. browser privacy mode) is harmless.
try{focusMode=localStorage.getItem('ronu-drafts-endurance-v1-focus')==='true';}catch(_){}
function route(){
 const bits=location.hash.replace(/^#/,'').split('/');
 let ci=CHAPTERS.findIndex(c=>c.id===bits[0]);if(ci<0)ci=0;
 let ti=CHAPTERS[ci].topics.findIndex(t=>t.id===bits[1]);if(ti<0)ti=0;
 return [ci,ti];
}
function routeTo(ci,ti=0,options={}){
 const hash=`#${CHAPTERS[ci].id}/${CHAPTERS[ci].topics[ti].id}`;
 if(location.hash!==hash){history.pushState(null,'',hash);}
 render(ci,ti,options);
}
function renderNav(){
 $('chapter-nav').innerHTML=CHAPTERS.map((c,i)=>`<a class="chapter-link" href="#${c.id}/overview" data-chapter="${i}"${i===currentChapter?' aria-current="page"':''}><span class="num">${pad(i+1)}</span><span class="chapter-name">${escapeHtml(c.label)}</span>${arrow}</a>`).join('');
 $('chapter-select').innerHTML=CHAPTERS.map((c,i)=>`<option value="${i}"${i===currentChapter?' selected':''}>${pad(i+1)} / ${escapeHtml(c.label)}</option>`).join('');
}
function render(ci,ti,options={}){
 const chapterChanged=ci!==currentChapter;
 currentChapter=ci;currentTopic=ti;
 const c=CHAPTERS[ci],t=c.topics[ti];
 renderNav();
 $('section-number').textContent=pad(ci+1);$('section-label').textContent=c.short;
 $('section-count').textContent=`${pad(ci+1)} / ${pad(CHAPTERS.length)}`;
 $('section-title').textContent=c.title;$('section-description').textContent=c.description;
 $('progress-label').textContent=`${pad(ci+1)} / ${pad(CHAPTERS.length)}`;
 $('progress-fill').style.width=`${100*(ci+1)/CHAPTERS.length}%`;
 $('chapter-progress').setAttribute('aria-valuenow',String(ci+1));
 $('topic-tabs').innerHTML=c.topics.map((topic,i)=>`<button class="topic-tab" id="tab-${c.id}-${topic.id}" role="tab" aria-controls="topic-panel" aria-selected="${i===ti}" tabindex="${i===ti?0:-1}" data-topic="${i}">${escapeHtml(topic.label)}</button>`).join('');
 $('topic-panel').setAttribute('aria-labelledby',`tab-${c.id}-${t.id}`);
 $('topic-eyebrow').textContent=`${pad(ci+1)}.${ti+1} / ${t.label}`;
 $('topic-heading').textContent=t.heading;$('topic-copy').innerHTML=paragraphs(t.paragraphs);
 $('mechanism-copy').innerHTML=paragraphs(t.mechanism);$('connection-copy').innerHTML=paragraphs(t.connection);
 $('mechanism-detail').open=false;$('connection-detail').open=false;
 currentStep=t.step;renderConcept();
 $('previous-chapter').disabled=ci===0;
 const next=(ci+1)%CHAPTERS.length;
 $('next-chapter').href=`#${CHAPTERS[next].id}/overview`;
 $('next-label').textContent=ci===CHAPTERS.length-1?'Back to the beginning':`Up next / ${pad(next+1)}`;
 $('next-title').textContent=CHAPTERS[next].label;
 document.title=`${c.short} — Endurance / Ronu`;
 $('route-announcement').textContent=`Chapter ${ci+1}: ${c.label}. Topic: ${t.label}.`;
 setRail(false);
 // Keep the topic panel steady. Chapter changes can return to the section heading.
 if(options.focusTab){document.querySelector('.topic-tab[aria-selected="true"]').focus({preventScroll:true});}
 if(options.focusHeading){$('section-title').focus({preventScroll:true});}
 if(options.scroll && chapterChanged){requestAnimationFrame(()=>{
   const y=$('section-title').getBoundingClientRect().top+window.scrollY-125;
   if(window.scrollY>80||mq.matches)window.scrollTo({top:Math.max(0,y),behavior:reduced.matches?'instant':'smooth'});
 });}
}
function renderConcept(){
 const c=CHAPTERS[currentChapter];
 $('concept-steps').innerHTML=c.steps.map((s,i)=>(i?'<span class="connector" aria-hidden="true"><svg viewBox="0 0 24 12" fill="none" stroke="currentColor" stroke-width="1"><path d="M0 6h21m-5-4 5 4-5 4"/></svg></span>':'')+`<button class="concept-node" data-step="${i}" aria-pressed="${i===currentStep}" aria-label="Explore ${escapeHtml(s.label.toLowerCase())}"><svg class="node-icon" viewBox="0 0 24 24" aria-hidden="true">${iconPaths[s.icon]}</svg><span class="node-title">${escapeHtml(s.label)}</span></button>`).join('');
 // Caption HTML is authored in this file; user input never reaches innerHTML.
 $('concept-caption').innerHTML=c.steps[currentStep].caption;
}
function setRail(open){
 const isOpen=Boolean(open&&mq.matches&&!focusMode);
 document.body.classList.toggle('rail-open',isOpen);
 $('menu-toggle').setAttribute('aria-expanded',String(isOpen));
 $('sidebar').inert=focusMode||(mq.matches&&!isOpen);
 if(isOpen){requestAnimationFrame(()=>$('sidebar').querySelector('[aria-current="page"]').focus());}
}
function applyFocus(){
 document.body.classList.toggle('focus-mode',focusMode);
 $('focus-toggle').setAttribute('aria-pressed',String(focusMode));
 $('focus-toggle').querySelector('span').textContent=focusMode?'Exit focus':'Focus view';
 setRail(false);
}
function showSearch(){
 lastSearchTrigger=document.activeElement;
 setRail(false);
 $('topic-search').value='';renderSearch('');
 $('search-dialog').showModal();$('topic-search').focus();
}
function closeDialog(id){$(id).close();}
function renderSearch(q){
 const term=q.trim().toLowerCase();let count=0;let html='';
 CHAPTERS.forEach((c,ci)=>{
  const hits=c.topics.map((t,ti)=>({t,ti})).filter(({t})=>!term||`${c.label} ${c.title} ${t.label} ${t.heading} ${t.paragraphs.join(' ')}`.toLowerCase().includes(term));
  if(hits.length){html+=`<div class="result-group">${pad(ci+1)} / ${escapeHtml(c.label)}</div>`;hits.forEach(({t,ti})=>{count++;html+=`<button class="result" data-search-chapter="${ci}" data-search-topic="${ti}"><span>${escapeHtml(t.label)}</span><small>${escapeHtml(c.short)} →</small></button>`;});}
 });
 $('search-results').innerHTML=html||`<p class="search-empty">No topics found. Try “energy”, “signal”, or “cycle”.</p>`;
}
$('chapter-nav').addEventListener('click',e=>{const a=e.target.closest('[data-chapter]');if(a){e.preventDefault();routeTo(Number(a.dataset.chapter),0,{focusHeading:true,scroll:true});}});
$('chapter-select').addEventListener('change',e=>routeTo(Number(e.target.value),0,{focusHeading:true}));
$('topic-tabs').addEventListener('click',e=>{const b=e.target.closest('[data-topic]');if(b)routeTo(currentChapter,Number(b.dataset.topic),{focusTab:true});});
$('topic-tabs').addEventListener('keydown',e=>{
 const n=CHAPTERS[currentChapter].topics.length;let target;
 if(e.key==='ArrowRight')target=(currentTopic+1)%n;
 else if(e.key==='ArrowLeft')target=(currentTopic+n-1)%n;
 else if(e.key==='Home')target=0;else if(e.key==='End')target=n-1;
 if(target!==undefined){e.preventDefault();routeTo(currentChapter,target,{focusTab:true});}
});
$('concept-steps').addEventListener('click',e=>{const b=e.target.closest('[data-step]');if(!b)return;currentStep=Number(b.dataset.step);renderConcept();document.querySelector(`[data-step="${currentStep}"]`).focus({preventScroll:true});});
$('previous-chapter').addEventListener('click',()=>{if(currentChapter>0)routeTo(currentChapter-1,0,{focusHeading:true,scroll:true});});
$('next-chapter').addEventListener('click',e=>{e.preventDefault();routeTo((currentChapter+1)%CHAPTERS.length,0,{focusHeading:true,scroll:true});});
document.querySelector('.brand').addEventListener('click',e=>{e.preventDefault();routeTo(0,0);window.scrollTo({top:0,behavior:reduced.matches?'instant':'smooth'});});
$('focus-toggle').addEventListener('click',()=>{focusMode=!focusMode;applyFocus();try{localStorage.setItem('ronu-drafts-endurance-v1-focus',String(focusMode));}catch(_){}window.scrollTo({top:0,behavior:'instant'});});
$('menu-toggle').addEventListener('click',()=>setRail(!document.body.classList.contains('rail-open')));
$('close-menu').addEventListener('click',()=>{setRail(false);$('menu-toggle').focus();});
$('rail-overlay').addEventListener('click',()=>{setRail(false);$('menu-toggle').focus();});
mq.addEventListener('change',()=>setRail(false));
$('open-search').addEventListener('click',showSearch);
$('topic-search').addEventListener('input',e=>renderSearch(e.target.value));
$('search-results').addEventListener('click',e=>{const b=e.target.closest('[data-search-chapter]');if(!b)return;closeDialog('search-dialog');routeTo(Number(b.dataset.searchChapter),Number(b.dataset.searchTopic),{focusHeading:true,scroll:true});});
$('search-dialog').addEventListener('keydown',e=>{
 const results=[...$('search-results').querySelectorAll('.result')];const active=results.indexOf(document.activeElement);
 if(e.key==='ArrowDown'){e.preventDefault();if(results.length)results[(active+1)%results.length].focus();}
 if(e.key==='ArrowUp'){e.preventDefault();if(active<=0)$('topic-search').focus();else results[active-1].focus();}
 if(e.key==='Enter'&&document.activeElement===$('topic-search')&&results.length){e.preventDefault();results[0].click();}
});
document.querySelectorAll('[data-open-about]').forEach(b=>b.addEventListener('click',()=>{setRail(false);$('about-dialog').showModal();}));
document.querySelectorAll('[data-close]').forEach(b=>b.addEventListener('click',()=>closeDialog(b.dataset.close)));
document.querySelectorAll('dialog').forEach(d=>d.addEventListener('click',e=>{if(e.target!==d)return;const r=d.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)d.close();}));
document.addEventListener('keydown',e=>{
 const openDialog=document.querySelector('dialog[open]');
 if(e.key==='Escape'&&openDialog){e.preventDefault();openDialog.close();return;}
 const editable=e.target instanceof HTMLElement&&e.target.matches('input,textarea,select,[contenteditable=true]');
 if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(!$('search-dialog').open)showSearch();}
 if(e.key==='/'&&!editable&&!document.querySelector('dialog[open]')){e.preventDefault();showSearch();}
 if(e.key==='Escape'&&document.body.classList.contains('rail-open')){setRail(false);$('menu-toggle').focus();}
 // Keep keyboard navigation inside the mobile navigation while the drawer is open.
 if(e.key==='Tab'&&document.body.classList.contains('rail-open')){
  const items=[...$('sidebar').querySelectorAll('button,a[href]')];const first=items[0],last=items[items.length-1];
  if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}
  else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
 }
});
window.addEventListener('popstate',()=>render(...route()));
window.addEventListener('hashchange',()=>{const [ci,ti]=route();if(ci!==currentChapter||ti!==currentTopic)render(ci,ti);});
if(/Mac|iPhone|iPad/.test(navigator.platform))document.querySelector('.key').textContent='⌘ K';
applyFocus();render(...route());
