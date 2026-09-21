'use strict';
// The preview changes presentation, not the original article or simulation code.
let articleController=null;
const articleCache=new Map();
let pendingAnchor='';
const SOURCE_READER_STYLE=`
:host{display:block;color:#a5aba1;font:15px/1.85 -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;--bg:#101211;--panel:#101211;--panel2:#171a18;--line:#30362f;--line2:#3c4637;--text:#a5aba1;--soft:#c3c9bd;--bright:#efeee8;--gold:#c69257;--blue:#7898b2;--green:#829b77;--red:#a87369;--violet:#9384a5}
*,*::before,*::after{box-sizing:border-box} .page,main,.imported-body{width:100%!important;max-width:none!important;padding:0!important;margin:0!important;border:0!important;background:transparent!important;border-radius:0!important;display:block!important;min-height:0!important}
p{font-family:inherit!important;font-size:15px!important;line-height:1.85!important;color:#a5aba1!important;margin:0 0 1.1em!important}h2{font-family:Georgia,'Times New Roman',serif!important;font-size:29px!important;font-weight:400!important;line-height:1.2!important;color:#efeee8!important;margin:2.1em 0 .8em!important;letter-spacing:-.5px;scroll-margin-top:140px}h3{font-family:inherit!important;color:#d5dbc9!important;font-size:14px!important;line-height:1.5!important;margin:0 0 8px!important}strong,b{color:#d5dbc9!important}a{position:static!important;color:#c1d6a4!important;font-size:inherit!important;letter-spacing:normal!important;text-transform:none!important;text-decoration:underline!important;text-underline-offset:4px}a:focus-visible,summary:focus-visible{outline:2px solid #c1d6a4;outline-offset:4px}.back,nav,.kicker{display:none!important}.deck{display:none!important}.line,.sep{height:1px;background:#30362f!important;margin:2em 0!important}
pre{max-width:100%;overflow:auto;padding:20px!important;border:1px solid #30362f!important;border-radius:9px!important;background:#171a18!important;color:#d5dbc9!important;font:13px/1.7 Consolas,monospace!important}table{display:block;width:100%;overflow-x:auto;border-collapse:collapse;font-size:12px;line-height:1.7;margin:24px 0}td,th{padding:12px!important;border-bottom:1px solid #30362f!important;text-align:left;color:#a5aba1!important}th{color:#d5dbc9!important}svg,img{max-width:100%;height:auto}svg{display:block}.stage{background:#171a18!important;border-color:#30362f!important}.scale{font:10px/1.6 Consolas,monospace!important;color:#8d997f!important;text-transform:uppercase;letter-spacing:.12em;margin-top:3em!important}.stage p,.energy-card p,.path-node p,.cycle-card p,.signal-col p,.signal-note p,.caveat p{font-size:12px!important;line-height:1.75!important}.sources{padding-top:24px!important;margin-top:36px!important}.sources summary{font-size:12px!important;color:#c1d6a4!important;line-height:1.7;cursor:pointer}.sources li{font-size:12px!important;line-height:1.8!important;color:#a5aba1!important;margin:14px 0!important}.sources ol{padding-left:20px}.signal-list{padding-left:0}.signal-list li{font-size:12px!important;line-height:1.75!important;color:#a5aba1!important}.name,.change,.key-row,.caption,.small{color:#9aa591!important}.state .name{font-size:12px}.state .change{font-size:11px}.chain{border-left:2px solid #526344!important;padding-left:20px!important;color:#c3c9bd!important}.hero{margin:0!important}.hero-sub{color:#a5aba1!important}.intro{margin:0!important}
@media(max-width:600px){p{font-size:14px!important}h2{font-size:25px!important}.signal-path,.cycle,.energy-grid,.state-grid,.signal-map,.two-signals,.race,.atp-layout{grid-template-columns:1fr!important}.network{grid-template-columns:1fr 1fr!important}.path-node:after,.network-node:after{display:none!important}.stage{padding:16px!important}table{font-size:11px}}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
`;
function sourceURL(topic){return SOURCE.site+topic.path;}
function stopArticle(){if(articleController){articleController.abort();articleController=null;}}
function goAnchor(id){
 if(!id)return;
 const target=$('#source-content')?.shadowRoot?.getElementById(id)||document.getElementById(id);
 if(target){target.scrollIntoView({behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'instant':'smooth',block:'start'});pendingAnchor='';}
 else pendingAnchor=id;
}
function sourceFragment(html,topic){
 const doc=new DOMParser().parseFromString(html,'text/html');
 if(!doc.querySelector('h1')&&!topic.inline)throw new Error('Unexpected source document');
 const css=Array.from(doc.querySelectorAll('style')).map(n=>n.textContent).join('\n').replace(/:root/g,':host').replace(/@import[^;]+;/gi,'');
 doc.querySelectorAll('script,link,meta,base,iframe,object,embed,form,input,button,textarea,select,nav,header nav,foreignObject').forEach(n=>n.remove());
 doc.querySelectorAll('h1,.kicker,.deck,.back').forEach(n=>n.remove());
 // Only import the actual article, not the source page's global navigation.
 const source=doc.querySelector('main')||doc.body;
 source.querySelectorAll('a').forEach(a=>{
  try{const u=new URL(a.getAttribute('href')||'',sourceURL(topic));if(!['https:','http:'].includes(u.protocol)){a.removeAttribute('href');return;}if(u.origin===SOURCE.site&&u.pathname==='/'){a.remove();return;}a.href=u.href;a.target='_blank';a.rel='noopener noreferrer';}catch{a.removeAttribute('href');}
 });
 source.querySelectorAll('*').forEach(n=>{Array.from(n.attributes).forEach(a=>{if(/^on/i.test(a.name)||['srcdoc','autofocus','formaction'].includes(a.name))n.removeAttribute(a.name);});if(n.hasAttribute('src')){try{const u=new URL(n.getAttribute('src'),sourceURL(topic));if(u.protocol==='https:')n.setAttribute('src',u.href);else n.removeAttribute('src');}catch{n.removeAttribute('src');}}});
 source.querySelectorAll('style').forEach(n=>n.remove());
 const host=$('#source-content');host.replaceChildren();const root=host.shadowRoot||host.attachShadow({mode:'open'});root.replaceChildren();
 const style=document.createElement('style');style.textContent=css+'\n'+SOURCE_READER_STYLE;root.append(style);
 const body=document.createElement('div');body.className='imported-body';body.append(...Array.from(source.childNodes).map(n=>document.importNode(n,true)));root.append(body);
 const headings=Array.from(root.querySelectorAll('h2'));headings.forEach((h,i)=>h.id='part-'+(i+1));
 $('#reader-toc').innerHTML='<a href="#topic/'+topic.id+'/article-start" data-anchor="article-start" class="is-active">Overview</a>'+headings.map(h=>`<a href="#topic/${topic.id}/${h.id}" data-anchor="${h.id}">${esc(h.textContent.trim())}</a>`).join('');
 if(topic.status)host.insertAdjacentHTML('afterend',`<p class="intro-only">${esc(topic.status)}. This is all the content currently published on this page.</p>`);
 if('IntersectionObserver' in window){if(observer)observer.disconnect();observer=new IntersectionObserver(entries=>{const e=entries.find(x=>x.isIntersecting);if(e)$$('.toc a').forEach(a=>a.classList.toggle('is-active',a.dataset.anchor===e.target.id));},{rootMargin:'-15% 0px -65% 0px'});headings.forEach(h=>observer.observe(h));}
 if(pendingAnchor)goAnchor(pendingAnchor);updateProgress();
}
async function loadArticle(topic){
 stopArticle();const controller=new AbortController();articleController=controller;const currentHost=$('#source-content');
 currentHost.innerHTML='<p class="load-state" role="status">Loading the published article…</p>';
 try{
  let html=topic.inline||articleCache.get(topic.id);
  if(!html){
   let lastError;
   const paths=[`https://raw.githubusercontent.com/${SOURCE.repo}/${SOURCE.revision}/${topic.file}`,`https://raw.githack.com/${SOURCE.repo}/${SOURCE.revision}/${topic.file}`];
   for(const url of paths){
    try{const response=await fetch(url,{signal:AbortSignal.any([controller.signal,AbortSignal.timeout(12000)]),credentials:'omit',referrerPolicy:'no-referrer'});if(!response.ok)throw new Error('Source HTTP '+response.status);const text=await response.text();if(text.length>200000||!/<h1[\s>]/i.test(text))throw new Error('Invalid source response');html=text;break;}catch(error){lastError=error;if(controller.signal.aborted)throw error;}
   }
   if(!html)throw lastError;articleCache.set(topic.id,html);
  }
  if(controller.signal.aborted||currentTopic?.id!==topic.id||$('#source-content')!==currentHost)return;
  sourceFragment(html,topic);
 }catch(error){
  if(controller.signal.aborted||currentTopic?.id!==topic.id||$('#source-content')!==currentHost)return;
  currentHost.innerHTML=`<div class="load-state" role="status"><p>The original article could not be loaded in this preview. You can retry or read the published page directly.</p><button type="button" id="retry-source">Retry article</button> <a class="source-link" href="${esc(sourceURL(topic))}" target="_blank" rel="noopener noreferrer">Open original page ↗</a></div>`;
  $('#retry-source').addEventListener('click',()=>loadArticle(topic));
 }
}
function renderReader(topic){
 stopArticle();pendingAnchor='';if(observer){observer.disconnect();observer=null;}currentTopic=topic;
 const s=SECTIONS[topic.section],related=TOPICS.find(t=>t.section===topic.section&&t.id!==topic.id);
 $('#browse-page').hidden=true;$('#reader-page').hidden=false;$('#reading-progress').hidden=false;
 $$('.nav a').forEach(a=>a.dataset.section===topic.section?a.setAttribute('aria-current','page'):a.removeAttribute('aria-current'));
 $('#reader-page').innerHTML=`<div class="reader-top"><a class="quiet-link" href="${esc(lastBrowseHash)}"><svg class="icon" aria-hidden="true"><use href="#i-back"/></svg>Back to exploring</a><span class="reader-demo-label">Site preview / Current content</span></div><div class="reader-layout"><aside class="reader-sidebar" aria-label="Article navigation"><p class="eyebrow">In this topic</p><nav class="toc" id="reader-toc" aria-label="On this page"><a href="#topic/${topic.id}/article-start" data-anchor="article-start" class="is-active">Overview</a>${topic.type==='Interactive'?`<a href="#topic/${topic.id}/experiment" data-anchor="experiment">The experiment</a>`:''}</nav><div class="reader-status">${topic.type==='Article'?'Published text.<br>New reading layout.':'Original model.<br>Unchanged controls.'}<br><br><a class="source-link" href="${esc(sourceURL(topic))}" target="_blank" rel="noopener noreferrer">Open original ↗</a></div></aside><article aria-labelledby="reader-title" id="article-start"><header class="reader-head"><div class="eyebrow">${s.number} / ${s.name} &nbsp; — &nbsp; ${topic.type}</div><h1 id="reader-title" tabindex="-1">${esc(topic.title)}</h1><p class="reader-deck">${esc(topic.description)}</p><div class="reader-byline"><span>Rafael Renó</span><span>${topic.status||'Published on Ronu'}</span></div></header><div class="article-body">${topic.type==='Article'?'<div class="source-content" id="source-content"></div>':`<section id="experiment" class="article-section"><div class="reader-hero-art">${artSvg(topic.art)}<span class="art-label">Schematic illustration / Not model output</span></div><p class="launch-copy">Open the existing experiment inside this preview. Its original controls and model code are preserved.</p><button type="button" class="launch-button" id="launch-experiment">Load interactive ${arrow}</button><div id="app-mount"></div></section>`}<p class="source-credit">Content source: <a href="${esc(sourceURL(topic))}" target="_blank" rel="noopener noreferrer">the published Ronu page ↗</a>.<br>${topic.type==='Article'?`Article source snapshot: ${SOURCE.date}. Text, references and original diagrams are retained; the layout is restyled. Content has not been re-reviewed here.`:'The experiment is loaded from its existing live page. Its original design remains inside the frame.'}</p>${related?`<div class="related"><div class="eyebrow">Keep exploring / ${s.name}</div><a class="related-link" href="#topic/${related.id}"><div><strong>${esc(related.title)}</strong><p>${esc(related.description)}</p></div>${arrow}</a></div>`:''}</div></article></div>`;
 document.title=topic.title+' — Ronu preview';
 if(topic.type==='Article')loadArticle(topic);else $('#launch-experiment').addEventListener('click',()=>launchExperiment(topic));
 updateProgress();
}
function launchExperiment(topic){
 $('#launch-experiment').hidden=true;
 $('#app-mount').innerHTML=`<div class="original-app" id="original-app"><div class="app-controls"><span class="reader-progress-text">Original interactive</span><div><button type="button" id="expand-app" aria-pressed="false">Expand</button> <button type="button" id="close-app">Close</button> <a class="source-link" href="${esc(sourceURL(topic))}" target="_blank" rel="noopener noreferrer">Open separately ↗</a></div></div><iframe title="${esc(topic.title)} — original interactive" src="${esc(sourceURL(topic))}" referrerpolicy="no-referrer" sandbox="allow-scripts allow-same-origin allow-forms allow-popups allow-popups-to-escape-sandbox allow-downloads" allow="fullscreen"></iframe><p class="app-note">Original page from ronu.one. If it does not display correctly, use “Open separately”. Escape exits expanded view.</p></div>`;
 $('#expand-app').addEventListener('click',()=>{const expanded=$('#original-app').classList.toggle('is-expanded');$('#expand-app').textContent=expanded?'Reduce':'Expand';$('#expand-app').setAttribute('aria-pressed',String(expanded));});
 $('#close-app').addEventListener('click',()=>{$('#app-mount').replaceChildren();$('#launch-experiment').hidden=false;$('#launch-experiment').focus();});
}
