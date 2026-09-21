'use strict';
const $ = (selector,root=document) => root.querySelector(selector);
const $$ = (selector,root=document) => Array.from(root.querySelectorAll(selector));
const esc = value => String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const arrow = '<svg class="icon" aria-hidden="true"><use href="#i-arrow"/></svg>';
let storedView='grid';try{storedView=localStorage.getItem('ronu-layout-view')==='list'?'list':'grid';}catch{/* File previews can block browser storage. */}
let state={section:'all',type:'all',query:'',view:storedView};
let currentTopic=null,lastBrowseHash='#explore',observer=null,scrollTick=false;
const searchDialog=$('#search-dialog'),aboutDialog=$('#about-dialog');
function browseHash(){
 const base=state.section==='all'?'#explore':'#section/'+state.section;
 const p=new URLSearchParams();if(state.type!=='all')p.set('format',state.type);if(state.query)p.set('q',state.query);if(state.view==='list')p.set('view','list');return base+(p.size?'?'+p.toString():'');
}
function updateBrowseUrl(){lastBrowseHash=browseHash();try{history.replaceState(null,'',lastBrowseHash);}catch{/* Filtering remains functional without URL history access. */}}
function matchTopics(topics,query){const q=query.trim().toLocaleLowerCase();if(!q)return topics;const words=q.split(/\s+/);return topics.filter(t=>words.every(w=>(t.title+' '+t.description+' '+SECTIONS[t.section].name+' '+t.type+' '+(t.tags||'')).toLocaleLowerCase().includes(w)));}
function renderFeature(){
 const topic=TOPICS.find(t=>t.id===SECTIONS[state.section].featured),f=$('#featured-topic'); if(!topic){f.hidden=true;f.innerHTML='';return;} const section=SECTIONS[topic.section];
 f.href='#topic/'+topic.id;
 f.innerHTML=`<div class="feature-copy"><div class="feature-top"><span class="feature-badge">Featured exploration</span><span class="divider"></span><span class="eyebrow">${section.name}</span></div><h2 id="feature-title">${esc(topic.title)}</h2><p class="feature-description">${esc(topic.description)}</p><span class="feature-cta">${topic.type==='Interactive'?'Explore the model':'Read the article'} ${arrow}</span></div><div class="feature-art">${artSvg(topic.art)}<span class="feature-number">${section.number} / EXPLORATION</span><span class="art-label">An idea in motion / schematic illustration</span></div>`;
}
function card(topic){const s=SECTIONS[topic.section];return `<a class="topic-card" href="#topic/${topic.id}" aria-labelledby="title-${topic.id}"><div class="card-art">${artSvg(topic.art)}<span class="card-section-number">${s.number} / ${topic.type==='Project'?'BUILD':'EXPLORE'}</span></div><div class="card-content"><div class="card-meta"><span class="eyebrow">${s.name}</span><span class="card-type">${topic.type}</span></div><h3 class="card-title" id="title-${topic.id}">${esc(topic.title)}</h3><p class="card-description">${esc(topic.description)}</p><div class="card-bottom"><span>${esc(topic.status||(topic.type==='Interactive'?'Open experiment':'Read article'))}</span>${arrow}</div></div></a>`;}
function renderLibrary(){
 let topics=TOPICS.filter(t=>state.section==='all'||t.section===state.section);if(state.type!=='all')topics=topics.filter(t=>t.type===state.type);topics=matchTopics(topics,state.query);
 $('#topic-grid').innerHTML=topics.map(card).join('');$('#topic-grid').classList.toggle('is-list',state.view==='list');$('#empty-state').hidden=topics.length!==0;
 $('#topic-count').textContent=String(topics.length).padStart(2,'0')+' '+(topics.length===1?'topic':'topics');
 $$('.filter-button').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.type===state.type)));
 $('#grid-view').setAttribute('aria-pressed',String(state.view==='grid'));$('#list-view').setAttribute('aria-pressed',String(state.view==='list'));
 $('#topic-filter').value=state.query;
 // Keep search results close to the controls rather than above a repeated feature.
 $('#featured-topic').hidden=!SECTIONS[state.section].featured||state.type!=='all'||state.query.trim()!=='';
 const sectionEmpty=!TOPICS.some(t=>state.section==='all'||t.section===state.section);
 $('#empty-state').innerHTML=sectionEmpty?`<div class="no-content"><h3>${esc(SECTIONS[state.section].name)}</h3><p>${esc(SECTIONS[state.section].empty||'No published topics are listed here.')}</p><a href="#explore">Explore the available topics →</a></div>`:`<h3>No topics found.</h3><p>Try another word or reset the filters.</p><button type="button" id="clear-filters">Clear filters</button>`;
 if(!sectionEmpty)$('#clear-filters').addEventListener('click',clearFilters);
 $('.browse-tools').hidden=sectionEmpty;
}
function renderBrowse(){
 stopArticle();
 // Remove the reader so hidden models do not continue running after navigation.
 $('#reader-page').replaceChildren();
 if(observer){observer.disconnect();observer=null;}currentTopic=null;
 const s=SECTIONS[state.section];$('#browse-page').hidden=false;$('#reader-page').hidden=true;$('#reading-progress').hidden=true;
 $('#page-title').innerHTML=s.headline;$('#intro-description').innerHTML=s.description;$('#intro-note').textContent=s.note;$('#intro-eyebrow').innerHTML='<span class="tiny-dot"></span>'+s.eyebrow;
 $('#library-title').textContent=state.section==='all'?'Find your next rabbit hole.':'Explore '+s.name.toLowerCase()+'.';
 $$('.nav a').forEach(a=>a.dataset.section===state.section?a.setAttribute('aria-current','page'):a.removeAttribute('aria-current'));
 document.title='Ronu.one — '+s.name;renderFeature();renderLibrary();lastBrowseHash=browseHash();
}
function navigate(){
 if(searchDialog.open)searchDialog.close();if(aboutDialog.open)aboutDialog.close();
 const hash=location.hash||'#explore';
 // The skip link is an in-page accessibility target, not a route.
 if(hash==='#main'){document.getElementById('main').focus();return;}
 const parts=hash.slice(1).split('?'),path=parts[0],params=new URLSearchParams(parts[1]||'');
 if(path.startsWith('topic/')){
   const [,id,anchor]=path.split('/'),topic=TOPICS.find(t=>t.id===id);
   if(topic){if(!currentTopic||currentTopic.id!==id){renderReader(topic);window.scrollTo({top:0,behavior:'instant'});$('#reader-title').setAttribute('tabindex','-1');$('#reader-title').focus({preventScroll:true});}if(anchor&&/^(part-\d+|article-start|experiment)$/.test(anchor))goAnchor(anchor);return;}
 }
 const proposed=path.startsWith('section/')?path.split('/')[1]:'all';state.section=Object.hasOwn(SECTIONS,proposed)?proposed:'all';
 state.type=['Article','Interactive','Project'].includes(params.get('format'))?params.get('format'):'all';state.query=params.get('q')||'';
 state.view=params.get('view')==='list'?'list':storedView;
 renderBrowse();window.scrollTo({top:0,behavior:'instant'});
}
function updateProgress(){if(!currentTopic)return;const max=document.documentElement.scrollHeight-window.innerHeight;const pct=max>0?Math.min(100,Math.max(0,window.scrollY/max*100)):100;$('#reading-progress span').style.width=pct+'%';}
function renderSearch(){const q=$('#global-search').value,topics=matchTopics(TOPICS,q);$('#search-results').innerHTML=topics.length?topics.map(t=>`<a class="search-result" href="#topic/${t.id}"><span class="result-symbol"><svg class="icon" aria-hidden="true"><use href="#i-${t.section}"/></svg></span><span class="result-copy"><strong>${esc(t.title)}</strong><span>${SECTIONS[t.section].name} · ${t.type}</span></span>${arrow}</a>`).join(''):'<div class="search-no-results">No topics found. Try a different word.</div>';$('#search-result-count').textContent=topics.length+' '+(topics.length===1?'topic':'topics')+' across all sections';}
function openSearch(){if(aboutDialog.open)aboutDialog.close();$('#global-search').value='';renderSearch();searchDialog.showModal();$('#global-search').focus();}
function openAbout(){if(searchDialog.open)searchDialog.close();aboutDialog.showModal();}
$('#search-open').addEventListener('click',openSearch);$('#search-close').addEventListener('click',()=>searchDialog.close());$('#global-search').addEventListener('input',renderSearch);
$('#search-results').addEventListener('click',e=>{if(e.target.closest('a'))searchDialog.close();});
$('#about-open').addEventListener('click',openAbout);$('#design-notes-open').addEventListener('click',openAbout);$('#about-close').addEventListener('click',()=>aboutDialog.close());
[searchDialog,aboutDialog].forEach(d=>d.addEventListener('click',e=>{const r=d.getBoundingClientRect();if(e.target===d&&(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom))d.close();}));
document.addEventListener('keydown',e=>{if(e.key==='/'&&!['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)&&!document.activeElement.isContentEditable){e.preventDefault();if(!searchDialog.open)openSearch();}});
$('#topic-filter').addEventListener('input',e=>{state.query=e.target.value;renderLibrary();updateBrowseUrl();});
$$('.filter-button').forEach(b=>b.addEventListener('click',()=>{state.type=b.dataset.type;renderLibrary();updateBrowseUrl();}));
function clearFilters(){state.type='all';state.query='';renderLibrary();updateBrowseUrl();$('#topic-filter').focus();}
function changeView(view){state.view=view;storedView=view;try{localStorage.setItem('ronu-layout-view',view);}catch{}renderLibrary();updateBrowseUrl();}
$('#grid-view').addEventListener('click',()=>changeView('grid'));$('#list-view').addEventListener('click',()=>changeView('list'));
window.addEventListener('hashchange',navigate);
window.addEventListener('scroll',()=>{if(!scrollTick){requestAnimationFrame(()=>{updateProgress();scrollTick=false;});scrollTick=true;}},{passive:true});
window.addEventListener('resize',updateProgress);
document.addEventListener('keydown',e=>{if(e.key==='Escape'){if(searchDialog.open){e.preventDefault();searchDialog.close();}if(aboutDialog.open){e.preventDefault();aboutDialog.close();}}if(e.key==='Escape'&&$('#original-app')?.classList.contains('is-expanded')){$('#original-app').classList.remove('is-expanded');$('#expand-app').textContent='Expand';$('#expand-app').setAttribute('aria-pressed','false');$('#expand-app').focus();}});
navigate();
