'use strict';
let svgCounter=0;
/** Schematic artwork only; these paths are not data or model outputs. */
function artSvg(kind,seed=0){
 const uid='art-'+(++svgCounter);
 let lines='',dots='',shapes='';
 const grid='<path d="M0 50H500M0 100H500M0 150H500M0 200H500M0 250H500M50 0V300M100 0V300M150 0V300M200 0V300M250 0V300M300 0V300M350 0V300M400 0V300M450 0V300" stroke="#b5c99e" stroke-opacity=".045" fill="none"/>';
 if(kind==='motion'){
   for(let i=0;i<9;i++){const angle=-35+i*9;lines+=`<ellipse cx="263" cy="147" rx="145" ry="${41+i*3}" transform="rotate(${angle} 263 147)" fill="none" stroke="#bbd29d" stroke-width="${i===4?1.8:1}" stroke-opacity="${i===4?.85:.16+i*.04}"/>`;}
   lines+='<path d="M80 147H448M263 23V276" stroke="#b6c99e" stroke-opacity=".11" stroke-dasharray="3 6"/>';
   dots='<circle cx="263" cy="147" r="4" fill="#d4e5bb"/><circle cx="381" cy="78" r="5" fill="#d4e5bb"/><circle cx="381" cy="78" r="13" stroke="#bdd69f" fill="none" opacity=".2"/><circle cx="129" cy="191" r="3" fill="#b9ce9d"/>';
 }else if(kind==='network'){
   const nodes=[];for(let r=0;r<6;r++)for(let c=0;c<10;c++)nodes.push([83+c*38,53+r*38]);
   nodes.forEach(([x,y],i)=>{const on=(i*13+7)%17<10;if(i%10<9)lines+=`<path d="M${x} ${y}h38" stroke="#b8ce9c" stroke-opacity="${on?.28:.07}"/>`;if(i<50)lines+=`<path d="M${x} ${y}v38" stroke="#b8ce9c" stroke-opacity="${on?.2:.055}"/>`;dots+=`<circle cx="${x}" cy="${y}" r="${on?3.7:2.4}" fill="${on?'#bacf9f':'#47503e'}"/>`;});
   lines+='<path d="M159 129h38v38h76v-38h38v-38h38v38h76" stroke="#d2e5b6" stroke-opacity=".8" stroke-width="2" fill="none"/>';
 }else if(kind==='curves'||kind==='branches'){
   lines='<path d="M65 53V245H447" stroke="#b8cf9e" stroke-opacity=".25" fill="none"/>';
   if(kind==='curves'){
    shapes='<path d="M70 219C140 219 154 135 228 139S321 175 435 66L435 245H70Z" fill="url(#'+uid+'-fade)"/>';
    lines+='<path d="M70 219C140 219 154 135 228 139S321 175 435 66" stroke="#c8dda9" stroke-width="2" fill="none"/><path d="M70 216C133 175 177 167 240 172S333 96 435 102" stroke="#819275" stroke-width="1.6" fill="none"/><path d="M70 218C169 188 270 149 435 161" stroke="#91a382" stroke-width="1.2" stroke-dasharray="4 6" fill="none"/>';dots='<circle cx="228" cy="139" r="4" fill="#d5e9b7"/><circle cx="435" cy="66" r="4" fill="#d5e9b7"/>';
   }else{for(let i=0;i<7;i++)lines+=`<path d="M70 175C145 183 170 129 238 142S350 ${55+i*25} 435 ${48+i*26}" stroke="#bad39f" stroke-width="${i===2?2:1}" stroke-opacity="${i===2?.95:.2+i*.05}" fill="none"/>`;lines+='<path d="M238 58V246" stroke="#b8cf9e" stroke-opacity=".3" stroke-dasharray="3 5"/>';dots='<circle cx="238" cy="142" r="4" fill="#d5e9b7"/>';}
 }else if(kind==='cube'||kind==='blocks'){
   const cube=(x,y,s,o)=>`<g transform="translate(${x} ${y})" opacity="${o}"><path d="M0 0 ${s} ${-s*.53} ${s*2} 0 ${s} ${s*.53}Z" fill="#a9c38c" fill-opacity=".10"/><path d="M0 0V${s*1.12}L${s} ${s*1.65}V${s*.53}M${s} ${s*1.65} ${s*2} ${s*1.12}V0" fill="#accc88" fill-opacity=".035"/><path d="M0 0 ${s} ${-s*.53} ${s*2} 0V${s*1.12}L${s} ${s*1.65} 0 ${s*1.12}V0L${s} ${s*.53} ${s*2} 0M${s} ${s*.53}V${s*1.65}" stroke="#bfd2a7" stroke-width="1.5" fill="none"/></g>`;
   if(kind==='cube'){shapes=cube(165,96,86,1);lines='<path d="M104 222 279 124 438 211M250 26V279" stroke="#a9bd96" stroke-opacity=".2" stroke-dasharray="3 6" fill="none"/>';shapes+=cube(116,106,86,.2);dots='<circle cx="251" cy="50" r="3" fill="#d6e6c0"/>';}else{shapes=cube(72,125,51,.32)+cube(207,103,51,.62)+cube(341,81,51,1);lines='<path d="M184 171h18m-5-4 5 4-5 4M318 147h18m-5-4 5 4-5 4" fill="none" stroke="#c1d5a5" opacity=".45"/>';}
 }else if(kind==='waves'||kind==='rhythm'){
   for(let row=0;row<8;row++){let points=[];for(let x=38;x<=468;x+=3){let y=45+row*28+Math.sin(x*.038+row*.47)*(kind==='rhythm'?14:18)*Math.sin((x-38)/430*Math.PI);points.push(`${x},${y.toFixed(1)}`);}lines+=`<polyline points="${points.join(' ')}" fill="none" stroke="#c2d7a6" stroke-width="${row===3?2:1}" stroke-opacity="${row===3?.9:.18+row*.06}"/>`;}
 }else if(kind==='pendulum'){
   lines='<path d="M90 65H417M254 58V260" stroke="#9cb18b" stroke-opacity=".2" stroke-dasharray="3 6" fill="none"/>';
   for(let i=0;i<6;i++){const a=(-50+i*19)*Math.PI/180,x=254+Math.sin(a)*91,y=72+Math.cos(a)*91,x2=x+Math.sin(a+.6)*81,y2=y+Math.cos(a+.6)*81;lines+=`<path d="M254 72L${x} ${y} ${x2} ${y2}" stroke="#c3d9a8" stroke-width="${i===3?2:1}" stroke-opacity="${i===3?1:.18}" fill="none"/>`;dots+=`<circle cx="${x}" cy="${y}" r="${i===3?4:2.5}" fill="#bfd5a2" opacity="${i===3?1:.25}"/><circle cx="${x2}" cy="${y2}" r="${i===3?5:3}" fill="#bfd5a2" opacity="${i===3?1:.25}"/>`;}
   dots+='<circle cx="254" cy="72" r="5" fill="#cfe3b4"/>';
 }else if(kind==='lanes'){
   for(let row=0;row<6;row++){const y=50+row*39;lines+=`<path d="M49 ${y+13}H454" stroke="#9eaf8d" opacity=".14" stroke-dasharray="3 6"/>`;for(let j=0;j<12;j++){const x=50+j*34+(row%2)*9;shapes+=`<rect x="${x}" y="${y}" width="${(j+row)%5===0?21:11}" height="5" rx="2.5" fill="#c4d8a9" opacity="${j>6&&j<10?.8:.22+row*.06}"/>`;}}
 }else if(kind==='bars'){
   lines='<path d="M63 244H448" stroke="#b8cf9e" opacity=".2"/>';[97,144,112,176,138,191,156].forEach((h,i)=>{shapes+=`<rect x="${81+i*49}" y="${244-h}" width="26" height="${h}" rx="3" fill="#b8cf9b" fill-opacity="${i===3?.65:.11+i*.055}"/><path d="M${81+i*49} ${244-h}h26" stroke="#c5d9a9" stroke-opacity=".8"/>`;});
 }else if(kind==='space'){
  shapes='<path d="M94 234 265 283 426 187 255 139Z" fill="#aec68e" opacity=".035"/>';
  lines='<g stroke="#bace9f" fill="none"><path d="M94 234V82L265 30 426 87V237L265 283Z" opacity=".75"/><path d="M94 82 265 139 426 87M265 139V283M265 30V139" opacity=".45"/><path d="M94 234 265 182 426 237M265 139V182" opacity=".20" stroke-dasharray="3 5"/><path d="m139 119 57 19v65l-57-19ZM319 136l61-20v73l-61 20Z" opacity=".6"/><path d="m113 242 152-46 142 47M120 252l147-47 122 41" opacity=".17"/></g>';
 }
 return `<svg viewBox="0 0 500 300" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false" preserveAspectRatio="xMidYMid meet"><defs><radialGradient id="${uid}-glow"><stop stop-color="#9fbc77" stop-opacity=".075"/><stop offset="1" stop-color="#9fbc77" stop-opacity="0"/></radialGradient><linearGradient id="${uid}-fade" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#a5c281" stop-opacity=".15"/><stop offset="1" stop-color="#a5c281" stop-opacity="0"/></linearGradient></defs><rect width="500" height="300" fill="url(#${uid}-glow)"/>${grid}${shapes}${lines}${dots}</svg>`;
}
