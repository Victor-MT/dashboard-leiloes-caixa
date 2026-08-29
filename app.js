const COL={id:'ID',mode:'Modo Venda',state:'Estado',city:'Cidade',hood:'Bairro',address:'Endereço',type:'Tipo',p1:'Preço Venda 1º Leilão',p2:'Preço Venda 2º Leilão',date1:'Fim 1º Leilão',date2:'Fim 2º Leilão',valuation:'Preço Avaliação',built:'Área construída (m²)',land:'Área terreno (m²)',discount:'Desconto',finance:'Aceita Financiamento',condo:'Responsabilidade Condomínio',tax:'Responsabilidade Tributos',site:'Site',rent:'Link Zap Imóveis Aluguel',buy:'Link Zap Imóveis Compra'};
let data=[], filtered=[], page=1, charts={}, heatLayer=null, markerLayer=null, scoredCache=null,timelineOffset=0,selectedTimelineKey=null,timelineDetailPage=1;
let processingToken=0,filterTimer=null,fileImportStarted=false,detectedLocation=null,locationDefaultsApplied=false,locationFiltersTouched=false;
const $=id=>document.getElementById(id);
const nextFrame=()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
function setGlobalLoading(visible,title='Processando dados…',detail='Isso pode levar alguns instantes.'){$('globalLoading').hidden=!visible;document.body.classList.toggle('is-processing',visible);$('globalLoadingTitle').textContent=title;$('globalLoadingDetail').textContent=detail}
async function beginProcessing(title,detail){setGlobalLoading(true,title,detail);await nextFrame()}
function endProcessing(token){if(token==null||token===processingToken)setGlobalLoading(false)}
const normalize=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const aliases={
 id:['id','codigo','numero do imovel','numero imovel'],mode:['modo venda','modalidade de venda','modalidade'],state:['estado','uf'],city:['cidade','municipio'],hood:['bairro'],address:['endereco','logradouro'],type:['tipo','tipo de imovel'],p1:['preco venda 1 leilao','preco 1 leilao','valor 1 leilao'],p2:['preco venda 2 leilao','preco 2 leilao','valor 2 leilao'],date1:['fim 1 leilao','fim do 1 leilao','data fim 1 leilao'],date2:['fim 2 leilao','fim do 2 leilao','data fim 2 leilao'],valuation:['preco avaliacao','preco de avaliacao','valor avaliacao','valor de avaliacao'],built:['area construida m2','area construida'],land:['area terreno m2','area do terreno','area terreno'],discount:['desconto','desconto percentual','percentual desconto'],finance:['aceita financiamento','financiamento','financiavel'],condo:['responsabilidade condominio'],tax:['responsabilidade tributos'],site:['site','link caixa','url caixa'],rent:['link zap imoveis aluguel','link aluguel'],buy:['link zap imoveis compra','link compra']
};
const num=v=>{if(v==null||v==='')return null;if(typeof v==='number')return Number.isFinite(v)?v:null;let s=String(v).trim().replace(/R\$|%/g,'').replace(/\s/g,'');if(!s)return null;if(s.includes(',')&&s.includes('.'))s=s.lastIndexOf(',')>s.lastIndexOf('.')?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');else if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');let n=Number(s);return Number.isFinite(n)?n:null};
function setUploadState(state,message,details){$('uploadBox').dataset.state=state;$('uploadMessage').textContent=message;$('uploadDetails').textContent=details;$('fileInput').disabled=state==='loading'}
function findHeaderRow(matrix){let best={index:-1,score:0};matrix.slice(0,15).forEach((row,index)=>{let names=(row||[]).map(normalize),score=Object.values(aliases).filter(list=>list.some(a=>names.includes(normalize(a)))).length;if(score>best.score)best={index,score}});return best.score>=3?best:null}
function inspectSheet(ws){if(!ws['!ref'])return null;let range=XLSX.utils.decode_range(ws['!ref']);range.e.r=Math.min(range.e.r,range.s.r+14);let rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true,range:XLSX.utils.encode_range(range)}),found=findHeaderRow(rows);return found?{score:found.score,headerRow:found.index}:null}
function normalizeSheet(ws){let matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true}),found=findHeaderRow(matrix);if(!found)throw new Error('HEADER_NOT_FOUND');let headers=matrix[found.index].map(normalize),mapping={};Object.entries(aliases).forEach(([key,list])=>{let index=headers.findIndex(h=>list.some(a=>h===normalize(a)));if(index>=0)mapping[key]=index});if(mapping.built==null)mapping.built=10;if(mapping.land==null)mapping.land=11;let missing=['id','state','city','address'].filter(k=>mapping[k]==null);if(missing.length)throw new Error(`MISSING_COLUMNS:${missing.map(k=>COL[k]).join(', ')}`);let rows=matrix.slice(found.index+1).filter(row=>row.some(v=>v!=null&&v!=='')).map(row=>{let out={};Object.entries(mapping).forEach(([key,index])=>out[COL[key]]=row[index]);return out});let ds=rows.map(r=>num(r[COL.discount])).filter(Number.isFinite);if(ds.length&&quant(ds,.9)<=1)rows.forEach(r=>{let d=num(r[COL.discount]);if(d!=null)r[COL.discount]=d*100});return {rows,headerRow:found.index+1,rawCount:matrix.length-found.index-1}}
function excelWorkerRuntime(){
  importScripts('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
  const normalize=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const numberValue=value=>{if(value==null||value==='')return null;if(typeof value==='number')return Number.isFinite(value)?value:null;let text=String(value).trim().replace(/R\$|%/g,'').replace(/\s/g,'');if(!text)return null;if(text.includes(',')&&text.includes('.'))text=text.lastIndexOf(',')>text.lastIndexOf('.')?text.replace(/\./g,'').replace(',','.'):text.replace(/,/g,'');else if(text.includes(','))text=text.replace(/\./g,'').replace(',','.');let result=Number(text);return Number.isFinite(result)?result:null};
  const quantile=(values,q)=>{values=values.filter(Number.isFinite).sort((a,b)=>a-b);return values.length?values[Math.floor((values.length-1)*q)]:null};
  self.onmessage=event=>{try{const {buffer,aliases,columns}=event.data;const findHeader=matrix=>{let best={index:-1,score:0};matrix.slice(0,15).forEach((row,index)=>{let names=(row||[]).map(normalize),score=Object.values(aliases).filter(list=>list.some(alias=>names.includes(normalize(alias)))).length;if(score>best.score)best={index,score}});return best.score>=3?best:null};const workbook=XLSX.read(buffer,{type:'array',dense:true});let sheetName='',bestScore=-1;for(const name of workbook.SheetNames){let sheet=workbook.Sheets[name];if(!sheet['!ref'])continue;let range=XLSX.utils.decode_range(sheet['!ref']);range.e.r=Math.min(range.e.r,range.s.r+14);let found=findHeader(XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,range:XLSX.utils.encode_range(range)}));if(found&&found.score>bestScore){bestScore=found.score;sheetName=name}}if(!sheetName)throw new Error('HEADER_NOT_FOUND');let matrix=XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:null,raw:true}),found=findHeader(matrix),headers=matrix[found.index].map(normalize),mapping={};Object.entries(aliases).forEach(([key,list])=>{let index=headers.findIndex(header=>list.some(alias=>header===normalize(alias)));if(index>=0)mapping[key]=index});if(mapping.built==null)mapping.built=10;if(mapping.land==null)mapping.land=11;let missing=['id','state','city','address'].filter(key=>mapping[key]==null);if(missing.length)throw new Error(`MISSING_COLUMNS:${missing.map(key=>columns[key]).join(', ')}`);let rawCount=matrix.length-found.index-1,rows=matrix.slice(found.index+1).filter(row=>row.some(value=>value!=null&&value!=='')).map(row=>{let output={};Object.entries(mapping).forEach(([key,index])=>output[columns[key]]=row[index]);return output});let discounts=rows.map(row=>numberValue(row[columns.discount])).filter(Number.isFinite);if(discounts.length&&quantile(discounts,.9)<=1)rows.forEach(row=>{let discount=numberValue(row[columns.discount]);if(discount!=null)row[columns.discount]=discount*100});self.postMessage({rows,rawCount,headerRow:found.index+1,sheetName})}catch(error){self.postMessage({error:error.message||'PARSE_ERROR'})}};
}
function parseExcelInWorker(buffer){return new Promise((resolve,reject)=>{let worker,workerUrl;try{workerUrl=URL.createObjectURL(new Blob([`(${excelWorkerRuntime.toString()})()`],{type:'text/javascript'}));worker=new Worker(workerUrl)}catch(error){if(workerUrl)URL.revokeObjectURL(workerUrl);reject(new Error(`WORKER_UNAVAILABLE:${error.message}`));return}let finish=()=>{worker.terminate();URL.revokeObjectURL(workerUrl)};worker.onmessage=event=>{finish();event.data.error?reject(new Error(event.data.error)):resolve(event.data)};worker.onerror=event=>{finish();reject(new Error(`WORKER_UNAVAILABLE:${event.message||'erro desconhecido'}`))};worker.postMessage({buffer,aliases,columns:COL},[buffer])})}
const DB_NAME='radar-oportunidades',DB_VERSION=1,DATA_SCHEMA_VERSION=1,STORE_NAME='datasets';
function openDatabase(){return new Promise((resolve,reject)=>{if(!('indexedDB'in window)){reject(new Error('INDEXEDDB_UNAVAILABLE'));return}let request=indexedDB.open(DB_NAME,DB_VERSION);request.onupgradeneeded=()=>{let db=request.result;if(!db.objectStoreNames.contains(STORE_NAME))db.createObjectStore(STORE_NAME,{keyPath:'id'})};request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error)})}
async function readSavedDataset(){let db=await openDatabase();try{return await new Promise((resolve,reject)=>{let request=db.transaction(STORE_NAME,'readonly').objectStore(STORE_NAME).get('latest');request.onsuccess=()=>resolve(request.result||null);request.onerror=()=>reject(request.error)})}finally{db.close()}}
async function saveDataset(rows,metadata){let db=await openDatabase();try{await new Promise((resolve,reject)=>{let transaction=db.transaction(STORE_NAME,'readwrite');transaction.objectStore(STORE_NAME).put({id:'latest',schemaVersion:DATA_SCHEMA_VERSION,rows,metadata});transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error);transaction.onabort=()=>reject(transaction.error)})}finally{db.close()}}
async function deleteSavedDataset(){let db=await openDatabase();try{await new Promise((resolve,reject)=>{let transaction=db.transaction(STORE_NAME,'readwrite');transaction.objectStore(STORE_NAME).delete('latest');transaction.oncomplete=resolve;transaction.onerror=()=>reject(transaction.error)})}finally{db.close()}}
const savedDate=value=>new Intl.DateTimeFormat('pt-BR',{dateStyle:'short',timeStyle:'short'}).format(new Date(value));
function showSavedMetadata(metadata){$('removeSavedData').hidden=false;$('uploadDetails').textContent=`${metadata.rowCount.toLocaleString('pt-BR')} imóveis · salva em ${savedDate(metadata.importedAt)} · somente neste navegador.`}
const STATE_UF={'acre':'AC','alagoas':'AL','amapa':'AP','amazonas':'AM','bahia':'BA','ceara':'CE','distrito federal':'DF','espirito santo':'ES','goias':'GO','maranhao':'MA','mato grosso':'MT','mato grosso do sul':'MS','minas gerais':'MG','para':'PA','paraiba':'PB','parana':'PR','pernambuco':'PE','piaui':'PI','rio de janeiro':'RJ','rio grande do norte':'RN','rio grande do sul':'RS','rondonia':'RO','roraima':'RR','santa catarina':'SC','sao paulo':'SP','sergipe':'SE','tocantins':'TO'};
function browserCoordinates(){return new Promise((resolve,reject)=>{if(!navigator.geolocation){reject(new Error('GEOLOCATION_UNAVAILABLE'));return}navigator.geolocation.getCurrentPosition(position=>resolve(position.coords),reject,{enableHighAccuracy:false,timeout:8000,maximumAge:3600000})})}
async function detectUserLocation(){try{let cached=sessionStorage.getItem('dashboardLocationV1');if(cached)return JSON.parse(cached);let coordinates=await browserCoordinates(),url=`https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${encodeURIComponent(coordinates.latitude)}&lon=${encodeURIComponent(coordinates.longitude)}&zoom=10&addressdetails=1`,response=await fetch(url,{headers:{'Accept-Language':'pt-BR'}});if(!response.ok)throw new Error(`REVERSE_GEOCODING_${response.status}`);let address=(await response.json()).address||{},iso=address['ISO3166-2-lvl4']||address['ISO3166-2-lvl3']||'',state=iso.match(/BR-([A-Z]{2})/i)?.[1]?.toUpperCase()||STATE_UF[normalize(address.state)],city=address.city||address.municipality||address.town||address.village;if(!state||!city)return null;let location={state,city};sessionStorage.setItem('dashboardLocationV1',JSON.stringify(location));return location}catch(error){console.info('Localização automática não aplicada:',error.message);return null}}
async function applyLocationDefaults(){if(locationDefaultsApplied||locationFiltersTouched||!detectedLocation||!data.length)return false;let stateOption=[...$('fState').options].find(option=>normalize(option.value)===normalize(detectedLocation.state));if(!stateOption)return false;$('fState').value=stateOption.value;refreshDependent();let cityOption=[...$('fCity').options].find(option=>normalize(option.value)===normalize(detectedLocation.city));if(!cityOption){$('fState').value='';refreshDependent();return false}$('fCity').value=cityOption.value;refreshDependent();locationDefaultsApplied=true;return true}
async function startLocationDetection(){detectedLocation=await detectUserLocation();if(!detectedLocation||!data.length||locationFiltersTouched)return;let applied=await applyLocationDefaults();if(applied)await applyFilters()}
const money=v=>v==null?'—':new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:0}).format(v);
const pct=v=>v==null?'—':`${new Intl.NumberFormat('pt-BR',{maximumFractionDigits:1}).format(v)}%`;
const pctHtml=v=>{let n=num(v);return `<span class="${n<0?'negative-percent':''}">${pct(n)}</span>`};
const med=a=>{a=a.filter(Number.isFinite).sort((x,y)=>x-y); if(!a.length)return null; let m=Math.floor(a.length/2);return a.length%2?a[m]:(a[m-1]+a[m])/2};
const quant=(a,q)=>{a=a.filter(Number.isFinite).sort((x,y)=>x-y);if(!a.length)return null;return a[Math.floor((a.length-1)*q)]};
const startDay=value=>{let d=new Date(value);return new Date(d.getFullYear(),d.getMonth(),d.getDate())};
const addDays=(date,days)=>{let d=new Date(date);d.setDate(d.getDate()+days);return startDay(d)};
const dayKey=date=>`${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
const dayDiff=(a,b)=>Math.round((startDay(a)-startDay(b))/86400000);
function parseAuctionDate(value){if(value==null||value==='')return null;if(value instanceof Date&&!Number.isNaN(value.getTime()))return startDay(value);if(typeof value==='number'){let parsed=typeof XLSX!=='undefined'&&XLSX.SSF?XLSX.SSF.parse_date_code(value):null;return parsed?new Date(parsed.y,parsed.m-1,parsed.d):null}let text=String(value).trim(),match=text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);if(match){let year=+match[3];if(year<100)year+=2000;let date=new Date(year,+match[2]-1,+match[1]);return Number.isNaN(date.getTime())?null:date}let date=new Date(text);return Number.isNaN(date.getTime())?null:startDay(date)}
function salePrice(r){let m=r[COL.mode],p1=num(r[COL.p1]),p2=num(r[COL.p2]);return m==='2º Leilão'&&p2!=null?p2:(p1??p2)}
function calculatedDiscount(r,sale){let valuation=num(r[COL.valuation]),price=num(sale);return valuation>0&&price!=null?100*(valuation-price)/valuation:num(r[COL.discount])}
function propertyArea(r){let type=normalize(r[COL.type]);let built=num(r[COL.built]),land=num(r[COL.land]);return type.includes('terreno')?(land??built):(built??land)}
function sqm(r){let p=salePrice(r),a=propertyArea(r);return p!=null&&a>0?p/a:null}
function street(r){let a=String(r[COL.address]||'').split(',')[0].trim();return a||'Não informado'}
function enrich(rows){return rows.map(r=>({...r,_price:salePrice(r),_sqm:sqm(r),_street:street(r)}))}
function uniq(col,rows=data){return [...new Set(rows.map(r=>r[col]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'pt-BR'))}
const multiFilters={fNeighborhood:{toggle:'fNeighborhoodToggle',menu:'fNeighborhoodMenu',all:'Todos os bairros',selected:'bairros selecionados'},fType:{toggle:'fTypeToggle',menu:'fTypeMenu',all:'Todos os tipos',selected:'tipos selecionados'},fMode:{toggle:'fModeToggle',menu:'fModeMenu',all:'Todos os modos',selected:'modos selecionados'}};
function setOptions(id,vals,keep=true){let el=$(id),old=keep?(el.multiple?[...el.selectedOptions].map(o=>o.value):el.value):(el.multiple?[]:'');el.innerHTML=(el.multiple?'':'<option value="">Todos</option>')+vals.map(v=>`<option>${esc(v)}</option>`).join('');if(el.multiple)[...el.options].forEach(o=>o.selected=old.includes(o.value));else if(vals.includes(old))el.value=old;if(multiFilters[id])renderMultiDropdown(id)}
function renderMultiDropdown(id){let config=multiFilters[id],select=$(id),selected=[...select.selectedOptions].map(o=>o.value),menu=$(config.menu),toggle=$(config.toggle),allSelected=select.options.length>0&&selected.length===select.options.length;toggle.firstChild.textContent=!selected.length||allSelected?`${config.all} `:selected.length===1?selected[0]:`${selected.length} ${config.selected} `;menu.innerHTML=`<div class="multi-actions"><button type="button" data-action="all">Selecionar todos</button><button type="button" data-action="clear">Limpar</button></div>`+[...select.options].map(option=>`<label class="multi-option"><input type="checkbox" value="${esc(option.value)}" ${option.selected?'checked':''}> <span>${esc(option.textContent)}</span></label>`).join('');let commit=()=>{renderMultiDropdown(id);select.dispatchEvent(new Event('change',{bubbles:true}))};menu.querySelector('[data-action="all"]').onclick=()=>{[...select.options].forEach(option=>option.selected=true);commit()};menu.querySelector('[data-action="clear"]').onclick=()=>{[...select.options].forEach(option=>option.selected=false);commit()};menu.querySelectorAll('.multi-option input').forEach(input=>input.onchange=()=>{let option=[...select.options].find(item=>item.value===input.value);if(option)option.selected=input.checked;commit()})}
function esc(s){return String(s??'').replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]))}
function safeUrl(value){try{let text=String(value||'').trim();if(/^www\./i.test(text))text=`https://${text}`;let url=new URL(text);return ['http:','https:'].includes(url.protocol)?url.href:null}catch{return null}}
async function init(rows){let valid=rows.filter(r=>r&&r[COL.id]!=null&&String(r[COL.id]).trim()!=='');data=[];for(let i=0;i<valid.length;i+=1000){data.push(...enrich(valid.slice(i,i+1000)));await nextFrame()}$('sourceNotice').hidden=true;$('clearFilters').disabled=false;$('printReport').disabled=false;$('totalCount').textContent=`${data.length.toLocaleString('pt-BR')} na base`;setOptions('fState',uniq(COL.state),false);setOptions('fType',uniq(COL.type),false);setOptions('fMode',uniq(COL.mode),false);refreshDependent();await applyLocationDefaults();await applyFilters({loading:false})}
async function applyFilters({loading=true}={}){const token=++processingToken;if(loading)await beginProcessing('Aplicando filtros…','Atualizando indicadores, gráficos e resultados.');let st=$('fState').value,ci=$('fCity').value,hoods=new Set([...$('fNeighborhood').selectedOptions].map(o=>o.value)),types=new Set([...$('fType').selectedOptions].map(o=>o.value)),modes=new Set([...$('fMode').selectedOptions].map(o=>o.value)),fi=$('fFinance').value,min=num($('fMin').value),max=num($('fMax').value),disc=num($('fDiscount').value),area=num($('fArea').value),result=[];
 for(let i=0;i<data.length;i+=1500){if(token!==processingToken)return;result.push(...data.slice(i,i+1500).filter(r=>(!st||r[COL.state]===st)&&(!ci||r[COL.city]===ci)&&(!hoods.size||hoods.has(r[COL.hood]))&&(!types.size||types.has(r[COL.type]))&&(!modes.size||modes.has(r[COL.mode]))&&(!fi||r[COL.finance]===fi)&&(min==null||r._price>=min)&&(max==null||r._price<=max)&&(disc==null||num(r[COL.discount])>=disc)&&(area==null||propertyArea(r)>=area)));await nextFrame()}
 filtered=result;page=1;scoredCache=null;$('filteredCount').textContent=filtered.length.toLocaleString('pt-BR');await renderAll(token);if(loading)endProcessing(token)}
function refreshDependent(){let st=$('fState').value;setOptions('fCity',uniq(COL.city,data.filter(r=>!st||r[COL.state]===st)));let ci=$('fCity').value;setOptions('fNeighborhood',uniq(COL.hood,data.filter(r=>(!st||r[COL.state]===st)&&(!ci||r[COL.city]===ci))))}
async function renderAll(token=processingToken){for(const step of [renderKPIs,renderTimeline,renderStreets,renderCharts,renderNeighborhoodFinance,renderOpportunities,renderInsights,renderTable]){if(token!==processingToken)return;step();await nextFrame()}}
function renderKPIs(){let prices=filtered.map(r=>r._price),discs=filtered.map(r=>num(r[COL.discount])),sqms=filtered.map(r=>r._sqm),medianDiscount=med(discs);$('kMedianPrice').textContent=money(med(prices));$('kMedianDiscount').textContent=pct(medianDiscount);$('kMedianDiscount').classList.toggle('negative-percent',medianDiscount<0);$('kMedianSqm').textContent=money(med(sqms));let f=filtered.length?100*filtered.filter(r=>r[COL.finance]==='Sim').length/filtered.length:0;$('kFinance').textContent=pct(f)}
function auctionEvents(){let events=[];filtered.forEach(r=>{[[COL.date1,1,COL.p1],[COL.date2,2,COL.p2]].forEach(([dateCol,stage,priceCol])=>{let date=parseAuctionDate(r[dateCol]);if(date)events.push({date,stage,price:num(r[priceCol]),property:r})})});return events}
function renderTimelineDetail(bucket,today){
  let detail=$('timelineDetail');
  if(!bucket){detail.innerHTML='<div class="timeline-placeholder">Clique em uma barra para ver os leilões daquela data.</div>';return}
  let events=[...bucket.events].sort((a,b)=>a.stage-b.stage||(b.property._score||0)-(a.property._score||0)),status=bucket.start<=today&&bucket.end>=today?'Encerra hoje':bucket.end<today?'Encerrado':'Próximo',size=5,pages=Math.max(1,Math.ceil(events.length/size));timelineDetailPage=Math.min(Math.max(1,timelineDetailPage),pages);let visible=events.slice((timelineDetailPage-1)*size,timelineDetailPage*size);
  let title=bucket.weekly?`${bucket.start.toLocaleDateString('pt-BR')} a ${bucket.end.toLocaleDateString('pt-BR')}`:bucket.start.toLocaleDateString('pt-BR',{dateStyle:'long'});
  detail.innerHTML=`<div class="timeline-detail-head"><h3>${events.length} ${events.length===1?'leilão':'leilões'} · ${title}</h3><span class="timeline-status">${status}</span></div><div class="table-wrap timeline-table"><table><thead><tr><th>Etapa</th><th>Imóvel</th><th>Local</th><th>Prazo final</th><th>Avaliação</th><th>Venda</th><th>Desconto</th><th>Financ.</th><th></th></tr></thead><tbody>${visible.map(event=>{let r=event.property,url=safeUrl(r[COL.site]),discount=calculatedDiscount(r,event.price);return `<tr><td>${event.stage}º leilão</td><td>${esc(r[COL.type])}<br><small>${esc(r[COL.address])}</small></td><td>${esc(r[COL.city])}/${esc(r[COL.state])}<br><small>${esc(r[COL.hood])}</small></td><td>${event.date.toLocaleDateString('pt-BR')}</td><td>${money(num(r[COL.valuation]))}</td><td>${money(event.price)}</td><td>${pctHtml(discount)}</td><td>${esc(r[COL.finance])}</td><td>${url?`<a class="caixa-link" href="${esc(url)}" target="_blank" rel="noopener noreferrer">Ver na Caixa ↗</a>`:'—'}</td></tr>`}).join('')}</tbody></table></div>${pages>1?`<div class="pager"><button id="timelinePagePrev" ${timelineDetailPage===1?'disabled':''}>←</button><span>Página ${timelineDetailPage} de ${pages} · exibindo até 5 por página</span><button id="timelinePageNext" ${timelineDetailPage===pages?'disabled':''}>→</button></div>`:''}`;
  if(pages>1){$('timelinePagePrev').onclick=()=>{if(timelineDetailPage>1){timelineDetailPage--;renderTimelineDetail(bucket,today)}};$('timelinePageNext').onclick=()=>{if(timelineDetailPage<pages){timelineDetailPage++;renderTimelineDetail(bucket,today)}}}
}
function renderTimeline(){
  let events=auctionEvents(),today=startDay(new Date()),range=+$('timelineRange').value||90,weekly=range===365;
  $('timePast30').textContent=events.filter(e=>dayDiff(e.date,today)>=-30&&dayDiff(e.date,today)<0).length.toLocaleString('pt-BR');
  $('timeToday').textContent=events.filter(e=>dayDiff(e.date,today)===0).length.toLocaleString('pt-BR');
  $('timeNext7').textContent=events.filter(e=>dayDiff(e.date,today)>0&&dayDiff(e.date,today)<=7).length.toLocaleString('pt-BR');
  $('timeNext30').textContent=events.filter(e=>dayDiff(e.date,today)>0&&dayDiff(e.date,today)<=30).length.toLocaleString('pt-BR');
  $('timeNoDate').textContent=filtered.filter(r=>!parseAuctionDate(r[COL.date1])&&!parseAuctionDate(r[COL.date2])).length.toLocaleString('pt-BR');
  let before=range===30?10:range===90?30:91,start=addDays(today,-before+timelineOffset),bucketCount=weekly?53:range,buckets=[];
  if(weekly){let weekday=(start.getDay()+6)%7;start=addDays(start,-weekday);for(let i=0;i<bucketCount;i++){let s=addDays(start,i*7);buckets.push({key:dayKey(s),start:s,end:addDays(s,6),events:[],weekly:true})}}
  else for(let i=0;i<bucketCount;i++){let date=addDays(start,i);buckets.push({key:dayKey(date),start:date,end:date,events:[],weekly:false})}
  let last=buckets.at(-1).end;events.forEach(event=>{if(event.date<start||event.date>last)return;let index=weekly?Math.floor(dayDiff(event.date,start)/7):dayDiff(event.date,start);if(buckets[index])buckets[index].events.push(event)});
  let max=Math.max(1,...buckets.map(b=>b.events.length)),bars=$('timelineBars');bars.style.minWidth=weekly?'1320px':range===90?'2250px':'100%';
  bars.innerHTML=buckets.map((bucket,index)=>{let first=bucket.events.filter(e=>e.stage===1).length,second=bucket.events.length-first,height=Math.max(bucket.events.length?5:1,150*bucket.events.length/max),firstHeight=bucket.events.length?100*first/bucket.events.length:0,isToday=bucket.start<=today&&bucket.end>=today,label=weekly?`${bucket.start.getDate()}/${bucket.start.getMonth()+1}`:(index%Math.max(1,Math.ceil(range/18))===0||isToday?`${bucket.start.getDate()}/${bucket.start.getMonth()+1}`:'');return `<button class="time-bucket ${bucket.end<today?'past':'future'} ${isToday?'today':''} ${selectedTimelineKey===bucket.key?'selected':''}" data-time-key="${bucket.key}" title="${bucket.events.length} eventos · ${bucket.start.toLocaleDateString('pt-BR')}"><span class="time-count">${bucket.events.length||''}</span><span class="time-volume" style="height:${height}px"><i class="time-first" style="height:${firstHeight}%"></i><i class="time-second" style="height:${100-firstHeight}%"></i></span><span class="time-label">${label}</span></button>`}).join('');
  bars.querySelectorAll('[data-time-key]').forEach(button=>button.onclick=()=>{selectedTimelineKey=button.dataset.timeKey;timelineDetailPage=1;renderTimelineDetail(buckets.find(b=>b.key===selectedTimelineKey),today);bars.querySelectorAll('.selected').forEach(el=>el.classList.remove('selected'));button.classList.add('selected')});
  let selected=buckets.find(b=>b.key===selectedTimelineKey);renderTimelineDetail(selected,today);
  if(timelineOffset===0)setTimeout(()=>bars.querySelector('.today')?.scrollIntoView({behavior:'smooth',inline:'center',block:'nearest'}),0)
}
function renderStreets(){let c={};filtered.forEach(r=>c[r._street]=(c[r._street]||0)+1);let arr=Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,10),mx=arr[0]?.[1]||1;$('streetRanking').innerHTML=arr.map(([n,v])=>`<div class="rank-row"><div><div class="rank-name" title="${esc(n)}">${esc(n)}</div><div class="bar"><i style="width:${100*v/mx}%"></i></div></div><b>${v}</b></div>`).join('')||'<small>Sem dados.</small>'}
function chart(id,type,labels,values,label){if(typeof Chart==='undefined')return;if(charts[id])charts[id].destroy();charts[id]=new Chart($(id),{type,data:{labels,datasets:[{label,data:values,borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:type==='doughnut'}},scales:type==='doughnut'?{}:{x:{grid:{display:false},ticks:{maxTicksLimit:8}},y:{beginAtZero:true,grid:{color:'#eef1f3'}}}}})}
function renderCharts(){let ps=filtered.map(r=>r._price).filter(Number.isFinite);if(ps.length){let lo=quant(ps,.01),hi=quant(ps,.99),bins=12,w=(hi-lo)/bins||1,counts=Array(bins).fill(0);ps.forEach(p=>{if(p>=lo&&p<=hi)counts[Math.min(bins-1,Math.floor((p-lo)/w))]++});chart('priceChart','bar',counts.map((_,i)=>money(lo+i*w)),counts,'Imóveis')}else chart('priceChart','bar',[],[],'Imóveis');
 let by={};filtered.forEach(r=>{let t=r[COL.type]||'Outros';(by[t]??=[]).push(num(r[COL.discount]))});let d=Object.entries(by).map(([k,v])=>[k,med(v)]).filter(x=>x[1]!=null).sort((a,b)=>b[1]-a[1]).slice(0,8);chart('discountChart','bar',d.map(x=>x[0]),d.map(x=>x[1]),'Desconto mediano (%)');let mc={};filtered.forEach(r=>mc[r[COL.mode]||'N/I']=(mc[r[COL.mode]||'N/I']||0)+1);chart('modeChart','doughnut',Object.keys(mc),Object.values(mc),'Imóveis')}
function renderNeighborhoodFinance(){
  let state=$('fState').value,empties=[...document.querySelectorAll('.neighborhood-empty')],wraps=[...document.querySelectorAll('.neighborhood-chart-wrap')],badges=[...document.querySelectorAll('.neighborhood-state-badge')];
  ['neighborhood','financeNeighborhood'].forEach(key=>{if(charts[key]){charts[key].destroy();delete charts[key]}});
  if(!state){badges.forEach(b=>b.textContent='Selecione um estado');empties.forEach(e=>e.hidden=false);wraps.forEach(w=>w.hidden=true);return}
  badges.forEach(b=>b.textContent=state);
  let groups={};
  filtered.forEach(r=>{let hood=String(r[COL.hood]||'Não informado').trim()||'Não informado',finance=normalize(r[COL.finance]);groups[hood]??={yes:0,no:0};if(['sim','s','yes'].includes(finance))groups[hood].yes++;else groups[hood].no++});
  let allRows=Object.entries(groups).map(([name,value])=>({name,...value,total:value.yes+value.no})),rows=[...allRows].sort((a,b)=>b.total-a.total).slice(0,12),financeRows=[...allRows].filter(r=>r.yes>0).sort((a,b)=>b.yes-a.yes||b.total-a.total).slice(0,12);
  if(!rows.length||typeof Chart==='undefined'){empties.forEach(e=>{e.textContent=typeof Chart==='undefined'?'O gráfico está indisponível porque a biblioteca Chart.js não foi carregada.':'Nenhum bairro encontrado para os filtros atuais.';e.hidden=false});wraps.forEach(w=>w.hidden=true);return}
  empties[0].hidden=true;wraps[0].hidden=false;
  charts.neighborhood=new Chart($('neighborhoodChart'),{type:'bar',data:{labels:rows.map(r=>r.name),datasets:[{label:'Financiáveis',data:rows.map(r=>r.yes),backgroundColor:'#126e5a'},{label:'Não financiáveis',data:rows.map(r=>r.no),backgroundColor:'#d8a13b'}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top'},tooltip:{callbacks:{footer:items=>`Total: ${items.reduce((sum,item)=>sum+item.parsed.x,0)}`}}},scales:{x:{stacked:true,beginAtZero:true,ticks:{precision:0},grid:{color:'#eef1f3'}},y:{stacked:true,grid:{display:false}}}}})
  if(financeRows.length){
    empties[1].hidden=true;wraps[1].hidden=false;
    charts.financeNeighborhood=new Chart($('financeNeighborhoodChart'),{type:'bar',data:{labels:financeRows.map(r=>r.name),datasets:[{label:'Imóveis financiáveis',data:financeRows.map(r=>r.yes),backgroundColor:'#126e5a'}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{precision:0},grid:{color:'#eef1f3'}},y:{grid:{display:false}}}}})
  }else{empties[1].textContent='Nenhum imóvel financiável encontrado para os filtros atuais.';empties[1].hidden=false;wraps[1].hidden=true}
}
function scored(){if(scoredCache)return scoredCache;let sq=filtered.map(r=>r._sqm).filter(Number.isFinite),sv=filtered.map(r=>Math.max(0,(num(r[COL.valuation])||0)-(r._price||0))).filter(Number.isFinite),qSq=quant(sq,.9)||1,qSav=quant(sv,.9)||1;scoredCache=filtered.map(r=>{let d=Math.max(0,Math.min(100,num(r[COL.discount])||0))/100;let sqmScore=r._sqm!=null?1-Math.min(1,r._sqm/qSq):0;let saving=Math.min(1,Math.max(0,(num(r[COL.valuation])||0)-(r._price||0))/qSav);let fin=r[COL.finance]==='Sim'?1:0;r._score=100*(.42*d+.30*sqmScore+.18*saving+.10*fin);return r}).sort((a,b)=>b._score-a._score);return scoredCache}
function renderOpportunities(){let arr=scored().slice(0,12);$('opportunityTable').innerHTML=arr.map(r=>{let url=safeUrl(r[COL.site]);return `<tr><td class="score">${r._score.toFixed(0)}</td><td>${esc(r[COL.type])}</td><td>${esc(r[COL.city])}/${esc(r[COL.state])}<br><small>${esc(r[COL.hood])}</small></td><td>${money(r._price)}</td><td>${pctHtml(r[COL.discount])}</td><td>${money(r._sqm)}</td><td>${url?`<a class="caixa-link" href="${esc(url)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a>`:'—'}</td></tr>`}).join('')}
function renderInsights(){let n=filtered.length;if(!n){$('insights').innerHTML='<p>Sem dados para os filtros atuais.</p>';return}let below=med(filtered.map(r=>r._sqm)),cheap=filtered.filter(r=>r._sqm!=null&&r._sqm<below&&num(r[COL.discount])>=med(filtered.map(x=>num(x[COL.discount])))).length;let cond=filtered.filter(r=>String(r[COL.cond]||'').toLowerCase().includes('caixa')).length;let high=filtered.filter(r=>(num(r[COL.discount])||0)>=50).length;let cities=new Set(filtered.map(r=>r[COL.city]).filter(Boolean)).size;let items=[[` ${pct(100*high/n)} com desconto ≥ 50%`,`${high.toLocaleString('pt-BR')} imóveis estão na faixa de desconto mais agressiva.`],[`${cheap.toLocaleString('pt-BR')} oportunidades preço/m² + desconto`,`Abaixo da mediana de R$/m² e simultaneamente acima da mediana de desconto da seleção.`],[`${pct(100*cond/n)} com vantagem condominial`,`Registros em que a descrição indica alguma responsabilidade da Caixa sobre condomínio.`],[`${cities.toLocaleString('pt-BR')} cidades cobertas`,`Ajuda a medir diversificação geográfica da seleção atual.`]];$('insights').innerHTML=items.map(x=>`<div class="insight"><strong>${x[0]}</strong><p>${x[1]}</p></div>`).join('')}
function renderTable(){let q=$('search').value.toLowerCase().trim(),arr=scored().filter(r=>!q||[r[COL.id],r[COL.address],r[COL.hood],r[COL.city]].some(v=>String(v||'').toLowerCase().includes(q))),size=15,pages=Math.max(1,Math.ceil(arr.length/size));page=Math.min(page,pages);let slice=arr.slice((page-1)*size,page*size);$('propertyTable').innerHTML=slice.map(r=>`<tr><td>${esc(r[COL.id])}</td><td>${esc(r[COL.type])}</td><td>${esc(r[COL.address])}</td><td>${money(r._price)}</td><td>${money(num(r[COL.valuation]))}</td><td>${pctHtml(r[COL.discount])}</td><td>${propertyArea(r)??'—'} m²</td><td>${money(r._sqm)}</td><td>${esc(r[COL.finance])}</td><td><button class="linkbtn" onclick="showProperty('${String(r[COL.id]).replace(/'/g,'')}')">Detalhes</button></td></tr>`).join('');$('pageInfo').textContent=`Página ${page} de ${pages} · ${arr.length.toLocaleString('pt-BR')} resultados`;window._tableRows=arr}
function showProperty(id){let r=data.find(x=>String(x[COL.id])===String(id));if(!r)return;let links=[[r[COL.site],'Caixa'],[r[COL.rent],'Zap · aluguel'],[r[COL.buy],'Zap · compra']].filter(x=>x[0]);let areaLabel=normalize(r[COL.type]).includes('terreno')?'Área do terreno':'Área construída';$('modalContent').innerHTML=`<h2>${esc(r[COL.type])} · ${esc(r[COL.city])}/${esc(r[COL.state])}</h2><p>${esc(r[COL.address])}</p><div class="modal-grid"><div><span>Preço de venda</span><b>${money(r._price)}</b></div><div><span>Avaliação</span><b>${money(num(r[COL.valuation]))}</b></div><div><span>Desconto</span><b>${pctHtml(r[COL.discount])}</b></div><div><span>R$/m²</span><b>${money(r._sqm)}</b></div><div><span>${areaLabel}</span><b>${propertyArea(r)??'—'} m²</b></div><div><span>Financiamento</span><b>${esc(r[COL.finance])}</b></div><div><span>Modalidade</span><b>${esc(r[COL.mode])}</b></div><div><span>Score</span><b>${(r._score||0).toFixed(0)}/100</b></div></div><div class="modal-links">${links.map(([u,t])=>`<a href="${esc(u)}" target="_blank" rel="noopener">${t} ↗</a>`).join('')}</div>`;$('propertyModal').showModal()}
window.showProperty=showProperty;
let map=null;
if($('map')&&typeof L!=='undefined'){
  try{map=L.map('map').setView([-14.2,-51.9],4);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap contributors'}).addTo(map)}
  catch(err){console.warn('Mapa indisponível:',err)}
}else if($('mapNote')){
  $('mapNote').textContent='O mapa está indisponível, mas o Excel e o restante do dashboard continuam funcionando.';
}
const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms));
async function nominatimSearch(query){
  let url=`https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=br&q=${encodeURIComponent(query)}`;
  try{
    let response=await fetch(url,{headers:{'Accept-Language':'pt-BR'}});
    if(!response.ok)throw new Error(`Geocodificação respondeu ${response.status}`);
    let result=await response.json();
    return result[0]?[+result[0].lat,+result[0].lon]:null;
  }finally{await wait(1100)}
}
async function geocode(){
  let btn=$('geocodeBtn');
  if(!map||typeof L==='undefined')return;
  btn.disabled=true;
  let sample=scored().filter(r=>String(r[COL.address]||'').trim()).slice(0,40),pts=[],found=0,failures=0;
  let cache=JSON.parse(localStorage.getItem('geoPinCacheV1')||'{}');
  if(heatLayer){map.removeLayer(heatLayer);heatLayer=null}
  if(markerLayer)map.removeLayer(markerLayer);
  markerLayer=L.layerGroup().addTo(map);
  try{
    for(let i=0;i<sample.length;i++){
      let r=sample[i];
      let parts=[r[COL.address],r[COL.hood],r[COL.city],r[COL.state],'Brasil'].map(v=>String(v||'').trim()).filter(Boolean);
      let fullQuery=parts.join(', '),fullKey=normalize(fullQuery),g=cache[fullKey];
      $('mapNote').textContent=`Adicionando pins ${i+1}/${sample.length}… ${found} endereços encontrados.`;
      try{
        if(!g&&fullQuery)g=await nominatimSearch(fullQuery);
        if(g){cache[fullKey]=g;found++}else failures++;
      }catch(error){failures++;console.warn('Falha ao geocodificar:',fullQuery,error)}
      if(g){
        pts.push([g[0],g[1]]);
        let marker=L.marker(g).addTo(markerLayer);
        marker.bindTooltip(`${esc(r[COL.type])} · ${esc(r[COL.hood])} · ${money(r._price)}`);
        marker.on('click',()=>showProperty(String(r[COL.id])));
      }
      localStorage.setItem('geoPinCacheV1',JSON.stringify(cache));
    }
    if(pts.length){
      map.fitBounds(pts,{padding:[20,20]});
      $('mapNote').textContent=`${found} endereços encontrados e ${failures} não localizados. Clique em um pin para abrir os detalhes do imóvel.`;
    }else{
      $('mapNote').textContent=location.protocol==='file:'?'Nenhum endereço foi localizado. Abra o dashboard por um servidor local (por exemplo, Live Server), pois o navegador pode bloquear consultas feitas por file://.':'Nenhum endereço foi localizado. Confira se Cidade e Estado estão preenchidos na planilha.';
    }
  }finally{btn.disabled=false}
}
if($('geocodeBtn'))$('geocodeBtn').onclick=geocode;
function activeFilterSummary(){
  let items=[],hoods=[...$('fNeighborhood').selectedOptions].map(o=>o.value),types=[...$('fType').selectedOptions].map(o=>o.value),modes=[...$('fMode').selectedOptions].map(o=>o.value);
  let add=(label,value)=>{if(value!=null&&value!=='')items.push(`${label}: ${value}`)};
  add('Estado',$('fState').value);add('Cidade',$('fCity').value);if(hoods.length)add('Bairros',hoods.join(', '));if(types.length)add('Tipos',types.join(', '));if(modes.length)add('Modos de venda',modes.join(', '));add('Financiamento',$('fFinance').value);
  let min=num($('fMin').value),max=num($('fMax').value),discount=num($('fDiscount').value),area=num($('fArea').value);
  if(min!=null)add('Valor mínimo',money(min));if(max!=null)add('Valor máximo',money(max));if(discount!=null)add('Desconto mínimo',pct(discount));if(area!=null)add('Área mínima',`${area} m²`);add('Busca',$('search').value.trim());
  return items.length?items.join(' · '):'Todos os imóveis, sem filtros adicionais';
}
function preparePrintReport(){
  $('printDate').textContent=`Extraído em ${new Intl.DateTimeFormat('pt-BR',{dateStyle:'long',timeStyle:'short'}).format(new Date())}`;
  $('printFilters').textContent=activeFilterSummary();
  let sections=[...document.querySelectorAll('main > section')];
  sections.forEach(section=>{section.classList.remove('report-page','print-last');if(!section.classList.contains('report-exclude'))section.classList.add('report-page')});
  sections.filter(section=>section.classList.contains('report-page')).at(-1)?.classList.add('print-last');
  Object.values(charts).forEach(instance=>instance?.resize());
}
$('printReport').onclick=()=>{preparePrintReport();setTimeout(()=>window.print(),180)};
window.addEventListener('beforeprint',preparePrintReport);
window.addEventListener('afterprint',()=>setTimeout(()=>Object.values(charts).forEach(instance=>instance?.resize()),100));
$('timelineRange').onchange=()=>{timelineOffset=0;selectedTimelineKey=null;timelineDetailPage=1;renderTimeline()};
$('timelineToday').onclick=()=>{timelineOffset=0;selectedTimelineKey=null;timelineDetailPage=1;renderTimeline()};
$('timelinePrev').onclick=()=>{let range=+$('timelineRange').value;timelineOffset-=range===365?182:Math.round(range*.7);selectedTimelineKey=null;timelineDetailPage=1;renderTimeline()};
$('timelineNext').onclick=()=>{let range=+$('timelineRange').value;timelineOffset+=range===365?182:Math.round(range*.7);selectedTimelineKey=null;timelineDetailPage=1;renderTimeline()};
['fState','fCity','fNeighborhood','fType','fMode','fFinance','fMin','fMax','fDiscount','fArea'].forEach(id=>$(id).addEventListener(['fMin','fMax','fDiscount','fArea'].includes(id)?'input':'change',()=>{if(id==='fState'||id==='fCity')locationFiltersTouched=true;clearTimeout(filterTimer);setGlobalLoading(true,'Aplicando filtros…','Atualizando indicadores, gráficos e resultados.');filterTimer=setTimeout(()=>{if(id==='fState'||id==='fCity')refreshDependent();applyFilters()},180)}));
Object.values(multiFilters).forEach(config=>{$(config.toggle).onclick=()=>{let menu=$(config.menu),open=menu.hidden;Object.values(multiFilters).forEach(other=>{$(other.menu).hidden=true;$(other.toggle).setAttribute('aria-expanded','false')});menu.hidden=!open;$(config.toggle).setAttribute('aria-expanded',String(open))}});
document.addEventListener('click',event=>{if(event.target.closest('.multi-dropdown'))return;Object.values(multiFilters).forEach(config=>{$(config.menu).hidden=true;$(config.toggle).setAttribute('aria-expanded','false')})});
$('clearFilters').onclick=()=>{locationFiltersTouched=true;['fState','fCity','fNeighborhood','fType','fMode','fFinance','fMin','fMax','fDiscount','fArea'].forEach(id=>{let el=$(id);if(el.multiple)[...el.options].forEach(o=>o.selected=false);else el.value=''});refreshDependent();applyFilters()};$('search').oninput=()=>{clearTimeout(filterTimer);setGlobalLoading(true,'Aplicando busca…','Atualizando a lista de imóveis.');filterTimer=setTimeout(async()=>{page=1;await nextFrame();renderTable();endProcessing()},180)};$('prevPage').onclick=()=>{if(page>1){page--;renderTable()}};$('nextPage').onclick=()=>{page++;renderTable()};
$('fileInput').onchange=e=>{
  const f=e.target.files[0];
  if(!f)return;
  fileImportStarted=true;
  if(!locationFiltersTouched)locationDefaultsApplied=false;
  if(!/\.(xlsx?|csv)$/i.test(f.name)){setUploadState('error','Formato não suportado','Selecione um arquivo .xlsx, .xls ou .csv.');$('dataStatus').textContent='Arquivo inválido';e.target.value='';return}
  const isCsv=/\.csv$/i.test(f.name);
  const size=f.size<1048576?`${(f.size/1024).toFixed(0)} KB`:`${(f.size/1048576).toFixed(1)} MB`;
  setGlobalLoading(true,'Carregando arquivo…',`${f.name} · ${size}`);
  setUploadState('loading',`Lendo ${f.name}…`,`${size} · aguarde enquanto a planilha é processada.`);
  $('dataStatus').textContent=`Lendo ${f.name}…`;
  const rd=new FileReader();
  rd.onload=async ev=>{
    try{
      if(typeof XLSX==='undefined')throw new Error('XLSX_UNAVAILABLE');
      await beginProcessing('Processando planilha…',`${f.name} · ${size}`);
      $('dataStatus').textContent='Processando planilha…';
      setUploadState('loading','Processando planilha…',`${f.name} · ${size}`);
      const parsed=await parseExcelInWorker(ev.target.result),sheetName=parsed.sheetName;
      if(!parsed.rows.some(r=>r[COL.id]!=null&&String(r[COL.id]).trim()!==''))throw new Error('NO_VALID_ROWS');
      $('globalLoadingDetail').textContent='Preparando indicadores e gráficos…';
      await init(parsed.rows);
      refreshDependent();
      const ignored=Math.max(0,parsed.rawCount-data.length);
      const summary=`${data.length.toLocaleString('pt-BR')} imóveis válidos${ignored?` · ${ignored.toLocaleString('pt-BR')} linhas ignoradas`:''}`;
      const metadata={fileName:f.name,fileSize:f.size,importedAt:new Date().toISOString(),rowCount:data.length,sheetName,headerRow:parsed.headerRow};
      let persisted=true;
      $('globalLoadingDetail').textContent='Salvando a base neste navegador…';
      try{await saveDataset(parsed.rows,metadata);$('removeSavedData').hidden=false}catch(storageError){persisted=false;console.warn('Não foi possível salvar a base no navegador:',storageError);$('removeSavedData').hidden=true}
      $('dataStatus').textContent=`Base carregada · ${data.length.toLocaleString('pt-BR')} imóveis`;
      setUploadState('success',`${f.name} carregado com sucesso`,`${summary}${isCsv?'':` · aba “${sheetName}”`}${persisted?' · salva neste navegador':' · armazenamento local indisponível'}`);
      endProcessing();
    }catch(err){
      console.error(err);
      let detail=err.message==='XLSX_UNAVAILABLE'?'A biblioteca de leitura de planilhas não carregou. Verifique a internet e recarregue a página.':err.message==='HEADER_NOT_FOUND'?'Não foi possível localizar um cabeçalho compatível nas primeiras 15 linhas do arquivo. O cabeçalho pode estar na linha 2 normalmente.':err.message==='NO_VALID_ROWS'?'O cabeçalho foi reconhecido, mas nenhuma linha possui um ID de imóvel válido.':err.message.startsWith('MISSING_COLUMNS:')?`O cabeçalho foi encontrado, mas estas colunas não foram reconhecidas: ${err.message.slice('MISSING_COLUMNS:'.length)}.`:err.message.startsWith('WORKER_UNAVAILABLE:')?'O processamento em segundo plano não pôde ser iniciado. Recarregue a página e confirme que a biblioteca de planilhas foi carregada.':`Erro técnico: ${err.message||'não identificado'}.`;
      $('dataStatus').textContent='Não foi possível ler o arquivo';
      setUploadState('error','Não foi possível processar o arquivo',detail);
      endProcessing();
    }
  };
  rd.onerror=()=>{$('dataStatus').textContent='Erro ao ler o arquivo';setUploadState('error','Erro ao ler o arquivo','O navegador não conseguiu acessar o conteúdo selecionado. Tente selecionar o arquivo novamente.');endProcessing()};
  rd.readAsArrayBuffer(f);
};

async function restoreSavedDataset(){
  try{
    const saved=await readSavedDataset();
    if(!saved||fileImportStarted)return;
    if(saved.schemaVersion!==DATA_SCHEMA_VERSION){await deleteSavedDataset();return}
    await beginProcessing('Restaurando última base…','Carregando os dados salvos neste navegador.');
    if(fileImportStarted)return;
    await init(saved.rows);
    refreshDependent();
    $('dataStatus').textContent=`Base restaurada · ${data.length.toLocaleString('pt-BR')} imóveis`;
    setUploadState('success',saved.metadata.fileName||'Última base','Base restaurada automaticamente deste navegador.');
    showSavedMetadata(saved.metadata);
  }catch(error){console.warn('Não foi possível restaurar a base salva:',error)}finally{if(!fileImportStarted)endProcessing()}
}
$('removeSavedData').onclick=async()=>{
  if(!confirm('Remover a base salva deste navegador? Os dados continuarão disponíveis até esta página ser fechada ou recarregada.'))return;
  try{await beginProcessing('Removendo base salva…','Limpando os dados persistidos neste navegador.');await deleteSavedDataset();$('removeSavedData').hidden=true;$('uploadDetails').textContent='Base salva removida. Os dados desta sessão continuam disponíveis.'}catch(error){console.error(error);$('uploadDetails').textContent='Não foi possível remover a base salva. Tente novamente.'}finally{endProcessing()}
};

// Se existir uma base compatível no IndexedDB, ela é restaurada automaticamente.
$('totalCount').textContent='Nenhuma base carregada';
$('filteredCount').textContent='—';
if($('geocodeBtn'))$('geocodeBtn').disabled=true;
$('clearFilters').disabled=true;
restoreSavedDataset();
startLocationDetection();
