/* Gazetki → Excel: browser-only flyer reader. */
(() => {
  'use strict';
  const $ = (s) => document.querySelector(s);
  const input = $('#fileInput'), drop = $('#dropzone'), processBtn = $('#processBtn');
  const body = $('#offersBody'), filesBox = $('#fileList'), exportBtn = $('#exportBtn');
  let files = [], rows = [], worker = null, toastTimer;
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  function toast(message) { const el=$('#toast'); el.textContent=message; el.classList.add('show'); clearTimeout(toastTimer); toastTimer=setTimeout(()=>el.classList.remove('show'),2600); }
  function esc(s) { return String(s ?? '').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function addFiles(list) {
    const valid=[...list].filter(f=>/^(application\/pdf|image\/(png|jpeg|webp))$/.test(f.type));
    files=[...files,...valid].filter((f,i,a)=>a.findIndex(x=>x.name===f.name&&x.size===f.size)===i);
    renderFiles(); processBtn.disabled=!files.length;
    if(list.length!==valid.length) toast('Pominięto pliki w nieobsługiwanym formacie.');
  }
  function renderFiles(){ filesBox.innerHTML=files.map((f,i)=>`<span class="file-chip">${f.type==='application/pdf'?'PDF':'IMG'} · ${esc(f.name)} <button type="button" data-remove-file="${i}" aria-label="Usuń plik">×</button></span>`).join(''); }
  input.addEventListener('change',e=>addFiles(e.target.files));
  filesBox.addEventListener('click',e=>{const b=e.target.closest('[data-remove-file]');if(!b)return;files.splice(+b.dataset.removeFile,1);renderFiles();processBtn.disabled=!files.length;});
  ['dragenter','dragover'].forEach(type=>drop.addEventListener(type,e=>{e.preventDefault();drop.classList.add('dragover');}));
  ['dragleave','drop'].forEach(type=>drop.addEventListener(type,e=>{e.preventDefault();drop.classList.remove('dragover');}));
  drop.addEventListener('drop',e=>addFiles(e.dataTransfer.files));

  function parseAmount(text){ const m=text.match(/(?:^|\s)(\d+(?:[.,]\d+)?)\s*(kg|g|l|ml|szt\.?|sztuk|sztuki)\b/i); if(!m)return ''; let n=parseFloat(m[1].replace(',','.')); const u=m[2].toLowerCase().replace('.',''); if(u==='g')n/=1000; if(u==='ml')n/=1000; return Number.isFinite(n)&&n>0?n:''; }
  function parsePrice(token){ const m=token.match(/(\d{1,3}[,.]\d{2})\s*(?:zł|zl|pln)?/i); return m?parseFloat(m[1].replace(',','.')):null; }
  function linesToOffers(text, store, page, validTo){
    const lines=text.split(/\n+/).map(s=>s.replace(/\s+/g,' ').trim()).filter(Boolean), out=[];
    for(let i=0;i<lines.length;i++){
      const priceLine=lines[i].replace(/\b(?:od\s*)?\d{1,2}[.,]\d{2}\s*(?:do|[-–])\s*\d{1,2}[.,]\d{2}\b/gi,'');
      const priceMatches=[...priceLine.matchAll(/\d{1,3}[,.]\d{2}\s*(?:zł|zl|pln)?/gi)].filter(m=>!/^\s*(?:zł|zl|zt|pln)?\s*\/\s*(?:kg|l|100\s*g|100\s*ml)\b/i.test(priceLine.slice(m.index+m[0].length)));
      if(!priceMatches.length) continue;
      const promo=parsePrice(priceMatches[0][0]); if(promo===null||promo<=0||promo>9999)continue;
      const before=priceLine.slice(0,priceMatches[0].index).trim();
      const prev=lines.slice(Math.max(0,i-2),i).filter(x=>!/(gazetka|oferta|promocja|strona|ważna|od\s+\d)/i.test(x)&&!(/\d{1,3}[,.]\d{2}/.test(x)));
      let name=[...prev,before].filter(Boolean).join(' ').replace(/\b\d{1,2}%\s*(?:taniej|rabat)?\b/gi,'').trim();
      name=name.replace(/\b(?:zł|zl|pln)\b/gi,'').replace(/\s+/g,' ');
      if(name.length<3) name=`Oferta z gazetki — str. ${page}`;
      if(name.length>115)name=name.slice(-115).trim();
      const rest=priceLine.slice(priceMatches[0].index+priceMatches[0][0].length);
      const price2=rest.match(/\d{1,3}[,.]\d{2}\s*(?:zł|zl|pln)?/i);
      const amount=parseAmount(lines.slice(Math.max(0,i-1),Math.min(lines.length,i+2)).join(' '));
      const condition=(lines[i].match(/(?:\d\s*\+\s*\d|przy zakupie.{0,35}|z aplikacją.{0,35}|z kartą.{0,35}|drugi.{0,30}|limit.{0,20})/i)||[])[0]||'';
      const duplicate=out.some(o=>o.product.toLowerCase()===name.toLowerCase()&&o.promo===promo);
      if(!duplicate)out.push({store,product:name,category:'',promo,regular:price2?parsePrice(price2[0]):'',amount,conditions:condition.trim(),validTo,page,raw:lines[i],confidence:name.startsWith('Oferta z')?'Niski':'Do sprawdzenia'});
    }
    return out;
  }
  async function extractPdf(file, store, validTo, onPage){
    const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise, found=[];
    for(let p=1;p<=pdf.numPages;p++){
      onPage(p,pdf.numPages,file.name);
      const page=await pdf.getPage(p), content=await page.getTextContent();
      const grouped=new Map();
      for(const item of content.items){
        const y=Math.round(item.transform?.[5]||0), key=String(y);
        if(!grouped.has(key))grouped.set(key,[]);
        grouped.get(key).push(item);
      }
      let text=[...grouped.entries()].sort((a,b)=>Number(b[0])-Number(a[0])).map(([,items])=>items.sort((a,b)=>(a.transform?.[4]||0)-(b.transform?.[4]||0)).map(x=>x.str).join(' ')).join('\n');
      if(text.trim().length<45){ const viewport=page.getViewport({scale:1.8}), canvas=document.createElement('canvas'), ctx=canvas.getContext('2d'); canvas.width=viewport.width;canvas.height=viewport.height;await page.render({canvasContext:ctx,viewport}).promise;text=await recognize(canvas); }
      found.push(...linesToOffers(text,store,p,validTo));
    }
    return found;
  }
  async function recognize(image){
    if(!window.Tesseract)throw new Error('Nie udało się załadować modułu OCR. Sprawdź połączenie z internetem.');
    if(!worker) worker=await Tesseract.createWorker('pol+eng',1,{logger:m=>{if(m.status==='recognizing text')setProgress(Math.round(m.progress*100),`OCR · ${Math.round(m.progress*100)}%`);}});
    const result=await worker.recognize(image);return result.data.text;
  }
  function setProgress(p,label){$('#progressBar').style.width=`${p}%`;$('#progressPct').textContent=`${p}%`;$('#progressLabel').textContent=label;}
  processBtn.addEventListener('click',async()=>{
    if(!files.length)return;
    processBtn.disabled=true;$('#progressWrap').hidden=false;
    const store=$('#storeInput').value.trim(), validTo=$('#validToInput').value;let total=0,done=0;
    try{
      for(const file of files){
        const before=rows.length;
        if(file.type==='application/pdf') rows.push(...await extractPdf(file,store,validTo,(p,n,name)=>setProgress(Math.round((done+(p-1)/n)/files.length*100),`${name} · strona ${p}/${n}`)));
        else {setProgress(Math.round(done/files.length*100),`Rozpoznawanie ${file.name}`);const text=await recognize(file);rows.push(...linesToOffers(text,store,1,validTo));}
        total+=rows.length-before;done++;
      }
      renderRows();setProgress(100,'Gotowe');$('#statusText').textContent=`Odczytano ${total} ${total===1?'ofertę':'ofert'} · sprawdź zaznaczone wartości.`;toast(total?`Odczytano ${total} ofert. Popraw je i wyeksportuj.`:'Nie znaleziono ofert z cenami. Spróbuj wyraźniejszego PDF lub zdjęcia.');
    }catch(err){console.error(err);toast(err.message||'Nie udało się przetworzyć pliku.');$('#statusText').textContent='Wystąpił błąd podczas odczytu.';}
    finally{processBtn.disabled=!files.length;}
  });
  function unitPrice(row){const n=Number(row.amount),p=Number(row.promo);return n>0&&p>0?(p/n).toFixed(2):'';}
  function renderRows(){
    $('#rowCount').textContent=`${rows.length} ${rows.length===1?'oferta':rows.length>1&&rows.length<5?'oferty':'ofert'}`;exportBtn.disabled=!rows.length;
    if(!rows.length){body.innerHTML='<tr class="empty-row"><td colspan="11"><div class="empty-state"><span>↳</span><strong>Tu pojawią się oferty</strong><small>Dodaj gazetkę, aby rozpocząć analizę.</small></div></td></tr>';return;}
    body.innerHTML=rows.map((r,i)=>`<tr data-row="${i}">
      <td><input class="cell-store" data-key="store" value="${esc(r.store)}" aria-label="Sklep"></td>
      <td><input class="cell-product" data-key="product" value="${esc(r.product)}" aria-label="Produkt"></td>
      <td><select data-key="category" aria-label="Kategoria">${['','Nabiał','Mięso i ryby','Owoce i warzywa','Pieczywo','Napoje','Produkty suche','Mrożonki','Chemia i dom','Inne'].map(c=>`<option ${r.category===c?'selected':''}>${esc(c)}</option>`).join('')}</select></td>
      <td><input class="cell-price" data-key="promo" inputmode="decimal" value="${esc(r.promo)}" aria-label="Cena promocyjna"></td>
      <td><input class="cell-price" data-key="regular" inputmode="decimal" value="${esc(r.regular)}" aria-label="Cena regularna"></td>
      <td><input class="cell-small" data-key="amount" inputmode="decimal" value="${esc(r.amount)}" aria-label="Masa lub pojemność w kilogramach lub litrach" placeholder="kg / l"></td>
      <td><span class="unit-price">${unitPrice(r)?`${unitPrice(r)} zł`:'—'}</span></td>
      <td><input class="cell-condition" data-key="conditions" value="${esc(r.conditions)}" aria-label="Warunki promocji"></td>
      <td><input class="cell-small" data-key="validTo" type="date" value="${esc(r.validTo)}" aria-label="Ważna do"></td>
      <td><input class="cell-small" data-key="page" inputmode="numeric" value="${esc(r.page)}" aria-label="Strona gazetki"></td>
      <td><button class="delete-row" data-delete="${i}" aria-label="Usuń ofertę">×</button></td></tr>`).join('');
  }
  body.addEventListener('input',e=>{const inp=e.target.closest('[data-key]');if(!inp)return;const tr=inp.closest('[data-row]'),r=rows[+tr.dataset.row];r[inp.dataset.key]=inp.value;if(['promo','amount'].includes(inp.dataset.key))tr.querySelector('.unit-price').textContent=unitPrice(r)?`${unitPrice(r)} zł`:'—';});
  body.addEventListener('change',e=>{const inp=e.target.closest('[data-key]');if(inp)rows[+inp.closest('[data-row]').dataset.row][inp.dataset.key]=inp.value;});
  body.addEventListener('click',e=>{const b=e.target.closest('[data-delete]');if(!b)return;rows.splice(+b.dataset.delete,1);renderRows();});
  $('#addRowBtn').addEventListener('click',()=>{rows.push({store:$('#storeInput').value.trim(),product:'',category:'',promo:'',regular:'',amount:'',conditions:'',validTo:$('#validToInput').value,page:'',confidence:'Ręcznie'});renderRows();body.querySelector('tr:last-child [data-key="product"]').focus();});
  exportBtn.addEventListener('click',()=>{
    if(!window.XLSX){toast('Nie udało się załadować eksportu Excel.');return;}
    const data=rows.map(r=>({'Sklep':r.store,'Produkt / opis':r.product,'Kategoria':r.category,'Cena promocyjna (zł)':num(r.promo),'Cena regularna (zł)':num(r.regular),'Ilość (kg/l)':num(r.amount),'Cena za kg/l (zł)':unitPrice(r)?Number(unitPrice(r)):null,'Warunki promocji':r.conditions,'Ważna do':r.validTo,'Strona gazetki':num(r.page),'Uwagi do weryfikacji':r.confidence,'Tekst źródłowy':r.raw||''}));
    const ws=XLSX.utils.json_to_sheet(data);ws['!cols']=[{wch:18},{wch:38},{wch:21},{wch:18},{wch:18},{wch:14},{wch:17},{wch:28},{wch:14},{wch:13},{wch:20},{wch:55}];ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:data.length,c:11}})};
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Oferty');const summary=[['Gazetki → Excel'],['Liczba ofert',rows.length],['Data eksportu',new Date().toLocaleDateString('pl-PL')],['Uwaga','Wyniki OCR należy zweryfikować z gazetką źródłową.'],['Cena za kg/l','Wyliczona na podstawie ceny promocyjnej i ilości podanej w kg lub l.']];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),'Informacje');
    const name=($('#storeInput').value.trim()||'gazetki').toLowerCase().replace(/[^a-z0-9ąćęłńóśźż-]+/gi,'-').replace(/^-|-$/g,'');XLSX.writeFile(wb,`${name||'gazetki'}-oferty-${new Date().toISOString().slice(0,10)}.xlsx`);toast('Plik Excel został pobrany.');
  });
  function num(v){const n=parseFloat(String(v??'').replace(',','.'));return Number.isFinite(n)?n:null;}
})();
