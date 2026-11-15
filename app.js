/* IBEV | Gerador de comprovantes
 * 100% HTML/CSS/JavaScript, sem bibliotecas externas e sem backend.
 * Lê o XLSX (formato OpenXML) no dispositivo e produz DOCX (OpenXML) válido.
 * Requer navegador moderno com DecompressionStream('deflate-raw') (Chrome / Edge / Firefox atuais).
 */
(() => {
  'use strict';

  const CATEGORIES = [
    'FOLHA DE PAGAMENTO / PESSOAL',
    'ADMINISTRATIVO',
    'PATRIMONIO',
    'REDES',
    'DESPESA ALIMENTAÇÃO',
    'DESPESAS BANCÁRIAS'
  ];
  const BANK_CATEGORY = 'DESPESAS BANCÁRIAS';
  const MONTHS = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO',
    'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XMLNS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const CURRENCY = new Intl.NumberFormat('pt-BR',{ style:'currency',currency:'BRL' });
  const el = (id) => document.getElementById(id);
  const state = { fileName:'', workbook:null, month:'', parsed:null, selected:new Set(CATEGORIES) };

  function normalize(value) {
    return String(value == null ? '' : value).normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z0-9]/g,'');
  }
  function safeMoney(value) { return CURRENCY.format(value || 0); }
  function cents(value) { return Math.round((Number(value) + Number.EPSILON)*100); }
  function fromCents(value) { return value/100; }
  function parseNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string') return null;
    let t = value.trim().replace(/\s|R\$/gi,'');
    if (!t || !/^-?[\d.,]+$/.test(t)) return null;
    if (t.includes(',')) t = t.replace(/\./g,'').replace(',','.');
    else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t=t.replace(/\./g,'');
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
  }
  function xmlText(s) {
    return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;')
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'');
  }
  function parseXml(text,path) {
    const doc=new DOMParser().parseFromString(text,'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('XML inválido no arquivo: '+path);
    return doc;
  }
  function directChild(parent,local) {
    return Array.from(parent?.children||[]).find(node=>node.localName===local) || null;
  }

  // ZIP reader: central directory, store and Deflate; limits protect against malformed uploads.
  class ZipReader {
    constructor(buffer) {
      this.bytes=new Uint8Array(buffer);
      this.view=new DataView(buffer);
      this.files=new Map();
      this.scan();
    }
    scan() {
      const b=this.bytes,v=this.view;
      if(b.length<22) throw new Error('O arquivo não é um ZIP/XLSX válido.');
      let end=-1;
      for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--){
        if(v.getUint32(i,true)===0x06054b50){end=i;break;}
      }
      if(end<0) throw new Error('O arquivo não contém o diretório ZIP esperado.');
      const count=v.getUint16(end+10,true), offset=v.getUint32(end+16,true);
      if(count>25000 || offset>b.length) throw new Error('Arquivo ZIP excessivamente grande ou inválido.');
      let p=offset;
      for(let i=0;i<count;i++){
        if(p+46>b.length||v.getUint32(p,true)!==0x02014b50) throw new Error('Diretório ZIP danificado.');
        const flags=v.getUint16(p+8,true), method=v.getUint16(p+10,true);
        const comp=v.getUint32(p+20,true), uncomp=v.getUint32(p+24,true);
        const nameLen=v.getUint16(p+28,true),extraLen=v.getUint16(p+30,true),commentLen=v.getUint16(p+32,true);
        const loc=v.getUint32(p+42,true);
        if(comp===0xffffffff||uncomp===0xffffffff||loc===0xffffffff)throw new Error('Arquivos ZIP64 não são compatíveis.');
        const endName=p+46+nameLen;
        if(endName>b.length)throw new Error('Nome de arquivo ZIP inválido.');
        const name=new TextDecoder('utf-8').decode(b.subarray(p+46,endName));
        if((flags&1)!==0) throw new Error('Planilhas protegidas por senha não são compatíveis.');
        if(this.files.has(name)) throw new Error('ZIP com entradas duplicadas não é aceito.');
        this.files.set(name,{method,comp,uncomp,loc});
        p=endName+extraLen+commentLen;
      }
    }
    has(path){return this.files.has(path);}
    async read(path){
      const info=this.files.get(path);
      if(!info)throw new Error('A planilha não contém a parte necessária: '+path);
      if(info.uncomp>32_000_000)throw new Error('Uma parte interna do Excel é grande demais.');
      const v=this.view,b=this.bytes,p=info.loc;
      if(p+30>b.length||v.getUint32(p,true)!==0x04034b50)throw new Error('Entrada ZIP inválida.');
      const begin=p+30+v.getUint16(p+26,true)+v.getUint16(p+28,true);
      if(begin+info.comp>b.length)throw new Error('Arquivo ZIP truncado.');
      const chunk=b.subarray(begin,begin+info.comp);
      let out;
      if(info.method===0){out=chunk;}
      else if(info.method===8){
        if(typeof DecompressionStream==='undefined')throw new Error('Atualize seu navegador para abrir planilhas Excel.');
        let stream;
        try{stream=new Blob([chunk]).stream().pipeThrough(new DecompressionStream('deflate-raw'));}
        catch(e){throw new Error('Seu navegador não oferece descompactação ZIP. Use Chrome, Edge ou Firefox atualizados.');}
        const buffer=await new Response(stream).arrayBuffer();
        out=new Uint8Array(buffer);
      }else{throw new Error('O ZIP usa um método de compactação não suportado ('+info.method+').');}
      if(out.length!==info.uncomp)throw new Error('Tamanho interno de planilha inconsistente.');
      return new TextDecoder('utf-8').decode(out);
    }
  }

  function pathFromTarget(base,target){
    if(!target||/^[a-z]+:/i.test(target))throw new Error('Referência externa não suportada no Excel.');
    const parts=(target.startsWith('/') ? target.slice(1):base+'/'+target).split('/');
    const normalized=[];
    for(const p of parts){if(p==='..')normalized.pop();else if(p&&p!=='.')normalized.push(p);}
    const result=normalized.join('/');
    if(!result.startsWith('xl/'))throw new Error('Relação de planilha fora da pasta XL.');
    return result;
  }
  async function readWorkbook(buffer) {
    const zip=new ZipReader(buffer);
    const [workbookXML,relsXML]=await Promise.all([
      zip.read('xl/workbook.xml'),zip.read('xl/_rels/workbook.xml.rels')
    ]);
    const wbDoc=parseXml(workbookXML,'workbook.xml');
    const relDoc=parseXml(relsXML,'workbook.xml.rels');
    const targets=new Map();
    for(const rel of relDoc.getElementsByTagNameNS('*','Relationship')){
      targets.set(rel.getAttribute('Id'),rel.getAttribute('Target'));
    }
    const sheets=new Map();
    for(const sh of wbDoc.getElementsByTagNameNS(NS_MAIN,'sheet')){
      const name=sh.getAttribute('name');
      const rid=sh.getAttributeNS(NS_OFFICE_REL,'id')||sh.getAttribute('r:id');
      const target=targets.get(rid);
      if(name&&target)sheets.set(name,pathFromTarget('xl',target));
    }
    if(!sheets.size)throw new Error('Nenhuma aba válida foi encontrada no Excel.');
    const shared=[];
    if(zip.has('xl/sharedStrings.xml')){
      const stringsDoc=parseXml(await zip.read('xl/sharedStrings.xml'),'sharedStrings.xml');
      for(const si of stringsDoc.getElementsByTagNameNS(NS_MAIN,'si')){
        shared.push(Array.from(si.getElementsByTagNameNS(NS_MAIN,'t')).map(t=>t.textContent).join(''));
      }
    }
    return {zip,sheets,shared};
  }
  async function readSheet(workbook,sheetName){
    const path=workbook.sheets.get(sheetName);
    if(!path)throw new Error('Mês não encontrado na planilha: '+sheetName);
    const doc=parseXml(await workbook.zip.read(path),path);
    const rows=[];
    for(const rowNode of doc.getElementsByTagNameNS(NS_MAIN,'row')){
      const cells={}, formulas={};
      for(const cell of rowNode.getElementsByTagNameNS(NS_MAIN,'c')){
        const address=cell.getAttribute('r')||'';
        const column=(address.match(/^[A-Z]+/i)||[])[0]?.toUpperCase();
        if(!column)continue;
        const kind=cell.getAttribute('t');
        const v=directChild(cell,'v');
        const inline=directChild(cell,'is');
        const formula=directChild(cell,'f');
        if(formula?.textContent)formulas[column]=formula.textContent;
        let value=null;
        if(kind==='s')value=workbook.shared[Number(v?.textContent)]??'';
        else if(kind==='inlineStr')value=Array.from(inline?.getElementsByTagNameNS(NS_MAIN,'t')||[]).map(t=>t.textContent).join('');
        else if(kind==='str')value=v?.textContent??'';
        else if(kind==='b')value=(v?.textContent==='1');
        else if(kind==='e')value=null;
        else if(v?.textContent!=null){
          const n=Number(v.textContent);
          value=Number.isFinite(n)&&v.textContent.trim()!==''?n:v.textContent;
        }
        if(value!==null&&value!=='')cells[column]=value;
      }
      rows.push({index:Number(rowNode.getAttribute('r'))||rows.length+1,cells,formulas});
    }
    return rows;
  }

  function extractMonth(rows,month){
    const catalog=new Set(CATEGORIES.map(normalize));
    // Janeiro/2026 has a typo: the expense block is labeled RESUMO ENTRADA.
    // Choose the marker followed by the largest number of real category headers.
    const candidates=[];
    rows.forEach((row,i)=>{
      const marker=normalize(row.cells.A);
      if(marker==='RESUMOSAIDA'||marker==='RESUMOENTRADA'){
        const following=new Set();
        for(let j=i+1;j<rows.length;j++){
          const r=rows[j];
          if(normalize(r.cells.A)==='TOTALRESUMO')break;
          if(catalog.has(normalize(r.cells.A))&&parseNumber(r.cells.B)===null)following.add(normalize(r.cells.A));
        }
        candidates.push({index:i,marker,score:following.size});
      }
    });
    candidates.sort((a,b)=>b.score-a.score || (a.marker==='RESUMOSAIDA'?-1:1) || b.index-a.index);
    const winner=candidates[0];
    if(!winner||winner.score<2)throw new Error('Não encontrei o bloco RESUMO SAÍDA e suas categorias na aba '+month+'.');
    const start=winner.index;
    const headers=[];
    let overall=null;
    for(let i=start+1;i<rows.length;i++){
      const row=rows[i];
      const name=normalize(row.cells.A);
      if(name==='TOTALRESUMO'){
        overall=parseNumber(row.cells.B);
        break;
      }
      // Crucial: item "PATRIMONIO" with amount in B is NOT a new category header.
      if(catalog.has(name)&&parseNumber(row.cells.B)===null){
        headers.push({index:i,key:name,title:CATEGORIES.find(c=>normalize(c)===name)});
      }
    }
    const headerMap=new Map(headers.map(h=>[h.key,h]));
    const categories=[];
    const warnings=[];
    if(winner.marker==='RESUMOENTRADA')warnings.push('O cabeçalho de saída nesta aba está escrito "RESUMO ENTRADA"; a seção correta foi reconhecida pelas categorias.');
    for(const name of CATEGORIES){
      const key=normalize(name),head=headerMap.get(key);
      if(!head){
        categories.push({name,items:[],total:0,reported:null,difference:null,missing:true});
        warnings.push('Categoria não localizada: '+name+'.');
        continue;
      }
      const next=headers.filter(h=>h.index>head.index).sort((a,b)=>a.index-b.index)[0];
      let end=next?next.index:rows.findIndex((r,i)=>i>head.index&&normalize(r.cells.A)==='TOTALRESUMO');
      if(end<0)end=rows.length;
      const items=[], classifications=[];
      let explicit=null,subtotals=[];
      for(let i=head.index+1;i<end;i++){
        const c=rows[i].cells;
        const description=typeof c.A==='string'?c.A.trim():'';
        const n=normalize(description);
        const amount=parseNumber(c.B),colC=parseNumber(c.C);
        if(n==='TOTAL') {explicit=amount;break;}
        if(colC!==null)subtotals.push(colC);
        if(!description || n==='TOTALRESUMO' || amount===null)continue;
        classifications.push({description,amount:fromCents(cents(amount)),formula:rows[i].formulas?.B||''});
        if(cents(amount)===0)continue;
        items.push({description,amount:fromCents(cents(amount))});
      }
      const sumCents=items.reduce((sum,item)=>sum+cents(item.amount),0);
      const reported=explicit!==null?explicit:(subtotals.length?subtotals.at(-1):null);
      const difference=reported===null?null:sumCents-cents(reported);
      categories.push({name,items,classifications,total:fromCents(sumCents),reported,difference,missing:false});
    }
    const sumAllCents=categories.reduce((total,c)=>total+cents(c.total),0);
    const overallDifference=overall===null?null:sumAllCents-cents(overall);
    const issues=categories.filter(c=>c.missing||(c.difference!==null&&c.difference!==0));
    if(overallDifference!==null&&overallDifference!==0){
      warnings.push('Soma das categorias: '+safeMoney(fromCents(sumAllCents))+'; TOTAL RESUMO da planilha: '+safeMoney(overall)+'.');
    }
    const detail=extractDetails(rows,candidates.reduce((min,c)=>Math.min(min,c.index),start),categories);
    return {month,categories,overall,overallDifference,grandTotal:fromCents(sumAllCents),issues,warnings,...detail};
  }

  // Read the individual expense rows, not the aggregated summary. Historical
  // SUMIF formulas can use a classification name different from the displayed label.
  function extractDetails(rows,limit,categories){
    const tags=[], lookup=new Map(), detailWarnings=[];
    const matchKey=s=>String(s??'').trim().toLocaleUpperCase('pt-BR');
    for(const cat of categories){
      const grouped=new Map();
      for(const item of cat.classifications||[]){
        const key=normalize(item.description);
        let tag=grouped.get(key);
        if(!tag){
          tag={category:cat.name,name:item.description,expectedCents:0,items:[]};
          grouped.set(key,tag);tags.push(tag);
        }
        tag.expectedCents+=cents(item.amount);
        const criterion=item.formula.match(/SUMIF\s*\(\s*\$?F\$?\d+:\$?F\$?\d+\s*[,;]\s*"((?:[^"]|"")*)"/i);
        const name=criterion?criterion[1].replace(/""/g,'"'):item.description;
        for(const alias of [name,item.description]){
          const k=matchKey(alias);
          if(!lookup.has(k))lookup.set(k,new Set());
          lookup.get(k).add(tag);
        }
      }
    }
    const unmatched=new Map();
    let started=false;
    for(const row of rows.slice(0,limit)){
      const c=row.cells;
      if(normalize(c.A)==='DATA'&&normalize(c.B)==='HISTORICO'){started=true;continue;}
      if(!started)continue;
      // Only journal rows: date + amount, not balance/summary or repeated headings.
      const date=c.A, amount=parseNumber(c.D);
      if(amount===null||cents(amount)===0||!(typeof date==='number'&&date>1||typeof date==='string'&&/^(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{4}-\d{2}-\d{2})$/.test(date.trim())))continue;
      const classification=String(c.F??'').trim()||'SEM CLASSIFICAÇÃO';
      const matches=lookup.get(matchKey(classification));
      let tag;
      if(matches?.size===1)tag=[...matches][0];
      else{
        const key=normalize(classification);
        if(!unmatched.has(key))unmatched.set(key,{category:null,name:classification,expectedCents:null,items:[]});
        tag=unmatched.get(key);
      }
      tag.items.push({description:String(c.B??'').trim()||'(Sem histórico)',amount:fromCents(cents(amount)),row:row.index,date});
    }
    for(const tag of [...tags,...unmatched.values()]){
      tag.total=fromCents(tag.items.reduce((sum,item)=>sum+cents(item.amount),0));
      const invalidDates=tag.items.filter(i=>typeof i.date==='number'&&i.date>100000);
      if(invalidDates.length)detailWarnings.push({category:tag.category,text:tag.name+': data a conferir na(s) linha(s) '+invalidDates.map(i=>i.row).join(', ')+'. O valor foi preservado.'});
      if(tag.expectedCents!==null&&tag.expectedCents!==cents(tag.total)){
        detailWarnings.push({category:tag.category,text:tag.name+': lançamentos '+safeMoney(tag.total)+'; resumo '+safeMoney(fromCents(tag.expectedCents))+'.'});
      }
      if(tag.category===null)detailWarnings.push({category:null,text:tag.name+': '+tag.items.length+' lançamento(s) fora das classificações do resumo ('+safeMoney(tag.total)+').'});
    }
    return {tags:[...tags.filter(t=>t.items.length||t.expectedCents!==0),...unmatched.values()],detailWarnings};
  }

  // Minimal self-contained ZIP writer for DOCX. Files use ZIP "store" compression,
  // intentionally avoiding any third-party script / CDN / backend dependency.
  const ENCODER=new TextEncoder();
  const CRC_TABLE=(()=>{
    const table=new Uint32Array(256);
    for(let i=0;i<256;i++){
      let crc=i;
      for(let j=0;j<8;j++)crc=(crc&1)?(0xedb88320^(crc>>>1)):(crc>>>1);
      table[i]=crc>>>0;
    }
    return table;
  })();
  function crc32(bytes){
    let c=0xffffffff;
    for(const byte of bytes)c=CRC_TABLE[(c^byte)&255]^(c>>>8);
    return (c^0xffffffff)>>>0;
  }
  function zipPack(entries){
    const local=[],central=[];
    let offset=0;
    const now=new Date(),dosTime=((now.getHours()&31)<<11)|((now.getMinutes()&63)<<5)|((Math.floor(now.getSeconds()/2))&31);
    const dosDate=(((now.getFullYear()-1980)&127)<<9)|(((now.getMonth()+1)&15)<<5)|(now.getDate()&31);
    for(const [name,raw] of Object.entries(entries)){
      const nameBytes=ENCODER.encode(name),data=ENCODER.encode(raw),hash=crc32(data);
      const header=new Uint8Array(30+nameBytes.length),h=new DataView(header.buffer);
      h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);
      h.setUint16(8,0,true);h.setUint16(10,dosTime,true);h.setUint16(12,dosDate,true);
      h.setUint32(14,hash,true);h.setUint32(18,data.length,true);h.setUint32(22,data.length,true);
      h.setUint16(26,nameBytes.length,true);header.set(nameBytes,30);
      local.push(header,data);
      const cen=new Uint8Array(46+nameBytes.length),c=new DataView(cen.buffer);
      c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);
      c.setUint16(8,0x0800,true);c.setUint16(10,0,true);c.setUint16(12,dosTime,true);
      c.setUint16(14,dosDate,true);c.setUint32(16,hash,true);
      c.setUint32(20,data.length,true);c.setUint32(24,data.length,true);
      c.setUint16(28,nameBytes.length,true);c.setUint32(42,offset,true);cen.set(nameBytes,46);
      central.push(cen);offset+=header.length+data.length;
    }
    const centralSize=central.reduce((a,b)=>a+b.length,0),end=new Uint8Array(22),v=new DataView(end.buffer);
    v.setUint32(0,0x06054b50,true);v.setUint16(8,central.length,true);
    v.setUint16(10,central.length,true);v.setUint32(12,centralSize,true);v.setUint32(16,offset,true);
    return new Blob([...local,...central,end],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  }

  function paragraph(text='',options={}){
    const {bold=false,align='left',size=20,spaceBefore=0,spaceAfter=0,breakBefore=false}=options;
    const attrs=`<w:jc w:val="${align}"/><w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}" w:line="240" w:lineRule="auto"/>${breakBefore?'<w:pageBreakBefore/>':''}`;
    return `<w:p><w:pPr>${attrs}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size}"/>${bold?'<w:b/>':''}</w:rPr><w:t xml:space="preserve">${xmlText(text)}</w:t></w:r></w:p>`;
  }
  function cell(text,width,{bold=false,center=false,right=false,fill=false,font=20,marginL=250,marginR=190}={}){
    return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/><w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:left w:w="${marginL}" w:type="dxa"/><w:right w:w="${marginR}" w:type="dxa"/></w:tcMar>${fill?'<w:shd w:val="clear" w:fill="E8E8E8"/>':''}</w:tcPr>${paragraph(text,{bold,size:font,align:center?'center':right?'right':'left'})}</w:tc>`;
  }
  function makeExpenseTable(category,page){
    const rows=[...page.items];
    if(page.showTotal) rows.push({description:'TOTAL',amount:category.total,total:true});
    const col1=7519,col2=2121,tableWidth=9640;
    const n=rows.length+1;
    const rowH=Math.floor(9260/n);
    const font=n>16?16:n>12?18:20;
    const line=(data,total=false,head=false)=>{
      const opts={bold:head||total,fill:head,font,center:head,marginL:head?100:620,marginR:200};
      const vOpts={bold:head||total,fill:head,font,center:head,right:!head,marginL:50,marginR:190};
      const first=head?'DESCRIÇÃO':data.description;
      const second=head?'VALOR':safeMoney(data.amount);
      return `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${rowH}" w:hRule="exact"/></w:trPr>`+
        cell(first,col1,opts)+cell(second,col2,vOpts)+'</w:tr>';
    };
    return `<w:tbl><w:tblPr><w:tblW w:w="${tableWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/>`+
      `<w:tblBorders><w:top w:val="single" w:sz="5" w:color="202020"/><w:left w:val="single" w:sz="5" w:color="202020"/><w:bottom w:val="single" w:sz="5" w:color="202020"/><w:right w:val="single" w:sz="5" w:color="202020"/><w:insideH w:val="single" w:sz="5" w:color="202020"/><w:insideV w:val="single" w:sz="5" w:color="202020"/></w:tblBorders>`+
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>`+
      `<w:tblGrid><w:gridCol w:w="${col1}"/><w:gridCol w:w="${col2}"/></w:tblGrid>`+
      line({},false,true)+rows.map(r=>line(r,!!r.total)).join('')+'</w:tbl>';
  }
  function pageTable(title,table,continued=false){
    const widths=[5500,785,9640,673];
    const outerCell=(width,contents,border=false)=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>${border?'<w:tcBorders><w:right w:val="dashed" w:sz="4" w:color="888888"/></w:tcBorders>':''}</w:tcPr>${contents}</w:tc>`;
    const heading=paragraph(title+(continued?' (CONTINUAÇÃO)':''),{bold:true,align:'center',size:30,spaceBefore:460,spaceAfter:580});
    const left=outerCell(widths[0],paragraph('',{size:2}),true);
    const gap=outerCell(widths[1],paragraph('',{size:2}));
    const main=outerCell(widths[2],heading+table+paragraph('',{size:2}));
    const right=outerCell(widths[3],paragraph('',{size:2}));
    return `<w:tbl><w:tblPr><w:tblW w:w="16598" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="nil"/><w:bottom w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/><w:insideV w:val="nil"/><w:insideH w:val="nil"/></w:tblBorders></w:tblPr>`+
      `<w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`+
      `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="11000" w:hRule="exact"/></w:trPr>${left}${gap}${main}${right}</w:tr></w:tbl>`;
  }
  function overviewPage(categories,month){
    const rows=categories.map(c=>({description:c.name,amount:c.total}));
    const total=fromCents(categories.reduce((s,c)=>s+cents(c.total),0));
    const fake={total};
    return pageTable('RESUMO GERAL DE SAÍDAS - '+month,
      makeExpenseTable(fake,{items:rows,showTotal:true}));
  }
  function selectedTags(data,selected){
    // Uma única nota já pode ser conferida diretamente pela folha do caderno.
    // O filtro também controla a prévia de quantidade e a paginação do Word.
    return (data.tags||[]).filter(t=>t.items.length>1&&(t.category===null?selected.size===CATEGORIES.length:selected.has(t.category)));
  }
  function tagDate(value){
    if(typeof value!=='number')return String(value??'');
    if(value>100000)return 'A conferir';
    const d=new Date(Date.UTC(1899,11,30)+Math.floor(value)*86400000);
    return d.toLocaleDateString('pt-BR',{timeZone:'UTC'});
  }
  // A4 portrait: two equal slots, or a full sheet when the list exceeds
  // the half-page budget. Long histories are conservatively wrapped.
  const TAG_HALF=7300, TAG_FULL=15100, TAG_WIDTH=10520;
  function tagWrap(value,max=43){
    const lines=[];
    for(const part of String(value).split(/\r?\n/)){
      let line='';
      for(let word of part.split(/\s+/)){
        if(!word)continue;
        while(word.length>max){
          if(line){lines.push(line);line='';}
          lines.push(word.slice(0,max));word=word.slice(max);
        }
        if(line.length+word.length+1>max){lines.push(line);line=word;}
        else line+=(line?' ':'')+word;
      }
      if(line)lines.push(line);
    }
    return lines.length?lines:[''];
  }
  function tagRowHeight(item){return 360+Math.max(0,tagWrap(item.description).length-1)*240;}
  function tagPlan(tags){
    const pages=[];let pending=null;
    const flush=()=>{if(pending){pages.push({slots:[pending],full:false});pending=null;}};
    for(const tag of tags){
      const cost=2500+tag.items.reduce((s,i)=>s+tagRowHeight(i),0);
      if(cost<=TAG_HALF-400){
        const slot={tag,items:tag.items,start:0,part:1,parts:1,last:true};
        if(pending){pages.push({slots:[pending,slot],full:false});pending=null;}else pending=slot;
      }else{
        flush();
        const chunks=[];let chunk=[],used=2500;
        for(const item of tag.items){
          const h=tagRowHeight(item);
          if(chunk.length&&used+h>TAG_FULL-850){chunks.push(chunk);chunk=[];used=2500;}
          chunk.push(item);used+=h;
        }
        if(chunk.length||!chunks.length)chunks.push(chunk);
        let start=0;
        chunks.forEach((items,i)=>{
          pages.push({full:true,slots:[{tag,items,start,part:i+1,parts:chunks.length,last:i===chunks.length-1}]});
          start+=items.length;
        });
      }
    }
    flush();return pages;
  }
  function tagCell(text,width,opts={}){
    const {bold=false,right=false,center=false,fill=false,wrap=false}=opts;
    const lines=wrap?tagWrap(text):[text];
    return '<w:tc><w:tcPr><w:tcW w:w="'+width+'" w:type="dxa"/>'+
      (fill?'<w:shd w:val="clear" w:fill="E8E8E8"/>':'')+
      '<w:tcMar><w:top w:w="45" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="45" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>'+
      lines.map(line=>paragraph(line,{size:19,bold,align:center?'center':right?'right':'left'})).join('')+'</w:tc>';
  }
  function detailTable(slot){
    const {tag,items,start,last}=slot;
    const widths=[620,1400,6200,2300],w=widths.reduce((s,n)=>s+n,0);
    const row=(values,header=false,total=false)=>'<w:tr><w:trPr><w:cantSplit/>'+
      (header?'<w:tblHeader/>':'')+'</w:trPr>'+
      values.map((v,i)=>tagCell(v,widths[i],{bold:header||total,center:header||i===0,right:i===3&&!header,fill:header,wrap:i===2&&!header})).join('')+'</w:tr>';
    const sum=fromCents(items.reduce((s,i)=>s+cents(i.amount),0));
    let content=row(['Nº','DATA','HISTÓRICO / DESPESA','VALOR'],true);
    content+=items.map((item,i)=>row([String(start+i+1),tagDate(item.date),item.description,safeMoney(item.amount)])).join('');
    if(slot.parts>1)content+=row(['','','SUBTOTAL DESTA FOLHA',safeMoney(sum)],false,true);
    if(last)content+=row(['','','TOTAL DA CLASSIFICAÇÃO',safeMoney(tag.total)],false,true);
    return '<w:tbl><w:tblPr><w:tblW w:w="'+w+'" w:type="dxa"/>'+
      '<w:tblBorders>'+['top','left','bottom','right','insideH','insideV'].map(side=>'<w:'+side+' w:val="single" w:sz="4" w:color="999999"/>').join('')+'</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'+
      '<w:tblGrid>'+widths.map(n=>'<w:gridCol w:w="'+n+'"/>').join('')+'</w:tblGrid>'+content+'</w:tbl>';
  }
  function detailSlot(slot,month){
    const tag=slot.tag;
    const subtitle=(tag.category||'FORA DO RESUMO DE SAÍDAS')+' • '+month;
    const continuation=slot.parts>1?' ('+slot.part+'/'+slot.parts+')':'';
    let content=paragraph('IBEV | '+subtitle,{size:18,bold:true,spaceAfter:100})+
      paragraph(tag.name.toUpperCase()+continuation,{size:28,bold:true,spaceAfter:130})+
      paragraph(tag.items.length+' lançamento(s) • Ordem da planilha',{size:18,spaceAfter:140})+
      detailTable(slot);
    if(slot.last&&tag.expectedCents!==null&&tag.expectedCents!==cents(tag.total)){
      content+=paragraph('ATENÇÃO: resumo da planilha = '+safeMoney(fromCents(tag.expectedCents))+'; conferir a diferença.',{size:18,bold:true,spaceBefore:140});
    }
    if(tag.category===null)content+=paragraph('Lançamento fora do resumo. Não incluído no total das folhas do caderno.',{size:18,spaceBefore:140});
    return content+paragraph('',{size:2});
  }
  function detailPage(page,month){
    const slots=page.full?page.slots:[page.slots[0],page.slots[1]||null];
    const height=page.full?TAG_FULL:TAG_HALF;
    return '<w:tbl><w:tblPr><w:tblW w:w="'+TAG_WIDTH+'" w:type="dxa"/><w:tblBorders><w:insideH w:val="dashed" w:sz="5" w:color="777777"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'+
      '<w:tblGrid><w:gridCol w:w="'+TAG_WIDTH+'"/></w:tblGrid>'+
      slots.map(slot=>'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="'+height+'" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="'+TAG_WIDTH+'" w:type="dxa"/><w:tcMar><w:top w:w="230" w:type="dxa"/><w:bottom w:w="230" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar><w:vAlign w:val="top"/></w:tcPr>'+
        (slot?detailSlot(slot,month):paragraph('',{size:2}))+'</w:tc></w:tr>').join('')+'</w:tbl>';
  }
  function buildDocx(data,selected,withOverview,withTags=false){
    const pages=[];
    for(const c of data.categories){
      if(!selected.has(c.name))continue;
      // For long, future categories, create a continuation page rather than unreadably small cells.
      const chunks=[];
      if(!c.items.length)chunks.push([]);
      else for(let i=0;i<c.items.length;i+=13)chunks.push(c.items.slice(i,i+13));
      chunks.forEach((chunk,i)=>pages.push(pageTable(c.name,
        makeExpenseTable(c,{items:chunk,showTotal:i===chunks.length-1}),i>0)));
    }
    if(withOverview)pages.push(overviewPage(data.categories.filter(c=>selected.has(c.name)),data.month));
    if(!pages.length)throw new Error('Escolha pelo menos uma categoria.');
    const breaks=pages.map((page,i)=>(i===0?'':paragraph('',{breakBefore:true,size:2}))+page).join('');
    const landscape='<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="180" w:right="120" w:bottom="180" w:left="120" w:header="0" w:footer="0" w:gutter="0"/><w:cols w:space="0"/><w:docGrid w:linePitch="240"/></w:sectPr>';
    const portrait='<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/><w:pgMar w:top="570" w:right="693" w:bottom="570" w:left="693" w:header="0" w:footer="0" w:gutter="0"/><w:cols w:space="0"/></w:sectPr>';
    const tagPages=withTags?tagPlan(selectedTags(data,selected)):[];
    const sectionBreak='<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'+landscape+'</w:pPr></w:p>';
    const pageSection='<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr>'+portrait+'</w:pPr></w:p>';
    const tagContent=tagPages.map((page,i)=>detailPage(page,data.month)+(i<tagPages.length-1?pageSection:'')).join('');
    const doc=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
      `<w:document xmlns:w="${XMLNS_W}"><w:body>${breaks}${tagPages.length?sectionBreak+tagContent+portrait:landscape}</w:body></w:document>`;
    const types=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const rels=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
    const wordRels=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${XMLNS_W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
    const currentDate=new Date().toISOString();
    const core=`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>IBEV - Caderno de comprovantes - ${xmlText(data.month)}</dc:title><dc:creator>Gerador IBEV</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${currentDate}</dcterms:created></cp:coreProperties>`;
    const app=`<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>IBEV Gerador</Application></Properties>`;
    return zipPack({'[Content_Types].xml':types,'_rels/.rels':rels,'word/document.xml':doc,
      'word/_rels/document.xml.rels':wordRels,'word/styles.xml':styles,
      'docProps/core.xml':core,'docProps/app.xml':app});
  }

  function alertValidation(){
    const output=el('validation'),data=state.parsed;
    if(!data){output.className='validation neutral';output.textContent='Aguardando a seleção da planilha.';return;}
    const relevant=data.issues.filter(i=>state.selected.has(i.name));
    const lines=[];
    for(const c of relevant){
      if(c.missing)lines.push(c.name+': categoria ausente na planilha.');
      else lines.push(c.name+': itens '+safeMoney(c.total)+'; subtotal registrado '+safeMoney(c.reported)+' (diferença '+safeMoney(c.difference/100)+').');
    }
    // A global mismatch also matters even if subset of categories is selected. Not block subset output on overall difference alone.
    if(data.overallDifference!==null&&data.overallDifference!==0){
      lines.push('O TOTAL RESUMO da aba é '+safeMoney(data.overall)+', mas as categorias calculadas somam '+safeMoney(data.grandTotal)+'.');
    }
    const detailIssues=el('include-tags').checked?(data.detailWarnings||[]).filter(w=>w.category===null?state.selected.size===CATEGORIES.length:state.selected.has(w.category)):[];
    lines.push(...detailIssues.map(w=>w.text));
    const tagCount=selectedTags(data,state.selected);
    el('tag-count').textContent=el('include-tags').checked?(tagCount.length?tagCount.length+' etiqueta(s) • '+tagPlan(tagCount).length+' folha(s) A4 retrato. Somente classificações com dois ou mais lançamentos.':'Nenhuma classificação selecionada tem dois ou mais lançamentos. Apenas as folhas do caderno serão emitidas.'):'Etiquetas desativadas. Apenas as folhas do caderno serão emitidas.';
    if(relevant.length||detailIssues.length||(data.overallDifference!==null&&data.overallDifference!==0)){
      output.className='validation warning';
      output.innerHTML='<strong>Conferência necessária: há diferenças nos registros.</strong><ul>'+lines.map(t=>'<li>'+xmlText(t)+'</li>').join('')+'</ul>';
      el('acknowledgement-wrap').hidden=false;
    }else if(data.warnings.length){
      output.className='validation warning';
      output.innerHTML='<strong>Leitura concluída com observações.</strong><ul>'+data.warnings.map(t=>'<li>'+xmlText(t)+'</li>').join('')+'</ul>';
      el('acknowledgement-wrap').hidden=true;
    }else{
      output.className='validation ok';
      output.innerHTML='<strong>Valores conferidos.</strong> As categorias coincidem com os subtotais registrados no Excel. Pronto para gerar o Word.';
      el('acknowledgement-wrap').hidden=true;
    }
    const needAck=relevant.length>0||detailIssues.length>0||(data.overallDifference!==null&&data.overallDifference!==0);
    el('download-btn').disabled=state.selected.size===0||(needAck&&!el('acknowledge').checked);
  }
  function showCategory(name){
    const cat=state.parsed?.categories.find(c=>c.name===name);
    if(!cat)return;
    el('paper-title').textContent=cat.name;
    const body=el('paper-rows');body.replaceChildren();
    let items=cat.items;
    if(items.length>13)items=items.slice(0,13);
    for(const item of [...items,{description:'TOTAL',amount:cat.total,total:true}]){
      const tr=document.createElement('tr');if(item.total)tr.className='total';
      const a=document.createElement('td'),b=document.createElement('td');
      a.textContent=item.description;b.textContent=safeMoney(item.amount);tr.append(a,b);body.appendChild(tr);
    }
    const count=cat.items.length;
    el('preview-counter').textContent='A4 • '+count+' item'+(count===1?'':'s');
  }
  function refreshSummary(){
    const data=state.parsed;
    if(!data)return;
    const list=el('categories');list.replaceChildren();
    for(const cat of data.categories){
      if(cat.name===BANK_CATEGORY){
        el('include-bank').checked=state.selected.has(BANK_CATEGORY);
        el('bank-total').textContent=safeMoney(cat.total)+(cat.missing?' ⚠':'');
        el('bank-option').classList.toggle('has-issue',!!(cat.missing||cat.difference));
        continue;
      }
      const row=document.createElement('label');row.className='category-row';
      if(cat.missing||cat.difference)row.classList.add('has-issue');
      const check=document.createElement('input');check.type='checkbox';check.checked=state.selected.has(cat.name);
      check.addEventListener('change',()=>{
        if(check.checked)state.selected.add(cat.name);else state.selected.delete(cat.name);
        el('acknowledge').checked=false;refreshSummary();
      });
      const name=document.createElement('span');name.className='category-title';name.textContent=cat.name;
      const val=document.createElement('span');val.className='category-value';val.textContent=safeMoney(cat.total)+(cat.missing?' ⚠':'');
      row.append(check,name,val);list.appendChild(row);
    }
    el('selected-total').textContent=safeMoney(data.categories.filter(c=>state.selected.has(c.name)).reduce((s,c)=>s+c.total,0));
    el('toggle-all').textContent=state.selected.size===CATEGORIES.length?'Desmarcar todas':'Marcar todas';
    const chosen=el('preview-select').value;
    if(!data.categories.some(c=>c.name===chosen))el('preview-select').value=CATEGORIES[0];
    showCategory(el('preview-select').value);
    alertValidation();
  }
  async function selectMonth(month){
    state.month=month;
    el('month-select').disabled=true;el('download-btn').disabled=true;el('include-bank').disabled=true;
    el('validation').className='validation neutral';el('validation').textContent='Conferindo os lançamentos e os subtotais...';
    try{
      const rows=await readSheet(state.workbook,month);
      state.parsed=extractMonth(rows,month);
      el('acknowledge').checked=false;
      el('preview-select').disabled=false;
      el('include-overview').disabled=false;
      el('include-tags').disabled=false;
      el('include-bank').disabled=false;
      el('toggle-all').disabled=false;
      refreshSummary();
    }catch(err){state.parsed=null;el('validation').className='validation error';el('validation').textContent=err.message;
      el('download-btn').disabled=true;el('categories').innerHTML='<p class="empty-message">Não foi possível identificar as despesas deste mês.</p>';}
    finally{el('month-select').disabled=false;}
  }
  async function loadExcel(file){
    if(!file)return;
    if(!/\.xlsx$/i.test(file.name)){
      el('validation').className='validation error';el('validation').textContent='Escolha um arquivo .xlsx. Arquivos .xls e .xlsm não são aceitos.';return;
    }
    if(file.size>35_000_000){el('validation').className='validation error';el('validation').textContent='Planilha maior que 35 MB. Verifique o arquivo.';return;}
    el('file-primary').textContent=file.name;
    el('file-secondary').textContent='Lendo a planilha com segurança...';
    el('validation').className='validation neutral';el('validation').textContent='Abrindo o Excel...';
    el('download-btn').disabled=true;
    el('include-bank').disabled=true;
    try{
      const wb=await readWorkbook(await file.arrayBuffer());
      const candidates=MONTHS.filter(m=>Array.from(wb.sheets.keys()).some(n=>normalize(n)===normalize(m)));
      if(!candidates.length)throw new Error('A planilha não contém abas mensais reconhecíveis.');
      const resolve=m=>Array.from(wb.sheets.keys()).find(n=>normalize(n)===normalize(m));
      const select=el('month-select');select.replaceChildren();
      for(const m of candidates){const option=document.createElement('option');option.value=resolve(m);option.textContent=m;select.appendChild(option);}
      state.workbook=wb;state.fileName=file.name;state.selected=new Set(CATEGORIES);
      el('file-secondary').textContent='Arquivo aberto • '+candidates.length+' meses encontrados';
      el('month-hint').textContent='Selecione qualquer mês disponível. O sistema lê os valores salvos no arquivo.';
      select.disabled=false;
      // Prefer August in the sample workbook; otherwise most recently selected month if available, else first.
      const preferred=(candidates.includes('AGOSTO')?'AGOSTO':candidates[0]);
      select.value=resolve(preferred);
      await selectMonth(select.value);
    }catch(err){state.workbook=null;state.parsed=null;
      el('validation').className='validation error';el('validation').textContent='Não consegui abrir o arquivo: '+(err.message||err);
      el('file-secondary').textContent='Verifique se é um .xlsx válido, não protegido por senha.';
      el('month-select').disabled=true;
    }
  }

  function init(){
    const upload=el('file-input'),drop=el('dropzone');
    upload.addEventListener('change',async()=>{await loadExcel(upload.files?.[0]);});
    for(const event of ['dragenter','dragover'])drop.addEventListener(event,e=>{e.preventDefault();drop.classList.add('dragover');});
    for(const event of ['dragleave','drop'])drop.addEventListener(event,e=>{e.preventDefault();drop.classList.remove('dragover');});
    drop.addEventListener('drop',async e=>{const file=e.dataTransfer.files?.[0];if(file)await loadExcel(file);});
    el('month-select').addEventListener('change',e=>selectMonth(e.target.value));
    el('preview-select').addEventListener('change',e=>showCategory(e.target.value));
    const preview=el('preview-select');preview.replaceChildren();
    for(const title of CATEGORIES){const opt=document.createElement('option');opt.value=title;opt.textContent=title;preview.appendChild(opt);}
    el('toggle-all').addEventListener('click',()=>{state.selected=state.selected.size===CATEGORIES.length?new Set():new Set(CATEGORIES);el('acknowledge').checked=false;refreshSummary();});
    el('include-bank').addEventListener('change',e=>{
      if(e.target.checked)state.selected.add(BANK_CATEGORY);else state.selected.delete(BANK_CATEGORY);
      el('acknowledge').checked=false;refreshSummary();
    });
    el('acknowledge').addEventListener('change',alertValidation);
    el('include-tags').addEventListener('change',()=>{el('acknowledge').checked=false;alertValidation();});
    el('download-btn').addEventListener('click',async()=>{
      if(!state.parsed||!state.selected.size)return;
      const relevant=state.parsed.issues.filter(i=>state.selected.has(i.name));
      const globalMismatch=state.parsed.overallDifference!==null&&state.parsed.overallDifference!==0;
      const detailIssues=el('include-tags').checked&&(state.parsed.detailWarnings||[]).some(w=>w.category===null?state.selected.size===CATEGORIES.length:state.selected.has(w.category));
      if((relevant.length||globalMismatch||detailIssues)&&!el('acknowledge').checked)return;
      const button=el('download-btn');button.disabled=true;button.textContent='Gerando documento Word...';
      try{
        const blob=buildDocx(state.parsed,state.selected,el('include-overview').checked,el('include-tags').checked);
        const url=URL.createObjectURL(blob);
        const a=document.createElement('a');a.href=url;
        a.download=`IBEV_Caderno_${normalize(state.month)}_${(state.fileName.match(/20\d{2}/)||[])[0]||'Financeiro'}.docx`;
        document.body.appendChild(a);a.click();a.remove();
        setTimeout(()=>URL.revokeObjectURL(url),5000);
        el('validation').className='validation ok';el('validation').innerHTML='<strong>Documento gerado.</strong> Confira a pasta Downloads e imprima respeitando a orientação de cada página do Word.';
      }catch(err){el('validation').className='validation error';el('validation').textContent='Falha ao gerar o Word: '+(err.message||err);}
      finally{button.textContent='↓  Baixar caderno em Word (.docx)';alertValidation();}
    });
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
  // Exposed only to help with local automated validation, no network calls.
  window.IBEV_TEST={readWorkbook,readSheet,extractMonth,buildDocx,normalize,tagPlan,selectedTags};
})();
