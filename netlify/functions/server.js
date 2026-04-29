'use strict';
const https = require('https');
const zlib = require('zlib');

function nReq(ep, method, payload, key) {
  return new Promise((resolve, reject) => {
    const data = payload ? JSON.stringify(payload) : undefined;
    const opts = {
      hostname: 'api.notion.com',
      path: '/v1' + ep,
      method: method || 'POST',
      headers: Object.assign({
        'Authorization': 'Bearer ' + key,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json'
      }, data ? { 'Content-Length': Buffer.byteLength(data) } : {})
    };
    const req = https.request(opts, res => {
      let d = '';
      res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { resolve({ error: d }); } });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

function crc32(buf) {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[i] = c; }
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function u32(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0, 0); return b; }
function u16(n) { const b = Buffer.alloc(2); b.writeUInt16LE(n & 0xFFFF, 0); return b; }

function buildZip(files) {
  const locals = []; const centrals = []; let offset = 0;
  files.forEach(f => {
    const name = Buffer.from(f.name, 'utf8');
    const raw = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, 'utf8');
    const comp = zlib.deflateRawSync(raw, { level: 6 });
    const crc = crc32(raw);
    const now = new Date();
    const dt = ((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
    const tm = (now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1);
    const lh = Buffer.concat([Buffer.from([0x50,0x4B,0x03,0x04]),u16(20),u16(0),u16(8),u16(tm),u16(dt),u32(crc),u32(comp.length),u32(raw.length),u16(name.length),u16(0),name]);
    const ch = Buffer.concat([Buffer.from([0x50,0x4B,0x01,0x02]),u16(20),u16(20),u16(0),u16(8),u16(tm),u16(dt),u32(crc),u32(comp.length),u32(raw.length),u16(name.length),u16(0),u16(0),u16(0),u16(0),u32(0),u32(offset),name]);
    locals.push(Buffer.concat([lh,comp])); centrals.push(ch);
    offset += lh.length + comp.length;
  });
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.concat([Buffer.from([0x50,0x4B,0x05,0x06]),u16(0),u16(0),u16(files.length),u16(files.length),u32(cd.length),u32(offset),u16(0)]);
  return Buffer.concat([...locals, cd, eocd]);
}

function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function run(text, opts) {
  opts = opts || {};
  const col = (opts.color||'#212121').replace('#','').toUpperCase();
  const sz = opts.size || 22;
  const rpr = [
    '<w:rFonts w:ascii="Montserrat" w:hAnsi="Montserrat" w:cs="Montserrat"/>',
    '<w:color w:val="'+col+'"/>',
    '<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/>',
    opts.bold?'<w:b/><w:bCs/>':'',
    opts.italic?'<w:i/><w:iCs/>':'',
    opts.underline?'<w:u w:val="single"/>':''
  ].filter(Boolean).join('');
  const t = esc(text);
  const sp = (t&&(t[0]===' '||t[t.length-1]===' '))?' xml:space="preserve"':'';
  return '<w:r><w:rPr>'+rpr+'</w:rPr><w:t'+sp+'>'+t+'</w:t></w:r>';
}

function hr() {
  // Horizontal rule
  return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CCCCCC"/></w:pBdr><w:spacing w:after="80"/></w:pPr></w:p>';
}

function p(runs, after) {
  const a = after !== undefined ? after : 80;
  return '<w:p><w:pPr><w:spacing w:after="'+a+'"/></w:pPr>'+(Array.isArray(runs)?runs.join(''):runs||'')+'</w:p>';
}
function bp(runs) { return '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="80"/></w:pPr>'+(Array.isArray(runs)?runs.join(''):runs)+'</w:p>'; }
function np(runs) { return '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:spacing w:after="80"/></w:pPr>'+(Array.isArray(runs)?runs.join(''):runs)+'</w:p>'; }
function ep() { return p([run(' ')], 80); }

const COLORS = {
  'AACC Preparatory Choir':'#D7282F','AACC Prep & Concert':'#D7282F','AACC Concert Choir':'#D7282F',
  'Anderson Area Youth Chorale':'#D7282F','AACC Youth Chorale':'#D7282F',
  'CICC Preparatory Choir':'#D7282F','CICC Concert Choir':'#D7282F','CICC Descant Choir':'#D7282F',
  'CICC Concert & Descant':'#D7282F','Foundations':'#D7282F','Early Childhood':'#D7282F',
  'ICC Preparatory Choirs':'#62A744','Preparatory Choir':'#62A744',
  'ICC Beginning Level Choirs':'#02793F','Beginning Level Choir':'#02793F',
  'ICC Lyric Choirs':'#462E8D','Lyric Choir':'#462E8D',
  'ICC IndyVoice':'#0083C1','Indy Voice':'#0083C1','Neighborhood Choir Academy':'#0083C1',
  'ICC Master Chorale':'#F6A800','Master Chorale':'#F6A800',
  'Jubilate':'#E63E51','Alumni Choir':'#E63E51',
  'Indianapolis Indiana Children\'s Choir':'#0083C1',
  'Anderson Area Children\'s Choir':'#D7282F',
  'Columbus Indiana Children\'s Choir':'#D7282F'
};
function cc(name) { return COLORS[name]||'#212121'; }

function rtRuns(rt, extra) {
  return (rt||[]).map(r => { const a=r.annotations||{}; return run(r.plain_text, Object.assign({bold:a.bold,italic:a.italic,underline:a.underline},extra||{})); });
}

function blocksToXml(blocks) {
  return (blocks||[]).map(b => {
    const t = b.type;
    if (t==='paragraph') { const rt=(b.paragraph||{}).rich_text||[]; return rt.length?p(rtRuns(rt)):ep(); }
    if (t==='bulleted_list_item') return bp(rtRuns((b.bulleted_list_item||{}).rich_text||[]));
    if (t==='numbered_list_item') return np(rtRuns((b.numbered_list_item||{}).rich_text||[]));
    if (t==='heading_1') return p([run((b.heading_1.rich_text||[]).map(r=>r.plain_text).join(''),{bold:true,size:32})],120);
    if (t==='heading_2') return p([run((b.heading_2.rich_text||[]).map(r=>r.plain_text).join(''),{bold:true,size:28})],100);
    if (t==='heading_3') return p([run((b.heading_3.rich_text||[]).map(r=>r.plain_text).join(''),{bold:true,size:24})],80);
    return '';
  }).join('');
}

function buildDocx(d) {
  const pub = new Date(d.publishDate+'T12:00:00');
  const mn = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const fileName = 'TNMB_-_'+mn[pub.getMonth()]+'_'+pub.getDate()+'_'+pub.getFullYear()+'.docx';
  const dl = d.deadline || 'EOD Monday';
  const body = [];

  body.push(p([run('Hi Team!')]));
  body.push(ep());
  body.push(p([run('I hope your week is off to a wonderful start. '),run('Please take the time to review',{bold:true}),run(' and '),run('respond with any changes or additions',{bold:true}),run(' '),run('by ',{bold:true}),run(dl,{bold:true}),run('.',{bold:true})]));
  body.push(ep());
  body.push(p([run('Thank you for your help in making these happen!')]));
  body.push(ep());

  body.push(p([run('Breaks & Schedule Changes:',{bold:true,size:36,color:'#000000'})],120));
  body.push(ep());
  if (!d.breaksSection||d.breaksSection.length===0) {
    body.push(p([run('No Changes This Week',{italic:true,color:'#212121'})]));
  } else {
    d.breaksSection.forEach(item => {
      const col = cc(item.choirName);
      body.push(p([run(item.choirName,{color:col}),run(' | ',{color:col}),run(item.entryName,{italic:true,color:col})]));
    });
  }
  body.push(ep());
  body.push(hr());

  body.push(p([run('Upcoming Events:',{bold:true,size:36,color:'#000000'})],120));

  (d.choirSections||[]).forEach(sec => {
    body.push(hr());
    const col = cc(sec.label);
    body.push(p([run(sec.label+':',{bold:true,size:28,color:col})],60));
    if (!sec.events||sec.events.length===0) {
      body.push(p([run('No Upcoming Events',{italic:true,color:'#212121'})]));
    } else {
      sec.events.forEach((ev,i) => {
        if (i>0) body.push(p([run(' ')], 40));
        body.push(p([run(ev.name+' | '+ev.formattedDate,{bold:true})],60));
        if (ev.venueName) body.push(p([run('Location: '+(ev.venueName+(ev.venueAddress?' | '+ev.venueAddress:'')))],40));
        if (ev.uniform) body.push(p([run('Uniform: '+ev.uniform)],40));
        if (ev.callTime) body.push(p([run('Call Time: '+ev.callTime)],40));
        if (ev.perfTime) body.push(p([run('Performance: '+ev.perfTime)],40));
        if (ev.dismissal) body.push(p([run('Approx. Dismissal: '+ev.dismissal)],40));
        if (ev.bodyBlocks&&ev.bodyBlocks.length>0) body.push(blocksToXml(ev.bodyBlocks));
      });
    }
  });

  body.push(ep());
  body.push(p([run('Marketing, Development & Admin:',{bold:true,size:28,color:'#000000'})],120));
  body.push(ep());
  body.push(p([run('Topic:',{bold:true,size:21,color:'#212121'})],40));
  body.push(p([run('Topic Heading',{bold:true,size:21,color:'#212121'})],40));
  body.push(p([run('Body Copy',{size:20,color:'#212121'})],80));
  body.push(ep());
  body.push(p([run('Thank you!')]));
  body.push(ep());
  body.push(p([run('Thank you for taking the time to proof & provide your additions!')]));
  body.push(ep());
  body.push(p([run('Best,')]));

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
${body.join('\n')}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="Montserrat" w:hAnsi="Montserrat" w:cs="Montserrat"/>
  <w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="212121"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="80"/></w:pPr></w:pPrDefault>
</w:docDefaults></w:styles>`;

  const numXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">
  <w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#x2022;"/>
  <w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Montserrat" w:hAnsi="Montserrat"/></w:rPr>
</w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">
  <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
  <w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
</w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;

  const buf = buildZip([
    {name:'[Content_Types].xml',data:ct},
    {name:'_rels/.rels',data:rootRels},
    {name:'word/document.xml',data:docXml},
    {name:'word/styles.xml',data:stylesXml},
    {name:'word/numbering.xml',data:numXml},
    {name:'word/_rels/document.xml.rels',data:docRels}
  ]);
  return {buffer:buf, fileName};
}

exports.handler = async (event) => {
  const cors = {'Access-Control-Allow-Origin':'*','Access-Control-Allow-Headers':'Content-Type'};
  if (event.httpMethod==='OPTIONS') return {statusCode:200,headers:cors,body:''};
  if (event.httpMethod!=='POST') return {statusCode:405,headers:cors,body:'Method not allowed'};
  let body; try { body=JSON.parse(event.body); } catch { return {statusCode:400,headers:cors,body:'Invalid JSON'}; }
  const {action,notionKey,payload} = body;

  if (action==='notion') {
    try {
      const r = await nReq(payload.endpoint,payload.method,payload.data,notionKey);
      return {statusCode:200,headers:{...cors,'Content-Type':'application/json'},body:JSON.stringify(r)};
    } catch(e) { return {statusCode:500,headers:cors,body:e.message}; }
  }

  if (action==='generate_doc') {
    try {
      const {buffer,fileName} = buildDocx(payload);
      return {statusCode:200,headers:{...cors,'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':'attachment; filename="'+fileName+'"'},body:buffer.toString('base64'),isBase64Encoded:true};
    } catch(e) { return {statusCode:500,headers:cors,body:'Doc error: '+e.message}; }
  }

  return {statusCode:400,headers:cors,body:'Unknown action'};
};
