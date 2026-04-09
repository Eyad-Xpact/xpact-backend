'use strict';
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// Load assets relative to this file
const IMGS = JSON.parse(fs.readFileSync(path.join(__dirname, 'assets/images.json'), 'utf8'));


const C = {
  dark: '171614', card: '1E1E1C', teal: '224F5B',
  mint: '00BD9C', white: 'FFFFFF', light: 'F5F5F3',
  g1: 'E2E2E2', g2: '9D9D9D', g3: '2A2A28',
};
const FONT = 'Montserrat';
const FONT_AR = 'Arial';

function isArabicText(text) {
  if (!text) return false;
  const arabicChars = (text.match(/[\u0600-\u06FF]/g) || []).length;
  return arabicChars > text.length * 0.3;
}

function txt(text, opts) {
  const arabic = isArabicText(text);
  return Object.assign({
    rtlMode: arabic,
    fontFace: arabic ? FONT_AR : FONT,
    align: arabic ? 'right' : (opts.align || 'left'),
  }, opts, { rtlMode: arabic || opts.rtlMode });
}

function footer(slide, pageNum, total) {
  slide.addShape('rect', { x:0, y:5.43, w:10, h:0.195, fill:{color:C.dark}, line:{color:C.dark} });
  slide.addText('xpact.net  |  e.matar[at]xpact.net  |  +966 53 587 9603', {
    x:0.3, y:5.44, w:7, h:0.18, fontSize:7, fontFace:FONT, color:C.g2, valign:'middle', margin:0
  });
  if (pageNum) {
    slide.addText(pageNum + ' / ' + total, {
      x:8.5, y:5.44, w:1.3, h:0.18, fontSize:7, fontFace:FONT, color:C.g2, align:'right', valign:'middle', margin:0
    });
  }
}

function logo(slide) {
  slide.addImage({ data: IMGS.logo, x:0.25, y:0.18, w:0.45, h:0.45 });
}

function sectionTag(slide, num, label) {
  const arabic = isArabicText(label);
  slide.addText(num + '  /  ' + label.toUpperCase(), {
    x:0.4, y:0.06, w:9.3, h:0.28, fontSize:7.5,
    fontFace: arabic ? FONT_AR : FONT,
    color:C.mint, bold:true, charSpacing: arabic ? 0 : 2,
    align: arabic ? 'left' : 'right', valign:'middle', margin:0, rtlMode: arabic
  });
}

function buildProposal(data, outputPath) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title = (data.event_name || 'Proposal') + ' | XPACT';

  const eventName = data.event_name || 'Event';
  const client    = data.client || '';
  const date      = data.date || '';
  const gen       = data.generated_sections || {};
  const content   = data.content || {};
  const fixed     = data.fixed_sections || {};
  const objectives = content.event_objectives || [];
  const objIntro   = content.objectives_intro || '';
  const isArabic   = isArabicText(gen.executive_summary || gen.understanding || objIntro);
  const TOTAL = 10;

  // Labels in Arabic or English
  const L = isArabic ? {
    technical_proposal: 'المقترح الفني',
    table_of_contents: 'فهرس المحتويات',
    event_brief: 'نبذة عن الفعالية',
    event_brief_objectives: 'نبذة عن الفعالية والأهداف',
    event_objectives: 'أهداف الفعالية',
    key_objectives: 'الأهداف الرئيسية',
    event_overview: 'نبذة عن الفعالية',
    project_mgmt: 'إدارة المشروع',
    project_mgmt_overview: 'نظرة عامة على إدارة المشروع',
    technical_production: 'الإنتاج الفني',
    our_approach: 'منهجيتنا',
    why_xpact: 'لماذا إكسباكت',
    about_xpact: 'عن إكسباكت',
    contact: 'تواصل معنا',
    project_label: 'المشروع',
    prepared_for: 'مقدم إلى',
    date_label: 'التاريخ',
    pre_event: 'ما قبل الفعالية',
    event_day: 'يوم الفعالية',
    post_event: 'ما بعد الفعالية',
    events_delivered: 'فعالية منجزة',
    attendees: 'حاضر وصلنا إليهم',
    satisfaction: 'رضا العملاء',
    get_in_touch: 'تواصل معنا',
    vision: 'We look forward to bringing your vision to life.',
    why_title: 'لماذا تختار إكسباكت',
    about_title: 'عن إكسباكت',
  } : {
    technical_proposal: 'TECHNICAL PROPOSAL',
    table_of_contents: 'TABLE OF CONTENTS',
    event_brief: 'Event Brief',
    event_brief_objectives: 'Event Brief & Objectives',
    event_objectives: 'Event Objectives',
    key_objectives: 'Key Objectives',
    event_overview: 'Event Overview',
    project_mgmt: 'Project Management',
    project_mgmt_overview: 'Project Management Overview',
    technical_production: 'Technical Production & AV',
    our_approach: 'Our Approach',
    why_xpact: 'Why Choose XPACT',
    about_xpact: 'About XPACT',
    contact: 'Get In Touch',
    project_label: 'PROJECT',
    prepared_for: 'PREPARED FOR',
    date_label: 'DATE',
    pre_event: 'Pre-Event',
    event_day: 'Event Day',
    post_event: 'Post-Event',
    events_delivered: 'Events Delivered',
    attendees: 'Attendees Reached',
    satisfaction: 'Client Satisfaction',
    get_in_touch: 'GET IN TOUCH',
    vision: 'We look forward to bringing your vision to life.',
    why_title: 'Why Choose XPACT',
    about_title: 'ABOUT XPACT',
  };

  // ── SLIDE 1: COVER ──
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x:0, y:0, w:5.5, h:5.625, fill:{color:C.dark}, line:{color:C.dark} });
    s.addImage({ data: IMGS.cover, x:5.5, y:0, w:4.5, h:5.625, sizing:{type:'cover',w:4.5,h:5.625} });
    s.addShape('rect', { x:0, y:5.45, w:5.5, h:0.05, fill:{color:C.mint}, line:{color:C.mint} });
    s.addImage({ data: IMGS.logo, x:0.35, y:0.25, w:0.5, h:0.5 });

    const titleLine1 = isArabic ? 'المقترح' : 'TECHNICAL';
    const titleLine2 = isArabic ? 'الفني' : 'PROPOSAL';
    s.addText(titleLine1, { x:0.35, y:1.2, w:4.8, h:0.45, fontSize:28, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, charSpacing:isArabic?0:4, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addText(titleLine2, { x:0.35, y:1.65, w:4.8, h:0.45, fontSize:28, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.mint, charSpacing:isArabic?0:4, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:0.35, y:2.25, w:4.8, h:0.04, fill:{color:C.g3}, line:{color:C.g3} });

    s.addText(L.project_label, { x:0.35, y:2.42, w:4.8, h:0.18, fontSize:7, fontFace:FONT, color:C.mint, bold:true, charSpacing:2, margin:0 });
    s.addText(eventName, { x:0.35, y:2.6, w:4.8, h:0.42, fontSize:18, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addText(L.prepared_for, { x:0.35, y:3.15, w:4.8, h:0.18, fontSize:7, fontFace:FONT, color:C.mint, bold:true, charSpacing:2, margin:0 });
    s.addText(client || '', { x:0.35, y:3.33, w:4.8, h:0.35, fontSize:16, fontFace:isArabic?FONT_AR:FONT, color:C.white, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addText(L.date_label, { x:0.35, y:3.82, w:4.8, h:0.18, fontSize:7, fontFace:FONT, color:C.mint, bold:true, charSpacing:2, margin:0 });
    s.addText(date || '', { x:0.35, y:4.0, w:4.8, h:0.28, fontSize:13, fontFace:FONT, color:C.white, margin:0 });
    s.addText('Events Management & Advisory  |  xpact.net', { x:0.35, y:5.28, w:5, h:0.18, fontSize:7.5, fontFace:FONT, color:C.g2, margin:0 });
  }

  // ── SLIDE 2: TABLE OF CONTENTS ──
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    s.addShape('rect', { x:0, y:0, w:3.2, h:5.625, fill:{color:C.dark}, line:{color:C.dark} });
    s.addImage({ data: IMGS.logo, x:0.3, y:0.25, w:0.5, h:0.5 });
    s.addText(L.table_of_contents.replace(' ', '\n'), { x:0.3, y:2.2, w:2.6, h:0.9, fontSize:20, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, margin:0, rtlMode:isArabic });
    s.addShape('rect', { x:0.3, y:3.18, w:1.2, h:0.05, fill:{color:C.mint}, line:{color:C.mint} });

    const items = [
      {n:'01', t: L.event_brief_objectives},
      {n:'02', t: L.project_mgmt},
      {n:'03', t: L.technical_production},
      {n:'04', t: L.our_approach},
      {n:'05', t: L.why_xpact},
      {n:'06', t: L.about_xpact},
      {n:'07', t: L.contact},
    ];
    items.forEach((item, i) => {
      const y = 0.55 + i * 0.7;
      const isFirst = i === 0;
      s.addShape('ellipse', { x:3.5, y:y+0.08, w:0.38, h:0.38, fill:{color:isFirst?C.mint:C.g1}, line:{color:isFirst?C.mint:C.g1} });
      s.addText(item.n, { x:3.5, y:y+0.08, w:0.38, h:0.38, fontSize:9, fontFace:FONT, bold:true, color:isFirst?C.white:'555555', align:'center', valign:'middle', margin:0 });
      s.addText(item.t, { x:4.1, y:y+0.06, w:5.5, h:0.42, fontSize:13, fontFace:isArabic?FONT_AR:FONT, bold:isFirst, color:isFirst?C.dark:'777777', valign:'middle', margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
      if (i < items.length-1) s.addShape('line', { x:3.5, y:y+0.5, w:6.1, h:0, line:{color:C.g1, width:0.5} });
    });
    footer(s, 2, TOTAL);
  }

  // ── SLIDE 3: EVENT BRIEF ──
  {
    const s = pres.addSlide();
    s.background = { color: C.light };
    s.addShape('rect', { x:0, y:0, w:10, h:0.72, fill:{color:C.dark}, line:{color:C.dark} });
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '01', L.event_brief_objectives);

    s.addText(L.event_brief, { x:0.4, y:0.9, w:9.2, h:0.35, fontSize:20, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:0.4, y:1.3, w:0.06, h:3.8, fill:{color:C.mint}, line:{color:C.mint} });

    // Left: overview
    s.addShape('rect', { x:0.6, y:1.35, w:4.3, h:3.7, fill:{color:C.white}, line:{color:C.g1,width:0.5} });
    s.addText(L.event_overview, { x:0.75, y:1.55, w:4.0, h:0.3, fontSize:11, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:0.75, y:1.87, w:4.0, h:0.04, fill:{color:C.mint}, line:{color:C.mint} });

    const overviewText = gen.understanding || gen.executive_summary || '';
    s.addText(overviewText.slice(0,500), { x:0.75, y:1.98, w:4.0, h:2.9, fontSize:10, fontFace:isArabic?FONT_AR:FONT, color:'444444', valign:'top', margin:0, lineSpacingMultiple:1.5, rtlMode:isArabic, align:isArabic?'right':'left' });

    // Right: objectives
    s.addShape('rect', { x:5.1, y:1.35, w:4.5, h:3.7, fill:{color:C.dark}, line:{color:C.dark} });
    s.addText(L.key_objectives, { x:5.3, y:1.55, w:4.1, h:0.3, fontSize:11, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.mint, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:5.3, y:1.87, w:4.1, h:0.04, fill:{color:C.teal}, line:{color:C.teal} });

    objectives.slice(0,5).forEach((obj, i) => {
      const title = typeof obj === 'object' ? obj.title : obj;
      const y = 2.05 + i * 0.6;
      s.addImage({ data: IMGS.check, x:isArabic?8.8:5.3, y:y+0.03, w:0.28, h:0.28 });
      s.addText(title, { x:5.72, y:y, w:isArabic?3.0:3.7, h:0.42, fontSize:9.5, fontFace:isArabic?FONT_AR:FONT, color:C.white, valign:'middle', margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    });

    footer(s, 3, TOTAL);
  }

  // ── SLIDE 4: PROJECT MANAGEMENT ──
  {
    const s = pres.addSlide();
    s.background = { color: C.light };
    s.addShape('rect', { x:0, y:0, w:10, h:0.72, fill:{color:C.dark}, line:{color:C.dark} });
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '02', L.project_mgmt);

    s.addText(L.project_mgmt_overview, { x:0.4, y:0.9, w:9.2, h:0.35, fontSize:20, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:0.4, y:1.3, w:0.06, h:0.8, fill:{color:C.mint}, line:{color:C.mint} });

    const pmText = content.pm_overview_para1 || gen.understanding || '';
    s.addText(pmText.slice(0,350), { x:0.6, y:1.38, w:9.0, h:1.1, fontSize:10.5, fontFace:isArabic?FONT_AR:FONT, color:'333333', valign:'top', margin:0, lineSpacingMultiple:1.6, rtlMode:isArabic, align:isArabic?'right':'left' });

    const phases = isArabic ? [
      {label:L.pre_event,  items:['انطلاقة المشروع وإحاطة الفريق','تطوير المفهوم الإبداعي','تنسيق الموقع والتصاريح','توفير المورّدين وإحاطتهم','الامتثال للوائح SCEGA']},
      {label:L.event_day,  items:['الإدارة الكاملة في الموقع','الإشراف على الإنتاج الفني','إدارة الضيوف وكبار الشخصيات','حل المشكلات الطارئة فورياً','تنسيق الإعلام والصحافة']},
      {label:L.post_event, items:['تفكيك المنشآت وتسليم الموقع','المصالحة المالية','تسليم تقرير ما بعد الفعالية','جمع آراء الحضور','توثيق الدروس المستفادة']},
    ] : [
      {label:L.pre_event,  items:['Project kick-off & team briefing','Creative concept development','Venue coordination & permits','Vendor procurement & briefing','SCEGA compliance & licensing']},
      {label:L.event_day,  items:['Full on-site management','AV & production supervision','Guest & VIP management','Real-time issue resolution','Media & press coordination']},
      {label:L.post_event, items:['Venue dismantling & handover','Financial reconciliation','Post-event report delivery','Attendee feedback collection','Lessons learned documentation']},
    ];

    phases.forEach((ph, i) => {
      const x = 0.55 + i * 3.15;
      s.addShape('rect', { x, y:2.75, w:2.95, h:2.65, fill:{color:C.white}, line:{color:C.g1,width:0.5} });
      s.addShape('rect', { x, y:2.75, w:2.95, h:0.42, fill:{color:i===1?C.mint:C.teal}, line:{color:i===1?C.mint:C.teal} });
      s.addText(ph.label, { x, y:2.75, w:2.95, h:0.42, fontSize:10, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, align:'center', valign:'middle', margin:0 });
      ph.items.forEach((item, j) => {
        s.addText([{text:item, options:{bullet:true}}], { x:x+0.1, y:3.23+j*0.4, w:2.75, h:0.36, fontSize:8.5, fontFace:isArabic?FONT_AR:FONT, color:'444444', margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
      });
    });
    footer(s, 4, TOTAL);
  }

  // ── SLIDE 5: OBJECTIVES DETAIL ──
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '01', L.event_objectives);

    s.addText(L.event_objectives, { x:0.4, y:0.8, w:9.2, h:0.38, fontSize:22, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    if (objIntro) {
      s.addText(objIntro.slice(0,200), { x:0.4, y:1.25, w:9.2, h:0.4, fontSize:10, fontFace:isArabic?FONT_AR:FONT, color:C.g2, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    }
    s.addShape('rect', { x:0.4, y:1.7, w:9.2, h:0.04, fill:{color:C.teal}, line:{color:C.teal} });

    objectives.slice(0,5).forEach((obj, i) => {
      const title = typeof obj === 'object' ? obj.title : obj;
      const row = Math.floor(i/3), col = i%3;
      const cardW = 3.0, gap = (9.2 - 3*cardW)/4;
      const x = 0.4 + gap + col*(cardW+gap);
      const y = 1.85 + row*1.7;
      s.addShape('rect', { x, y, w:cardW, h:1.55, fill:{color:C.teal}, line:{color:C.teal} });
      s.addShape('rect', { x, y, w:0.5, h:0.5, fill:{color:C.mint}, line:{color:C.mint} });
      s.addText(String(i+1).padStart(2,'0'), { x, y, w:0.5, h:0.5, fontSize:11, fontFace:FONT, bold:true, color:C.dark, align:'center', valign:'middle', margin:0 });
      s.addText(title, { x:x+0.12, y:y+0.55, w:cardW-0.25, h:0.95, fontSize:9.5, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, valign:'top', margin:0, lineSpacingMultiple:1.4, rtlMode:isArabic, align:isArabic?'right':'left' });
    });
    footer(s, 5, TOTAL);
  }

  // ── SLIDE 6: PRODUCTION ──
  {
    const s = pres.addSlide();
    s.background = { color: C.light };
    s.addShape('rect', { x:0, y:0, w:10, h:0.72, fill:{color:C.dark}, line:{color:C.dark} });
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '03', L.technical_production);

    s.addText(L.technical_production, { x:0.4, y:0.9, w:9.2, h:0.35, fontSize:20, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });

    const prodText = content.production_overview || gen.production_av || '';
    s.addText(prodText.slice(0,300), { x:0.4, y:1.35, w:9.2, h:0.7, fontSize:10, fontFace:isArabic?FONT_AR:FONT, color:'444444', margin:0, lineSpacingMultiple:1.5, rtlMode:isArabic, align:isArabic?'right':'left' });

    const services = isArabic ? [
      {title:'تصميم المسرح والديكور',    desc:'تصميم مسرح مخصص يتوافق مع هوية الفعالية وحجم الجمهور'},
      {title:'أنظمة الصوت',              desc:'نظام صوت احترافي يضمن وضوحاً تاماً في جميع أرجاء القاعة'},
      {title:'شاشات LED والعروض',        desc:'جدران LED رئيسية وشاشات ثقة وشاشات IMAG'},
      {title:'تصميم الإضاءة',            desc:'إضاءة محيطية وتأكيدية وديناميكية لإبراز الأجواء'},
      {title:'الإخراج الفني والتحكم',    desc:'مخرجون فنيون متخصصون وغرف تحكم طوال فترة الفعالية'},
      {title:'التوثيق والبث',            desc:'تصوير احترافي وفيديو وبث مباشر للفعالية'},
    ] : [
      {title:'Stage & Set Design',       desc:'Custom stage design aligned with event theme and audience size'},
      {title:'Audio Systems',            desc:'Professional line array speakers for crystal clear sound'},
      {title:'LED & Screens',            desc:'Main LED walls, confidence monitors, and IMAG screens'},
      {title:'Lighting Design',          desc:'Ambient, accent, and dynamic lighting for atmosphere'},
      {title:'AV Control & Direction',   desc:'Dedicated technical directors and switchers throughout'},
      {title:'Documentation & Broadcast',desc:'Professional photography, videography, and live streaming'},
    ];

    services.forEach((svc, i) => {
      const col = i%2, row = Math.floor(i/2);
      const x = 0.4 + col*4.85;
      const y = 2.2 + row*1.05;
      s.addShape('rect', { x, y, w:4.6, h:0.92, fill:{color:C.white}, line:{color:C.g1,width:0.5} });
      s.addShape('rect', { x, y, w:0.08, h:0.92, fill:{color:C.mint}, line:{color:C.mint} });
      s.addText(svc.title, { x:x+0.2, y:y+0.08, w:4.2, h:0.28, fontSize:10, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
      s.addText(svc.desc, { x:x+0.2, y:y+0.38, w:4.2, h:0.46, fontSize:8.5, fontFace:isArabic?FONT_AR:FONT, color:'555555', margin:0, lineSpacingMultiple:1.3, rtlMode:isArabic, align:isArabic?'right':'left' });
    });
    footer(s, 6, TOTAL);
  }

  // ── SLIDE 7: OUR APPROACH ──
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '04', L.our_approach);

    s.addText(L.our_approach, { x:0.4, y:0.8, w:9.2, h:0.38, fontSize:22, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });

    const steps = isArabic ? [
      {n:'01', title:'الاستكشاف\nوالإحاطة',    time:'أسبوع 1-2'},
      {n:'02', title:'الاستراتيجية\nوالتخطيط', time:'أسبوع 2-4'},
      {n:'03', title:'التصميم\nوالإنتاج',      time:'أسبوع 4-8'},
      {n:'04', title:'التنفيذ\nوإدارة الموقع', time:'يوم الفعالية'},
      {n:'05', title:'مراجعة\nما بعد الفعالية',time:'أسبوع +1'},
    ] : [
      {n:'01', title:'Discovery\n& Briefing',   time:'Week 1-2'},
      {n:'02', title:'Strategy\n& Planning',    time:'Week 2-4'},
      {n:'03', title:'Design &\nProduction',    time:'Week 4-8'},
      {n:'04', title:'Execution\n& On-site',    time:'Event Day'},
      {n:'05', title:'Post-Event\nReview',      time:'Week +1'},
    ];

    s.addShape('line', { x:0.9, y:2.8, w:8.2, h:0, line:{color:C.teal,width:1.5} });
    steps.forEach((step, i) => {
      const x = 0.5 + i*2.0;
      const isActive = i===2;
      s.addShape('ellipse', { x, y:2.5, w:0.6, h:0.6, fill:{color:isActive?C.mint:C.teal}, line:{color:isActive?C.mint:C.teal} });
      s.addText(step.n, { x, y:2.5, w:0.6, h:0.6, fontSize:11, fontFace:FONT, bold:true, color:C.white, align:'center', valign:'middle', margin:0 });
      s.addText(step.title, { x:x-0.5, y:3.25, w:1.6, h:0.7, fontSize:9.5, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, align:'center', margin:0 });
      s.addText(step.time, { x:x-0.5, y:1.75, w:1.6, h:0.35, fontSize:8, fontFace:isArabic?FONT_AR:FONT, color:C.mint, align:'center', margin:0 });
    });

    const diffs = isArabic ? [
      {t:'سجل حافل بالإنجازات',  d:'أكثر من 20 فعالية رفيعة المستوى في المملكة العربية السعودية'},
      {t:'خبرة شاملة من الألف إلى الياء', d:'فريق واحد. ستة محاور خدمية. صفر من فجوات التنسيق.'},
      {t:'قاعدة سعودية ومرونة إقليمية',  d:'معرفة عميقة بالسوق المحلي مع معايير تنفيذ دولية'},
      {t:'ثقافة تضع العميل أولاً',        d:'كل مقترح وخطة تنفيذ مصمّمة خصيصاً لأهدافك'},
    ] : [
      {t:'Proven Track Record', d:'20+ high-profile events delivered across Saudi Arabia'},
      {t:'End-to-End Expertise', d:'One team. Six service pillars. Zero coordination gaps.'},
      {t:'KSA-Based & Agile',    d:'Deep local knowledge with international execution standards'},
      {t:'Client-First Culture', d:'Every proposal and plan tailored to your specific objectives'},
    ];

    diffs.forEach((d, i) => {
      const x = 0.4 + (i%2)*4.85;
      const y = 4.15 + Math.floor(i/2)*0.65;
      s.addShape('rect', { x, y, w:4.6, h:0.55, fill:{color:C.card}, line:{color:C.g3} });
      s.addText(d.t, { x:x+0.15, y:y+0.04, w:2.8, h:0.25, fontSize:9, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.mint, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
      s.addText(d.d, { x:x+0.15, y:y+0.28, w:4.3, h:0.22, fontSize:8, fontFace:isArabic?FONT_AR:FONT, color:C.g2, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    });
    footer(s, 7, TOTAL);
  }

  // ── SLIDE 8: WHY XPACT ──
  {
    const s = pres.addSlide();
    s.background = { color: C.light };
    s.addShape('rect', { x:0, y:0, w:10, h:0.72, fill:{color:C.dark}, line:{color:C.dark} });
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    logo(s);
    sectionTag(s, '05', L.why_xpact);

    s.addText(L.why_title, { x:0.4, y:0.9, w:6, h:0.35, fontSize:20, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });

    const whyText = gen.why_xpact || '';
    const whyLines = whyText.split('\n').filter(l=>l.trim().startsWith('-')).slice(0,4);
    const points = whyLines.length >= 4 ? whyLines.map(l=>({ t:l.replace(/^-\s*/,'').split(':')[0]||l.replace(/^-\s*/,'').slice(0,50), d:l.replace(/^-\s*/,'') })) : (isArabic ? [
      {t:'سجل حافل بالإنجازات',        d:'أكثر من 20 فعالية رفيعة المستوى في المملكة لجهات حكومية وقطاع خاص.'},
      {t:'خبرة شاملة من الألف إلى الياء',d:'فريق واحد. ستة محاور خدمية. صفر من فجوات التنسيق.'},
      {t:'قاعدة سعودية ومرونة إقليمية', d:'معرفة محلية عميقة مع معايير تنفيذ دولية.'},
      {t:'ثقافة تضع العميل أولاً',       d:'كل مقترح وخطة تنفيذ مصمّمة خصيصاً لأهدافك.'},
    ] : [
      {t:'Proven Track Record',    d:'20+ high-profile events across Saudi Arabia for government entities and private sector leaders.'},
      {t:'End-to-End Expertise',   d:'From strategy through post-event review - one team, zero coordination gaps.'},
      {t:'KSA-Based & Agile',      d:'Deep local knowledge combined with international best practices.'},
      {t:'Client-First Culture',   d:'Every proposal and execution plan is tailored to your specific objectives.'},
    ]);

    points.slice(0,4).forEach((pt, i) => {
      const x = 0.4 + (i%2)*3.4;
      const y = 1.45 + Math.floor(i/2)*1.6;
      s.addShape('rect', { x, y, w:3.15, h:1.45, fill:{color:C.white}, line:{color:C.g1,width:0.5} });
      s.addShape('rect', { x, y, w:3.15, h:0.08, fill:{color:i<2?C.mint:C.teal}, line:{color:i<2?C.mint:C.teal} });
      s.addImage({ data: IMGS.check, x:x+0.15, y:y+0.2, w:0.35, h:0.35 });
      s.addText(pt.t, { x:x+0.62, y:y+0.18, w:2.4, h:0.38, fontSize:10, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.dark, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
      s.addText(pt.d.slice(0,120), { x:x+0.15, y:y+0.65, w:2.85, h:0.72, fontSize:8.5, fontFace:isArabic?FONT_AR:FONT, color:'555555', margin:0, lineSpacingMultiple:1.4, rtlMode:isArabic, align:isArabic?'right':'left' });
    });

    s.addShape('rect', { x:6.95, y:1.45, w:2.65, h:3.65, fill:{color:C.dark}, line:{color:C.dark} });
    const stats = [
      {v:'20+',    l:L.events_delivered},
      {v:'5,000+', l:L.attendees},
      {v:'100%',   l:L.satisfaction},
    ];
    stats.forEach((st, i) => {
      s.addText(st.v, { x:6.95, y:1.65+i*1.1, w:2.65, h:0.6, fontSize:28, fontFace:FONT, bold:true, color:C.mint, align:'center', margin:0 });
      s.addText(st.l, { x:6.95, y:2.28+i*1.1, w:2.65, h:0.3, fontSize:8.5, fontFace:isArabic?FONT_AR:FONT, color:C.g2, align:'center', margin:0 });
      if(i<2) s.addShape('rect', { x:7.3, y:2.62+i*1.1, w:1.95, h:0.03, fill:{color:C.g3}, line:{color:C.g3} });
    });
    footer(s, 8, TOTAL);
  }

  // ── SLIDE 9: ABOUT XPACT ──
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addImage({ data: IMGS.about, x:0, y:0, w:5.5, h:5.625, sizing:{type:'cover',w:5.5,h:5.625} });
    s.addShape('rect', { x:0, y:0, w:5.5, h:5.625, fill:{color:'000000', transparency:50}, line:{color:'000000', transparency:50} });
    s.addShape('rect', { x:5.5, y:0, w:4.5, h:5.625, fill:{color:C.dark}, line:{color:C.dark} });
    s.addShape('rect', { x:5.5, y:0, w:4.5, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    s.addImage({ data: IMGS.logo, x:5.7, y:0.2, w:0.5, h:0.5 });

    s.addText(L.about_title, { x:5.7, y:0.85, w:4.0, h:0.75, fontSize:22, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:5.7, y:1.68, w:3.6, h:0.04, fill:{color:C.mint}, line:{color:C.mint} });

    const aboutText = fixed.about_us || (isArabic
      ? 'إكسباكت شركة متخصصة في إدارة الفعاليات والاستشارات، تسعى إلى تقديم تجارب استثنائية ومصمّمة بعناية. نجمع بين التفكير الاستراتيجي والتميز الإبداعي والتنفيذ السلس لخدمة عملائنا في المملكة العربية السعودية والمنطقة.'
      : 'XPACT is an events management and advisory company dedicated to delivering exceptional, tailor-made event experiences. We bring together strategic thinking, creative excellence, and seamless execution to serve clients across Saudi Arabia and the region.');

    s.addText(aboutText.slice(0,350), { x:5.7, y:1.8, w:3.9, h:2.0, fontSize:9.5, fontFace:isArabic?FONT_AR:FONT, color:C.g1, margin:0, lineSpacingMultiple:1.6, rtlMode:isArabic, align:isArabic?'right':'left' });

    const stats = [
      {v:'20+',    l:L.events_delivered},
      {v:'5,000+', l:L.attendees},
      {v:'100%',   l:L.satisfaction},
    ];
    stats.forEach((st, i) => {
      const x = 5.65 + i*1.45;
      s.addShape('rect', { x, y:4.0, w:1.3, h:1.05, fill:{color:C.teal}, line:{color:C.teal} });
      s.addText(st.v, { x, y:4.08, w:1.3, h:0.48, fontSize:18, fontFace:FONT, bold:true, color:C.mint, align:'center', margin:0 });
      s.addText(st.l, { x, y:4.58, w:1.3, h:0.38, fontSize:7.5, fontFace:isArabic?FONT_AR:FONT, color:C.white, align:'center', margin:0 });
    });
    footer(s, 9, TOTAL);
  }

  // ── SLIDE 10: CONTACT ──
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x:0, y:0, w:10, h:0.06, fill:{color:C.mint}, line:{color:C.mint} });
    s.addImage({ data: IMGS.logo, x:0.35, y:0.25, w:0.55, h:0.55 });

    s.addText('07', { x:0.35, y:1.5, w:1, h:0.45, fontSize:14, fontFace:FONT, color:C.mint, bold:true, margin:0 });
    s.addText(L.get_in_touch, { x:0.35, y:2.0, w:5, h:0.55, fontSize:24, fontFace:isArabic?FONT_AR:FONT, bold:true, color:C.white, charSpacing:isArabic?0:2, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addText(L.vision, { x:0.35, y:2.65, w:5.5, h:0.42, fontSize:12, fontFace:isArabic?FONT_AR:FONT, color:C.g2, margin:0, rtlMode:isArabic, align:isArabic?'right':'left' });
    s.addShape('rect', { x:0.35, y:3.2, w:3.5, h:0.04, fill:{color:C.mint}, line:{color:C.mint} });

    const contacts = ['+966 53 587 9603','e.matar[at]xpact.net','www.xpact.net','Al Wizarat Dist., Riyadh, KSA'];
    contacts.forEach((c, i) => {
      s.addShape('ellipse', { x:0.35, y:3.45+i*0.52, w:0.32, h:0.32, fill:{color:C.mint}, line:{color:C.mint} });
      s.addText(c, { x:0.82, y:3.45+i*0.52, w:5, h:0.32, fontSize:11, fontFace:FONT, color:C.white, valign:'middle', margin:0 });
    });

    s.addText('Experts Impact Company  |  CR: 7050428643  |  VAT: 313076658400003', { x:0.35, y:5.2, w:9.3, h:0.22, fontSize:7.5, fontFace:FONT, color:'555555', margin:0 });
    s.addShape('rect', { x:6, y:0.5, w:3.5, h:4.7, fill:{color:C.teal}, line:{color:C.teal} });
    s.addText('XPACT', { x:6, y:2.2, w:3.5, h:0.6, fontSize:36, fontFace:FONT, bold:true, color:C.white, align:'center', charSpacing:8, margin:0 });
    s.addText('Events Management\n& Advisory', { x:6, y:2.95, w:3.5, h:0.65, fontSize:11, fontFace:FONT, color:'9FCFDB', align:'center', margin:0 });
    s.addText(eventName, { x:6, y:3.75, w:3.5, h:0.55, fontSize:10, fontFace:isArabic?FONT_AR:FONT, color:C.mint, align:'center', bold:true, margin:0, rtlMode:isArabic });
  }

  pres.writeFile({ fileName: outputPath });
  console.log('Done: ' + outputPath);
}


module.exports = { buildProposal };
