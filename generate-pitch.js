/**
 * Weaver Fintech — Generic Pitch Deck
 * ─────────────────────────────────────
 * Run:  node setup.js && node generate-pitch.js
 *
 * Output: Weaver_Pitch.pptx
 *
 * ⚠️  Before running:
 *   1. Replace YOUR_USERNAME below with your GitHub username
 *   2. Ensure the repo is public (or update GH_TOKEN for private repos)
 *   3. Run `node setup.js` once per session to write logos to disk
 */

const PptxGenJS = require('pptxgenjs');
const pptx      = new PptxGenJS();
pptx.layout     = 'LAYOUT_WIDE'; // 13.33" × 7.5"

// ── BRAND ────────────────────────────────────────────────────────────────────
const C = {
  purple:'875EFA', lime:'C4FF00', green:'1DCD86', blue:'2A62F5',
  black:'1A1A1A', white:'FFFFFF', offwhite:'F4F4F4', grey:'DDDDDD', insightbg:'EFEFEF',
};
const TF = 'Aptos Display';
const BF = 'Aptos';

// ── ASSETS ───────────────────────────────────────────────────────────────────
const LOGO      = '/home/claude/assets/image1.svg'; // written by setup.js
const LOGO_MARK = '/home/claude/assets/image2.svg'; // written by setup.js

// ⚠️ Replace YOUR_USERNAME with your GitHub username
const GH = 'https://raw.githubusercontent.com/TheLeroyMeyer/weaver-fintech-skill/main/assets/';

const PHOTOS = {
  title: { path: GH+'image5.png',  w: 9.91 },
  div1:  { path: GH+'image6.png',  w: 6.89 },
  div2:  { path: GH+'image7.png',  w: 6.89 },
  div3:  { path: GH+'image8.png',  w: 7.99 },
  div4:  { path: GH+'image9.png',  w: 6.88 },
  div5:  { path: GH+'image10.png', w: 7.38 },
};

// ── HELPERS ──────────────────────────────────────────────────────────────────

function logoMark(s) {
  s.addImage({ path: LOGO_MARK, x: 12.55, y: 0.08, w: 0.62, h: 0.62 });
}

function slideNum(s, n) {
  s.addText(String(n), {
    x: 12.8, y: 7.25, w: 0.4, h: 0.2,
    fontSize: 9, bold: true, color: C.black, align: 'right', fontFace: BF
  });
}

function slideTitle(s, title, subtitle) {
  s.addText(title, {
    x: 0.35, y: 0.08, w: 11.7, h: subtitle ? 0.44 : 0.52,
    fontSize: 20, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.35, y: 0.54, w: 11.7, h: 0.26,
      fontSize: 11, italic: true, color: C.black, fontFace: BF
    });
  }
}

function chartBox(s, x, y, w, h, title) {
  s.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: C.offwhite }, line: { color: C.grey, pt: 1 }
  });
  s.addText(title, {
    x: x+0.14, y: y+0.1, w: w-0.28, h: 0.27,
    fontSize: 9, bold: true, color: C.black, fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: x+0.14, y: y+0.38, w: w-0.28, h: 0.012,
    fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });
}

function insightBar(s, headline, impl, watch, action) {
  const Y = 5.72, H = 1.78;
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: Y, w: 13.33, h: H, fill: { color: C.insightbg }, line: { color: C.insightbg }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: Y, w: 0.07, h: H, fill: { color: C.purple }, line: { color: C.purple }
  });
  s.addText('INSIGHT', {
    x: 0.2, y: Y+0.1, w: 2, h: 0.22,
    fontSize: 8, bold: true, color: C.purple, fontFace: BF, charSpacing: 2
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.2, y: Y+0.33, w: 13.0, h: 0.012,
    fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });
  s.addText(headline, {
    x: 0.2, y: Y+0.38, w: 13.0, h: 0.42,
    fontSize: 10.5, bold: true, color: C.black, fontFace: BF, wrap: true
  });
  [
    { label: 'Strategic implication', body: impl  },
    { label: 'Risk or watch point',   body: watch  },
    { label: 'Recommended action',    body: action },
  ].forEach(({ label, body }, i) => {
    const cx = 0.2 + i * 4.35;
    s.addText(label, {
      x: cx, y: Y+0.88, w: 4.1, h: 0.2,
      fontSize: 8.5, bold: true, color: C.black, fontFace: BF
    });
    s.addText(body, {
      x: cx, y: Y+1.1, w: 4.1, h: 0.58,
      fontSize: 7.5, color: C.black, fontFace: BF, wrap: true
    });
  });
}

// ── SLIDE 1: TITLE ───────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.addImage({ path: LOGO, x: 0.35, y: 0.28, w: 2.5, h: 1.1 });
  s.addText('Powering the\nFuture of\nFinancial Access.', {
    x: 0.35, y: 2.0, w: 7.2, h: 2.9,
    fontSize: 44, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  s.addText('Company Overview & Investment Proposition', {
    x: 0.35, y: 5.1, w: 7.0, h: 0.38,
    fontSize: 15, bold: true, color: C.black, fontFace: BF
  });
  s.addText('June 2026', {
    x: 0.35, y: 5.58, w: 4, h: 0.3, fontSize: 13, color: C.black, fontFace: BF
  });
  s.addText('weaverfintech.com', {
    x: 0.35, y: 5.95, w: 4, h: 0.28, fontSize: 12, color: C.black, fontFace: BF
  });
  const p = PHOTOS.title;
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
}

// ── SLIDE 2: DIVIDER — Who We Are ────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  s.addText('Who\nWe Are.', {
    x: 0.35, y: 2.0, w: 5.8, h: 3.2,
    fontSize: 52, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  const p = PHOTOS.div1;
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
  slideNum(s, 2);
}

// ── SLIDE 3: COMPANY OVERVIEW ─────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s,
    'Weaver Fintech is redefining access to financial services across emerging markets.',
    'Company Overview  |  2026'
  );

  // Stat cards
  [
    { num: '2018',  lbl: 'Founded\nJohannesburg, ZA',    color: C.purple  },
    { num: '1.2M',  lbl: 'Active customers\nand growing', color: '1A1A1A' },
    { num: 'R2.4B', lbl: 'Loans originated\nto date',     color: C.purple  },
    { num: '84%',   lbl: 'Digitally acquired\ncustomers',  color: '1A1A1A' },
    { num: '24%',   lbl: 'Revenue growth\nyear-on-year',   color: C.purple  },
  ].forEach((st, i) => {
    const xx = 0.2 + i * 2.6;
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: xx, y: 0.95, w: 2.48, h: 2.0,
      fill: { color: st.color }, line: { color: st.color }, rectRadius: 0.1
    });
    s.addText(st.num, {
      x: xx+0.12, y: 1.05, w: 2.24, h: 0.9,
      fontSize: 30, bold: true,
      color: st.color === C.purple ? C.white : C.purple,
      fontFace: TF, align: 'center'
    });
    s.addText(st.lbl, {
      x: xx+0.1, y: 1.98, w: 2.28, h: 0.8,
      fontSize: 8.5, color: C.white, fontFace: BF, align: 'center', wrap: true
    });
  });

  // Evidence list
  [
    'Licensed and regulated — fully compliant with FSCA, NCA, and POPIA across all markets.',
    'Digital-first acquisition model delivers customers at R18 CPA vs R108 at branch — a 6× efficiency advantage.',
    'Proprietary credit scoring engine trained on 6M+ data points, enabling responsible lending to underserved segments.',
    'Omnichannel servicing via app, USSD, WhatsApp, and branch — meeting customers where they are.',
    'Profitable and capital-efficient: 9.1× LTV:CAC ratio, well above the 4–5× global benchmark.',
  ].forEach((txt, i) => {
    const yy = 3.2 + i * 0.48;
    s.addShape(pptx.shapes.OVAL, {
      x: 0.2, y: yy+0.04, w: 0.3, h: 0.3,
      fill: { color: '1A1A1A' }, line: { color: '1A1A1A' }
    });
    s.addText(String(i+1).padStart(2,'0'), {
      x: 0.2, y: yy+0.04, w: 0.3, h: 0.3,
      fontSize: 7, bold: true, color: C.white, align: 'center', fontFace: BF
    });
    s.addText(txt, {
      x: 0.65, y: yy, w: 12.5, h: 0.42,
      fontSize: 9, color: C.black, fontFace: BF, wrap: true
    });
  });

  insightBar(s,
    'Weaver has already proven product-market fit at scale — the question is no longer if, but how fast.',
    'A proven model with a 6× cost advantage and profitable unit economics is ready to absorb capital and grow.',
    'Execution risk is the primary variable — leadership depth and operational infrastructure must scale with growth.',
    'Deploy capital into the three highest-leverage channels: digital acquisition, product expansion, and geographic entry.'
  );
  slideNum(s, 3);
}

// ── SLIDE 4: DIVIDER — What We Do ────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  s.addText('What\nWe Do.', {
    x: 0.35, y: 2.0, w: 5.8, h: 3.2,
    fontSize: 52, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  const p = PHOTOS.div2;
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
  slideNum(s, 4);
}

// ── SLIDE 5: PRODUCTS & CHANNELS ─────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s,
    'A full-stack lending platform: origination, servicing, collections, and reinvestment — all in one.',
    'Products & Channels  |  2026'
  );

  // Left chart — product mix
  chartBox(s, 0.2, 0.88, 6.4, 4.72, 'Revenue by Product Line (%)');
  s.addChart(pptx.charts.DOUGHNUT, [{
    name: 'Revenue',
    labels: ['Personal Loans', 'Buy Now Pay Later', 'Insurance', 'Savings', 'Other'],
    values: [52, 28, 11, 6, 3]
  }], {
    x: 0.25, y: 1.32, w: 6.3, h: 4.22,
    chartColors: [C.purple, '1A1A1A', C.blue, C.green, C.grey],
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    showPercent: true, dataLabelFontSize: 9,
    showTitle: false, holeSize: 55
  });

  // Right chart — channel mix
  chartBox(s, 6.73, 0.88, 6.4, 4.72, 'Customer Acquisition by Channel (%)');
  s.addChart(pptx.charts.BAR, [{
    name: 'FY22', labels: ['App', 'USSD', 'WhatsApp', 'Branch', 'Partner'], values: [31, 20, 8, 35, 6]
  }, {
    name: 'FY26', labels: ['App', 'USSD', 'WhatsApp', 'Branch', 'Partner'], values: [54, 16, 14, 11, 5]
  }], {
    x: 6.78, y: 1.32, w: 6.3, h: 4.22,
    barGrouping: 'clustered',
    barDir: 'bar',
    chartColors: ['DDDDDD', C.purple],
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 9, showTitle: false,
    plotAreaFill: { color: C.offwhite },
  });

  insightBar(s,
    'The shift from branch to digital acquisition is the single biggest margin driver in the business.',
    'Every 10% shift from branch to digital drops blended CPA by ~R9, adding directly to net margin at scale.',
    'Digital channel mix is approaching saturation in current segments — new segments (youth, rural) must be activated.',
    'Accelerate USSD and WhatsApp origination in rural corridors where smartphone penetration lags but need is highest.'
  );
  slideNum(s, 5);
}

// ── SLIDE 6: DIVIDER — Why Weaver ────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  s.addText('Why\nWeaver.', {
    x: 0.35, y: 2.0, w: 5.8, h: 3.2,
    fontSize: 52, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  const p = PHOTOS.div3;
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
  slideNum(s, 6);
}

// ── SLIDE 7: COMPETITIVE ADVANTAGE ───────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s,
    'Three structural advantages that are durable, compounding, and difficult to replicate.',
    'Competitive Position  |  2026'
  );

  const yrs = ['FY22','FY23','FY24','FY25','FY26'];
  [
    { x: 0.2,  y: 0.88, w: 4.2, h: 4.72, title: 'LTV:CAC Ratio vs Peers',     vals: [5.1, 6.2, 7.4, 8.3, 9.1] },
    { x: 4.57, y: 0.88, w: 4.2, h: 4.72, title: 'Digital Acquisition Mix (%)', vals: [31, 48, 62, 75, 84] },
    { x: 8.93, y: 0.88, w: 4.2, h: 4.72, title: 'Revenue Growth YoY (%)',      vals: [9, 14, 18, 21, 24] },
  ].forEach(({ x, y, w, h, title, vals }) => {
    chartBox(s, x, y, w, h, title);
    s.addChart(pptx.charts.LINE, [{ name: title, labels: yrs, values: vals }], {
      x: x+0.05, y: y+0.44, w: w-0.1, h: h-0.5,
      chartColors: [C.purple],
      showLegend: false,
      showValue: true, dataLabelFontSize: 8,
      catAxisLabelFontSize: 8, showTitle: false,
      plotAreaFill: { color: C.offwhite },
      lineDataSymbol: 'circle',
    });
  });

  insightBar(s,
    'All three advantage metrics are accelerating — not plateauing — which means the moat is widening, not narrowing.',
    'A widening moat at this stage of scale is rare. It signals that technology and data advantages are compounding.',
    'Sustaining this trajectory requires continued R&D investment — any cost-cutting that reduces tech spend reverses the trend.',
    'Ring-fence the technology and data science budget as a non-negotiable capex line, not a discretionary opex item.'
  );
  slideNum(s, 7);
}

// ── SLIDE 8: RECOMMENDATION ──────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  logoMark(s);

  s.addText('Partner with Weaver Fintech to capture the R180B underserved lending opportunity across Southern Africa.', {
    x: 0.35, y: 0.05, w: 11.7, h: 0.55,
    fontSize: 16, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  s.addText('Investment Proposition  |  Strategic Partnership', {
    x: 0.35, y: 0.62, w: 11.7, h: 0.24,
    fontSize: 11, italic: true, color: C.black, fontFace: BF
  });

  // Evidence panel
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.2, y: 0.92, w: 6.0, h: 4.78,
    fill: { color: C.offwhite }, line: { color: C.grey, pt: 1 }
  });
  s.addText('The Case for Partnership', {
    x: 0.35, y: 1.0, w: 5.7, h: 0.28,
    fontSize: 10, bold: true, color: C.black, fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.35, y: 1.3, w: 5.7, h: 0.012,
    fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });

  [
    'R180B addressable market remains 60% underserved by traditional banks in Southern Africa alone.',
    'Weaver\'s 9.1× LTV:CAC ratio is more than double the global fintech benchmark of 4–5×.',
    'Regulatory standing is clean: fully licensed under FSCA, NCA, and POPIA with zero enforcement actions.',
    'Technology stack is proprietary and API-first — integration with partner systems is weeks, not months.',
    'Management team has 80+ years of combined financial services and technology leadership experience.',
  ].forEach((txt, i) => {
    const yy = 1.42 + i * 0.84;
    s.addShape(pptx.shapes.OVAL, {
      x: 0.35, y: yy+0.04, w: 0.34, h: 0.34,
      fill: { color: '1A1A1A' }, line: { color: '1A1A1A' }
    });
    s.addText(String(i+1).padStart(2,'0'), {
      x: 0.35, y: yy+0.04, w: 0.34, h: 0.34,
      fontSize: 7.5, bold: true, color: C.white, align: 'center', fontFace: BF
    });
    s.addText(txt, {
      x: 0.82, y: yy, w: 5.2, h: 0.72,
      fontSize: 8.5, color: C.black, fontFace: BF, wrap: true
    });
  });

  // Action cards
  [
    { n:'1', title:'Invest',     body:'Deploy capital into Weaver\'s Series B to fund geographic expansion into SADC markets.',  color: C.purple  },
    { n:'2', title:'Partner',    body:'White-label Weaver\'s credit engine to extend financial access through your distribution.',  color: '1A1A1A' },
    { n:'3', title:'Integrate',  body:'Connect your customer base to Weaver\'s lending, insurance, and savings product suite.',   color: C.purple  },
    { n:'4', title:'Co-build',   body:'Commission bespoke product development leveraging Weaver\'s technology and data assets.',   color: '1A1A1A' },
  ].forEach((a, i) => {
    const xx = 6.4 + (i%2)*3.45;
    const yy = 0.92 + Math.floor(i/2)*2.42;
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: xx, y: yy, w: 3.3, h: 2.2,
      fill: { color: a.color }, line: { color: a.color }, rectRadius: 0.1
    });
    s.addShape(pptx.shapes.OVAL, {
      x: xx+0.15, y: yy+0.12, w: 0.46, h: 0.46,
      fill: { color: a.color===C.purple ? C.white : '875EFA' },
      line: { color: a.color===C.purple ? C.white : '875EFA' }
    });
    s.addText(a.n, {
      x: xx+0.15, y: yy+0.12, w: 0.46, h: 0.46,
      fontSize: 13, bold: true,
      color: a.color===C.purple ? C.purple : C.white,
      align: 'center', fontFace: TF
    });
    s.addText(a.title, {
      x: xx+0.15, y: yy+0.68, w: 2.98, h: 0.44,
      fontSize: 9.5, bold: true, color: C.white, fontFace: TF, wrap: true
    });
    s.addText(a.body, {
      x: xx+0.15, y: yy+1.16, w: 2.98, h: 0.92,
      fontSize: 8,
      color: a.color===C.purple ? C.white : 'CCCCCC',
      fontFace: BF, wrap: true
    });
  });

  insightBar(s,
    'Weaver\'s window for category leadership is open now — first-mover advantage in SADC digital lending closes within 24 months.',
    'The partner that moves first gains preferred distribution rights, co-branded product exclusivity, and data network effects.',
    'Delay beyond Q3 2026 risks a competing capital deployment that locks Weaver into an exclusive arrangement with a rival.',
    'Schedule a due diligence session and term sheet discussion before end of Q2 2026.'
  );
  slideNum(s, 8);
}

// ── SLIDE 9: THANK YOU ───────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.addImage({ path: LOGO, x: 0.35, y: 0.28, w: 2.5, h: 1.1 });
  s.addText('Thank you.', {
    x: 0, y: 2.7, w: 13.33, h: 1.8,
    fontSize: 64, bold: true, color: C.purple, align: 'center', fontFace: TF
  });
  s.addText('weaverfintech.com', {
    x: 0, y: 6.7, w: 13.33, h: 0.38,
    fontSize: 12, color: C.black, align: 'center', fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 7.28, w: 13.33, h: 0.22,
    fill: { color: C.lime }, line: { color: C.lime }
  });
}

// ── SAVE ─────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: '/mnt/user-data/outputs/Weaver_Pitch.pptx' })
  .then(() => console.log('✅ Weaver_Pitch.pptx saved'))
  .catch(e => console.error('❌', e));
