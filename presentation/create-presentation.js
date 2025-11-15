const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// Create a new presentation
const pres = new PptxGenJS();

// Set presentation properties
pres.author = 'Mallacoota Black Lip Abalone Co-Operative';
pres.company = 'Mallacoota Co-Op';
pres.subject = 'Beauty Market Repositioning Strategy';
pres.title = 'From Crisis to Opportunity: The Beauty Pivot';

// Define custom layout for 1920x1080px (Full HD 16:9)
// At 96 DPI: 1920px = 20 inches, 1080px = 11.25 inches
pres.defineLayout({ name: 'LAYOUT_1920x1080', width: 20, height: 11.25 });
pres.layout = 'LAYOUT_1920x1080';

// Scaling factor to convert from design coordinates (10" width) to Full HD (20" width)
const SCALE = 2;

// Helper function to scale coordinates for Full HD
function s(value) {
  return value * SCALE;
}

// Helper function to scale options object
function scaleOpts(opts) {
  const scaled = { ...opts };
  if (opts.x !== undefined) scaled.x = s(opts.x);
  if (opts.y !== undefined) scaled.y = s(opts.y);
  if (opts.w !== undefined) scaled.w = s(opts.w);
  if (opts.h !== undefined) scaled.h = s(opts.h);
  if (opts.line && opts.line.width) scaled.line = { ...opts.line, width: s(opts.line.width) };
  return scaled;
}

// Color palette
const colors = {
  oceanBlue: '003366',
  pearlWhite: 'F8F8F8',
  goldAccent: 'D4AF37',
  seafoamGreen: '7FCDCD',
  darkText: '1a1a1a',
  lightText: 'ffffff',
  alertRed: 'dc3545',
  successGreen: '28a745'
};

// Helper function to add speaker notes
function addSpeakerNotes(slide, notes) {
  if (notes) {
    slide.addNotes(notes);
  }
}

// SLIDE 1: Title Slide
function createSlide1() {
  const slide = pres.addSlide();

  // Background gradient
  slide.background = { fill: colors.oceanBlue };

  // Main title
  slide.addText('From Crisis to Opportunity:\nThe Beauty Pivot', {
    x: s(0.5), y: s(1.5), w: s(9), h: s(1.5),
    fontSize: 48,
    bold: true,
    color: colors.lightText,
    align: 'center',
    valign: 'middle'
  });

  // Gold line
  slide.addShape('rect', {
    x: s(4), y: s(3.1), w: s(2), h: s(0.05),
    fill: { color: colors.goldAccent }
  });

  // Subtitle
  slide.addText('Repositioning Mallacoota Wild Abalone\nfor Premium Markets', {
    x: s(0.5), y: s(3.4), w: s(9), h: s(0.8),
    fontSize: 24,
    color: colors.seafoamGreen,
    align: 'center'
  });

  // Date
  slide.addText('November 2025', {
    x: s(0.5), y: s(4.5), w: s(9), h: s(0.3),
    fontSize: 18,
    color: colors.lightText,
    align: 'center'
  });

  // Company name
  slide.addText('MALLACOOTA BLACK LIP ABALONE CO-OPERATIVE', {
    x: s(0.5), y: s(4.9), w: s(9), h: s(0.4),
    fontSize: 20,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "Welcome to this strategic presentation on transforming our crisis into opportunity. Today we'll explore how repositioning our wild abalone from commodity food to premium beauty ingredient can secure the future of our 16 diving families. This is about survival and growth through smart, risk-managed innovation.");
}

// SLIDE 2: The Crisis
function createSlide2() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Crisis', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // 75% stat box
  slide.addShape('rect', {
    x: s(0.5), y: s(1.2), w: s(4), h: s(1),
    fill: { color: 'ffffff' },
    line: { color: colors.alertRed, width: s(4), type: 'solid' }
  });

  slide.addText('75%', {
    x: s(0.7), y: s(1.3), w: s(3.6), h: s(0.5),
    fontSize: 44,
    bold: true,
    color: colors.alertRed
  });

  slide.addText('Revenue Collapse in 2 Years', {
    x: s(0.7), y: s(1.8), w: s(3.6), h: s(0.3),
    fontSize: 16,
    color: colors.darkText
  });

  // 12→3 containers stat box
  slide.addShape('rect', {
    x: s(0.5), y: s(2.4), w: s(4), h: s(1),
    fill: { color: 'ffffff' },
    line: { color: colors.alertRed, width: s(4), type: 'solid' }
  });

  slide.addText('12 → 3', {
    x: s(0.7), y: s(2.5), w: s(3.6), h: s(0.5),
    fontSize: 44,
    bold: true,
    color: colors.alertRed
  });

  slide.addText('Container Shipments to Singapore\n(Annual Volume)', {
    x: s(0.7), y: s(3.0), w: s(3.6), h: s(0.3),
    fontSize: 15,
    color: colors.darkText
  });

  // Price comparisons
  slide.addShape('rect', {
    x: s(5.0), y: s(1.2), w: s(4.5), h: s(0.8),
    fill: { color: 'ffffff' },
    line: { color: 'e0e0e0', width: s(1) }
  });

  slide.addText('Chinese Farmed Abalone', {
    x: s(5.2), y: s(1.3), w: s(4.1), h: s(0.3),
    fontSize: 14,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('$8-12/kg', {
    x: s(5.2), y: s(1.6), w: s(4.1), h: s(0.3),
    fontSize: 28,
    bold: true,
    color: colors.alertRed
  });

  slide.addShape('rect', {
    x: s(5.0), y: s(2.2), w: s(4.5), h: s(0.8),
    fill: { color: 'ffffff' },
    line: { color: 'e0e0e0', width: s(1) }
  });

  slide.addText('Our Wild Abalone', {
    x: s(5.2), y: s(2.3), w: s(4.1), h: s(0.3),
    fontSize: 14,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('$40-60/kg', {
    x: s(5.2), y: s(2.6), w: s(4.1), h: s(0.3),
    fontSize: 28,
    bold: true,
    color: colors.successGreen
  });

  // Impact box
  slide.addShape('rect', {
    x: s(0.5), y: s(3.8), w: s(9), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('16 Diving Families at Risk', {
    x: s(0.5), y: s(3.9), w: s(9), h: s(0.6),
    fontSize: 24,
    bold: true,
    color: colors.lightText,
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "We're facing an existential crisis. In just two years, our revenue has collapsed by 75%, with shipments dropping from 12 containers to just 3. Chinese farmed abalone at $8-12 per kilogram is undercutting our wild product priced at $40-60, making us uncompetitive in traditional food markets. The livelihoods of 16 diving families hang in the balance.");
}

// SLIDE 3: Why Traditional Strategies Won't Work
function createSlide3() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Why Traditional Strategies Won\'t Work', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Strategy boxes
  const strategies = [
    {
      y: s(1.1),
      title: '✗ Chef Education Programs',
      text: 'Too resource-intensive for small co-op; requires dedicated staff, travel, and years to build relationships'
    },
    {
      y: s(1.9),
      title: '✗ Retail Consumer Marketing',
      text: 'Consumers won\'t pay 5X premium for food; price sensitivity too high in commodity market'
    },
    {
      y: s(2.7),
      title: '✗ Premium Food Branding',
      text: 'No existing brand equity; would require $60K+ investment in packaging, marketing, distribution without guaranteed ROI'
    }
  ];

  strategies.forEach(strategy => {
    slide.addShape('rect', {
      x: s(0.5), y: strategy.y, w: s(9), h: s(0.7),
      fill: { color: 'ffffff' },
      line: { color: colors.alertRed, width: s(3), type: 'solid' }
    });

    slide.addText(strategy.title, {
      x: s(0.7), y: strategy.y + 0.05, w: s(8.6), h: s(0.25),
      fontSize: 18,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(strategy.text, {
      x: s(0.7), y: strategy.y + 0.32, w: s(8.6), h: s(0.35),
      fontSize: 15,
      color: colors.darkText
    });
  });

  // Quote box
  slide.addShape('rect', {
    x: s(0.5), y: s(3.7), w: s(9), h: s(0.9),
    fill: { color: 'f8f9fa' },
    line: { color: colors.seafoamGreen, width: s(3), type: 'solid' }
  });

  slide.addText('"We\'re a small operation. We can\'t compete with marketing budgets or farming scale. We need a strategy that works with what we have—exceptional wild product and minimal resources."', {
    x: s(0.7), y: s(3.8), w: s(8.6), h: s(0.5),
    fontSize: 16,
    italic: true,
    color: colors.darkText
  });

  slide.addText('— Charles, CEO', {
    x: s(0.7), y: s(4.35), w: s(8.6), h: s(0.2),
    fontSize: 14,
    bold: true,
    color: colors.oceanBlue
  });

  addSpeakerNotes(slide, "Traditional market strategies simply won't work for us. Chef programs demand resources we don't have. Consumers won't pay 5X for food regardless of quality. Premium branding requires major upfront investment with uncertain returns. We need a fundamentally different approach that leverages our strengths without resource constraints.");
}

// SLIDE 4: The Strategic Insight
function createSlide4() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Strategic Insight', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Insight boxes
  const insights = [
    { x: s(0.5), y: s(1.0), icon: '🧬', title: 'Cultural Beliefs Already Exist', text: 'Asian cultures associate abalone with health, virility, and longevity—beliefs embedded for generations' },
    { x: s(5.25), y: s(1.0), icon: '🔬', title: 'Lab Testing Opportunity', text: 'Prove wild has 30%+ more beneficial compounds than farmed—collagen, peptides, omega-3s' },
    { x: s(0.5), y: s(2.3), icon: '🚫', title: 'Zero Hormones Advantage', text: 'Wild abalone = hormone-free vs. farmed "full of hormones" perception among health-conscious consumers' },
    { x: s(5.25), y: s(2.3), icon: '💎', title: 'Category Shift', text: 'Not competing as food ingredient—repositioning as premium beauty/wellness product' }
  ];

  insights.forEach(insight => {
    slide.addShape('rect', {
      x: insight.x, y: insight.y, w: s(4.5), h: s(1.1),
      fill: { color: 'ffffff' },
      line: { color: 'e0e0e0', width: s(1) }
    });

    slide.addText(insight.icon, {
      x: insight.x + 0.2, y: insight.y + 0.1, w: s(0.5), h: s(0.3),
      fontSize: 24
    });

    slide.addText(insight.title, {
      x: insight.x + 0.2, y: insight.y + 0.35, w: s(4.1), h: s(0.25),
      fontSize: 17,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(insight.text, {
      x: insight.x + 0.2, y: insight.y + 0.62, w: s(4.1), h: s(0.4),
      fontSize: 14,
      color: colors.darkText
    });
  });

  // Highlight box
  slide.addShape('rect', {
    x: s(1.0), y: s(3.7), w: s(8), h: s(0.9),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Game Changer', {
    x: s(1.0), y: s(3.75), w: s(8), h: s(0.25),
    fontSize: 20,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  slide.addText('Stop fighting the food price war.\nStart winning in the beauty category.', {
    x: s(1.0), y: s(4.05), w: s(8), h: s(0.5),
    fontSize: 18,
    color: colors.lightText,
    align: 'center'
  });

  addSpeakerNotes(slide, "The breakthrough insight: we don't need to create new beliefs—they already exist in Asian culture. We just need to prove them scientifically and shift the category from food to beauty/wellness. Our wild advantage isn't taste, it's purity and potency. Lab testing can validate our 30%+ compound superiority and zero-hormone story, positioning us in a premium category where consumers happily pay multiples for proven benefits.");
}

// SLIDE 5: Market Opportunity
function createSlide5() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Market Opportunity', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Left column - market stats
  slide.addShape('rect', {
    x: s(0.5), y: s(1.0), w: s(4.5), h: s(0.7),
    fill: { color: 'ffffff' },
    line: { color: colors.goldAccent, width: s(3), type: 'solid' }
  });

  slide.addText('$82B', {
    x: s(0.7), y: s(1.05), w: s(4.1), h: s(0.35),
    fontSize: 36,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('Asian Beauty/Wellness Market Size (2024)', {
    x: s(0.7), y: s(1.42), w: s(4.1), h: s(0.2),
    fontSize: 14,
    color: colors.darkText
  });

  slide.addShape('rect', {
    x: s(0.5), y: s(1.85), w: s(4.5), h: s(0.7),
    fill: { color: 'ffffff' },
    line: { color: colors.goldAccent, width: s(3), type: 'solid' }
  });

  slide.addText('12.3%', {
    x: s(0.7), y: s(1.9), w: s(4.1), h: s(0.35),
    fontSize: 36,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('Annual Growth Rate (CAGR)', {
    x: s(0.7), y: s(2.27), w: s(4.1), h: s(0.2),
    fontSize: 14,
    color: colors.darkText
  });

  // Right column - trends
  const trends = [
    { y: s(1.0), title: '📈 "Eating Your Way to Beauty"', text: 'Collagen drinks, bird\'s nest, pearl powder—ingestible beauty mainstream in Asia' },
    { y: s(2.0), title: '🌿 Clean Beauty Movement', text: 'Hormone-free, wild-sourced, sustainable—exactly our positioning' },
    { y: s(3.0), title: '📱 Influencer-Driven Discovery', text: 'Beauty consumers trust micro-influencers over traditional ads; cost-effective reach' }
  ];

  trends.forEach(trend => {
    slide.addShape('rect', {
      x: s(5.25), y: trend.y, w: s(4.25), h: s(0.85),
      fill: { color: 'd4edda' },
      line: { color: colors.successGreen, width: s(3), type: 'solid' }
    });

    slide.addText(trend.title, {
      x: s(5.45), y: trend.y + 0.1, w: s(3.85), h: s(0.25),
      fontSize: 16,
      bold: true,
      color: '155724'
    });

    slide.addText(trend.text, {
      x: s(5.45), y: trend.y + 0.38, w: s(3.85), h: s(0.4),
      fontSize: 13,
      color: '155724'
    });
  });

  addSpeakerNotes(slide, "The Asian beauty and wellness market is massive—$82 billion and growing 12% annually. We're targeting Gen Z and Millennial consumers who are growing, digitally native, and willing to pay premiums for proven benefits, unlike the declining banquet market. The 'eating your way to beauty' trend is already mainstream, clean beauty is booming, and influencer marketing provides cost-effective access to our target audience.");
}

// SLIDE 6: Stage-Gate Strategy Overview
function createSlide6() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.75),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Stage-Gate Strategy Overview', {
    x: s(0.5), y: s(0.18), w: s(9), h: s(0.4),
    fontSize: 34,
    bold: true,
    color: colors.lightText
  });

  // Gate boxes
  const gates = [
    { y: s(0.95), gate: 'Gate 1: Lab Testing', desc: 'Test wild vs farmed for beneficial compounds & hormones', investment: '$15-25K' },
    { y: s(1.65), gate: 'Gate 2: Legal Review', desc: 'Validate permissible claims in target markets', investment: '$5-10K' },
    { y: s(2.35), gate: 'Gate 3: Micro-Influencer Test', desc: '5-8 influencers, measure engagement & sentiment', investment: '$10-15K' },
    { y: s(3.05), gate: 'Gate 4: Limited Launch', desc: '5K-10K units, premium line, mid-tier influencers', investment: '$30-50K' },
    { y: s(3.75), gate: 'Gate 5: Scale & Sustain', desc: 'Multi-market expansion, major influencers, retail partners', investment: '$50-100K' }
  ];

  gates.forEach((gate, index) => {
    // Gate box
    slide.addShape('rect', {
      x: s(0.5), y: gate.y, w: s(6.5), h: s(0.6),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(2) }
    });

    slide.addText(gate.gate, {
      x: s(0.65), y: gate.y + 0.05, w: s(4.5), h: s(0.2),
      fontSize: 16,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(gate.desc, {
      x: s(0.65), y: gate.y + 0.27, w: s(4.5), h: s(0.15),
      fontSize: 12,
      color: colors.darkText
    });

    slide.addText(gate.investment, {
      x: s(0.65), y: gate.y + 0.44, w: s(4.5), h: s(0.12),
      fontSize: 14,
      bold: true,
      color: colors.goldAccent
    });

    // GO/NO-GO badge
    slide.addShape('rect', {
      x: s(5.8), y: gate.y + 0.2, w: s(1.0), h: s(0.2),
      fill: { color: colors.seafoamGreen }
    });

    slide.addText('GO/NO-GO', {
      x: s(5.8), y: gate.y + 0.2, w: s(1.0), h: s(0.2),
      fontSize: 10,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });

    // Arrow (except after last gate)
    if (index < gates.length - 1) {
      slide.addText('▼', {
        x: s(3.2), y: gate.y + 0.62, w: s(0.6), h: s(0.25),
        fontSize: 20,
        color: colors.seafoamGreen,
        align: 'center'
      });
    }
  });

  // Total investment box
  slide.addShape('rect', {
    x: s(7.2), y: s(1.0), w: s(2.3), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Total Investment', {
    x: s(7.2), y: s(1.05), w: s(2.3), h: s(0.25),
    fontSize: 14,
    color: colors.seafoamGreen,
    align: 'center'
  });

  slide.addText('$110-200K', {
    x: s(7.2), y: s(1.32), w: s(2.3), h: s(0.4),
    fontSize: 24,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "Our stage-gate approach breaks the strategy into five manageable phases with clear GO/NO-GO decision points. Total investment ranges from $110K to $200K, but we only commit to each stage after the previous one succeeds. This minimizes risk and ensures we can exit at any point with controlled losses rather than committing everything upfront.");
}

// SLIDE 7-17: Continue with similar structure
// For brevity, I'll create simplified versions of the remaining slides

function createSlide7() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Gate 1: Lab Testing (Detail)', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  slide.addText('Test Wild vs Farmed For:', {
    x: s(0.5), y: s(1.0), w: s(4.5), h: s(0.3),
    fontSize: 20,
    bold: true,
    color: colors.oceanBlue
  });

  const tests = [
    'Collagen & Peptides - Key compounds for skin elasticity',
    'Omega-3 Fatty Acids - Essential for skin hydration',
    'Minerals & Trace Elements - Zinc, selenium, iron',
    'Hormone Testing - Verify zero hormones in wild'
  ];

  tests.forEach((test, i) => {
    slide.addText('▸ ' + test, {
      x: s(0.7), y: s(1.4) + (i * 0.4), w: s(4), h: s(0.35),
      fontSize: 14,
      color: colors.darkText
    });
  });

  // Success criteria
  slide.addShape('rect', {
    x: s(5.25), y: s(1.0), w: s(4.25), h: s(1.2),
    fill: { color: 'd4edda' },
    line: { color: colors.successGreen, width: s(2) }
  });

  slide.addText('Success Criteria (GO Decision)', {
    x: s(5.45), y: s(1.1), w: s(3.85), h: s(0.25),
    fontSize: 17,
    bold: true,
    color: '155724'
  });

  slide.addText('✓ ≥30% higher beneficial compounds\n✓ Zero detectable hormones\n✓ Results from certified lab', {
    x: s(5.45), y: s(1.4), w: s(3.85), h: s(0.7),
    fontSize: 14,
    color: '155724'
  });

  // Investment details
  slide.addText('Timeline: 6-8 Weeks\nInvestment: $15,000 - $25,000', {
    x: s(5.45), y: s(2.5), w: s(3.85), h: s(0.5),
    fontSize: 16,
    bold: true,
    color: colors.oceanBlue
  });

  addSpeakerNotes(slide, "Gate 1 is our foundation. We invest $15-25K to scientifically prove our wild abalone has at least 30% more beneficial compounds than farmed, with zero hormones. Testing covers collagen, peptides, omega-3s, and minerals—all key beauty ingredients. Timeline is 6-8 weeks. If results don't meet our criteria, we stop here with minimal financial exposure.");
}

function createSlide8() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Gates 2-3: Legal Review & Testing', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Gate 2
  slide.addShape('rect', {
    x: s(0.5), y: s(1.0), w: s(9), h: s(1.5),
    fill: { color: 'ffffff' },
    line: { color: colors.oceanBlue, width: s(2) }
  });

  slide.addText('Gate 2: Legal Review', {
    x: s(0.7), y: s(1.1), w: s(8.6), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('▸ Determine permissible claims in Singapore, Malaysia, Hong Kong\n▸ Approved claim language for packaging and marketing\n▸ Market-specific compliance requirements\n\nTimeline: 4-6 Weeks  |  Investment: $5-10K', {
    x: s(0.7), y: s(1.45), w: s(6), h: s(0.95),
    fontSize: 14,
    color: colors.darkText
  });

  // Gate 3
  slide.addShape('rect', {
    x: s(0.5), y: s(2.7), w: s(9), h: s(1.7),
    fill: { color: 'ffffff' },
    line: { color: colors.oceanBlue, width: s(2) }
  });

  slide.addText('Gate 3: Micro-Influencer Test', {
    x: s(0.7), y: s(2.8), w: s(8.6), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('▸ Partner with 5-8 micro-influencers (50K-150K followers)\n▸ Success: >5% engagement rate\n▸ Success: >80% positive sentiment\n▸ Validated messaging resonance\n\nTimeline: 6-10 Weeks  |  Investment: $10-15K', {
    x: s(0.7), y: s(3.15), w: s(6), h: s(1.15),
    fontSize: 14,
    color: colors.darkText
  });

  addSpeakerNotes(slide, "Gates 2 and 3 validate our legal positioning and market messaging. Legal review ensures compliance across target markets for $5-10K. Micro-influencer testing with 5-8 partners proves audience resonance before major investment, requiring greater than 5% engagement and 80% positive sentiment for a GO decision. Combined investment of $15-25K with 10-16 week timeline.");
}

function createSlide9() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Gate 4: Limited Launch', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Product name highlight
  slide.addShape('rect', {
    x: s(1.5), y: s(1.0), w: s(7), h: s(0.7),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Mallacoota Beauty Reserve', {
    x: s(1.5), y: s(1.1), w: s(7), h: s(0.35),
    fontSize: 28,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  slide.addText('Premium Wild Abalone for Beauty & Wellness', {
    x: s(1.5), y: s(1.45), w: s(7), h: s(0.2),
    fontSize: 16,
    italic: true,
    color: colors.seafoamGreen,
    align: 'center'
  });

  // Details
  slide.addText('Product Launch:\n5,000-10,000 units at $60-80 per can\n\nMarketing Strategy:\n15-20 mid-tier influencers (150K-500K followers)\n\nDistribution:\nE-commerce direct to consumer\n\nSuccess Criteria:\n70% sell-through in 3 months\n\nInvestment: $30-50K', {
    x: s(0.7), y: s(2.0), w: s(8.6), h: s(2.4),
    fontSize: 16,
    color: colors.darkText,
    lineSpacing: 20
  });

  addSpeakerNotes(slide, "Gate 4 launches our premium 'Mallacoota Beauty Reserve' line with 5,000-10,000 units at $60-80 per can. We partner with 15-20 mid-tier influencers and sell direct-to-consumer via e-commerce. Success means 70% sell-through in 3 months, validating both demand and premium pricing. Investment is $30-50K for packaging, influencer partnerships, and initial inventory.");
}

function createSlide10() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Gate 5: Scale & Sustain', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Strategy boxes
  const strategies = [
    { icon: '📱', title: 'Macro-Influencer Partnerships', text: 'Expand to 1M+ follower beauty influencers' },
    { icon: '🌏', title: 'Multi-Market Expansion', text: 'Singapore, Malaysia, Hong Kong, Taiwan' },
    { icon: '🏪', title: 'Premium Retail Partnerships', text: 'Sephora, Watsons, boutique retailers' }
  ];

  strategies.forEach((strat, i) => {
    slide.addText(strat.icon + ' ' + strat.title, {
      x: s(0.7), y: s(1.1) + (i * 0.6), w: s(8.6), h: s(0.25),
      fontSize: 18,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(strat.text, {
      x: s(0.7), y: s(1.37) + (i * 0.6), w: s(8.6), h: s(0.2),
      fontSize: 14,
      color: colors.darkText
    });
  });

  // Target
  slide.addShape('rect', {
    x: s(2.5), y: s(3.0), w: s(5), h: s(0.9),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Revenue Target: 20-30%', {
    x: s(2.5), y: s(3.1), w: s(5), h: s(0.3),
    fontSize: 28,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  slide.addText('of total revenue from premium beauty line', {
    x: s(2.5), y: s(3.45), w: s(5), h: s(0.25),
    fontSize: 14,
    color: colors.lightText,
    align: 'center'
  });

  slide.addText('Annual Investment: $50-100K', {
    x: s(2.5), y: s(4.1), w: s(5), h: s(0.3),
    fontSize: 18,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "Gate 5 focuses on scaling what works. We expand to macro-influencers with 1M+ followers, enter new markets across Asia, and secure premium retail partnerships. The target is generating 20-30% of total co-op revenue from the beauty line. Annual investment of $50-100K sustains marketing and expansion, but at this stage we have proven demand and are growing profitably.");
}

function createSlide11() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Messaging Framework', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Core message
  slide.addShape('rect', {
    x: s(1.0), y: s(1.0), w: s(8), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue },
    line: { color: colors.goldAccent, width: s(3) }
  });

  slide.addText('CORE MESSAGE', {
    x: s(1.0), y: s(1.05), w: s(8), h: s(0.2),
    fontSize: 14,
    color: colors.seafoamGreen,
    align: 'center'
  });

  slide.addText('"Beauty from the 40°S\nSouthern Ocean Reserve"', {
    x: s(1.0), y: s(1.3), w: s(8), h: s(0.45),
    fontSize: 26,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  // Four pillars
  const pillars = [
    { x: s(0.5), y: s(2.0), icon: '🌊', title: 'Purity', text: 'Zero hormones, pristine Southern Ocean, UNESCO reserve' },
    { x: s(5.0), y: s(2.0), icon: '💪', title: 'Potency', text: 'Lab-tested [X%] more collagen, peptides, omega-3s' },
    { x: s(0.5), y: s(3.0), icon: '🏛️', title: 'Provenance', text: '56-year heritage, 16 families, EX1191 license' },
    { x: s(5.0), y: s(3.0), icon: '🔬', title: 'Proven', text: 'Scientific validation of traditional beliefs' }
  ];

  pillars.forEach(pillar => {
    slide.addShape('rect', {
      x: pillar.x, y: pillar.y, w: s(4.25), h: s(0.8),
      fill: { color: 'ffffff' },
      line: { color: colors.seafoamGreen, width: s(2) }
    });

    slide.addText(pillar.icon + ' ' + pillar.title, {
      x: pillar.x + 0.2, y: pillar.y + 0.1, w: s(3.85), h: s(0.25),
      fontSize: 18,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(pillar.text, {
      x: pillar.x + 0.2, y: pillar.y + 0.38, w: s(3.85), h: s(0.35),
      fontSize: 13,
      color: colors.darkText
    });
  });

  // Tagline
  slide.addShape('rect', {
    x: s(2.0), y: s(4.0), w: s(6), h: s(0.6),
    fill: { color: colors.goldAccent }
  });

  slide.addText('Pure. Wild. Proven.', {
    x: s(2.0), y: s(4.1), w: s(6), h: s(0.4),
    fontSize: 32,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "Our messaging framework builds on existing cultural beliefs while adding scientific credibility. The core message positions us as premium wild product from pristine Australian waters. Four pillars—Purity, Potency, Provenance, Proven—create a compelling brand story. The tagline 'Pure. Wild. Proven.' encapsulates our competitive advantage in three powerful words that resonate across Asian beauty markets.");
}

function createSlide12() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Financial Projections', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Table header
  slide.addShape('rect', {
    x: s(0.5), y: s(1.0), w: s(9), h: s(0.4),
    fill: { color: colors.oceanBlue }
  });

  const headers = [
    { text: 'Year', x: s(0.6), w: s(1) },
    { text: 'Revenue', x: s(2.0), w: s(2.5) },
    { text: 'Units Sold', x: s(4.8), w: s(2) },
    { text: 'Gross Margin', x: s(7.1), w: s(2) }
  ];

  headers.forEach(h => {
    slide.addText(h.text, {
      x: h.x, y: s(1.1), w: h.w, h: s(0.2),
      fontSize: 14,
      bold: true,
      color: colors.lightText
    });
  });

  // Table rows
  const rows = [
    { year: 'Year 1', revenue: '$1.2-1.8M', units: '20-30K units', margin: '60%', y: s(1.5) },
    { year: 'Year 2', revenue: '$3-5M', units: '50-80K units', margin: '60%', y: s(2.0) },
    { year: 'Year 3', revenue: '$6-10M', units: '100-150K units', margin: '60%', y: s(2.5) }
  ];

  rows.forEach(row => {
    slide.addText(row.year, {
      x: s(0.6), y: row.y, w: s(1), h: s(0.3),
      fontSize: 14,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(row.revenue, {
      x: s(2.0), y: row.y, w: s(2.5), h: s(0.3),
      fontSize: 18,
      bold: true,
      color: colors.successGreen
    });

    slide.addText(row.units, {
      x: s(4.8), y: row.y, w: s(2), h: s(0.3),
      fontSize: 14,
      color: colors.darkText
    });

    slide.addText(row.margin, {
      x: s(7.1), y: row.y, w: s(2), h: s(0.3),
      fontSize: 14,
      color: colors.darkText
    });
  });

  // ROI box
  slide.addShape('rect', {
    x: s(5.5), y: s(3.2), w: s(4), h: s(1.2),
    fill: { color: colors.goldAccent }
  });

  slide.addText('3-Year ROI', {
    x: s(5.5), y: s(3.3), w: s(4), h: s(0.25),
    fontSize: 16,
    color: 'ffffff',
    align: 'center'
  });

  slide.addText('1,800-3,173%', {
    x: s(5.5), y: s(3.6), w: s(4), h: s(0.5),
    fontSize: 32,
    bold: true,
    color: 'ffffff',
    align: 'center'
  });

  slide.addText('On $110-200K total investment', {
    x: s(5.5), y: s(4.15), w: s(4), h: s(0.2),
    fontSize: 13,
    color: 'ffffff',
    align: 'center'
  });

  // Margin comparison
  slide.addText('Beauty Line Margin: 60%\nvs\nCommodity Margin: 30-40%', {
    x: s(0.7), y: s(3.3), w: s(4), h: s(0.8),
    fontSize: 16,
    bold: true,
    color: colors.oceanBlue
  });

  addSpeakerNotes(slide, "Financial projections show strong growth potential: Year 1 revenue of $1.2-1.8M, scaling to $6-10M by Year 3. The beauty line delivers 60% gross margins versus 30-40% on commodity sales. Three-year ROI ranges from 1,800% to 3,173% on total investment of $110-200K. These are conservative estimates based on proven beauty market dynamics and influencer marketing performance data.");
}

function createSlide13() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Risk Management', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Key points
  slide.addText('🛡️ Gate System Minimizes Exposure', {
    x: s(0.7), y: s(1.0), w: s(8.6), h: s(0.3),
    fontSize: 20,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('Clear GO/NO-GO decision points mean you never commit to the full strategy upfront. Exit at any stage with controlled, predictable losses.', {
    x: s(0.7), y: s(1.35), w: s(8.6), h: s(0.4),
    fontSize: 15,
    color: colors.darkText
  });

  // Gate risks
  slide.addText('Maximum Risk per Gate:', {
    x: s(0.7), y: s(1.9), w: s(8.6), h: s(0.25),
    fontSize: 18,
    bold: true,
    color: colors.oceanBlue
  });

  const gateRisks = [
    'Gate 1: $15-25K',
    'Gate 2: $5-10K',
    'Gate 3: $10-15K',
    'Gate 4: $30-50K',
    'Gate 5: $50-100K'
  ];

  gateRisks.forEach((risk, i) => {
    slide.addText(risk, {
      x: s(1.0) + (i % 3) * 2.8, y: s(2.25) + Math.floor(i / 3) * 0.4, w: s(2.5), h: s(0.3),
      fontSize: 16,
      bold: true,
      color: colors.goldAccent
    });
  });

  // Exit strategy highlight
  slide.addShape('rect', {
    x: s(0.7), y: s(3.2), w: s(8.6), h: s(0.7),
    fill: { color: 'fff3cd' },
    line: { color: 'ffc107', width: s(2) }
  });

  slide.addText('Exit Anytime with Controlled Losses\nIf lab results fail: out $25K max. If influencer test fails: out $50K total.\nCompare to status quo: losing $6M annually with no plan.', {
    x: s(0.9), y: s(3.3), w: s(8.2), h: s(0.5),
    fontSize: 15,
    color: '856404',
    align: 'center'
  });

  // Probability-adjusted ROI
  slide.addShape('rect', {
    x: s(2.5), y: s(4.1), w: s(5), h: s(0.5),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Probability-Adjusted ROI: 378%', {
    x: s(2.5), y: s(4.2), w: s(5), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "Our stage-gate approach is designed to minimize financial risk. Maximum exposure at any single gate ranges from $15-25K initially to $50-100K at scale—but only after proving success at earlier stages. You can exit anytime with controlled losses. Even with probability-adjusted returns accounting for failure risk at each gate, ROI is still 378%—dramatically better than the status quo of declining revenue with no recovery plan.");
}

function createSlide14() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Implementation Timeline', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Timeline items
  const timeline = [
    { period: 'Months 1-2', title: 'Research & Legal [GATE 1-2]', desc: 'Lab testing, legal review. Investment: $20-35K', y: s(1.1) },
    { period: 'Months 3-4', title: 'Content Testing [GATE 3]', desc: 'Micro-influencer partnerships. Investment: $10-15K', y: s(2.0) },
    { period: 'Months 5-7', title: 'Limited Launch [GATE 4]', desc: 'Premium product, 5K-10K units. Investment: $30-50K', y: s(2.9) },
    { period: 'Months 8-12', title: 'Scale & Expand [GATE 5]', desc: 'Multi-market, macro-influencers. Investment: $50-100K', y: s(3.8) }
  ];

  timeline.forEach(item => {
    // Timeline dot
    slide.addShape('ellipse', {
      x: s(0.7), y: item.y, w: s(0.3), h: s(0.3),
      fill: { color: colors.goldAccent },
      line: { color: 'ffffff', width: s(2) }
    });

    // Content box
    slide.addShape('rect', {
      x: s(1.2), y: item.y - 0.05, w: s(8.3), h: s(0.7),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(2) }
    });

    slide.addText(item.period, {
      x: s(1.4), y: item.y, w: s(2), h: s(0.2),
      fontSize: 13,
      bold: true,
      color: colors.goldAccent
    });

    slide.addText(item.title, {
      x: s(1.4), y: item.y + 0.22, w: s(7.9), h: s(0.2),
      fontSize: 16,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(item.desc, {
      x: s(1.4), y: item.y + 0.44, w: s(7.9), h: s(0.15),
      fontSize: 12,
      color: colors.darkText
    });
  });

  // Key milestone
  slide.addShape('rect', {
    x: s(1.5), y: s(4.7), w: s(7), h: s(0.3),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('First GO/NO-GO Decision in 8 Weeks — Lab results determine next step', {
    x: s(1.5), y: s(4.75), w: s(7), h: s(0.2),
    fontSize: 14,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "The implementation timeline spans 12 months with clear milestones and decision points. Months 1-2 focus on research and legal validation. Months 3-4 test messaging with micro-influencers. Months 5-7 launch limited product run. Months 8-12 scale what works. Most importantly, we reach our first critical GO/NO-GO decision in just 8 weeks with lab results, allowing for quick validation or exit.");
}

function createSlide15() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Why This Strategy Works', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Reasons
  const reasons = [
    { icon: '🌊', title: 'Leverages Existing Assets', text: 'Wild quality, pristine waters, heritage story' },
    { icon: '⚙️', title: 'Works Within Constraints', text: 'Uses bulk processing, minimal new infrastructure' },
    { icon: '📱', title: 'Scalable Through Influencers', text: 'Reach millions without resource-intensive programs' },
    { icon: '📈', title: 'Targets Growth Market', text: 'Gen Z/Millennial beauty consumers expanding' },
    { icon: '🛡️', title: 'Risk-Managed Approach', text: 'Stage-gates allow validation at every phase' },
    { icon: '💰', title: 'Higher Margins', text: '60% gross margins vs 30-40% commodity' }
  ];

  reasons.forEach((reason, i) => {
    const x = i % 2 === 0 ? 0.5 : 5.0;
    const y = 1.0 + Math.floor(i / 2) * 0.7;

    slide.addShape('rect', {
      x: x, y: y, w: s(4.5), h: s(0.6),
      fill: { color: 'ffffff' },
      line: { color: colors.goldAccent, width: s(2) }
    });

    slide.addText(reason.icon + ' ' + reason.title, {
      x: x + 0.2, y: y + 0.08, w: s(4.1), h: s(0.22),
      fontSize: 16,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(reason.text, {
      x: x + 0.2, y: y + 0.32, w: s(4.1), h: s(0.2),
      fontSize: 13,
      color: colors.darkText
    });
  });

  // Summary
  slide.addShape('rect', {
    x: s(1.0), y: s(3.4), w: s(8), h: s(0.9),
    fill: { type: 'solid', color: colors.oceanBlue },
    line: { color: colors.goldAccent, width: s(2) }
  });

  slide.addText('The Bottom Line', {
    x: s(1.0), y: s(3.5), w: s(8), h: s(0.25),
    fontSize: 20,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  slide.addText('This isn\'t a risky bet—it\'s the logical evolution of our business.\nWe\'re using what we have to reach who\'s buying, in a category where we can win.', {
    x: s(1.0), y: s(3.8), w: s(8), h: s(0.4),
    fontSize: 16,
    color: colors.lightText,
    align: 'center'
  });

  addSpeakerNotes(slide, "This strategy succeeds because it's designed around our reality. We leverage existing assets—wild quality, heritage, pristine environment. It works within our constraints using bulk processing and influencer marketing instead of resource-intensive programs. We target growing Gen Z beauty consumers, not declining banquets. Stage-gates manage risk. And we capture 60% margins instead of 30-40%. This is smart repositioning, not wishful thinking.");
}

function createSlide16() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Immediate Next Steps', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 36,
    bold: true,
    color: colors.lightText
  });

  // Steps
  const steps = [
    { num: '1', title: 'Board Approval for Gate 1', timeline: 'Week 1', desc: 'Secure $15-25K investment approval', y: s(1.1) },
    { num: '2', title: 'Select Independent Lab', timeline: 'Weeks 1-2', desc: 'Identify certified lab for marine product analysis', y: s(1.8) },
    { num: '3', title: 'Collect Samples', timeline: 'Week 3', desc: 'Gather wild and farmed comparison samples', y: s(2.5) },
    { num: '4', title: 'Begin Testing', timeline: 'Weeks 4-10', desc: 'Comprehensive compound and hormone analysis', y: s(3.2) },
    { num: '5', title: 'First GO/NO-GO Decision', timeline: 'Week 10', desc: 'Review results: Proceed or exit with minimal loss', y: s(3.9) }
  ];

  steps.forEach(step => {
    // Number badge
    slide.addShape('ellipse', {
      x: s(0.6), y: step.y, w: s(0.4), h: s(0.4),
      fill: { color: colors.goldAccent }
    });

    slide.addText(step.num, {
      x: s(0.6), y: step.y, w: s(0.4), h: s(0.4),
      fontSize: 20,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });

    // Content
    slide.addText(step.title + ' - ' + step.timeline, {
      x: s(1.2), y: step.y + 0.02, w: s(8), h: s(0.2),
      fontSize: 17,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(step.desc, {
      x: s(1.2), y: step.y + 0.24, w: s(8), h: s(0.15),
      fontSize: 14,
      color: colors.darkText
    });
  });

  addSpeakerNotes(slide, "The path forward is clear and concrete. Step one is board approval for the $15-25K Gate 1 investment. Within two weeks we select a certified lab. By week three we collect samples. Testing runs for 6-8 weeks. By week 10—just 10 weeks from today—we'll have our first GO/NO-GO decision based on hard scientific data, not guesswork.");
}

function createSlide17() {
  const slide = pres.addSlide();

  // Background gradient
  slide.background = { fill: colors.oceanBlue };

  // Main headline
  slide.addText('Let\'s Test Our Hypothesis', {
    x: s(1), y: s(1.0), w: s(8), h: s(0.8),
    fontSize: 42,
    bold: true,
    color: colors.lightText,
    align: 'center'
  });

  slide.addText('Hypothesis', {
    x: s(4.2), y: s(1.05), w: s(1.6), h: s(0.7),
    fontSize: 42,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  // Stats boxes
  const stats = [
    { label: 'Initial Commitment', value: '$15-25K', x: s(0.75) },
    { label: 'Potential Upside', value: '$6-10M', x: s(3.5) },
    { label: 'Risk Management', value: 'Controlled\nat Every Stage', x: s(6.25) }
  ];

  stats.forEach(stat => {
    slide.addShape('rect', {
      x: stat.x, y: s(2.2), w: s(2.8), h: s(0.8),
      fill: { color: 'ffffff' },
      line: { color: colors.goldAccent, width: s(2) }
    });

    slide.addText(stat.label, {
      x: stat.x, y: s(2.28), w: s(2.8), h: s(0.2),
      fontSize: 13,
      bold: true,
      color: colors.oceanBlue,
      align: 'center'
    });

    slide.addText(stat.value, {
      x: stat.x, y: s(2.52), w: s(2.8), h: s(0.4),
      fontSize: 24,
      bold: true,
      color: colors.oceanBlue,
      align: 'center',
      valign: 'middle'
    });
  });

  // Question box
  slide.addShape('rect', {
    x: s(1.5), y: s(3.3), w: s(7), h: s(0.7),
    fill: { color: colors.goldAccent }
  });

  slide.addText('Ready to Transform\nCrisis into Opportunity?', {
    x: s(1.5), y: s(3.35), w: s(7), h: s(0.6),
    fontSize: 26,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  // Closing text
  slide.addText('The only risk greater than trying this strategy\nis continuing to do nothing while our market disappears.', {
    x: s(1), y: s(4.2), w: s(8), h: s(0.4),
    fontSize: 16,
    color: colors.seafoamGreen,
    align: 'center',
    italic: true
  });

  addSpeakerNotes(slide, "This is the moment of decision. We're asking for just $15-25K to test our hypothesis with lab results. The potential upside is $6-10M in annual revenue. Risk is controlled at every stage through our gate system. The real question isn't whether we can afford to try this—it's whether we can afford not to. Let's approve Gate 1 and take the first step toward securing our future.");
}

// Create all slides
createSlide1();
createSlide2();
createSlide3();
createSlide4();
createSlide5();
createSlide6();
createSlide7();
createSlide8();
createSlide9();
createSlide10();
createSlide11();
createSlide12();
createSlide13();
createSlide14();
createSlide15();
createSlide16();
createSlide17();

// Save the presentation
pres.writeFile({ fileName: 'mallacoota-beauty-pivot.pptx' })
  .then(() => {
    console.log('✓ Presentation created successfully: mallacoota-beauty-pivot.pptx');
  })
  .catch(err => {
    console.error('Error creating presentation:', err);
  });
