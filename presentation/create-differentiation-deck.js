const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// Create a new presentation
const pres = new PptxGenJS();

// Set presentation properties
pres.author = 'Mallacoota Black Lip Abalone Co-Operative';
pres.company = 'Mallacoota Co-Op';
pres.subject = 'Scientific Differentiation Strategy';
pres.title = 'The Science of Differentiation';

// Define custom layout for 1920x1080px (Full HD 16:9)
pres.defineLayout({ name: 'LAYOUT_1920x1080', width: 20, height: 11.25 });
pres.layout = 'LAYOUT_1920x1080';

// Scaling factor to convert from design coordinates (10" width) to Full HD (20" width)
const SCALE = 2;

// Helper function to scale coordinates for Full HD
function s(value) {
  return value * SCALE;
}

// Color palette - Updated for scientific/clinical feel
const colors = {
  oceanBlue: '003366',
  pearlWhite: 'F8F8F8',
  goldAccent: 'D4AF37',
  seafoamGreen: '7FCDCD',
  darkText: '1a1a1a',
  lightText: 'ffffff',
  alertRed: 'dc3545',
  successGreen: '28a745',
  scientificBlue: '0066cc',
  clinicalGray: 'e8e8e8'
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
  slide.background = { fill: colors.oceanBlue };

  // Main title
  slide.addText('The Science of\nDifferentiation', {
    x: s(0.5), y: s(1.5), w: s(9), h: s(1.5),
    fontSize: 54,
    bold: true,
    color: colors.lightText,
    align: 'center',
    valign: 'middle'
  });

  // Gold line
  slide.addShape('rect', {
    x: s(4), y: s(3.2), w: s(2), h: s(0.05),
    fill: { color: colors.goldAccent }
  });

  // Subtitle
  slide.addText('Why Wild EX1191 and Chinese Farmed\nAre Completely Different Products', {
    x: s(0.5), y: s(3.5), w: s(9), h: s(0.8),
    fontSize: 24,
    color: colors.seafoamGreen,
    align: 'center'
  });

  // Date
  slide.addText('November 2025', {
    x: s(0.5), y: s(4.6), w: s(9), h: s(0.3),
    fontSize: 18,
    color: colors.lightText,
    align: 'center'
  });

  // Company name
  slide.addText('MALLACOOTA BLACK LIP ABALONE CO-OPERATIVE', {
    x: s(0.5), y: s(5.0), w: s(9), h: s(0.4),
    fontSize: 18,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  addSpeakerNotes(slide, "This presentation shifts from premium pricing to scientific differentiation. We're not asking consumers to pay more—we're showing them they're buying fundamentally different products. Our defense is science, not luxury positioning.");
}

// SLIDE 2: The Market Reality
function createSlide2() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Market Reality', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 40,
    bold: true,
    color: colors.lightText
  });

  // Crisis stat
  slide.addShape('rect', {
    x: s(0.5), y: s(1.1), w: s(9), h: s(0.7),
    fill: { color: 'ffffff' },
    line: { color: colors.alertRed, width: s(4) }
  });

  slide.addText('75% Revenue Collapse from Chinese Farmed Flooding', {
    x: s(0.7), y: s(1.25), w: s(8.6), h: s(0.4),
    fontSize: 32,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  // Pricing comparison
  slide.addText('Current Market Pricing:', {
    x: s(0.7), y: s(2.0), w: s(8.6), h: s(0.3),
    fontSize: 24,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('Chinese Farmed: $8-12/kg  •  Wild EX1191: $40-60/kg', {
    x: s(0.7), y: s(2.35), w: s(8.6), h: s(0.3),
    fontSize: 22,
    color: colors.darkText,
    align: 'center'
  });

  // The problem box
  slide.addShape('rect', {
    x: s(0.5), y: s(2.9), w: s(4.4), h: s(1.5),
    fill: { color: 'fff3cd' },
    line: { color: 'ffc107', width: s(2) }
  });

  slide.addText('THE PROBLEM', {
    x: s(0.7), y: s(3.0), w: s(4.0), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: '856404'
  });

  slide.addText('Consumers think\n"abalone is abalone"\n\nWe\'re competing on price\nin the same category', {
    x: s(0.7), y: s(3.35), w: s(4.0), h: s(1.0),
    fontSize: 18,
    color: '856404',
    align: 'center'
  });

  // The solution box
  slide.addShape('rect', {
    x: s(5.1), y: s(2.9), w: s(4.4), h: s(1.5),
    fill: { color: 'd4edda' },
    line: { color: colors.successGreen, width: s(2) }
  });

  slide.addText('THE SOLUTION', {
    x: s(5.3), y: s(3.0), w: s(4.0), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: '155724'
  });

  slide.addText('Prove they\'re buying\nDIFFERENT products\n\nCompete on PURPOSE,\nnot on price', {
    x: s(5.3), y: s(3.35), w: s(4.0), h: s(1.0),
    fontSize: 18,
    color: '155724',
    align: 'center'
  });

  addSpeakerNotes(slide, "The market perceives all abalone as identical. This perception forces us into price competition we cannot win. Our strategy: prove scientifically that wild EX1191 and Chinese farmed are fundamentally different products serving different consumer needs. We defend market share through education, not price cuts.");
}

// SLIDE 3: The Category Shift
function createSlide3() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Category Shift', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 40,
    bold: true,
    color: colors.lightText
  });

  // Left box - Chinese Farmed
  slide.addShape('rect', {
    x: s(0.5), y: s(1.2), w: s(4.4), h: s(3.0),
    fill: { color: 'ffebee' },
    line: { color: colors.alertRed, width: s(3) }
  });

  slide.addText('Chinese Farmed', {
    x: s(0.7), y: s(1.35), w: s(4.0), h: s(0.35),
    fontSize: 26,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(1.0), y: s(1.85), w: s(3.4), h: s(0.5),
    fill: { color: 'ffffff' }
  });

  slide.addText('Basic Protein Food', {
    x: s(1.0), y: s(1.95), w: s(3.4), h: s(0.3),
    fontSize: 24,
    bold: true,
    color: colors.darkText,
    align: 'center'
  });

  slide.addText('✗ Hormones\n✗ Pellet-fed\n✗ Commodity\n✗ Price-driven', {
    x: s(0.9), y: s(2.55), w: s(3.6), h: s(1.4),
    fontSize: 20,
    color: colors.alertRed,
    lineSpacing: 28
  });

  // Right box - Wild EX1191
  slide.addShape('rect', {
    x: s(5.1), y: s(1.2), w: s(4.4), h: s(3.0),
    fill: { color: 'e8f5e9' },
    line: { color: colors.successGreen, width: s(3) }
  });

  slide.addText('Wild EX1191', {
    x: s(5.3), y: s(1.35), w: s(4.0), h: s(0.35),
    fontSize: 26,
    bold: true,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(5.6), y: s(1.85), w: s(3.4), h: s(0.5),
    fill: { color: 'ffffff' }
  });

  slide.addText('Beauty Supplement', {
    x: s(5.6), y: s(1.95), w: s(3.4), h: s(0.3),
    fontSize: 24,
    bold: true,
    color: colors.darkText,
    align: 'center'
  });

  slide.addText('✓ Hormone-free\n✓ Wild algae diet\n✓ Beauty compounds\n✓ Benefit-driven', {
    x: s(5.5), y: s(2.55), w: s(3.6), h: s(1.4),
    fontSize: 20,
    color: colors.successGreen,
    lineSpacing: 28
  });

  // Bottom message
  slide.addShape('rect', {
    x: s(1.5), y: s(4.5), w: s(7), h: s(0.6),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Stop competing on PRICE. Start competing on PURPOSE.', {
    x: s(1.5), y: s(4.55), w: s(7), h: s(0.5),
    fontSize: 26,
    bold: true,
    color: colors.goldAccent,
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "This is the fundamental shift: we're not premium abalone competing against cheap abalone. We're a beauty supplement that happens to be abalone, competing against basic protein. Different categories, different purposes, different consumers. Price becomes irrelevant when you're solving different problems.");
}

// SLIDE 4: Hormone-Free Advantage
function createSlide4() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Science: Hormone-Free Advantage', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  // Split screen comparison
  // Left - Wild
  slide.addShape('rect', {
    x: s(0.5), y: s(1.2), w: s(4.4), h: s(2.5),
    fill: { color: 'e8f5e9' },
    line: { color: colors.successGreen, width: s(3) }
  });

  slide.addText('Wild EX1191', {
    x: s(0.7), y: s(1.35), w: s(4.0), h: s(0.3),
    fontSize: 28,
    bold: true,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addText('✓', {
    x: s(1.5), y: s(1.85), w: s(2.4), h: s(0.8),
    fontSize: 96,
    bold: true,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addText('ZERO\nGrowth Hormones', {
    x: s(0.9), y: s(2.75), w: s(3.6), h: s(0.7),
    fontSize: 24,
    bold: true,
    color: colors.successGreen,
    align: 'center'
  });

  // Right - Farmed
  slide.addShape('rect', {
    x: s(5.1), y: s(1.2), w: s(4.4), h: s(2.5),
    fill: { color: 'ffebee' },
    line: { color: colors.alertRed, width: s(3) }
  });

  slide.addText('Chinese Farmed', {
    x: s(5.3), y: s(1.35), w: s(4.0), h: s(0.3),
    fontSize: 28,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addText('✗', {
    x: s(6.1), y: s(1.85), w: s(2.4), h: s(0.8),
    fontSize: 96,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addText('Industry-Standard\nHormone Use', {
    x: s(5.5), y: s(2.75), w: s(3.6), h: s(0.7),
    fontSize: 24,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  // Consumer question box
  slide.addShape('rect', {
    x: s(1.5), y: s(4.0), w: s(7), h: s(0.8),
    fill: { color: colors.goldAccent }
  });

  slide.addText('Consumer Question:\n"Would you put hormones on your face?"', {
    x: s(1.5), y: s(4.1), w: s(7), h: s(0.6),
    fontSize: 26,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  // Lab test note
  slide.addText('Lab Test: Simple binary hormone panel = Clear YES/NO result', {
    x: s(1.0), y: s(5.0), w: s(8.0), h: s(0.3),
    fontSize: 18,
    italic: true,
    color: colors.oceanBlue,
    align: 'center'
  });

  addSpeakerNotes(slide, "Hormones are the clearest, most binary differentiator. Wild abalone has ZERO growth hormones because they grow naturally over years. Farmed abalone uses hormones to accelerate growth for commercial viability. For beauty consumers, this is non-negotiable. Simple lab test proves it. Game over.");
}

// SLIDE 5: Beauty Compounds Comparison
function createSlide5() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Science: Beauty Compounds Comparison', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  // Table header
  const tableY = s(1.2);
  const rowHeight = s(0.65);

  slide.addShape('rect', {
    x: s(0.5), y: tableY, w: s(9), h: rowHeight,
    fill: { color: colors.oceanBlue }
  });

  slide.addText('Compound', {
    x: s(0.6), y: tableY + s(0.15), w: s(2.5), h: s(0.35),
    fontSize: 18,
    bold: true,
    color: colors.lightText
  });

  slide.addText('Wild EX1191', {
    x: s(3.2), y: tableY + s(0.15), w: s(2.0), h: s(0.35),
    fontSize: 18,
    bold: true,
    color: colors.lightText,
    align: 'center'
  });

  slide.addText('Chinese Farmed', {
    x: s(5.3), y: tableY + s(0.15), w: s(2.0), h: s(0.35),
    fontSize: 18,
    bold: true,
    color: colors.lightText,
    align: 'center'
  });

  slide.addText('Beauty Benefit', {
    x: s(7.4), y: tableY + s(0.15), w: s(2.0), h: s(0.35),
    fontSize: 18,
    bold: true,
    color: colors.lightText,
    align: 'center'
  });

  // Table rows
  const compounds = [
    { compound: '400-600 Da Peptides', wild: '✓ HIGH', farmed: '✗ Low', benefit: 'Optimal absorption' },
    { compound: 'Tyrosinase Inhibitors', wild: '✓ Present', farmed: '✗ Minimal', benefit: 'Skin brightening' },
    { compound: 'MMP Inhibitors', wild: '✓ Strong', farmed: '✗ Weak', benefit: 'Anti-wrinkle' },
    { compound: 'Sulfated Polysaccharides', wild: '✓ High', farmed: '✗ None', benefit: 'Anti-inflammatory' }
  ];

  compounds.forEach((row, i) => {
    const y = tableY + rowHeight + (i * rowHeight);
    const bgColor = i % 2 === 0 ? 'ffffff' : colors.clinicalGray;

    slide.addShape('rect', {
      x: s(0.5), y: y, w: s(9), h: rowHeight,
      fill: { color: bgColor },
      line: { color: 'cccccc', width: s(0.5) }
    });

    slide.addText(row.compound, {
      x: s(0.6), y: y + s(0.12), w: s(2.5), h: s(0.4),
      fontSize: 16,
      color: colors.darkText
    });

    slide.addText(row.wild, {
      x: s(3.2), y: y + s(0.12), w: s(2.0), h: s(0.4),
      fontSize: 16,
      bold: true,
      color: colors.successGreen,
      align: 'center'
    });

    slide.addText(row.farmed, {
      x: s(5.3), y: y + s(0.12), w: s(2.0), h: s(0.4),
      fontSize: 16,
      bold: true,
      color: colors.alertRed,
      align: 'center'
    });

    slide.addText(row.benefit, {
      x: s(7.4), y: y + s(0.12), w: s(2.0), h: s(0.4),
      fontSize: 14,
      color: colors.darkText,
      align: 'center'
    });
  });

  addSpeakerNotes(slide, "Each compound serves a specific beauty function. Wild abalone contains significantly higher levels because of their natural diet and environment. These aren't marginal differences—these are presence vs absence in key categories. Lab testing will quantify these differences and create undeniable proof points for marketing.");
}

// SLIDE 6: Natural Bio-Accumulation
function createSlide6() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  // Header
  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('The Natural Bio-Accumulation Story', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  // Wild diet flowchart
  slide.addText('Wild EX1191:', {
    x: s(0.7), y: s(1.2), w: s(8.6), h: s(0.3),
    fontSize: 24,
    bold: true,
    color: colors.successGreen
  });

  // Flowchart boxes
  slide.addShape('rect', {
    x: s(1.0), y: s(1.7), w: s(2.2), h: s(0.6),
    fill: { color: colors.successGreen }
  });

  slide.addText('Complex\nMarine Algae', {
    x: s(1.0), y: s(1.75), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  slide.addText('→', {
    x: s(3.4), y: s(1.85), w: s(0.6), h: s(0.3),
    fontSize: 36,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(4.2), y: s(1.7), w: s(2.2), h: s(0.6),
    fill: { color: colors.successGreen }
  });

  slide.addText('Wild Abalone\n(56 years)', {
    x: s(4.2), y: s(1.75), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  slide.addText('→', {
    x: s(6.6), y: s(1.85), w: s(0.6), h: s(0.3),
    fontSize: 36,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(7.4), y: s(1.7), w: s(2.2), h: s(0.6),
    fill: { color: colors.successGreen }
  });

  slide.addText('Beauty\nCompounds', {
    x: s(7.4), y: s(1.75), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  // Farmed diet flowchart
  slide.addText('Chinese Farmed:', {
    x: s(0.7), y: s(2.7), w: s(8.6), h: s(0.3),
    fontSize: 24,
    bold: true,
    color: colors.alertRed
  });

  slide.addShape('rect', {
    x: s(1.0), y: s(3.2), w: s(2.2), h: s(0.6),
    fill: { color: colors.alertRed }
  });

  slide.addText('Commercial\nPellets', {
    x: s(1.0), y: s(3.25), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  slide.addText('→', {
    x: s(3.4), y: s(3.35), w: s(0.6), h: s(0.3),
    fontSize: 36,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(4.2), y: s(3.2), w: s(2.2), h: s(0.6),
    fill: { color: colors.alertRed }
  });

  slide.addText('Farmed Abalone\n(6-12 months)', {
    x: s(4.2), y: s(3.25), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  slide.addText('→', {
    x: s(6.6), y: s(3.35), w: s(0.6), h: s(0.3),
    fontSize: 36,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addShape('rect', {
    x: s(7.4), y: s(3.2), w: s(2.2), h: s(0.6),
    fill: { color: colors.alertRed }
  });

  slide.addText('Basic\nNutrition Only', {
    x: s(7.4), y: s(3.25), w: s(2.2), h: s(0.5),
    fontSize: 18,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  // Bottom quote
  slide.addShape('rect', {
    x: s(1.0), y: s(4.2), w: s(8.0), h: s(0.8),
    fill: { color: colors.oceanBlue }
  });

  slide.addText('"You can\'t fake 56 years of wild Australian waters"', {
    x: s(1.0), y: s(4.3), w: s(8.0), h: s(0.6),
    fontSize: 28,
    bold: true,
    italic: true,
    color: colors.goldAccent,
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "This is the story behind the science. Wild abalone bio-accumulate beauty compounds from complex marine algae over decades. Farmed abalone eat pellets designed for fast growth, not compound richness. You literally cannot replicate 56 years of natural accumulation in a 6-month farming cycle. This is our moat.");
}

// Continue with remaining slides...
// For brevity, I'll create simplified versions of the remaining slides

function createSlide7() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Lab Testing Strategy (Gate 1)', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  slide.addText('Investment: $15,000 - $25,000', {
    x: s(1.0), y: s(1.1), w: s(8.0), h: s(0.4),
    fontSize: 28,
    bold: true,
    color: colors.goldAccent
  });

  const tests = [
    '1. Hormone Panel (MOST CRITICAL) - Binary yes/no result',
    '2. Peptide Size Distribution - 400-600 Da range analysis',
    '3. Beauty Compound Screening - Tyrosinase, MMP inhibitors',
    '4. Comparative Analysis - Side-by-side wild vs farmed'
  ];

  tests.forEach((test, i) => {
    slide.addText('✓ ' + test, {
      x: s(1.2), y: s(1.7 + i * 0.5), w: s(7.6), h: s(0.4),
      fontSize: 20,
      color: colors.darkText
    });
  });

  slide.addText('Timeline: 6-8 Weeks', {
    x: s(1.0), y: s(3.9), w: s(3.8), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addShape('rect', {
    x: s(1.0), y: s(4.4), w: s(8.0), h: s(0.6),
    fill: { color: 'd4edda' },
    line: { color: colors.successGreen, width: s(2) }
  });

  slide.addText('GO/NO-GO: Clear superiority in 3+ markers proves differentiation', {
    x: s(1.2), y: s(4.5), w: s(7.6), h: s(0.4),
    fontSize: 20,
    bold: true,
    color: '155724',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "Lab testing is our foundation. $15-25K investment proves the science. Hormone panel is non-negotiable - this alone differentiates us. Peptide and compound analysis adds depth. We need clear superiority in at least 3 markers. If results don't show meaningful differentiation, we pivot to sustainability story with minimal sunk cost.");
}

function createSlide8() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Consumer Messaging Framework', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  // Three pillars
  const pillars = [
    { num: '1', title: 'Hormone-Free Beauty', text: 'Clean, pure, safe for your skin' },
    { num: '2', title: 'Scientifically Different', text: 'Lab-proven compound superiority' },
    { num: '3', title: 'EX1191 Legacy', text: '56 years wild, fully traceable' }
  ];

  pillars.forEach((pillar, i) => {
    const x = s(0.5 + i * 3.2);

    slide.addShape('rect', {
      x: x, y: s(1.3), w: s(2.8), h: s(2.5),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(3) }
    });

    slide.addShape('ellipse', {
      x: x + s(0.9), y: s(1.5), w: s(1.0), h: s(1.0),
      fill: { color: colors.goldAccent }
    });

    slide.addText(pillar.num, {
      x: x + s(0.9), y: s(1.5), w: s(1.0), h: s(1.0),
      fontSize: 48,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });

    slide.addText(pillar.title, {
      x: x + s(0.2), y: s(2.7), w: s(2.4), h: s(0.4),
      fontSize: 20,
      bold: true,
      color: colors.oceanBlue,
      align: 'center'
    });

    slide.addText(pillar.text, {
      x: x + s(0.2), y: s(3.15), w: s(2.4), h: s(0.5),
      fontSize: 16,
      color: colors.darkText,
      align: 'center'
    });
  });

  // QR code strategy
  slide.addShape('rect', {
    x: s(1.5), y: s(4.2), w: s(7.0), h: s(0.7),
    fill: { color: colors.goldAccent }
  });

  slide.addText('QR Code Strategy: Direct link to lab results on every package', {
    x: s(1.5), y: s(4.3), w: s(7.0), h: s(0.5),
    fontSize: 24,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "Three pillars only—don't overcomplicate. Hormone-free resonates with clean beauty movement. Scientific proof gives credibility. EX1191 provides traceability and heritage. QR codes on packaging link directly to lab results—radical transparency. Let consumers verify our claims themselves.");
}

function createSlide9() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Influencer Testing Approach', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  slide.addText('Phase 1: 5-8 Beauty Micro-Influencers (50K-150K followers)', {
    x: s(0.7), y: s(1.1), w: s(8.6), h: s(0.3),
    fontSize: 22,
    color: colors.oceanBlue
  });

  const steps = [
    'Give them BOTH products (wild + farmed) with lab results',
    'Message: "I tested both - here\'s the science"',
    'Focus: Hormone-free + measurable beauty benefits',
    'Content format: Before/after, side-by-side comparison',
    'Authenticity: Let them share genuine findings'
  ];

  steps.forEach((step, i) => {
    slide.addText('▸ ' + step, {
      x: s(1.0), y: s(1.6 + i * 0.45), w: s(8.0), h: s(0.35),
      fontSize: 18,
      color: colors.darkText
    });
  });

  slide.addShape('rect', {
    x: s(1.0), y: s(3.9), w: s(8.0), h: s(0.6),
    fill: { color: 'd4edda' },
    line: { color: colors.successGreen, width: s(2) }
  });

  slide.addText('Success Metric: >5% engagement on differentiation posts', {
    x: s(1.2), y: s(4.0), w: s(7.6), h: s(0.4),
    fontSize: 20,
    bold: true,
    color: '155724',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "Influencer testing validates market messaging. We give them both products and lab results—let them compare and educate their audience. This isn't sponsorship, it's scientific demonstration. If they can't generate engagement around the differences, our differentiation story needs work. 5% engagement threshold proves market interest.");
}

function createSlide10() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Market Positioning Strategy', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  // STOP box
  slide.addShape('rect', {
    x: s(0.5), y: s(1.2), w: s(4.4), h: s(1.8),
    fill: { color: 'ffebee' },
    line: { color: colors.alertRed, width: s(3) }
  });

  slide.addText('STOP ✗', {
    x: s(0.7), y: s(1.35), w: s(4.0), h: s(0.35),
    fontSize: 28,
    bold: true,
    color: colors.alertRed,
    align: 'center'
  });

  slide.addText('Fighting in seafood aisle\nfor "premium abalone"\n\nCompeting on price\nagainst farming scale', {
    x: s(0.9), y: s(1.85), w: s(3.6), h: s(1.0),
    fontSize: 18,
    color: colors.alertRed,
    align: 'center'
  });

  // START box
  slide.addShape('rect', {
    x: s(5.1), y: s(1.2), w: s(4.4), h: s(1.8),
    fill: { color: 'e8f5e9' },
    line: { color: colors.successGreen, width: s(3) }
  });

  slide.addText('START ✓', {
    x: s(5.3), y: s(1.35), w: s(4.0), h: s(0.35),
    fontSize: 28,
    bold: true,
    color: colors.successGreen,
    align: 'center'
  });

  slide.addText('Creating "Beauty Foods"\ncategory\n\nRetail placement near\ncollagen/supplements', {
    x: s(5.5), y: s(1.85), w: s(3.6), h: s(1.0),
    fontSize: 18,
    color: colors.successGreen,
    align: 'center'
  });

  // Price comparison
  slide.addText('Price Comparison Shift:', {
    x: s(0.7), y: s(3.3), w: s(8.6), h: s(0.3),
    fontSize: 22,
    bold: true,
    color: colors.oceanBlue
  });

  slide.addText('Not: Wild abalone ($60) vs Farmed abalone ($12)\n\nBut: Wild abalone ($60) vs Collagen drink ($30-50/month)', {
    x: s(1.0), y: s(3.7), w: s(8.0), h: s(0.8),
    fontSize: 20,
    color: colors.darkText,
    align: 'center'
  });

  addSpeakerNotes(slide, "Positioning is everything. We lose in the seafood aisle competing against farmed. We win in the beauty supplements aisle where our price is competitive with collagen drinks and peptide supplements. Same product, different shelf, different comparison set, different outcome. This isn't about raising prices—it's about changing context.");
}

// Continuing with remaining slides...

function createSlide11() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Defense Against Price Competition', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  slide.addText('When they say: "Why pay more?"', {
    x: s(1.0), y: s(1.1), w: s(8.0), h: s(0.4),
    fontSize: 26,
    bold: true,
    italic: true,
    color: colors.oceanBlue
  });

  const responses = [
    '"Same price point, different benefits"',
    '"Protein vs Beauty - your choice"',
    '"Lab results don\'t lie - scan the QR code"'
  ];

  responses.forEach((response, i) => {
    slide.addShape('rect', {
      x: s(1.5), y: s(1.8 + i * 0.6), w: s(7.0), h: s(0.5),
      fill: { color: colors.successGreen }
    });

    slide.addText(response, {
      x: s(1.7), y: s(1.85 + i * 0.6), w: s(6.6), h: s(0.4),
      fontSize: 20,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });
  });

  // Comparison table
  slide.addText('Simple Comparison:', {
    x: s(1.0), y: s(3.6), w: s(8.0), h: s(0.3),
    fontSize: 20,
    bold: true,
    color: colors.oceanBlue
  });

  // Table headers
  slide.addShape('rect', {
    x: s(1.0), y: s(4.0), w: s(8.0), h: s(0.4),
    fill: { color: colors.oceanBlue }
  });

  ['', 'Chinese Farmed', 'Wild EX1191'].forEach((header, i) => {
    slide.addText(header, {
      x: s(1.0 + i * 2.67), y: s(4.05), w: s(2.67), h: s(0.3),
      fontSize: 16,
      bold: true,
      color: colors.lightText,
      align: 'center'
    });
  });

  // Table rows
  const rows = [
    ['Price', '$$', '$$'],
    ['Protein', '✓', '✓'],
    ['Hormones', '✗ YES', '✓ NO'],
    ['Beauty Compounds', '✗', '✓✓✓'],
    ['Traceability', '✗', '✓ EX1191']
  ];

  rows.forEach((row, i) => {
    const bgColor = i % 2 === 0 ? 'ffffff' : colors.clinicalGray;

    slide.addShape('rect', {
      x: s(1.0), y: s(4.4 + i * 0.35), w: s(8.0), h: s(0.35),
      fill: { color: bgColor },
      line: { color: 'cccccc', width: s(0.5) }
    });

    row.forEach((cell, j) => {
      const color = j === 0 ? colors.darkText : (j === 1 ? colors.alertRed : colors.successGreen);
      slide.addText(cell, {
        x: s(1.0 + j * 2.67), y: s(4.45 + i * 0.35), w: s(2.67), h: s(0.25),
        fontSize: 14,
        bold: j > 0,
        color: color,
        align: 'center'
      });
    });
  });

  addSpeakerNotes(slide, "This is our armor against price objections. We're not asking for more money—we're offering different value. Protein? Both have it. Price? Same ballpark. Everything else? Completely different. The comparison table makes this visual and undeniable. QR codes let consumers verify instantly.");
}

function createSlide12() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Financial Projections (Volume Recovery)', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  slide.addText('Based on market share defense, NOT price increases', {
    x: s(1.0), y: s(1.0), w: s(8.0), h: s(0.3),
    fontSize: 18,
    italic: true,
    color: colors.oceanBlue,
    align: 'center'
  });

  // Timeline
  const timeline = [
    { year: 'Year 1', target: 'Defend 50% of lost volume', volume: '6 containers', revenue: '$2.5M' },
    { year: 'Year 2', target: 'Recover 75% through differentiation', volume: '9 containers', revenue: '$3.8M' },
    { year: 'Year 3', target: 'Full recovery + beauty category', volume: '12+ containers', revenue: '$5M+' }
  ];

  timeline.forEach((item, i) => {
    slide.addShape('rect', {
      x: s(0.7), y: s(1.6 + i * 1.0), w: s(8.6), h: s(0.85),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(2) }
    });

    slide.addText(item.year, {
      x: s(0.9), y: s(1.7 + i * 1.0), w: s(1.5), h: s(0.3),
      fontSize: 24,
      bold: true,
      color: colors.goldAccent
    });

    slide.addText(item.target, {
      x: s(2.6), y: s(1.7 + i * 1.0), w: s(4.0), h: s(0.3),
      fontSize: 18,
      color: colors.darkText
    });

    slide.addText(item.volume, {
      x: s(6.8), y: s(1.75 + i * 1.0), w: s(1.2), h: s(0.2),
      fontSize: 16,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(item.revenue, {
      x: s(8.2), y: s(1.7 + i * 1.0), w: s(1.0), h: s(0.3),
      fontSize: 20,
      bold: true,
      color: colors.successGreen
    });
  });

  // Bottom message
  slide.addShape('rect', {
    x: s(1.5), y: s(4.7), w: s(7.0), h: s(0.6),
    fill: { color: colors.goldAccent }
  });

  slide.addText('NO PRICE INCREASE NEEDED - Growth through volume recovery', {
    x: s(1.5), y: s(4.75), w: s(7.0), h: s(0.5),
    fontSize: 22,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "Financial model is conservative and realistic. We're not projecting premium pricing—we're projecting volume recovery through differentiation. Year 1: stem the bleeding by defending half our lost share. Year 2: recover three-quarters through proven differentiation. Year 3: full recovery plus new beauty-focused consumers. ROI based on keeping customers we've lost, not charging more.");
}

function createSlide13() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Implementation Timeline', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  const phases = [
    { weeks: 'Weeks 1-2', task: 'Lab testing commissioned', desc: 'Select independent lab, submit samples' },
    { weeks: 'Weeks 3-8', task: 'Results analysis & verification', desc: 'Review data, create comparison infographics' },
    { weeks: 'Weeks 9-12', task: 'Influencer seeding', desc: '5-8 micro-influencers test both products' },
    { weeks: 'Weeks 13-16', task: 'Retail category discussions', desc: 'Position in beauty foods section' },
    { weeks: 'Month 5+', task: 'Full market rollout', desc: 'QR codes, packaging, education campaign' }
  ];

  phases.forEach((phase, i) => {
    const y = s(1.2 + i * 0.75);

    // Timeline dot
    slide.addShape('ellipse', {
      x: s(0.8), y: y + s(0.05), w: s(0.3), h: s(0.3),
      fill: { color: colors.goldAccent },
      line: { color: 'ffffff', width: s(2) }
    });

    // Content box
    slide.addShape('rect', {
      x: s(1.3), y: y, w: s(8.2), h: s(0.65),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(2) }
    });

    slide.addText(phase.weeks, {
      x: s(1.5), y: y + s(0.08), w: s(1.8), h: s(0.25),
      fontSize: 16,
      bold: true,
      color: colors.goldAccent
    });

    slide.addText(phase.task, {
      x: s(3.5), y: y + s(0.08), w: s(3.0), h: s(0.25),
      fontSize: 18,
      bold: true,
      color: colors.oceanBlue
    });

    slide.addText(phase.desc, {
      x: s(3.5), y: y + s(0.35), w: s(5.8), h: s(0.2),
      fontSize: 14,
      color: colors.darkText
    });
  });

  addSpeakerNotes(slide, "Timeline is aggressive but achievable. Lab results in 8 weeks provide ammunition for everything else. Influencer testing validates messaging in weeks 9-12. Retail discussions happen in parallel. Full rollout by month 5 means we're defending market share before losing more ground. Every week counts when revenue is collapsing.");
}

function createSlide14() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Risk Mitigation', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  const risks = [
    {
      risk: 'Lab results don\'t show difference',
      mitigation: 'Only $15K invested, pivot to sustainability story',
      likelihood: 'LOW'
    },
    {
      risk: 'Consumers don\'t care about hormones',
      mitigation: 'Focus on beauty benefits & peptides instead',
      likelihood: 'LOW'
    },
    {
      risk: 'Chinese competitors copy claims',
      mitigation: 'EX1191 traceability + continuous testing',
      likelihood: 'MEDIUM'
    }
  ];

  risks.forEach((item, i) => {
    const y = s(1.3 + i * 1.2);

    slide.addShape('rect', {
      x: s(0.7), y: y, w: s(8.6), h: s(1.0),
      fill: { color: 'ffffff' },
      line: { color: colors.oceanBlue, width: s(2) }
    });

    slide.addText('Risk:', {
      x: s(0.9), y: y + s(0.1), w: s(1.0), h: s(0.25),
      fontSize: 16,
      bold: true,
      color: colors.alertRed
    });

    slide.addText(item.risk, {
      x: s(2.0), y: y + s(0.1), w: s(6.8), h: s(0.25),
      fontSize: 18,
      color: colors.darkText
    });

    slide.addText('Mitigation:', {
      x: s(0.9), y: y + s(0.45), w: s(1.2), h: s(0.2),
      fontSize: 14,
      bold: true,
      color: colors.successGreen
    });

    slide.addText(item.mitigation, {
      x: s(2.0), y: y + s(0.45), w: s(6.2), h: s(0.2),
      fontSize: 16,
      color: colors.darkText
    });

    slide.addShape('rect', {
      x: s(8.4), y: y + s(0.35), w: s(0.8), h: s(0.3),
      fill: { color: item.likelihood === 'LOW' ? colors.successGreen : 'ffc107' }
    });

    slide.addText(item.likelihood, {
      x: s(8.4), y: y + s(0.37), w: s(0.8), h: s(0.26),
      fontSize: 12,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });
  });

  addSpeakerNotes(slide, "Risks are manageable. If lab results fail to show differentiation, we've only invested $15K and can pivot to sustainability messaging. If hormone messaging doesn't resonate, we have beauty compounds as backup. If competitors copy claims, our EX1191 traceability and ongoing testing creates authenticity they can't match. Every risk has a controlled exit.");
}

function createSlide15() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.pearlWhite };

  slide.addShape('rect', {
    x: s(0), y: s(0), w: s(10), h: s(0.8),
    fill: { type: 'solid', color: colors.oceanBlue }
  });

  slide.addText('Immediate Next Steps', {
    x: s(0.5), y: s(0.2), w: s(9), h: s(0.4),
    fontSize: 38,
    bold: true,
    color: colors.lightText
  });

  const steps = [
    'Approve $15-25K for comprehensive lab testing',
    'Send BOTH wild and farmed samples to independent lab',
    'Focus on hormone + beauty compound analysis',
    'Create comparison infographic from results',
    'Test messaging with 5 Singapore beauty influencers'
  ];

  steps.forEach((step, i) => {
    slide.addShape('ellipse', {
      x: s(0.8), y: s(1.4 + i * 0.75), w: s(0.5), h: s(0.5),
      fill: { color: colors.goldAccent }
    });

    slide.addText((i + 1).toString(), {
      x: s(0.8), y: s(1.4 + i * 0.75), w: s(0.5), h: s(0.5),
      fontSize: 28,
      bold: true,
      color: 'ffffff',
      align: 'center',
      valign: 'middle'
    });

    slide.addText(step, {
      x: s(1.5), y: s(1.5 + i * 0.75), w: s(7.8), h: s(0.35),
      fontSize: 20,
      color: colors.darkText
    });
  });

  addSpeakerNotes(slide, "Five concrete steps to start immediately. Board approval for lab testing is step one—everything else follows. We test both products to ensure comparative credibility. Hormone and compound analysis are non-negotiable. Results become infographics for consumer education. Singapore influencers validate messaging before wider rollout. Clear, actionable, achievable.");
}

function createSlide16() {
  const slide = pres.addSlide();
  slide.background = { fill: colors.oceanBlue };

  // Main headline
  slide.addText('We Don\'t Need Higher Prices.\nWe Need Different Positioning.', {
    x: s(0.7), y: s(1.2), w: s(8.6), h: s(1.0),
    fontSize: 42,
    bold: true,
    color: colors.goldAccent,
    align: 'center'
  });

  // Three key points
  const points = [
    'Wild EX1191 = Beauty Supplement\n(happens to be abalone)',
    'Chinese Farmed = Protein Food\n(just abalone)',
    'Let the science speak.\nLet consumers choose.'
  ];

  points.forEach((point, i) => {
    slide.addShape('rect', {
      x: s(1.5), y: s(2.5 + i * 0.75), w: s(7.0), h: s(0.6),
      fill: { color: 'ffffff', transparency: 10 }
    });

    slide.addText(point, {
      x: s(1.7), y: s(2.55 + i * 0.75), w: s(6.6), h: s(0.5),
      fontSize: 20,
      bold: true,
      color: colors.lightText,
      align: 'center',
      valign: 'middle'
    });
  });

  // Call to action
  slide.addShape('rect', {
    x: s(2.0), y: s(4.8), w: s(6.0), h: s(0.5),
    fill: { color: colors.goldAccent }
  });

  slide.addText('Ready to prove we\'re different?', {
    x: s(2.0), y: s(4.85), w: s(6.0), h: s(0.4),
    fontSize: 28,
    bold: true,
    color: 'ffffff',
    align: 'center',
    valign: 'middle'
  });

  addSpeakerNotes(slide, "This is the bottom line. We're not premium abalone—we're a beauty supplement. They're not cheap abalone—they're basic protein. Different products, different purposes, different value propositions. Science creates the separation. Transparency builds trust. Education drives choice. We defend our market through differentiation, not discounting. Ready to invest $15-25K and prove it?");
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

// Save the presentation
pres.writeFile({ fileName: 'mallacoota-differentiation-deck.pptx' })
  .then(() => {
    console.log('✓ Differentiation deck created successfully: mallacoota-differentiation-deck.pptx');
  })
  .catch(err => {
    console.error('Error creating presentation:', err);
  });
