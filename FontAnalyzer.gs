/**
 * Font Analyzer for Google Slides
 * This script analyzes fonts used in a Google Slides presentation
 * and creates a summary table on the first slide
 */

/**
 * Main function to analyze fonts in the presentation
 * @param {number} endSlideNumber - The slide number to stop analysis at (optional)
 */
function analyzeFonts(endSlideNumber) {
  try {
    // Get the active presentation
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    
    // Determine the range of slides to analyze
    const totalSlides = slides.length;
    const lastSlideToAnalyze = endSlideNumber && endSlideNumber <= totalSlides 
      ? endSlideNumber 
      : totalSlides;
    
    // Collect font information from each slide
    const fontData = [];
    
    for (let i = 0; i < lastSlideToAnalyze; i++) {
      const slide = slides[i];
      const slideNumber = i + 1;
      const slideFonts = analyzeSlidesFonts(slide, slideNumber);
      
      if (slideFonts.length > 0) {
        fontData.push({
          slideNumber: slideNumber,
          fonts: slideFonts
        });
      }
    }
    
    // Create summary table on the first slide
    if (fontData.length > 0) {
      createFontSummaryTable(slides[0], fontData);
      
      // Show success message
      SlidesApp.getUi().alert(
        'Font Analysis Complete', 
        `Analyzed ${lastSlideToAnalyze} slides. Summary table added to the first slide.`,
        SlidesApp.getUi().ButtonSet.OK
      );
    } else {
      SlidesApp.getUi().alert(
        'No Fonts Found', 
        'No text elements with fonts were found in the analyzed slides.',
        SlidesApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    SlidesApp.getUi().alert(
      'Error', 
      `An error occurred: ${error.toString()}`,
      SlidesApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Analyze fonts in a single slide
 * @param {Slide} slide - The slide to analyze
 * @param {number} slideNumber - The slide number
 * @returns {Array} Array of font information objects
 */
function analyzeSlidesFonts(slide, slideNumber) {
  const fontMap = new Map();
  const pageElements = slide.getPageElements();
  
  pageElements.forEach(element => {
    // Check if element contains text
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE ||
        element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
      
      try {
        const textRange = element.asShape().getText();
        if (textRange) {
          extractFontInfo(textRange, fontMap);
        }
      } catch (e) {
        // Handle table elements
        if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
          const table = element.asTable();
          const numRows = table.getNumRows();
          const numCols = table.getNumColumns();
          
          for (let row = 0; row < numRows; row++) {
            for (let col = 0; col < numCols; col++) {
              try {
                const cell = table.getCell(row, col);
                const textRange = cell.getText();
                if (textRange) {
                  extractFontInfo(textRange, fontMap);
                }
              } catch (cellError) {
                // Skip problematic cells
              }
            }
          }
        }
      }
    }
  });
  
  // Convert map to array
  const fontsArray = [];
  fontMap.forEach((sizes, fontFamily) => {
    fontsArray.push({
      fontFamily: fontFamily,
      sizes: Array.from(sizes).sort((a, b) => a - b)
    });
  });
  
  return fontsArray;
}

/**
 * Extract font information from a text range
 * @param {TextRange} textRange - The text range to analyze
 * @param {Map} fontMap - Map to store font information
 */
function extractFontInfo(textRange, fontMap) {
  const runs = textRange.getRuns();
  
  runs.forEach(run => {
    const textStyle = run.getTextStyle();
    const fontFamily = textStyle.getFontFamily() || 'Default';
    const fontSize = textStyle.getFontSize();
    
    if (fontSize) {
      if (!fontMap.has(fontFamily)) {
        fontMap.set(fontFamily, new Set());
      }
      fontMap.get(fontFamily).add(fontSize.getMagnitude());
    }
  });
}

/**
 * Create a summary table on the first slide
 * @param {Slide} firstSlide - The first slide where the table will be added
 * @param {Array} fontData - Array of font data for each slide
 */
function createFontSummaryTable(firstSlide, fontData) {
  // Calculate table dimensions
  let totalRows = 1; // Header row
  fontData.forEach(slideData => {
    totalRows += slideData.fonts.length;
  });
  
  // Create table
  const table = firstSlide.insertTable(totalRows, 3, 10, 10, 500, 20 * totalRows);
  
  // Style header row
  const headerRow = table.getRow(0);
  headerRow.getCell(0).getText().setText('Slide #').getTextStyle()
    .setBold(true).setFontSize(12);
  headerRow.getCell(1).getText().setText('Font Family').getTextStyle()
    .setBold(true).setFontSize(12);
  headerRow.getCell(2).getText().setText('Font Sizes (pt)').getTextStyle()
    .setBold(true).setFontSize(12);
  
  // Fill table with font data
  let currentRow = 1;
  fontData.forEach(slideData => {
    const slideNumber = slideData.slideNumber;
    
    slideData.fonts.forEach((fontInfo, index) => {
      const row = table.getRow(currentRow);
      
      // Only show slide number for the first font of each slide
      if (index === 0) {
        row.getCell(0).getText().setText(`Slide ${slideNumber}`);
      }
      
      row.getCell(1).getText().setText(fontInfo.fontFamily);
      row.getCell(2).getText().setText(fontInfo.sizes.join(', '));
      
      currentRow++;
    });
  });
  
  // Style the table
  for (let i = 0; i < totalRows; i++) {
    for (let j = 0; j < 3; j++) {
      const cell = table.getCell(i, j);
      cell.getFill().setSolidFill('#f8f9fa');
      cell.getBorder().getTop().setWeight(1);
      cell.getBorder().getBottom().setWeight(1);
      cell.getBorder().getLeft().setWeight(1);
      cell.getBorder().getRight().setWeight(1);
    }
  }
}

/**
 * Menu function to run the analyzer with a dialog
 */
function showAnalyzerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('dialog')
      .setWidth(400)
      .setHeight(200);
  SlidesApp.getUi()
      .showModalDialog(html, 'Font Analyzer Settings');
}

/**
 * Create menu when the presentation is opened
 */
function onOpen() {
  SlidesApp.getUi()
      .createMenu('Font Analyzer')
      .addItem('Analyze Fonts', 'showAnalyzerDialog')
      .addItem('Analyze All Slides', 'analyzeAllSlides')
      .addToUi();
}

/**
 * Analyze all slides (menu shortcut)
 */
function analyzeAllSlides() {
  analyzeFonts();
}