/**
 * Font Analyzer for Google Slides - Version 2
 * Enhanced version with better font size detection
 */

/**
 * Main function to analyze fonts in the presentation
 * @param {number} endSlideNumber - The slide number to stop analysis at (optional)
 */
function analyzeFonts(endSlideNumber) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    
    const totalSlides = slides.length;
    const lastSlideToAnalyze = endSlideNumber && endSlideNumber <= totalSlides 
      ? endSlideNumber 
      : totalSlides;
    
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
    
    if (fontData.length > 0) {
      createFontSummaryTable(slides[0], fontData);
      
      SlidesApp.getUi().alert(
        'Анализ завершен', 
        `Проанализировано ${lastSlideToAnalyze} слайдов. Таблица добавлена на первый слайд.`,
        SlidesApp.getUi().ButtonSet.OK
      );
    } else {
      SlidesApp.getUi().alert(
        'Шрифты не найдены', 
        'В проанализированных слайдах не найдено текстовых элементов со шрифтами.',
        SlidesApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    SlidesApp.getUi().alert(
      'Ошибка', 
      `Произошла ошибка: ${error.toString()}`,
      SlidesApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Analyze fonts in a single slide
 */
function analyzeSlidesFonts(slide, slideNumber) {
  const fontMap = new Map();
  const pageElements = slide.getPageElements();
  
  pageElements.forEach(element => {
    const elementType = element.getPageElementType();
    
    // Process shapes and text boxes
    if (elementType === SlidesApp.PageElementType.SHAPE) {
      try {
        const shape = element.asShape();
        if (shape.getText) {
          const textRange = shape.getText();
          if (textRange && textRange.asString().trim() !== '') {
            extractFontInfoV2(textRange, fontMap);
          }
        }
      } catch (e) {
        console.log('Error processing shape:', e);
      }
    }
    
    // Process tables
    else if (elementType === SlidesApp.PageElementType.TABLE) {
      try {
        const table = element.asTable();
        const numRows = table.getNumRows();
        const numCols = table.getNumColumns();
        
        for (let row = 0; row < numRows; row++) {
          for (let col = 0; col < numCols; col++) {
            try {
              const cell = table.getCell(row, col);
              const textRange = cell.getText();
              if (textRange && textRange.asString().trim() !== '') {
                extractFontInfoV2(textRange, fontMap);
              }
            } catch (cellError) {
              console.log('Error processing cell:', cellError);
            }
          }
        }
      } catch (e) {
        console.log('Error processing table:', e);
      }
    }
  });
  
  // Convert map to array and filter out empty entries
  const fontsArray = [];
  fontMap.forEach((sizes, fontFamily) => {
    if (sizes.size > 0) {
      fontsArray.push({
        fontFamily: fontFamily,
        sizes: Array.from(sizes).sort((a, b) => a - b)
      });
    }
  });
  
  return fontsArray;
}

/**
 * Enhanced font extraction with better error handling
 */
function extractFontInfoV2(textRange, fontMap) {
  try {
    // Get all paragraphs
    const paragraphs = textRange.getParagraphs();
    
    paragraphs.forEach(paragraph => {
      const paragraphText = paragraph.getRange();
      
      // Try to get runs
      try {
        const runs = paragraphText.getRuns();
        
        runs.forEach(run => {
          if (run.getLength() === 0) return;
          
          const textStyle = run.getTextStyle();
          let fontFamily = 'Default';
          let fontSize = 11; // Default size
          
          // Get font family
          try {
            fontFamily = textStyle.getFontFamily() || 'Default';
          } catch (e) {
            console.log('Could not get font family');
          }
          
          // Get font size - try multiple approaches
          try {
            const fontSizeObj = textStyle.getFontSize();
            if (fontSizeObj) {
              fontSize = fontSizeObj.getMagnitude ? fontSizeObj.getMagnitude() : fontSizeObj;
            }
          } catch (e) {
            // If that fails, use default
            console.log('Using default font size');
          }
          
          // Add to map
          if (!fontMap.has(fontFamily)) {
            fontMap.set(fontFamily, new Set());
          }
          fontMap.get(fontFamily).add(Math.round(fontSize));
        });
      } catch (e) {
        // If runs fail, try to get style from the whole paragraph
        try {
          const textStyle = paragraphText.getTextStyle();
          const fontFamily = textStyle.getFontFamily() || 'Default';
          const fontSize = 11; // Default if we can't get it
          
          if (!fontMap.has(fontFamily)) {
            fontMap.set(fontFamily, new Set());
          }
          fontMap.get(fontFamily).add(fontSize);
        } catch (e2) {
          console.log('Could not extract font info from paragraph');
        }
      }
    });
  } catch (e) {
    console.log('Error in extractFontInfoV2:', e);
  }
}

/**
 * Create a summary table on the first slide
 */
function createFontSummaryTable(firstSlide, fontData) {
  // Calculate table dimensions
  let totalRows = 1; // Header row
  fontData.forEach(slideData => {
    totalRows += slideData.fonts.length;
  });
  
  // Create table with better positioning
  const presentation = SlidesApp.getActivePresentation();
  const slideWidth = presentation.getPageWidth();
  const tableWidth = Math.min(500, slideWidth - 20);
  const rowHeight = 25;
  const table = firstSlide.insertTable(totalRows, 3, 10, 10, tableWidth, rowHeight * totalRows);
  
  // Style header row
  const headerRow = table.getRow(0);
  headerRow.getCell(0).getText().setText('Слайд №').getTextStyle()
    .setBold(true).setFontSize(12);
  headerRow.getCell(1).getText().setText('Шрифт').getTextStyle()
    .setBold(true).setFontSize(12);
  headerRow.getCell(2).getText().setText('Размеры (pt)').getTextStyle()
    .setBold(true).setFontSize(12);
  
  // Fill table with font data
  let currentRow = 1;
  fontData.forEach(slideData => {
    const slideNumber = slideData.slideNumber;
    
    slideData.fonts.forEach((fontInfo, index) => {
      if (currentRow < totalRows) {
        const row = table.getRow(currentRow);
        
        // Only show slide number for the first font of each slide
        if (index === 0) {
          row.getCell(0).getText().setText(`Слайд ${slideNumber}`);
        }
        
        row.getCell(1).getText().setText(fontInfo.fontFamily);
        row.getCell(2).getText().setText(fontInfo.sizes.join(', ') + ' pt');
        
        currentRow++;
      }
    });
  });
  
  // Style the table
  try {
    for (let i = 0; i < totalRows; i++) {
      for (let j = 0; j < 3; j++) {
        const cell = table.getCell(i, j);
        if (i === 0) {
          // Header row
          cell.getFill().setSolidFill('#4285f4');
          cell.getText().getTextStyle().setForegroundColor('#ffffff');
        } else {
          // Data rows
          cell.getFill().setSolidFill(i % 2 === 0 ? '#f8f9fa' : '#ffffff');
        }
      }
    }
  } catch (e) {
    console.log('Could not style table cells:', e);
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
      .showModalDialog(html, 'Настройки анализа шрифтов');
}

/**
 * Create menu when the presentation is opened
 */
function onOpen() {
  SlidesApp.getUi()
      .createMenu('Анализ шрифтов')
      .addItem('Анализировать шрифты', 'showAnalyzerDialog')
      .addItem('Анализировать все слайды', 'analyzeAllSlides')
      .addToUi();
}

/**
 * Analyze all slides (menu shortcut)
 */
function analyzeAllSlides() {
  analyzeFonts();
}

/**
 * Debug function to test font detection
 */
function debugFontDetection() {
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSlides()[0];
  const elements = slide.getPageElements();
  
  console.log(`Found ${elements.length} elements on first slide`);
  
  elements.forEach((element, index) => {
    console.log(`Element ${index}: ${element.getPageElementType()}`);
    
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      try {
        const shape = element.asShape();
        const text = shape.getText();
        console.log(`Text content: "${text.asString()}"`);
        
        const runs = text.getRuns();
        runs.forEach((run, runIndex) => {
          const style = run.getTextStyle();
          console.log(`Run ${runIndex}: Font=${style.getFontFamily()}, Size=${style.getFontSize()}`);
        });
      } catch (e) {
        console.log(`Error processing element ${index}:`, e);
      }
    }
  });
}