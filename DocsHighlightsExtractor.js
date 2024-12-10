function processMultipleDocsHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getActiveSheet();
  
  // Get document links
  const links = inputSheet.getRange('A:A')
    .getValues()
    .flat()
    .filter(link => link !== '' && link.includes('docs.google.com'));

  if (links.length === 0) {
    SpreadsheetApp.getUi().alert('No valid Google Doc links found in column A');
    return;
  }

  // Process each document
  links.forEach(link => {
    try {
      // Extract document ID and get content
      const docId = link.match(/\/d\/(.*?)\/|$/)[1];
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      
      // Store colors and their phrases
      const colorData = new Map(); // Map to store color -> phrases[]
      const totalElements = body.getNumChildren();

      // Process each element
      for (let i = 0; i < totalElements; i++) {
        const element = body.getChild(i);
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const text = element.asParagraph().getText();
          const textStyle = element.asParagraph().editAsText();
          
          let currentText = '';
          let currentColor = null;
          
          // Process each character
          for (let j = 0; j < text.length; j++) {
            const bgColor = textStyle.getBackgroundColor(j);
            
            if (bgColor !== null) {
              if (currentColor === bgColor) {
                currentText += text[j];
              } else {
                // Save previous text if exists
                if (currentText) {
                  if (!colorData.has(currentColor)) {
                    colorData.set(currentColor, []);
                  }
                  const trimmedText = currentText.trim();
                  if (trimmedText) {
                    colorData.get(currentColor).push(trimmedText);
                  }
                }
                currentText = text[j];
                currentColor = bgColor;
              }
            } else {
              // No background color, save accumulated text
              if (currentText) {
                if (!colorData.has(currentColor)) {
                  colorData.set(currentColor, []);
                }
                const trimmedText = currentText.trim();
                if (trimmedText) {
                  colorData.get(currentColor).push(trimmedText);
                }
                currentText = '';
                currentColor = null;
              }
            }
          }
          
          // Handle end of paragraph
          if (currentText) {
            if (!colorData.has(currentColor)) {
              colorData.set(currentColor, []);
            }
            const trimmedText = currentText.trim();
            if (trimmedText) {
              colorData.get(currentColor).push(trimmedText);
            }
          }
        }
      }

      // Get or create sheet based on document name
      let sheetName = doc.getName().substring(0, 30);
      let resultSheet = ss.getSheetByName(sheetName);
      
      // If sheet exists, clear it; if not, create it
      if (resultSheet) {
        resultSheet.clear();
      } else {
        resultSheet = ss.insertSheet(sheetName);
      }

      if (colorData.size > 0) {
        // Get all colors except null
        const colors = Array.from(colorData.keys()).filter(color => color !== null);
        
        // Find maximum number of phrases for any color
        const maxPhrases = Math.max(...colors.map(color => colorData.get(color).length));
        
        // Create data arrays for each color
        const data = [];
        for (let i = 0; i < maxPhrases; i++) {
          const row = [];
          colors.forEach(color => {
            const phrases = colorData.get(color);
            row.push(phrases[i] || '');
          });
          data.push(row);
        }

        // Set headers and color them
        resultSheet.getRange(1, 1, 1, colors.length).setValues([colors]);
        colors.forEach((color, index) => {
          resultSheet.getRange(1, index + 1).setBackground(color);
        });

        // Write phrases
        if (data.length > 0) {
          resultSheet.getRange(2, 1, data.length, colors.length).setValues(data);
        }

        // Format sheet
        resultSheet.autoResizeColumns(1, colors.length);
        resultSheet.setColumnWidths(1, colors.length, 200); // Set width for all columns
        resultSheet.getRange(1, 1, data.length + 1, colors.length).setWrap(true);
      } else {
        resultSheet.getRange('A1').setValue('No highlighted text found in this document');
      }
      
    } catch (error) {
      Logger.log(`Error processing document ${link}: ${error.toString()}`);
      console.error(error);
    }
  });
}
