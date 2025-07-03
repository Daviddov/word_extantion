// taskpane.js - קובץ ראשי המטפל ב-API ובממשק המשתמש

// הפונקציות הגלובליות
window.hideCitationInput = hideCitationInput;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("extractText").onclick = extractAndProcessWithAPI;
  }
});

// פיצול ציטוטים גדולים לציטוטים קטנים יותר
function splitLargeCitations(citations) {
  const refinedCitations = [];
  
  for (const citation of citations) {
    if (!citation.matches || citation.matches.length === 0) {
      continue;
    }
    
    // אם יש התאמה אחת בלבד, נשאיר את הציטוט כמו שהוא
    if (citation.matches.length === 1) {
      refinedCitations.push({
        ...citation,
        primaryMatch: citation.matches[0]
      });
      continue;
    }
    
    // אם יש מספר התאמות, ניצור ציטוט נפרד לכל התאמה
    citation.matches.forEach((match, index) => {
      // נמצא את המיקום הטוב ביותר להתאמה הזו בתוך הטקסט
      const cleanCitationText = stripHtmlTags(citation.text);
      const cleanMatchText = stripHtmlTags(match.matchedText);
      
      // נחפש את המיקום של ההתאמה בתוך הציטוט
      const matchPosition = cleanCitationText.indexOf(cleanMatchText.trim());
      
      let startPos = citation.startIChar;
      let searchText = cleanMatchText;
      
      // אם מצאנו את המיקום, נתאים את הפוזיציה
      if (matchPosition >= 0) {
        startPos = citation.startIChar + matchPosition;
        // נלקח חלק מהטקסט סביב ההתאמה לחיפוש טוב יותר
        const contextStart = Math.max(0, matchPosition - 10);
        const contextEnd = Math.min(cleanCitationText.length, matchPosition + cleanMatchText.length + 10);
        searchText = cleanCitationText.substring(contextStart, contextEnd);
      }
      
      refinedCitations.push({
        startIChar: startPos,
        endIChar: startPos + searchText.length,
        text: searchText,
        matches: [match],
        primaryMatch: match,
        originalCitation: citation
      });
    });
  }
  
  return refinedCitations;
}

// פונקציית עזר
function stripHtmlTags(html) {
  if (!html) return '';
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  return tmp.textContent || tmp.innerText || '';
};

// הפונקציה הראשית
async function extractAndProcessWithAPI() {
  const statusDiv = document.getElementById('status');
  const button = document.getElementById('extractText');
  
  button.disabled = true;
  statusDiv.innerHTML = '<div class="loading">מחלץ טקסט מהמסמך...</div>';
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      context.load(body, 'text');
      await context.sync();
      
      const documentText = body.text;
      
      if (!documentText || documentText.trim().length === 0) {
        throw new Error('המסמך ריק או לא נמצא טקסט');
      }
      
      // עיבוד הטקסט עם API של דיקטה בחלקים
      await processDictaAPIInChunks(documentText, context);
    });
  } catch (error) {
    console.error('Error:', error);
    statusDiv.innerHTML = `<div class="error">שגיאה: ${error.message}</div>`;
  } finally {
    button.disabled = false;
  }
}

// עיבוד עם API של דיקטה בחלקים
async function processDictaAPIInChunks(text, wordContext) {
  const statusDiv = document.getElementById('status');
  const MAX_CHUNK_SIZE = 9500; // השארנו מקום בטוח מתחת ל-10K
  
  try {
    // חלוקת הטקסט לחלקים
    const chunks = splitTextIntoChunks(text, MAX_CHUNK_SIZE);
    statusDiv.innerHTML = `<div class="loading">מעבד ${chunks.length} חלקים של הטקסט...</div>`;
    
    let allCitations = [];
    let totalCharactersProcessed = 0;
    
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      statusDiv.innerHTML = `<div class="loading">מעבד חלק ${i + 1} מתוך ${chunks.length}...</div>`;
      
      try {
        const chunkCitations = await processChunkWithAPI(chunk, totalCharactersProcessed);
        if (chunkCitations && chunkCitations.length > 0) {
          allCitations = allCitations.concat(chunkCitations);
        }
      } catch (chunkError) {
        console.warn(`שגיאה בחלק ${i + 1}:`, chunkError);
        // ממשיכים לחלק הבא גם אם יש שגיאה
      }
      
      totalCharactersProcessed += chunk.length;
      
      // הפסקה קצרה בין חלקים כדי לא להעמיס על השרת
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    if (allCitations.length === 0) {
      statusDiv.innerHTML = '<div class="error">לא נמצאו ציטוטים בטקסט</div>';
      return;
    }
    
    statusDiv.innerHTML = '<div class="loading">מוסיף ציטוטים למסמך...</div>';
    
    // פיצול ציטוטים גדולים לציטוטים קטנים יותר
    const refinedCitations = splitLargeCitations(allCitations);
    
    // הוספת הציטוטים למסמך (קריאה לפונקציה מהקובץ השני)
    await window.insertCitationsToDocument(refinedCitations, wordContext);
    
    statusDiv.innerHTML = `<div class="success">🎉 הושלם! נוספו ${refinedCitations.length} ציטוטים למסמך</div>`;
    
  } catch (error) {
    console.error('Error processing with API:', error);
    
    if (error.message.includes('cors') || error.message.includes('CORS')) {
      statusDiv.innerHTML = `
        <div class="error">בעיית CORS - האתר חוסם בקשות חיצוניות</div>
        <div style="margin-top: 10px;">
          <button onclick="showManualInput()" style="padding: 8px 15px; background-color: #28a745; color: white; border: none; border-radius: 4px;">
            הוסף ציטוטים ידנית
          </button>
        </div>
      `;
    } else {
      statusDiv.innerHTML = `<div class="error">שגיאה: ${error.message}</div>`;
    }
  }
}

// חלוקת הטקסט לחלקים
function splitTextIntoChunks(text, maxSize) {
  if (text.length <= maxSize) {
    return [text];
  }
  
  const chunks = [];
  let currentIndex = 0;
  
  while (currentIndex < text.length) {
    let endIndex = currentIndex + maxSize;
    
    // אם לא הגענו לסוף הטקסט, ננסה לחתוך במקום טבעי (רווח, נקודה, פסיק)
    if (endIndex < text.length) {
      const searchStart = Math.max(currentIndex + maxSize - 200, currentIndex);
      const chunkToSearch = text.substring(searchStart, endIndex + 200);
      
      // חיפוש נקודת חיתוך טובה (פסקה, משפט, מילה)
      const breakPoints = ['\n\n', '. ', '.\n', ', ', ' '];
      let bestBreakPoint = -1;
      
      for (const breakPoint of breakPoints) {
        const lastIndex = chunkToSearch.lastIndexOf(breakPoint);
        if (lastIndex > 0) {
          bestBreakPoint = searchStart + lastIndex + breakPoint.length;
          break;
        }
      }
      
      if (bestBreakPoint > currentIndex) {
        endIndex = bestBreakPoint;
      }
    }
    
    chunks.push(text.substring(currentIndex, Math.min(endIndex, text.length)));
    currentIndex = endIndex;
  }
  
  return chunks;
}

// עיבוד חלק יחיד
async function processChunkWithAPI(chunkText, offsetPosition) {
  try {
    // קריאה ראשונה - חיפוש התאמות
    const firstResponse = await fetch('https://cors-anywhere.herokuapp.com/https://talmudfinder-2-0.loadbalancer.dicta.org.il/TalmudFinder/api/markpsukim', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        mode: "tanakh",
        thresh: 0,
        fdirectonly: false,
        data: chunkText
      })
    });
    
    if (!firstResponse.ok) {
      throw new Error(`שגיאה בקריאה ראשונה: ${firstResponse.status}`);
    }
    
    const firstData = await firstResponse.json();
    
    if (!firstData.downloadId || !firstData.results || firstData.results.length === 0) {
      return []; // לא נמצאו התאמות בחלק הזה
    }
    
    // קריאה שנייה - קבלת הציטוטים המעוצבים
    const secondResponse = await fetch('https://cors-anywhere.herokuapp.com/https://talmudfinder-2-0.loadbalancer.dicta.org.il/TalmudFinder/api/parsetogroups?smin=22&smax=10000', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        downloadId: firstData.downloadId,
        results: firstData.results,
        allText: firstData.allText,
        failedPrefixes: firstData.failedPrefixes,
        keepredundant: true
      })
    });
    
    if (!secondResponse.ok) {
      throw new Error(`שגיאה בקריאה שנייה: ${secondResponse.status}`);
    }
    
    const citations = await secondResponse.json();
    
    if (!citations || citations.length === 0) {
      return [];
    }
    
    // התאמת מיקומים לטקסט המלא
    return citations.map(citation => ({
      ...citation,
      startIChar: citation.startIChar + offsetPosition,
      endIChar: citation.endIChar + offsetPosition
    }));
    
  } catch (error) {
    console.error('Error processing chunk:', error);
    return [];
  }
}

// הצגת חלון הוספה ידנית
function showManualInput() {
  let existingInput = document.getElementById('citation-input-container');
  if (existingInput) {
    existingInput.style.display = 'block';
    return;
  }
  
  const container = document.createElement('div');
  container.id = 'citation-input-container';
  container.style.cssText = `
    margin: 15px 0;
    padding: 15px;
    border: 1px solid #ddd;
    border-radius: 4px;
    background-color: #f8f9fa;
  `;
  
  container.innerHTML = `
    <h4 style="margin-top: 0;">הוסף ציטוטים ידנית:</h4>
    <div style="margin-bottom: 10px;">
      <label style="display: block; font-weight: bold; margin-bottom: 5px;">הטקסט למציאה:</label>
      <input type="text" id="search-text" 
             placeholder="לדוגמה: תוֹלְדוֹת הַשָּׁמַיִם וְהָאָרֶץ"
             style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; direction: rtl;">
    </div>
    <div style="margin-bottom: 10px;">
      <label style="display: block; font-weight: bold; margin-bottom: 5px;">הציטוט:</label>
      <textarea id="citation-text" 
                placeholder="לדוגמה: בראשית ב, ד: אֵלֶּה תוֹלְדוֹת הַשָּׁמַיִם וְהָאָרֶץ בְּהִבָּרְאָם"
                style="width: 100%; height: 80px; resize: vertical; direction: rtl; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></textarea>
    </div>
    <div>
      <button onclick="window.insertFromInput()" 
              style="margin-right: 10px; padding: 8px 15px; background-color: #28a745; color: white; border: none; border-radius: 4px;">
        הוסף למסמך
      </button>
      <button onclick="hideCitationInput()" 
              style="padding: 8px 15px; background-color: #6c757d; color: white; border: none; border-radius: 4px;">
        ביטול
      </button>
    </div>
  `;
  
  const mainContainer = document.querySelector('.container');
  mainContainer.appendChild(container);
  
  setTimeout(() => {
    document.getElementById('search-text').focus();
  }, 100);
}

// הסתרת חלון הקלט
function hideCitationInput() {
  const container = document.getElementById('citation-input-container');
  if (container) {
    container.style.display = 'none';
  }
}