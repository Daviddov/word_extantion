// הפונקציות הגלובליות
window.insertFromInput = insertFromInput;
window.hideCitationInput = hideCitationInput;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("extractText").onclick = extractAndProcessWithAPI;
  }
});

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
      
      // עיבוד הטקסט עם API של דיקטה
      await processDictaAPI(documentText, context);
    });
  } catch (error) {
    console.error('Error:', error);
    statusDiv.innerHTML = `<div class="error">שגיאה: ${error.message}</div>`;
  } finally {
    button.disabled = false;
  }
}
let proxy = 'https://dictaproxy-production.up.railway.app/';
proxy = 'https://carnelian-carnation-red.glitch.me/';
// עיבוד עם API של דיקטה
async function processDictaAPI(text, wordContext) {
  const statusDiv = document.getElementById('status');
  
  try {
    statusDiv.innerHTML = '<div class="loading">שלב 1: מחפש התאמות בטקסט...</div>';
    let url = 'https://talmudfinder-2-0.loadbalancer.dicta.org.il/TalmudFinder/api/markpsukim';
    // קריאה ראשונה - חיפוש התאמות
    const firstResponse = await fetch(proxy + 'markpsukim', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        mode: "tanakh",
        thresh: 0,
        fdirectonly: false,
        data: text
      })
    });
    
    if (!firstResponse.ok) {
      throw new Error(`שגיאה בקריאה ראשונה: ${firstResponse.status}`);
    }
    
    const firstData = await firstResponse.json();
    
    if (!firstData.downloadId || !firstData.results || firstData.results.length === 0) {
      statusDiv.innerHTML = '<div class="error">לא נמצאו התאמות בטקסט</div>';
      return;
    }
    
    statusDiv.innerHTML = '<div class="loading">שלב 2: מעבד את התוצאות...</div>';
    url = 'https://talmudfinder-2-0.loadbalancer.dicta.org.il/TalmudFinder/api';
    // קריאה שנייה - קבלת הציטוטים המעוצבים
    const secondResponse = await fetch(proxy + 'parsetogroups', {
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
      statusDiv.innerHTML = '<div class="error">לא נמצאו ציטוטים להוספה</div>';
      return;
    }
    
    statusDiv.innerHTML = '<div class="loading">שלב 3: מוסיף ציטוטים למסמך...</div>';
    
    // הוספת הציטוטים למסמך כהערות תוך הטקסט
    await insertInlineCitationsToDocument(citations, wordContext);
    
    statusDiv.innerHTML = `<div class="success">🎉 הושלם! נוספו ${citations.length} ציטוטים למסמך</div>`;
    
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

// הוספת ציטוטים תוך הטקסט (inline citations)
async function insertInlineCitationsToDocument(citations, context) {
  try {
    let addedCitations = 0;
    
    // מיון הציטוטים לפי מיקום בטקסט (מהסוף להתחלה כדי לא לשבש את המיקומים)
    const sortedCitations = citations.sort((a, b) => b.startIChar - a.startIChar);
    
    for (const citation of sortedCitations) {
      if (citation.matches && citation.matches.length > 0) {
        const statusDiv = document.getElementById('status');
        statusDiv.innerHTML = `<div class="loading">מוסיף ציטוט ${addedCitations + 1} מתוך ${citations.length}...</div>`;
        
        // הכנת טקסט הציטוט
        const originalText = stripHtmlTags(citation.text);
        const citationTexts = citation.matches.map(match => {
          const cleanMatchText = stripHtmlTags(match.matchedText);
          return `${match.verseDispHeb}`;
        });
        
        const inlineCitation = ` (${citationTexts.join('; ')})`;
        
        // חיפוש הטקסט במסמך
        const body = context.document.body;
        const searchResults = body.search(originalText, { 
          matchCase: false, 
          matchWildcards: false
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length > 0) {
          const foundRange = searchResults.items[0];
          
          // טעינת מאפייני הגופן של הטקסט המקורי
          context.load(foundRange.font, 'size');
          await context.sync();
          
          // הוספת הציטוט מיד אחרי הטקסט המקורי
          const citationRange = foundRange.insertText(inlineCitation, Word.InsertLocation.after);
          
          // עיצוב הציטוט - קטן יותר וצבע אפור
          const originalSize = foundRange.font.size || 12; // ברירת מחדל אם לא נמצא
          citationRange.font.size = originalSize - 2; // קטן יותר מהטקסט הרגיל
          citationRange.font.color = '#666666'; // אפור
          citationRange.font.italic = true; // נטוי
          
          await context.sync();
          addedCitations++;
        }
      }
    }
    
    if (addedCitations === 0) {
      throw new Error('לא הצליח למצוא את הטקסטים במסמך להוספת ציטוטים');
    }
    
  } catch (error) {
    console.error('Error inserting inline citations:', error);
    throw new Error(`שגיאה בהוספת ציטוטים: ${error.message}`);
  }
}

// פונקציות עזר
function stripHtmlTags(html) {
  if (!html) return '';
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  return tmp.textContent || tmp.innerText || '';
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
      <label style="display: block; font-weight: bold; margin-bottom: 5px;">הציטוט (יופיע בסוגריים):</label>
      <textarea id="citation-text" 
                placeholder="לדוגמה: בראשית ב, ד"
                style="width: 100%; height: 60px; resize: vertical; direction: rtl; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></textarea>
    </div>
    <div>
      <button onclick="insertFromInput()" 
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

// הוספה מהקלט הידני כציטוט תוך הטקסט
async function insertFromInput() {
  const searchText = document.getElementById('search-text')?.value.trim();
  const citationText = document.getElementById('citation-text')?.value.trim();
  
  if (!searchText || !citationText) {
    document.getElementById('status').innerHTML = '<div class="error">נא להזין גם טקסט לחיפוש וגם ציטוט</div>';
    return;
  }
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      
      // חיפוש הטקסט
      const searchResults = body.search(searchText, { 
        matchCase: false, 
        matchWildcards: false
      });
      context.load(searchResults, 'items');
      await context.sync();
      
      if (searchResults.items.length === 0) {
        document.getElementById('status').innerHTML = '<div class="error">לא נמצא הטקסט במסמך</div>';
        return;
      }
      
      const foundRange = searchResults.items[0];
      
      // טעינת מאפייני הגופן
      context.load(foundRange.font, 'size');
      await context.sync();
      
      // הוספת הציטוט מיד אחרי הטקסט
      const inlineCitation = ` (${citationText})`;
      const citationRange = foundRange.insertText(inlineCitation, Word.InsertLocation.after);
      
      // עיצוב הציטוט
      const originalSize = foundRange.font.size || 12;
      citationRange.font.size = originalSize - 2;
      citationRange.font.color = '#666666';
      citationRange.font.italic = true;
      
      await context.sync();
      
      document.getElementById('status').innerHTML = '<div class="success">✅ הציטוט נוסף בהצלחה תוך הטקסט!</div>';
      
      // ניקוי השדות
      document.getElementById('search-text').value = '';
      document.getElementById('citation-text').value = '';
      
      setTimeout(() => {
        hideCitationInput();
      }, 2000);
    });
  } catch (error) {
    console.error('Error inserting citation:', error);
    document.getElementById('status').innerHTML = `<div class="error">שגיאה: ${error.message}</div>`;
  }
}

// הסתרת חלון הקלט
function hideCitationInput() {
  const container = document.getElementById('citation-input-container');
  if (container) {
    container.style.display = 'none';
  }
}