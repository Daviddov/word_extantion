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
    
    // הוספת הציטוטים למסמך
    await insertCitationsToDocument(allCitations, wordContext);
    
    statusDiv.innerHTML = `<div class="success">🎉 הושלם! נוספו ${allCitations.length} ציטוטים למסמך</div>`;
    
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

// הוספת ציטוטים למסמך עם footnotes אמיתיים
async function insertCitationsToDocument(citations, context) {
  try {
    let addedCitations = 0;
    let footnoteCounter = 1;
    
    // מיון הציטוטים לפי מיקום בטקסט (מהסוף להתחלה)
    const sortedCitations = citations.sort((a, b) => b.startIChar - a.startIChar);
    
    const body = context.document.body;
    
    for (const citation of sortedCitations) {
      if (citation.matches && citation.matches.length > 0) {
        const statusDiv = document.getElementById('status');
        statusDiv.innerHTML = `<div class="loading">מוסיף ציטוט ${addedCitations + 1} מתוך ${citations.length}...</div>`;
        
        // הכנת טקסט הציטוט
        const originalText = stripHtmlTags(citation.text);
        const citationTexts = citation.matches.map(match => {
          const cleanMatchText = stripHtmlTags(match.matchedText);
          return `${match.verseDispHeb}: ${cleanMatchText}`;
        });
        
        const footnoteText = citationTexts.join('; ');
        
        // חיפוש הטקסט במסמך
        const searchResults = body.search(originalText, { 
          matchCase: false, 
          matchWildcards: false,
          matchWholeWord: false
        });
        context.load(searchResults, 'items');
        await context.sync();
        
        if (searchResults.items.length > 0) {
          const foundRange = searchResults.items[0];
          
          // הוספת footnote עם hyperlink (MSO style)
          await insertMSOFootnote(foundRange, footnoteText, footnoteCounter, context);
          
          addedCitations++;
          footnoteCounter++;
        }
      }
    }
    
    if (addedCitations === 0) {
      throw new Error('לא הצליח למצוא את הטקסטים במסמך להוספת ציטוטים');
    }
    
  } catch (error) {
    console.error('Error inserting citations:', error);
    throw new Error(`שגיאה בהוספת ציטוטים: ${error.message}`);
  }
}

// הוספת footnote בסגנון MSO עם hyperlinks
async function insertMSOFootnote(range, footnoteText, footnoteNumber, context) {
  try {
    // הוספת הקישור בטקסט הראשי
    const footnoteRefHtml = `<a href="#_ftn${footnoteNumber}" name="_ftnref${footnoteNumber}"><span style="mso-footnote-id:ftn${footnoteNumber}; vertical-align:super; color:blue; text-decoration:underline;">[${footnoteNumber}]</span></a>`;
    
    // הכנסת HTML עם insertHtml
    range.insertHtml(footnoteRefHtml, Word.InsertLocation.after);
    await context.sync();
    
    // חיפוש או יצירת אזור footnotes
    const body = context.document.body;
    let footnotesSection = await findOrCreateFootnotesSection(body, context);
    
    // הוספת ה-footnote עצמו
    const footnoteHtml = `
    <div style="mso-element:footnote;" id="ftn${footnoteNumber}">
        <p style="text-align:right; direction:rtl; font-size:10pt; margin:0; padding:2px 0;">
            <a href="#_ftnref${footnoteNumber}" name="_ftn${footnoteNumber}">
                <span style="mso-footnote-id:ftn${footnoteNumber}; color:blue; text-decoration:underline;">
                    [${footnoteNumber}]
                </span>
            </a>
            <span style="color:#666666; margin-right:5px;">${footnoteText}</span>
        </p>
    </div>`;
    
    footnotesSection.insertHtml(footnoteHtml, Word.InsertLocation.end);
    await context.sync();
    
  } catch (error) {
    console.warn('שגיאה בהוספת MSO footnote, משתמש בשיטה פשוטה:', error);
    
    // שיטה פשוטה יותר אם HTML לא עובד
    const linkText = `[${footnoteNumber}]`;
    const insertedRange = range.insertText(linkText, Word.InsertLocation.after);
    insertedRange.font.superscript = true;
    insertedRange.font.color = '#0066cc';
    insertedRange.hyperlink = `#_ftn${footnoteNumber}`;
    
    await context.sync();
    
    // הוספת footnote פשוט
    const body = context.document.body;
    const footnoteParagraph = body.insertParagraph('', Word.InsertLocation.end);
    
    const numberRange = footnoteParagraph.insertText(`[${footnoteNumber}] `, Word.InsertLocation.start);
    numberRange.font.color = '#0066cc';
    numberRange.hyperlink = `#_ftnref${footnoteNumber}`;
    
    const textRange = footnoteParagraph.insertText(footnoteText, Word.InsertLocation.end);
    textRange.font.size = 10;
    textRange.font.color = '#666666';
    
    footnoteParagraph.alignment = Word.Alignment.right;
    footnoteParagraph.leftIndent = 18;
    
    await context.sync();
  }
}

// מציאה או יצירה של אזור footnotes
async function findOrCreateFootnotesSection(body, context) {
  // חיפוש אזור footnotes קיים
  const footnotesSearch = body.search('mso-element:footnote-list', { matchCase: false });
  context.load(footnotesSearch, 'items');
  await context.sync();
  
  if (footnotesSearch.items.length > 0) {
    return footnotesSearch.items[0];
  }
  

  
  // יצירת container לfootnotes
  const footnotesContainer = body.insertParagraph('', Word.InsertLocation.end);
  footnotesContainer.insertHtml('<div style="mso-element:footnote-list;"></div>', Word.InsertLocation.start);
  
  await context.sync();
  return footnotesContainer;
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
      <label style="display: block; font-weight: bold; margin-bottom: 5px;">הציטוט:</label>
      <textarea id="citation-text" 
                placeholder="לדוגמה: בראשית ב, ד: אֵלֶּה תוֹלְדוֹת הַשָּׁמַיִם וְהָאָרֶץ בְּהִבָּרְאָם"
                style="width: 100%; height: 80px; resize: vertical; direction: rtl; padding: 8px; border: 1px solid #ccc; border-radius: 4px;"></textarea>
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

// הוספה מהקלט הידני
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
      
      // מציאת מספר footnote הבא
      const footnoteSearch = body.search(/\[(\d+)\]/, { matchWildcards: true });
      context.load(footnoteSearch, 'items');
      await context.sync();
      
      let footnoteNumber = footnoteSearch.items.length + 1;
      
      const foundRange = searchResults.items[0];
      
      // הוספת footnote עם MSO style
      await insertMSOFootnote(foundRange, citationText, footnoteNumber, context);
      
      document.getElementById('status').innerHTML = '<div class="success">✅ הציטוט נוסף בהצלחה!</div>';
      
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