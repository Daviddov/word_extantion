// Make functions globally accessible
window.insertFromInput = insertFromInput;
window.hideCitationInput = hideCitationInput;Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("extractText").onclick = extractAndSearch;
  }
});

function showDictaWebsite() {
  const iframeContainer = document.getElementById('iframe-container');
  const iframe = document.getElementById('dicta-frame');
  
  iframe.src = 'https://citation.dicta.org.il/';
  iframeContainer.style.display = 'block';
}

async function extractAndSearch() {
  const statusDiv = document.getElementById('status');
  const resultsDiv = document.getElementById('results');
  const button = document.getElementById('extractText');
  
  button.disabled = true;
  statusDiv.innerHTML = '<div class="loading">מחלץ טקסט מהמסמך...</div>';
  resultsDiv.style.display = 'none';
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      context.load(body, 'text');
      await context.sync();
      
      const documentText = body.text;
      
      if (!documentText || documentText.trim().length === 0) {
        throw new Error('המסמך ריק או לא נמצא טקסט');
      }
      
      statusDiv.innerHTML = '<div class="loading">פותח אתר דיקטה ושולח טקסט...</div>';
      
      showDictaWithText(documentText);
      
      statusDiv.innerHTML = '<div class="success">האתר נפתח עם הטקסט מהמסמך. בחר ציטוטים מהאתר.</div>';
    });
  } catch (error) {
    console.error('Error:', error);
    statusDiv.innerHTML = `<div class="error">שגיאה: ${error.message}</div>`;
  } finally {
    button.disabled = false;
  }
}

function showDictaWithText(text) {
  const iframeContainer = document.getElementById('iframe-container');
  const iframe = document.getElementById('dicta-frame');
  
  iframe.src = 'https://citation.dicta.org.il/';
  iframeContainer.style.display = 'block';
  
  iframe.onload = () => {
    setTimeout(() => {
      try {
        const message = {
          type: 'FILL_TEXT',
          text: text
        };
        iframe.contentWindow.postMessage(message, 'https://citation.dicta.org.il');
      } catch (error) {
        console.log('Cannot send message to iframe:', error);
      }
    }, 2000);
  };
  
  addCopyButton();
}

function addCopyButton() {
  const existingButton = document.getElementById('copyFromSite');
  if (existingButton) return;
  
  const button = document.createElement('button');
  button.id = 'copyFromSite';
  button.textContent = 'העתק ציטוט נבחר מהאתר';
  button.onclick = copySelectedCitation;
  
  const container = document.querySelector('.container');
  container.appendChild(button);
}

async function copySelectedCitation() {
  const iframe = document.getElementById('dicta-frame');
  if (!iframe) return;
  
  // יצירת אלמנט input להזנת ציטוט
  showCitationInput();
}

function showCitationInput() {
  // בדיקה אם כבר קיים
  let existingInput = document.getElementById('citation-input-container');
  if (existingInput) {
    existingInput.style.display = 'block';
    return;
  }
  
  // יצירת container להזנת הציטוט
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
    <h4 style="margin-top: 0;">הוסף ציטוט:</h4>
    <textarea id="citation-text" 
              placeholder="הדבק כאן את הציטוט מהאתר..."
              style="width: 100%; height: 80px; resize: vertical; margin-bottom: 10px;"></textarea>
    <div>
      <button onclick="insertFromInput()" style="margin-right: 10px;">הוסף למסמך</button>
      <button onclick="hideCitationInput()">ביטול</button>
    </div>
  `;
  
  const mainContainer = document.querySelector('.container');
  mainContainer.appendChild(container);
  
  // פוקוס על הטקסט
  setTimeout(() => {
    document.getElementById('citation-text').focus();
  }, 100);
}

function hideCitationInput() {
  const container = document.getElementById('citation-input-container');
  if (container) {
    container.style.display = 'none';
  }
}

async function insertFromInput() {
  const textArea = document.getElementById('citation-text');
  const citationText = textArea ? textArea.value.trim() : '';
  
  if (!citationText) {
    document.getElementById('status').innerHTML = '<div class="error">נא להזין טקסט ציטוט</div>';
    return;
  }
  
  await insertCitationToDocument({text: citationText});
  
  // ניקוי השדה
  textArea.value = '';
  hideCitationInput();
}

async function insertCitationToDocument(citation) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      
      const citationText = citation.text || citation.source || 'ציטוט';
      
      selection.insertText(citationText, Word.InsertLocation.end);
      
      await context.sync();
      
      document.getElementById('status').innerHTML = '<div class="success">הציטוט נוסף למסמך!</div>';
    });
  } catch (error) {
    console.error('Error inserting citation:', error);
    document.getElementById('status').innerHTML = `<div class="error">שגיאה בהוספת הציטוט: ${error.message}</div>`;
  }
}