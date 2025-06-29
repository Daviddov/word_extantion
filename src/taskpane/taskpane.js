// taskpane.js - גרסה מתוקנת וקצרה

// חיפוש בספרייה
async function searchSefaria(query) {
  if (!query || query.length < 2) return null;
  
  try {
    const encodedQuery = encodeURIComponent(query);
    const response = await fetch(`https://www.sefaria.org/api/name/${encodedQuery}?ref_only=1&limit=10`, {
      method: 'GET',
      mode: 'cors',
      headers: { 'Accept': 'application/json' }
    });
    
    return response.ok ? await response.json() : null;
  } catch (error) {
    console.error("Search error:", error);
    return null;
  }
}

// קבלת טקסט מספרייה
async function getSefariaText(ref) {
  try {
    const encodedRef = encodeURIComponent(ref);
    const response = await fetch(`https://www.sefaria.org/api/texts/${encodedRef}?stripItags=1`, {
      method: 'GET',
      mode: 'cors',
      headers: { 'Accept': 'application/json' }
    });
    
    if (response.ok) {
      return await response.json();
    }
  } catch (error) {
    console.error("Text fetch error:", error);
  }
  
  // נתונים לדוגמה במקרה של שגיאה
  return {
    he: [`טקסט לדוגמה עבור ${ref}`],
    heRef: ref,
    ref: ref,
    url: ref.replace(/\s+/g, ".")
  };
}

// ניקוי HTML tags
function stripHtml(text) {
  return text ? text.replace(/<[^>]*>/g, "").replace(/&nbsp;/g, " ").trim() : "";
}

// הצגת תוצאות חיפוש בזמן אמת
async function showDropdown(query) {
  const dropdown = document.getElementById("dropdownOptions");
  
  if (!query || query.length < 2) {
    dropdown.style.display = "none";
    return;
  }

  dropdown.innerHTML = '<div class="dropdown-option">מחפש...</div>';
  dropdown.style.display = "block";

  const results = await searchSefaria(query);
  if (!results?.completion_objects?.length) {
    dropdown.innerHTML = '<div class="dropdown-option">לא נמצאו תוצאות</div>';
    return;
  }

  let options = "";


  // אם יש is_ref ונמצאו heAddressExamples - הצג אפשרויות כתובות
  if (results.is_ref && results.heAddressExamples && results.heAddressExamples.length > 0) {
    const bookTitle = results.ref || results.book || query;
    const bookTitleHe = results.completions[0] || bookTitle; // השם העברי של הספר
    
    // הצג את הספר עצמו
    options += `<div class="dropdown-option" onclick="selectFromDropdown('${results.key || results.ref}', '${bookTitleHe}')">
      <div class="book-title">${bookTitleHe}</div>
      <div class="book-subtitle">הספר כולו</div>
    </div>`;

    // הצג דוגמאות כתובות
    const maxExamples = Math.min(results.heAddressExamples.length, 5);
    for (let i = 0; i < maxExamples; i++) {
      const heAddress = results.heAddressExamples[i];
      const enAddress = results.addressExamples?.[i] || heAddress;
      const fullRefHe = `${bookTitleHe} ${heAddress}`; // שם עברי מלא
      const fullKeyEn = `${results.key || results.ref} ${enAddress}`; // מפתח אנגלי לAPI
      
      options += `<div class="dropdown-option" onclick="selectFromDropdown('${fullKeyEn}', '${fullRefHe}')">
        <div class="book-title">${fullRefHe}</div>
        <div class="book-subtitle">פסוק/מקטע ספציפי</div>
      </div>`;
    }
  } else {
    // הצג רשימת ספרים רגילה
    options = results.completion_objects
      .filter(obj => obj.type === "ref" && obj.is_primary)
      .slice(0, 8)
      .map(opt => {
        const titleHe = opt.title || "ללא שם"; // השם העברי
        const keyEn = opt.key || opt.title || ""; // המפתח האנגלי
        return `<div class="dropdown-option" onclick="selectFromDropdown('${keyEn}', '${titleHe}')">
          <div class="book-title">${titleHe}</div>
          <div class="book-subtitle">${opt.type}</div>
        </div>`;
      }).join("");
  }

  dropdown.innerHTML = options || '<div class="dropdown-option">לא נמצאו ספרים מתאימים</div>';
}

// בחירת אפשרות מהתפריט הנפתח
function selectFromDropdown(key, displayName) {
  const searchInput = document.getElementById("sourceSearch");
  searchInput.value = displayName;
  searchInput.focus();
  hideDropdown();
}

// הסתרת התפריט הנפתח
function hideDropdown() {
  setTimeout(() => {
    const dropdown = document.getElementById("dropdownOptions");
    if (dropdown) {
      dropdown.style.display = "none";
    }
  }, 200);
}

// סריקת טקסט למקורות
async function scanTextForSources() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text?.trim()) {
        document.getElementById("output").innerHTML = "לא נבחר טקסט";
        return;
      }

      const response = await fetch("https://www.sefaria.org/api/find-refs/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          text: { body: selection.text, title: "" },
          lang: "he"
        })
      });

      if (!response.ok) throw new Error("שגיאה בשרת");

      const result = await response.json();
      const refs = result?.body?.results || [];

      if (!refs.length) {
        document.getElementById("output").innerHTML = "לא נמצאו מקורות";
        return;
      }

      let linksCreated = 0;
      for (const ref of refs) {
        try {
          const refKey = ref.refs[0];
          const refData = result.body.refData[refKey];
          if (!refData) continue;

          const url = "https://www.sefaria.org/" + (refData.url || refKey.replace(/\s+/g, "."));
          const refText = selection.text.substring(ref.startChar, ref.endChar);

          const searchResults = context.document.body.search(refText);
          searchResults.load("items");
          await context.sync();

          if (searchResults.items.length > 0) {
            searchResults.items[0].hyperlink = url;
            searchResults.items[0].font.color = "#0066cc";
            linksCreated++;
          }
        } catch (error) {
          console.error("Error processing ref:", error);
        }
      }

      document.getElementById("output").innerHTML = 
        linksCreated > 0 ? `הומרו ${linksCreated} מקורות לקישורים` : "לא ניתן היה ליצור קישורים";
      document.getElementById("output").style.color = linksCreated > 0 ? "green" : "red";

      await context.sync();
    });
  } catch (error) {
    console.error("Scan error:", error);
    document.getElementById("output").innerHTML = `שגיאה: ${error.message}`;
  }
}

// הצגת טקסט מקור
function showSourceText(textData, displayName) {
  let sourceHtml = `<h4>${textData.heRef || displayName}</h4><div class='source-sections'>`;
  console.log(textData);
  
  if (Array.isArray(textData.he)) {
    textData.he.forEach((section, index) => {
      if (typeof section === "string" && section.trim()) {
        const cleanText = stripHtml(section);
        const preview = cleanText.substring(0, 150) + (cleanText.length > 150 ? "..." : "");
        const sectionTitle = `${textData.heRef || displayName} ${index + 1}`;
        
        sourceHtml += `<div class="source-section" onclick="insertSourceText('${textData.heRef || displayName}', ${index}, '${textData.heRef || displayName}', '${sectionTitle}')">
          <strong>${sectionTitle}:</strong> ${preview}
        </div>`;
      }
    });
  } else if (typeof textData.he === "string") {
    const preview = stripHtml(textData.he).substring(0, 200) + "...";
    sourceHtml += `<div class="source-section" onclick="insertSourceText('${textData.heRef || displayName}', 0, '${textData.heRef || displayName}', '${textData.heRef || displayName}')">
      ${preview}
    </div>`;
  } else if (Array.isArray(textData.text)) {
    textData.text.forEach((section, index) => {
      if (typeof section === "string" && section.trim()) {
        const cleanText = stripHtml(section);
        const preview = cleanText.substring(0, 150) + (cleanText.length > 150 ? "..." : "");
        const sectionTitle = `${textData.heRef || displayName} ${index + 1}`;
        
        sourceHtml += `<div class="source-section" onclick="insertSourceText('${textData.heRef || displayName}', ${index}, '${textData.heRef || displayName}', '${sectionTitle}')">
          <strong>${sectionTitle}:</strong> ${preview}
        </div>`;
      }
    });
  } else if (typeof textData.text === "string") {
    const preview = stripHtml(textData.text).substring(0, 200) + "...";
    sourceHtml += `<div class="source-section" onclick="insertSourceText('${textData.heRef || displayName}', 0, '${textData.heRef || displayName}', '${textData.heRef || displayName}')">
      ${preview}
    </div>`;
  }

  sourceHtml += `</div><button onclick="goBackToSearch()" class="back-button">חזור לחיפוש</button>`;
  document.getElementById("searchResults").innerHTML = sourceHtml;
}

// חיפוש מקורות - קבלת טקסט ישירות
async function searchSources() {
  const query = document.getElementById("sourceSearch").value.trim();
  if (!query) {
    document.getElementById("searchResults").innerHTML = "אנא הכנס מונח לחיפוש";
    return;
  }

  document.getElementById("searchResults").innerHTML = "מחפש...";
  
  // קודם נבדוק אם זה ספר עם כתובת ספציפית
  const nameResults = await searchSefaria(query);
  
  // אם יש is_ref ו-heAddressExamples, זה אומר שזו כתובת ספציפית
  if (nameResults && nameResults.is_ref && nameResults.ref) {
    document.getElementById("searchResults").innerHTML = "טוען טקסט...";
    
    const textData = await getSefariaText(nameResults.ref);
    if (textData && (textData.he || textData.text)) {
      showSourceText(textData, nameResults.heRef || nameResults.ref);
      return;
    }
  }
  
  // אם לא מצאנו התאמה מדויקת, ננסה לקבל טקסט ישירות
  const textData = await getSefariaText(query);
  if (textData && (textData.he || textData.text)) {
    showSourceText(textData, query);
    return;
  }
  
  // אם כלום לא עבד, הצג רשימת אפשרויות
  if (nameResults?.completion_objects?.length) {
    const resultsHtml = "<h4>בחר מקור:</h4>" + 
      nameResults.completion_objects
        .filter(obj => obj.type === "ref" && obj.is_primary)
        .slice(0, 8)
        .map(result => {
          const title = result.title || result.key || "ללא שם";
          const key = result.key || result.title || "";
          return `<div class="search-result" onclick="selectSource('${key}', '${title}')">
            ${title} (${result.type || 'ספר'})
          </div>`;
        }).join("");

    document.getElementById("searchResults").innerHTML = resultsHtml;
  } else {
    document.getElementById("searchResults").innerHTML = "לא נמצאו תוצאות";
  }
}

// בחירת מקור והצגת חלקים (משמש רק כשחיפוש השמות מוצג)
async function selectSource(key, displayText) {
  document.getElementById("searchResults").innerHTML = "טוען מקור...";
  
  const textData = await getSefariaText(key);
  if (!textData) {
    document.getElementById("searchResults").innerHTML = "שגיאה בטעינת המקור";
    return;
  }

  showSourceText(textData, displayText);
}
// הוספת טקסט למסמך עם קישור לחיץ
async function insertSourceText(sourceKey, sectionIndex, displayName, sectionTitle) {
  try {
    await Word.run(async (context) => {
      const textData = await getSefariaText(sourceKey);
      if (!textData) throw new Error("לא ניתן לטעון את הטקסט");

      let textToInsert = "";
      if (Array.isArray(textData.he) && textData.he[sectionIndex]) {
        textToInsert = textData.he[sectionIndex];
      } else if (typeof textData.he === "string") {
        textToInsert = textData.he;
      }

      if (!textToInsert?.trim()) throw new Error("לא נמצא טקסט להוספה");

      const cleanText = stripHtml(textToInsert);
      const finalTitle = textData.heRef || sectionTitle || displayName;
      const sourceUrl = `https://www.sefaria.org/${textData.url || sourceKey.replace(/\s+/g, " ")}`;

      const selection = context.document.getSelection();
      
      // הוספת הכותרת והטקסט
      selection.insertText(cleanText, Word.InsertLocation.after);
      
      // הוספת קישור לחיץ
      const range = selection.getRange(Word.RangeLocation.after);
      const hyperlink = range.insertText(finalTitle, Word.InsertLocation.after);
      hyperlink.hyperlink = sourceUrl;
      
      // הוספת שורת רווח אחרי הקישור
      const afterHyperlink = hyperlink.getRange(Word.RangeLocation.after);
      afterHyperlink.insertText("\n", Word.InsertLocation.after);
      
      await context.sync();

      document.getElementById("searchResults").innerHTML = 
        `<div style="color: green;">הטקסט נוסף בהצלחה: ${finalTitle}</div>`;
    });
  } catch (error) {
    console.error("Insert error:", error);
    document.getElementById("searchResults").innerHTML = 
      `<div style="color: red;">שגיאה בהוספת הטקסט: ${error.message}</div>`;
  }
}

// אלטרנטיבה: יצירת היפרלינק עם טקסט מותאם אישית
async function insertSourceTextWithCustomLink(sourceKey, sectionIndex, displayName, sectionTitle, linkText = "מקור בספריא") {
  try {
    await Word.run(async (context) => {
      const textData = await getSefariaText(sourceKey);
      if (!textData) throw new Error("לא ניתן לטעון את הטקסט");

      let textToInsert = "";
      if (Array.isArray(textData.he) && textData.he[sectionIndex]) {
        textToInsert = textData.he[sectionIndex];
      } else if (typeof textData.he === "string") {
        textToInsert = textData.he;
      }

      if (!textToInsert?.trim()) throw new Error("לא נמצא טקסט להוספה");

      const cleanText = stripHtml(textToInsert);
      const finalTitle = textData.heRef || sectionTitle || displayName;
      const sourceUrl = `https://www.sefaria.org/${textData.url || sourceKey.replace(/\s+/g, ".")}`;

      const selection = context.document.getSelection();
      
      // הוספת הכותרת והטקסט
      const titleAndText = `${finalTitle}\n${cleanText}\n`;
      selection.insertText(titleAndText, Word.InsertLocation.after);
      
      // יצירת היפרלינק
      const range = selection.getRange(Word.RangeLocation.after);
      const paragraphRange = range.insertParagraph("", Word.InsertLocation.after);
      const hyperlink = paragraphRange.insertText(linkText, Word.InsertLocation.start);
      
      // הגדרת הקישור
      hyperlink.hyperlink = sourceUrl;
      
      // עיצוב הקישור (אופציונלי)
      hyperlink.font.color = "blue";
      hyperlink.font.underline = Word.UnderlineType.single;
      
      await context.sync();

      document.getElementById("searchResults").innerHTML = 
        `<div style="color: green;">הטקסט נוסף בהצלחה: ${finalTitle}</div>`;
    });
  } catch (error) {
    console.error("Insert error:", error);
    document.getElementById("searchResults").innerHTML = 
      `<div style="color: red;">שגיאה בהוספת הטקסט: ${error.message}</div>`;
  }
}

// חזרה לחיפוש
function goBackToSearch() {
  document.getElementById("sourceSearch").value = "";
  document.getElementById("searchResults").innerHTML = "";
  hideDropdown();
}

// אתחול
Office.onReady(() => {
  // אירועי לחיצה
  document.getElementById("scanText").onclick = scanTextForSources;
  document.getElementById("addSource").onclick = searchSources;

  // אירועי תיבת החיפוש
  const searchInput = document.getElementById("sourceSearch");
  searchInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
      hideDropdown();
      searchSources();
    }
  });
  
  searchInput.addEventListener("focus", (e) => showDropdown(e.target.value));
  searchInput.addEventListener("input", (e) => showDropdown(e.target.value));
  searchInput.addEventListener("blur", (e) => {
    // מניעת הסתרה מיידית כדי לאפשר לחיצה על אפשרויות
    setTimeout(() => hideDropdown(), 300);
  });

  // הסתרת תפריט בלחיצה מחוץ לאזור
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".search-container") && !e.target.closest("#dropdownOptions")) {
      hideDropdown();
    }
  });
});