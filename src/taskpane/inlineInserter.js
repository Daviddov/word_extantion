// inlineInserter.js - הכנסת ציטוטים inline (בתוך הטקסט) - מתוקן עם מיקום מדויק

/**
 * הסרת תגי HTML
 */
function stripHtmlTags(text) {
  return text ? text.replace(/<[^>]*>/g, '').trim() : '';
}

/**
 * יצירת תוכן ציטוט inline מההתאמות
 */
function createInlineCitationContent(matches) {
  return matches.map(match =>
    match.verseDispHeb
  ).join(' | ');
}

/**
 * מציאת מיקום מדויק לפי אינדקס התו - עם הוספת רווח אם צריך
 */
async function findPositionByCharIndex(context, charIndex, addSpaceBefore = false) {
  try {
    const body = context.document.body;
    const range = body.getRange();

    context.load(range, 'text');
    await context.sync();

    const fullText = range.text;

    // וידוא שהאינדקס תקין
    if (charIndex < 0 || charIndex > fullText.length) {
      console.warn(`אינדקס לא תקין: ${charIndex}, אורך הטקסט: ${fullText.length}`);
      return null;
    }

    // יצירת range באמצעות moveStart ו-moveEnd
    const targetRange = body.getRange('Start');

    // הזזת המיקום למקום הרצוי
    if (charIndex > 0) {
      targetRange.moveStart('Character', charIndex);
    }

    // אם צריך להוסיף רווח לפני הציטוט
    if (addSpaceBefore && charIndex > 0) {
      // בדיקה אם יש רווח לפני המיקום
      const prevChar = fullText.charAt(charIndex - 1);
      if (prevChar && prevChar !== ' ' && prevChar !== '\n' && prevChar !== '\t') {
        // הוספת רווח לפני הציטוט
        targetRange.insertText(' ', 'Before');
        await context.sync();
      }
    }

    return targetRange;

  } catch (error) {
    console.error(`שגיאה במציאת מיקום ${charIndex}:`, error);
    return null;
  }
}

/**
 * מציאת מיקום מדויק על ידי חיפוש חלק מסוים של הטקסט
 */
async function findPositionByTextSearch(context, citation) {
  try {
    const body = context.document.body;
    const range = body.getRange();

    context.load(range, 'text');
    await context.sync();

    const fullText = range.text;

    // חילוץ החלק הספציפי של הטקסט שצריך לחפש
    // משתמש ב-startPos ו-endPos כדי לקבל את הטקסט המדויק
    const searchText = fullText.substring(citation.startPos, citation.endPos);

    if (!searchText.trim()) {
      return null;
    }

    // חיפוש הטקסט הספציפי הזה
    const searchResults = body.search(searchText, {
      matchCase: false,
      matchWholeWord: false
    });

    context.load(searchResults, 'items');
    await context.sync();

    if (searchResults.items.length > 0) {
      // לקיחת הסוף של הטקסט שנמצא
      const foundRange = searchResults.items[0];
      return foundRange.getRange('End');
    }

    return null;

  } catch (error) {
    console.error('שגיאה בחיפוש טקסט:', error);
    return null;
  }
}

/**
 * הוספת ציטוט inline בודד - במיקום המדויק
 */
async function insertSingleInlineCitation(context, citation, citationNumber) {
  try {
    console.log(`מעבד ציטוט inline ${citationNumber}: startPos=${citation.startPos}, endPos=${citation.endPos}`);

    // ניסיון ראשון: מציאת מיקום על ידי חיפוש טקסט
    let targetPosition = await findPositionByTextSearch(context, citation);

    // ניסיון שני: מציאת מיקום לפי אינדקס תו
    if (!targetPosition) {
      console.log(`מנסה מיקום לפי אינדקס תו: ${citation.endPos}`);
      targetPosition = await findPositionByCharIndex(context, citation.endPos, false);
    }

    if (!targetPosition) {
      console.warn(`לא ניתן למצוא מיקום לציטוט inline במיקום ${citation.endPos}`);
      return false;
    }

    // יצירת תוכן הציטוט
    const citationContent = createInlineCitationContent(citation.matches);

    // בדיקה אם צריך רווח לפני הציטוט
    const body = context.document.body;
    const range = body.getRange();
    context.load(range, 'text');
    await context.sync();

    const fullText = range.text;
    const nextChar = citation.endPos < fullText.length ? fullText.charAt(citation.endPos) : '';

    // אם התו הבא לא רווח, נוסיף רווח לפני הציטוט
    if (nextChar && nextChar !== ' ' && nextChar !== '\n' && nextChar !== '\t') {
      targetPosition.insertText(' ', 'Before');
      await context.sync();
    }

    // הוספת הציטוט במיקום המדויק
    const textToInsert = `(${citationContent})`;
    targetPosition.insertText(textToInsert, 'After');

    await context.sync();
    console.log(`נוסף ציטוט inline ${citationNumber} במיקום ${citation.endPos}`);
    return true;

  } catch (error) {
    console.error(`שגיאה בהוספת ציטוט inline ${citationNumber}:`, error);
    return false;
  }
}

/**
 * הכנת רשימת ציטוטים תקינים
 */
function prepareValidInlineCitations(citations, minScore) {
  return citations
    .filter(citation => citation.matches && citation.matches.length > 0)
    .map(citation => ({
      startPos: citation.originalCitation?.startIChar || citation.startIChar,
      endPos: citation.originalCitation?.endIChar || citation.endIChar,
      originalText: stripHtmlTags(citation.text),
      matches: citation.matches.filter(match =>
        match.verseDispHeb && match.score >= minScore
      )
    }))
    .filter(citation => citation.matches.length > 0);
}

/**
 * הוספת כל הציטוטים inline למסמך
 */
async function insertInlineCitationsToDocument(citations, context, minScore = 22) {
  return await Word.run(async (context) => {
    console.log('מתחיל עיבוד ציטוטים inline:', citations);

    // הכנת רשימת ציטוטים
    const validCitations = prepareValidInlineCitations(citations, minScore);

    if (validCitations.length === 0) {
      throw new Error('לא נמצאו ציטוטים תקינים להוספה inline');
    }

    console.log('ציטוטים תקינים inline:', validCitations);

    // עיבוד מהסוף להתחלה כדי לשמור על המיקומים
    const sortedCitations = validCitations.sort((a, b) => b.endPos - a.endPos);

    let successCount = 0;

    // עיבוד ציטוט אחד בכל פעם כדי למנוע התנגשויות
    for (let i = 0; i < sortedCitations.length; i++) {
      const citation = sortedCitations[i];
      const citationNumber = sortedCitations.length - i; // מספור הפוך כדי לשמור על הסדר הנכון

      console.log(`מעבד ציטוט inline ${citationNumber}: ${citation.originalText}`);

      const success = await insertSingleInlineCitation(context, citation, citationNumber);
      if (success) {
        successCount++;
        // המתנה קצרה בין הוספות כדי למנוע בעיות
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    return successCount;
  });
}

/**
 * פונקציה ראשית להפעלה - inline citations
 */
async function processCitationsWithInline(apiResponse, minScore = 22) {
  try {
    console.log('מתחיל עיבוד ציטוטים inline...');
    const addedCount = await insertInlineCitationsToDocument(apiResponse, null, minScore);
    console.log(`הושלם! נוספו ${addedCount} ציטוטים inline`);
    return addedCount;
  } catch (error) {
    console.error('שגיאה בעיבוד ציטוטים inline:', error);
    throw error;
  }
}

/**
 * הוספת ציטוט inline ידני
 */
async function insertManualInlineCitation(searchText, citationText) {
  return await Word.run(async (context) => {
    try {
      const body = context.document.body;
      const searchResults = body.search(searchText, {
        matchCase: false,
        matchWholeWord: false
      });
      
      context.load(searchResults, 'items');
      await context.sync();
      
      if (searchResults.items.length === 0) {
        throw new Error('לא נמצא הטקסט במסמך');
      }
      
      const targetRange = searchResults.items[0];
      targetRange.insertText(` (${citationText})`, 'After');
      
      await context.sync();
      return true;
      
    } catch (error) {
      console.error('שגיאה בהוספת ציטוט ידני:', error);
      throw error;
    }
  });
}

// Export functions to global scope
if (typeof window !== 'undefined') {
  window.processCitationsWithInline = processCitationsWithInline;
  window.insertInlineCitationsToDocument = insertInlineCitationsToDocument;
  window.insertManualInlineCitation = insertManualInlineCitation;
  window.createInlineCitationContent = createInlineCitationContent;
}