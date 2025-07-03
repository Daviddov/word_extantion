// documentInserter.js - הכנסת ציטוטים עם OOXML footnotes (מתוקן)

/**
 * הסרת תגי HTML
 */
function stripHtmlTags(text) {
  return text ? text.replace(/<[^>]*>/g, '').trim() : '';
}

/**
 * יצירת OOXML עבור footnote reference בלבד (ללא החלפת טקסט)
 */
function createFootnoteReferenceOOXML(footnoteId, footnoteText) {
  const escapedText = footnoteText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:name="/_rels/.rels" pkg:padding="512">
      <pkg:xmlData>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="word/document.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
        </Relationships>
      </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:name="/word/_rels/document.xml.rels" pkg:padding="256">
      <pkg:xmlData>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
          <Relationship Id="rId2" Target="footnotes.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"/>
        </Relationships>
      </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" pkg:name="/word/document.xml">
      <pkg:xmlData>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:rPr>
                  <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteReference w:id="${footnoteId}"/>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
      </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" pkg:name="/word/footnotes.xml">
      <pkg:xmlData>
        <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:footnote w:id="-1" w:type="separator">
            <w:p>
              <w:r>
                <w:separator/>
              </w:r>
            </w:p>
          </w:footnote>
          <w:footnote w:id="0" w:type="continuationSeparator">
            <w:p>
              <w:r>
                <w:continuationSeparator/>
              </w:r>
            </w:p>
          </w:footnote>
          <w:footnote w:id="${footnoteId}">
            <w:p>
              <w:pPr>
                <w:pStyle w:val="FootnoteText"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rStyle w:val="FootnoteReference"/>
                </w:rPr>
                <w:footnoteRef/>
              </w:r>
              <w:r>
                <w:t xml:space="preserve">${escapedText}</w:t>
              </w:r>
            </w:p>
          </w:footnote>
        </w:footnotes>
      </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" pkg:name="/word/styles.xml">
      <pkg:xmlData>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:default="1" w:styleId="Normal" w:type="paragraph">
            <w:name w:val="Normal"/>
            <w:qFormat/>
          </w:style>
          <w:style w:default="1" w:styleId="DefaultParagraphFont" w:type="character">
            <w:name w:val="Default Paragraph Font"/>
            <w:uiPriority w:val="1"/>
            <w:semiHidden/>
            <w:unhideWhenUsed/>
          </w:style>
          <w:style w:styleId="FootnoteText" w:type="paragraph">
            <w:name w:val="footnote text"/>
            <w:basedOn w:val="Normal"/>
            <w:link w:val="FootnoteTextChar"/>
            <w:uiPriority w:val="99"/>
            <w:semiHidden/>
            <w:unhideWhenUsed/>
            <w:rPr>
              <w:sz w:val="20"/>
              <w:szCs w:val="20"/>
            </w:rPr>
          </w:style>
          <w:style w:customStyle="1" w:styleId="FootnoteTextChar" w:type="character">
            <w:name w:val="Footnote Text Char"/>
            <w:basedOn w:val="DefaultParagraphFont"/>
            <w:link w:val="FootnoteText"/>
            <w:uiPriority w:val="99"/>
            <w:semiHidden/>
            <w:rPr>
              <w:sz w:val="20"/>
              <w:szCs w:val="20"/>
            </w:rPr>
          </w:style>
          <w:style w:styleId="FootnoteReference" w:type="character">
            <w:name w:val="footnote reference"/>
            <w:basedOn w:val="DefaultParagraphFont"/>
            <w:uiPriority w:val="99"/>
            <w:semiHidden/>
            <w:unhideWhenUsed/>
            <w:rPr>
              <w:vertAlign w:val="superscript"/>
            </w:rPr>
          </w:style>
        </w:styles>
      </pkg:xmlData>
    </pkg:part>
  </pkg:package>`;
}

/**
 * יצירת תוכן footnote מההתאמות
 */
function createFootnoteContent(matches) {
  return matches.map(match =>
    `${stripHtmlTags(match.matchedText)} (${match.verseDispHeb})`
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
    // משתמש ב-startIChar ו-endIChar כדי לקבל את הטקסט המדויק
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
 * הוספת ציטוט בודד באמצעות OOXML - במיקום המדויק
 */
async function insertSingleCitation(context, citation, footnoteNumber) {
  try {
    console.log(`מעבד ציטוט ${footnoteNumber}: startPos=${citation.startPos}, endPos=${citation.endPos}`);

    // ניסיון ראשון: מציאת מיקום על ידי חיפוש טקסט
    let targetPosition = await findPositionByTextSearch(context, citation);

    // ניסיון שני: מציאת מיקום לפי אינדקס תו
    if (!targetPosition) {
      console.log(`מנסה מיקום לפי אינדקס תו: ${citation.endPos}`);
      targetPosition = await findPositionByCharIndex(context, citation.endPos, false);
    }

    if (!targetPosition) {
      console.warn(`לא ניתן למצוא מיקום לציטוט במיקום ${citation.endPos}`);
      return false;
    }

    // יצירת תוכן footnote
    const footnoteContent = createFootnoteContent(citation.matches);

    // יצירת OOXML רק להפניה
    const ooxmlContent = createFootnoteReferenceOOXML(footnoteNumber, footnoteContent);

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

    // הוספת ההפניה במיקום המדויק
    targetPosition.insertOoxml(ooxmlContent, 'After');

    await context.sync();
    console.log(`נוסף ציטוט ${footnoteNumber} במיקום ${citation.endPos}`);
    return true;

  } catch (error) {
    console.error(`שגיאה בהוספת ציטוט ${footnoteNumber}:`, error);
    return false;
  }
}

/**
 * הכנת רשימת ציטוטים תקינים
 */
function prepareValidCitations(citations, minScore) {
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
 * הפונקציה הראשית - הוספת כל הציטוטים
 */
async function insertCitationsToDocument(citations, context, minScore = 22) {
  return await Word.run(async (context) => {
    console.log('מתחיל עיבוד ציטוטים:', citations);

    // הכנת רשימת ציטוטים
    const validCitations = prepareValidCitations(citations, minScore);

    if (validCitations.length === 0) {
      throw new Error('לא נמצאו ציטוטים תקינים');
    }

    console.log('ציטוטים תקינים:', validCitations);

    // עיבוד מהסוף להתחלה כדי לשמור על המיקומים
    const sortedCitations = validCitations.sort((a, b) => b.endPos - a.endPos);

    let successCount = 0;

    // עיבוד ציטוט אחד בכל פעם כדי למנוע התנגשויות ID
    for (let i = 0; i < sortedCitations.length; i++) {
      const citation = sortedCitations[i];
      const footnoteNumber = sortedCitations.length - i; // מספור הפוך כדי לשמור על הסדר הנכון

      console.log(`מעבד ציטוט ${footnoteNumber}: ${citation.originalText}`);

      const success = await insertSingleCitation(context, citation, footnoteNumber);
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
 * פונקציה ראשית להפעלה
 */
async function processCitationsWithFootnotes(apiResponse, minScore = 22) {
  try {
    console.log('מתחיל עיבוד ציטוטים עם OOXML...');
    const addedCount = await insertCitationsToDocument(apiResponse, null, minScore);
    console.log(`הושלם! נוספו ${addedCount} footnotes באמצעות OOXML`);
    return addedCount;
  } catch (error) {
    console.error('שגיאה בעיבוד ציטוטים:', error);
    throw error;
  }
}

// Export
if (typeof window !== 'undefined') {
  window.processCitationsWithFootnotes = processCitationsWithFootnotes;
  window.insertCitationsToDocument = insertCitationsToDocument;
  window.createFootnoteReferenceOOXML = createFootnoteReferenceOOXML;
}