// commands.js - פונקציות פקודה עבור התוסף

Office.onReady(() => {
    console.log('Commands loaded');
});

// פונקציה לבדיקה מהירה של הטקסט הנבחר
async function quickCheck(event) {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            context.load(selection, 'text');
            await context.sync();
            
            if (!selection.text) {
                showNotification('אנא בחר טקסט לבדיקה');
                event.completed();
                return;
            }
            
            // כאן תוכל להוסיף קריאה מהירה לאתר Dicta
            showNotification('בודק מקורות עבור: ' + selection.text.substring(0, 50) + '...');
            
            // הדמיית בדיקה
            setTimeout(() => {
                showNotification('נמצאו 2 מקורות אפשריים');
            }, 2000);
        });
    } catch (error) {
        console.error('Error in quick check:', error);
        showNotification('שגיאה בבדיקה המהירה');
    }
    
    event.completed();
}

// פונקציה להכנסת ציטוט מהיר
async function insertQuickCitation(event) {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            
            // הוספת ציטוט בסוף הטקסט הנבחר
            selection.insertText(' [מקור: לבדיקה]', Word.InsertLocation.after);
            
            await context.sync();
            showNotification('סימון לבדיקה נוסף');
        });
    } catch (error) {
        console.error('Error inserting citation:', error);
        showNotification('שגיאה בהוספת הסימון');
    }
    
    event.completed();
}

// פונקציה להצגת התראות
function showNotification(message) {
    Office.ribbon.requestUpdate({
        tabs: [{
            id: "TabHome",
            groups: [{
                id: "CommandsGroup",
                controls: [{
                    id: "TaskpaneButton",
                    enabled: true
                }]
            }]
        }]
    });
    
    // אם יש תמיכה בהתראות מותאמות אישית
    if (Office.context.requirements.isSetSupported('CustomFunctions', '1.7')) {
        // הצג הודעה מותאמת אישית
        console.log('Notification:', message);
    }
}

// רישום הפונקציות
if (typeof Office !== 'undefined') {
    Office.actions.associate('quickCheck', quickCheck);
    Office.actions.associate('insertQuickCitation', insertQuickCitation);
}