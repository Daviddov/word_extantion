Office.onReady(() => {
  document.getElementById("scanText").onclick = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        const selectedText = selection.text;
        
        if (!selectedText || selectedText.trim() === "") {
          console.log("No text selected");
          return;
        }

        const body = {
          text: { body: selectedText, title: "" },
          lang: "he"
        };

        const response = await fetch("https://www.sefaria.org/api/find-refs/", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(body)
        });

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        console.log(result);

        // Check both possible locations for results
        const refs = result?.body?.results || result?.title?.results || [];

        if (refs.length === 0) {
          console.log("No references found");
          document.getElementById("output").innerHTML = "לא נמצאו מקורות בטקסט הנבחר";
          return;
        }

        // Collect all references for display
        const foundReferences = [];

        // Process each reference and convert to hyperlinks
        for (const ref of refs) {
          try {
            const refKey = ref.refs[0];
            const refData = result.body.refData[refKey] || result.title.refData[refKey];
            
            if (!refData) {
              console.log("No reference data found for:", refKey);
              continue;
            }

            const refText = refData.heRef || refData.en || refKey;
            const url = "https://www.sefaria.org/" + (refData.url || refKey.replace(/\s+/g, "."));

            // Get the portion of selected text that contains the reference
            const startChar = Math.max(0, ref.startChar);
            const endChar = Math.min(selectedText.length, ref.endChar);
            const referenceText = selectedText.substring(startChar, endChar);

            console.log("Converting to hyperlink:", referenceText, "->", url);

            try {
              // Method 1: Try to find and replace the specific text with a hyperlink
              const searchResults = context.document.body.search(referenceText, {
                matchCase: false,
                matchWholeWord: false
              });
              
              searchResults.load("items");
              await context.sync();

              if (searchResults.items.length > 0) {
                // Convert the first match to a hyperlink
                const range = searchResults.items[0];
                range.hyperlink = url;
                range.font.color = "#0066cc";
                range.font.underline = true;
                foundReferences.push({
                  text: referenceText,
                  source: refText,
                  url: url,
                  method: "hyperlink"
                });
              } else {
                console.log("Could not find text to convert:", referenceText);
              }
              
            } catch (searchError) {
              console.log("Search method failed, trying alternative approach");
              
              // Method 2: If search doesn't work, try to use the selection itself
              try {
                if (ref.startChar === 0 && ref.endChar >= selectedText.length - 5) {
                  // If the reference covers most/all of the selection
                  selection.hyperlink = url;
                  selection.font.color = "#0066cc";
                  selection.font.underline = true;
                  foundReferences.push({
                    text: referenceText,
                    source: refText,
                    url: url,
                    method: "selection_hyperlink"
                  });
                }
              } catch (selectionError) {
                console.log("Selection hyperlink failed too");
                
                // Method 3: Insert clickable text
                try {
                  const hyperlinkText = `${referenceText}`;
                  const insertedRange = selection.insertText(hyperlinkText, Word.InsertLocation.replace);
                  insertedRange.hyperlink = url;
                  insertedRange.font.color = "#0066cc";
                  insertedRange.font.underline = true;
                  
                  foundReferences.push({
                    text: referenceText,
                    source: refText,
                    url: url,
                    method: "inserted_hyperlink"
                  });
                } catch (insertError) {
                  console.log("All hyperlink methods failed for:", referenceText);
                  foundReferences.push({
                    text: referenceText,
                    source: refText,
                    url: url,
                    method: "failed"
                  });
                }
              }
            }

            await context.sync();

          } catch (refError) {
            console.error("Error processing reference:", ref, refError);
          }
        }

        // Display results
        if (foundReferences.length > 0) {
          const successfulLinks = foundReferences.filter(ref => ref.method !== "failed").length;
          const outputDiv = document.getElementById("output");
          
          if (successfulLinks > 0) {
            outputDiv.innerHTML = `הומרו ${successfulLinks} מקורות לקישורים`;
            outputDiv.style.color = "green";
          } else {
            let referencesHtml = "<h3>מקורות שנמצאו (לא ניתן היה להמיר לקישורים):</h3><ul>";
            foundReferences.forEach(ref => {
              referencesHtml += `<li>"${ref.text}" - <a href="${ref.url}" target="_blank">${ref.source}</a></li>`;
            });
            referencesHtml += "</ul>";
            outputDiv.innerHTML = referencesHtml;
            outputDiv.style.color = "orange";
          }
        } else {
          document.getElementById("output").innerHTML = "לא נמצאו מקורות";
          document.getElementById("output").style.color = "red";
        }

        await context.sync();
        console.log("Comments added successfully");

      });
    } catch (error) {
      console.error("Error:", error);
      // Optionally show error to user
      document.getElementById("output").innerHTML = `שגיאה: ${error.message}`;
    }
  };
});