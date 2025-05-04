Office.onReady(async (info) => {
  if (info.host !== Office.HostType.PowerPoint) return;
  const runButton = document.getElementById("runButton");
  if (!runButton) {
    console.warn("Run button not found.");
    return;
  }
  runButton.onclick = async () => {
    console.log("Starting slide processing...");
    let metadataComplete = false;
    let pptxComplete = false;
    let userInstructionComplete = false;
    let currentSlideIndex = null;
    let totalSlides = 0; 
    const instruction = document.getElementById("instructionInput").value;
    await Promise.allSettled([
      extractAllSlideShapes().then(() => (metadataComplete = true)).catch(() => {}),
      sendPptxAsBase64ToBackend().then(() => (pptxComplete = true)).catch(() => {}),
      getCurrentSlideIndex().then(({ currentSlideIndex: index, totalSlides: total }) => {
        currentSlideIndex = index;
        totalSlides = total;
        userInstructionComplete = true;
      }).catch(() => {})
    ]);

    if (metadataComplete && pptxComplete) {
      userInstructionComplete = true;
      try {
        const payload = {
          instruction: instruction,
          slide_index: currentSlideIndex,
          total_slides: totalSlides
        };

        const response = await sendInstructionToBackend(payload);
        
        if (response.status === 200) {
          const responseData = await response.json();
          console.log("Instruction sent with slide index:", currentSlideIndex);

          const codeToExecute = responseData.generated_code || responseData.code;
          if ((responseData.status === "success" || responseData.status === "partial_success") && codeToExecute) {
            console.log(`Received code from backend (Type: ${typeof codeToExecute}). Executing...`);
            await executeGeneratedOfficeJsCode(codeToExecute);
          } else {
            console.log(`No executable code returned from backend or status is not success/partial_success. Message: ${responseData.message || '(No message)'}`);
            console.warn("No code returned or backend processing failed:", responseData);
          }


        } else {
          console.error("Failed to send instruction to backend.");
        }
      } catch (err) {
        console.error("Error sending instruction to backend:", err);
      }
    } else {
      console.warn("Not all parallel tasks completed successfully.");
    }
    if (!metadataComplete) {
      try {
        console.log("Retrying slide metadata extraction...");
        await extractAllSlideShapes();
        metadataComplete = true;
      } catch (err) {
        console.error("Slide metadata extraction failed on retry:", err);
      }
    }
    if (!pptxComplete) {
      try {
        console.log("Retrying PPTX upload...");
        await sendPptxAsBase64ToBackend();
        pptxComplete = true;
      } catch (err) {
        console.error("PPTX upload failed on retry:", err);
      }
    }
    if (metadataComplete && pptxComplete && userInstructionComplete) {
      console.log("All data successfully sent to backend.");
    } else {
      console.warn("Some tasks did not complete successfully.");
    }
  };
});

async function getCurrentSlideIndex() {
  try {
      return await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items/id"); 

          // Get the currently selected slide
          const selectedSlides = context.presentation.getSelectedSlides();
          selectedSlides.load("items/id");
          await context.sync();

          if (!selectedSlides?.items?.length) {
              console.warn("[getCurrentSlideIndex] No slide selected.");
              return { currentSlideIndex: null, totalSlides: slides.items.length };
          }
          const activeId = selectedSlides.items[0].id;
          console.log(`[getCurrentSlideIndex] Active slide ID: ${activeId}`);
          let currentSlideIndex = null;
          for (let i = 0; i < slides.items.length; i++) {
              if (slides.items[i].id === activeId) {
                  console.log(`[getCurrentSlideIndex] Found active slide at index ${i}`);
                  currentSlideIndex = i;
                  break;
              }
          }
          if (currentSlideIndex === null) {
              console.warn("[getCurrentSlideIndex] Active slide ID not found in slides collection.");
          }
          console.log(`[getCurrentSlideIndex] Total number of slides: ${slides.items.length}`);
          return { currentSlideIndex, totalSlides: slides.items.length };
      });
  } catch (error) {
      console.error("[getCurrentSlideIndex] Error:", error);
      return { currentSlideIndex: null, totalSlides: 0 };
  }
}

async function sendInstructionToBackend(payload) { 
  const response = await fetch('http://localhost:8000/process_instruction', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      instruction: payload.instruction, 
      slide_index: payload.slide_index, 
      total_slides: payload.total_slides 
    }),
  });
  return response;
}

async function executeGeneratedOfficeJsCode(codeInput) {
  let codeToExecuteMap = {};
  let overallSuccess = true;

  if (!codeInput) {
      console.log("No code received from backend.");
      return false;
  } else if (typeof codeInput === 'string') {
      console.log("Received single code snippet. Executing assuming target slide 0 or iterative code.");
      codeToExecuteMap[0] = codeInput;
  } else if (typeof codeInput === 'object' && Object.keys(codeInput).length > 0) {
    console.log("Received code snippets by slide index.");  
    codeToExecuteMap = codeInput;
  }
  const slideIndices = Object.keys(codeToExecuteMap).map(Number).sort((a, b) => a - b);
  console.log(`Executing code for slide indices: ${slideIndices.join(', ')}...`);

  for (const slideIndex of slideIndices) {
      let codeStr = codeToExecuteMap[slideIndex];

      if (!codeStr || typeof codeStr !== 'string' || codeStr.trim().length < 5) {
        console.warn(`Warning: Code snippet for slide index ${slideIndex} is empty or too short. Skipping.`);
        overallSuccess = false; 
        continue;
    }
      console.log(`--- Executing code for Slide Index ${slideIndex} ---`);
      codeStr = codeStr.trim();
      if (codeStr.startsWith("```")) {
          codeStr = codeStr.replace(/^```[a-z]*\s*/i, '').replace(/\s*```$/, '').trim();
      } else if (codeStr.startsWith("`")) {
          codeStr = codeStr.replace(/^`\s*/, '').replace(/\s*`$/, '').trim();
      }

      if (!codeStr) {
        console.warn(`--- Code cleaning resulted in empty string for slide index ${slideIndex}. Skipping. ---`);
        overallSuccess = false;
        continue;
    }
      console.log(`[Cleaned Code for Slide ${slideIndex}]:\n${codeStr}`);

      try {
          // Create and execute the dynamic function 
          const dynamicFunc = new Function(`
              return (async () => {
                try {
                  await PowerPoint.run(async (context) => {
                    console.log('Executing generated code inside PowerPoint.run for slide index ${slideIndex}...');
                    ${codeStr}
                    console.log('Code block finished for slide index ${slideIndex}.');
                  });
                  console.log('PowerPoint.run completed successfully for slide index ${slideIndex}.');
                  return { success: true }; // Indicate success for this slide
                } catch(runError) {
                  console.error('Error during PowerPoint.run execution for slide index ${slideIndex}:', runError);
                  let errorMsg = runError.message || 'Unknown execution error';
                  // Include debug info if available (useful for Office.js errors)
                  if (runError.debugInfo) { errorMsg += ' | Debug Info: ' + JSON.stringify(runError.debugInfo); }
                  if (runError.stack) { console.error("Execution Stack Trace:", runError.stack); }
                  return { success: false, error: { message: errorMsg } }; // Indicate failure
                }
              })();
          `);
          const executionResult = await dynamicFunc();
          if (executionResult.success) {
              console.log(`--- Code executed successfully for Slide Index ${slideIndex} ---`);
          } else {
            console.log(`--- Code execution FAILED for Slide Index ${slideIndex}: ${executionResult.error?.message || 'Unknown'} ---`);
              overallSuccess = false; 
          }
      } catch (evalError) {
          console.log(`--- Code evaluation/setup error for Slide Index ${slideIndex}: ${evalError.message} ---`);
          overallSuccess = false;
      }
  } 
  console.log(`Finished executing all provided code snippets. Overall success: ${overallSuccess}`);
  return overallSuccess;
}

function logDebug(message) {
  const debugLog = document.getElementById("debugLog");
  debugLog.innerText += message + "\n";
}