// Office.onReady(async (info) => {
//   if (info.host !== Office.HostType.PowerPoint) return;

//   const runButton = document.getElementById("runButton");
//   if (!runButton) {
//     console.warn("Run button not found.");
//     return;
//   }

//   runButton.onclick = async () => {
//     console.log("Starting slide processing...");

//     let metadataComplete = false;
//     let pptxComplete = false;
//     let userInstructionComplete = false;
//     let currentSlideIndex = null;

//     const instruction = document.getElementById("instructionInput").value;

//     await Promise.allSettled([
//       extractAllSlideShapes().then(() => (metadataComplete = true)).catch(() => {}),
//       sendPptxAsBase64ToBackend().then(() => (pptxComplete = true)).catch(() => {}),
//       getCurrentSlideIndex().then(index => {
//         currentSlideIndex = index;
//         slideInfoComplete = true;
//       }).catch(() => {})
//     ]);

//     if (metadataComplete && pptxComplete) {
//       userInstructionComplete = true;
//       try {
//         const response = await sendInstructionToBackend(instruction, currentSlideIndex);
//         if (response.status === 200) {
//           console.log("Instruction sent with slide index:", currentSlideIndex);
//         } else {
//           console.error("Failed to send instruction to backend.");
//         }
//       } catch (err) {
//         console.error("Error sending instruction to backend:", err);
//       }
//     } else {
//       console.warn("Not all parallel tasks completed successfully.");
//     }

//     // Retry logic if needed
//     if (!metadataComplete) {
//       try {
//         console.log("Retrying slide metadata extraction...");
//         await extractAllSlideShapes();
//         metadataComplete = true;
//       } catch (err) {
//         console.error("Slide metadata extraction failed on retry:", err);
//       }
//     }

//     if (!pptxComplete) {
//       try {
//         console.log("Retrying PPTX upload...");
//         await sendPptxAsBase64ToBackend();
//         pptxComplete = true;
//       } catch (err) {
//         console.error("PPTX upload failed on retry:", err);
//       }
//     }

//     if (metadataComplete && pptxComplete && userInstructionComplete) {
//       console.log("All data successfully sent to backend.");
//     } else {
//       console.warn("Some tasks did not complete successfully.");
//     }
//   };
// });

// async function getCurrentSlideIndex() {
//   return new Promise((resolve, reject) => {
//     Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         const slideIndex = result.value.slides[0].index; 
//         // let slideIndex = result.value.slides[0].index;
//         // slideIndex = slideIndex - 1;
//         console.log("Current slide index:", slideIndex);
//         resolve(slideIndex);
//       } else {
//         console.error("Failed to get current slide index:", result.error.message);
//         reject(result.error.message);
//       }
//     });
//   });
// }

// async function sendInstructionToBackend(instruction, slideIndex) {
//   const response = await fetch('http://localhost:8000/process_instruction', {
//     method: 'POST',
//     headers: {
//       'Content-Type': 'application/json',
//     },
//     body: JSON.stringify({
//       instruction: instruction,
//       slide_index: slideIndex
//     }),
//   });
//   return response;
// }

// function logToDebugLog(message) {
//   const logBox = document.getElementById("debugLog");
//   if (logBox) {
//     const timestamp = new Date().toLocaleTimeString();
//     logBox.textContent += `[${timestamp}] ${message}\n`;
//     logBox.scrollTop = logBox.scrollHeight;
//   } else {
//     console.error("Debug log element (#debugLog) not found in HTML. Message:", message);
//   }
// }

// function logDebug(message) {
//   const debugLog = document.getElementById("debugLog");
//   debugLog.innerText += message + "\n";
// }


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

    const instruction = document.getElementById("instructionInput").value;

    await Promise.allSettled([
      extractAllSlideShapes().then(() => (metadataComplete = true)).catch(() => {}),
      sendPptxAsBase64ToBackend().then(() => (pptxComplete = true)).catch(() => {}),
      getCurrentSlideIndex().then(index => {
        currentSlideIndex = index;
        slideInfoComplete = true;
      }).catch(() => {})
    ]);

    if (metadataComplete && pptxComplete) {
      userInstructionComplete = true;
      try {

        const response = await sendInstructionToBackend(instruction, currentSlideIndex);
        if (response.status === 200) {
          const responseData = await response.json();
          console.log("Instruction sent with slide index:", currentSlideIndex);

          if (responseData.status === "success" && responseData.code) {
            await executeGeneratedOfficeJsCode(responseData.code);
          } else {
            console.warn("No code returned from backend or code generation failed:", responseData.message);
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

    // Retry logic if needed
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

// Office.onReady((info) => {
//   if (info.host === Office.HostType.PowerPoint) {
//     document.getElementById("runButton").onclick = process;
//   }
// });
// async function process() {
//   await PowerPoint.run(async (context) => {
//     const slide = context.presentation.slides.getItemAt(0);
//     const shapes = slide.shapes;
//     shapes.load("items/id, items/left, items/top, items/width, items/height");
//     await context.sync();

//     const shape463 = shapes.items.find(s => s.id === "463");
//     if (!shape463) {
//       console.error(`Critical: Shape with ID '463' not found. Skipping related operations.`);
//     }
//     const shape468 = shapes.items.find(s => s.id === "468");
//     if (!shape468) {
//       console.error(`Critical: Shape with ID '468' not found. Skipping related operations.`);
//     }
//     const shape479 = shapes.items.find(s => s.id === "479");
//     if (!shape479) {
//       console.error(`Critical: Shape with ID '479' not found. Skipping related operations.`);
//     }
    
//     if (shape463) {
//       shape463.width = 250;
//     }

//     if (shape468) {
//       shape468.textFrame.horizontalAlignment = "Center";
//     }

//     if (shape479) {
//       shape479.left = 50;
//       shape479.top = 100;
//     }

//     await context.sync();
//   });
// }






async function getCurrentSlideIndex() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const slideIndex = result.value.slides[0].index; 
        // let slideIndex = result.value.slides[0].index;
        // slideIndex = slideIndex - 1;
        console.log("Current slide index:", slideIndex);
        resolve(slideIndex);
      } else {
        console.error("Failed to get current slide index:", result.error.message);
        reject(result.error.message);
      }
    });
  });
}

async function sendInstructionToBackend(instruction, slideIndex) {
  const response = await fetch('http://localhost:8000/process_instruction', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      instruction: instruction,
      slide_index: slideIndex
    }),
  });
  return response;
}


async function executeGeneratedOfficeJsCode(codeStr) {
  if (!codeStr || codeStr.trim().length < 5) {
    console.warn("Warning: Generated code is empty or too short.");
    return;
  }

  let rawCode = codeStr.trim();
  if (rawCode.startsWith("```")) {
    rawCode = rawCode.replace(/^```[a-z]*\s*/i, '').replace(/\s*```$/, '').trim();
  } else if (rawCode.startsWith("`")) {
    rawCode = rawCode.replace(/^`\s*/, '').replace(/\s*`$/, '').trim();
  }

  console.log("Executing generated Office.js code:\n" + rawCode);

  try {
    const dynamicFunc = new Function(`
      return (async () => {
        try {
          await PowerPoint.run(async (context) => {
            ${rawCode}
          });
          return { success: true };
        } catch(runError) {
          console.error("Execution Error:", runError.message);
          return { success: false, error: runError };
        }
      })();
    `);

    const result = await dynamicFunc();

    if (result.success) {
      console.log("Code executed successfully.");
    } else {
      console.error("Execution failed:", result.error?.message || "Unknown error");
    }

  } catch (e) {
    console.error("Code evaluation failed:", e.message);
  }
}

function logDebug(message) {
  const debugLog = document.getElementById("debugLog");
  debugLog.innerText += message + "\n";
}

