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

    // Attempt both tasks in parallel
    await Promise.allSettled([
      extractAllSlideShapes().then(() => (metadataComplete = true)).catch(() => {}),
      sendPptxAsBase64ToBackend().then(() => (pptxComplete = true)).catch(() => {})
    ]);

    // Fallback: retry in series if needed
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

    // Proceed only if both succeeded
    if (metadataComplete && pptxComplete) {
      try {
        console.log("Processing instructions...");
        await processInstruction();
        // await sendInstructionToBackend();
        console.log("Slide processing complete.");
      } catch (err) {
        console.error("Instruction processing failed:", err);
      }
    } else {
      console.error("Operation aborted: required steps failed.");
    }
  };
});

// async function getCurrentSlideIndex() {
//   return new Promise((resolve, reject) => {
//     Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         const slideIndex = result.value?.slides?.[0]?.index;
//         if (slideIndex !== undefined) {
          
//           resolve(slideIndex);
//         } else {
//           reject("Slide index not found.");
//         }
//       } else {
//         reject("Failed to get selected slide.");
//       }
//     });
//   });
// }

// async function handleInstructionAction(result) {
//   const task = result.task;

//   if (task === "change_font_size") {
//       // await changeTitleFontSize(result.font_size);
//       console.log("change_font_size triggered")
//   } else if (task === "add_shape") {
//       // await addShape(result.shape, result.position, result.insert_title);
//       console.log("add_shape triggered")
//   } else if (task === "cleanup_slide") {
//       console.log("cleanup triggered")
//       await processInstruction(); 
//   } else {
//       console.warn("Unrecognized task. Please handle manually.");
//   }
// }

// async function sendInstructionToBackend() {
//   const instruction = document.getElementById("instructionInput")?.value;
//   if (!instruction) {
//     console.warn("Instruction input is empty.");
//     return;
//   }

//   let slideNumber;
//   try {
//     slideNumber = await getCurrentSlideIndex();
//   } catch (err) {
//     console.error("Could not determine current slide index:", err);
//     return;
//   }

//   try {
//     const response = await fetch("http://localhost:8000/agent/classify-instruction", {
//       method: "POST",
//       headers: { "Content-Type": "application/json" },
//       body: JSON.stringify({
//         instruction: instruction.trim(),
//         slide_number: slideNumber
//       })
//     });
//     const data = await response.json();
//     console.log("Backend Response:", data);

//     if (data?.result?.task) {
//       handleInstructionAction(data.result);
//     }

//   } catch (err) {
//     console.error("Failed to send instruction to backend:", err);
//   }
// }

async function processInstruction() {
  const instruction = document.getElementById("instructionInput").value.toLowerCase();
  const log = document.getElementById("debugLog");
  log.textContent = "Triggered: " + instruction + "\nStarting cleanup agent...\n";

  try {
    const response = await fetch("http://localhost:8000/agent/cleanup", {
      method: "POST",
    });

    const data = await response.json();

    if (data.status === "success") {
      log.textContent += "Cleanup instructions received:\n\n";
      log.textContent += data.instructions + "\n";

      const codeResponse = await fetch("http://localhost:8000/agent/generate_code", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ instructions: data.instructions })
      });

      const codeData = await codeResponse.json();

      if (codeData.code) {
        let rawCode = codeData.code.trim();

        if (rawCode.startsWith("```")) {
          rawCode = rawCode
            .replace(/^```[a-z]*\n?/i, "")
            .replace(/```$/, "")
            .trim();
        }

        console.log("Cleaned Generated Code:\n", rawCode);
        log.textContent += "\nExecuting generated code...\n";

        if (!rawCode || rawCode.length < 10) {
          log.textContent += " Warning: Code is empty or suspicious.\n";
          return;
        }        
        try {
          const dynamicFunc = new Function(`
            return (async () => {
              await PowerPoint.run(async (context) => {
                ${rawCode}
              });
            })();
          `);
          await dynamicFunc();
          log.textContent += "Code executed successfully.\n";
        } catch (e) {
          log.textContent += "Code execution error: " + e.message + "\n";
        }
      } else {
        log.textContent += "No code returned by generator.\n";
      }
    } else {
      log.textContent += "Cleanup failed: " + data.message + "\n";
    }

  } catch (err) {
    log.textContent += "Error contacting backend: " + err.message + "\n";
  }
}

async function logToDebugLog(message) {
  const logBox = document.getElementById("debug-log");
  logBox.value += '${message}\n\n';
  logBox.scrollTop = logBox.scrollHeight;
}

function logDebug(message) {
  const debugLog = document.getElementById("debugLog");
  debugLog.innerText += message + "\n";
}