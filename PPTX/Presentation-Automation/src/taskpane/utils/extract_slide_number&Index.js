// async function getCurrentSlideIndex() {
//     return new Promise((resolve, reject) => {
//       Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           try {
//             const slideIndex = result.value?.slides?.[0]?.index;
//             if (slideIndex === undefined) {
//               console.warn("Slide index not found.");
//               return reject("Slide index not found.");
//             }
//             const slideNumber = slideIndex + 1;
//             const payload = {
//               slideIndex: slideIndex,
//               slideNumber: slideNumber,
//               capturedAt: new Date().toISOString()
//             };
//             fetch("http://localhost:8000/upload-slide-info", {
//               method: "POST",
//               headers: { "Content-Type": "application/json" },
//               body: JSON.stringify(payload)
//             })
//             .then(response => response.json())
//             .then(data => {
//               console.log("Slide info sent to backend:", data);
//               resolve(data);
//             })
//             .catch(err => {
//               console.error("Failed to send slide info:", err);
//               reject(err);
//             });
  
//           } catch (e) {
//             console.error("Error extracting slide info:", e);
//             reject("Error extracting slide info.");
//           }
//         } else {
//           console.error("Failed to get selected slide:", result.error.message);
//           reject("Could not get slide info.");
//         }
//       });
//     });
//   }