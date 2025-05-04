
// Office.onReady((info) => {
//   if (info.host === Office.HostType.PowerPoint) {
//     console.log("Office.js is ready!");

//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";

//     document.getElementById("run").onclick = extractStickyNotes;
//   }
// });

// export async function extractStickyNotes() {
//   try {
//     await PowerPoint.run(async (context) => {
//       let slides = context.presentation.slides;
//       slides.load("items");
//       await context.sync();

//       let allText = [];

//       for (let i = 0; i < slides.items.length; i++) {
//         let slide = slides.items[i];
//         let shapes = slide.shapes;
//         shapes.load("items/textFrame/textRange/text");
//         await context.sync();

//         let slideText = [];
//         for (let j = 0; j < shapes.items.length; j++) {
//           let text = shapes.items[j].textFrame.textRange.text;
//           if (text && text.trim() !== "") {
//             slideText.push(text.trim());
//           }
//         }

//         if (slideText.length > 0) {
//           allText.push(`Slide ${i + 1}: ${slideText.join(" ")}`);
//         }
//       }

//       // Step 3: Send text to Groq API for filtering
//       let filteredStickyNotes = await validateStickyNotes(allText);

//       console.log("Extracted Sticky Notes:\n", filteredStickyNotes);

//       // Step 4: Display in Task Pane
//       document.getElementById("item-subject").innerText = filteredStickyNotes || "No sticky notes found.";
//     });
//   } catch (error) {
//     console.error("Error extracting sticky notes:", error);
//   }
// }

// // Step 3: Function to send extracted text to Groq API
// async function validateStickyNotes(textArray) {
//   const API_KEY = "API-key"; 
//   const GROQ_URL = "https://api.groq.com/openai/v1/chat/completions";

//   const prompt = `Extract and return only the instructional content from the following text.
//   Instructional content includes tasks, action points, to-do lists, reminders, or directives.
//   Do not return general text, just instructions as plain text.\n\nText: ${textArray.join("\n")}\n\nExtracted Instructions:`;

//   const payload = {
//     model: "llama3-70b-8192",
//     messages: [{ role: "user", content: prompt }],
//     temperature: 0.2,
//   };

//   try {
//     const response = await fetch(GROQ_URL, {
//       method: "POST",
//       headers: {
//         "Authorization": `Bearer ${API_KEY}`,
//         "Content-Type": "application/json",
//       },
//       body: JSON.stringify(payload),
//     });

//     if (!response.ok) {
//       throw new Error("Failed to fetch from Groq API");
//     }

//     const data = await response.json();
//     return data.choices[0].message.content.trim();
//   } catch (error) {
//     console.error("Error in Groq API:", error);
//     return "Validation Failed";
//   }
// }

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
      console.log("Office.js is ready!");
  
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
  
      document.getElementById("run").onclick = extractStickyNotes;
    }
  });
  
  export async function extractStickyNotes() {
  try {
    await PowerPoint.run(async (context) => {
      let slides = context.presentation.slides;
      slides.load("items"); // Load all slides
      await context.sync();
  
      let allText = [];
  
      // Loop through each slide and process shapes
      for (let i = 0; i < slides.items.length; i++) {
        let slide = slides.items[i];
        let shapes = slide.shapes;
        shapes.load("items"); // Load shapes for this slide
        await context.sync();
  
        let slideText = [];
        
        // Process each shape to extract text if available
        for (let j = 0; j < shapes.items.length; j++) {
          let shape = shapes.items[j];
  
          // Check if the shape has a textFrame
          if (shape.textFrame) {
            try {
              shape.textFrame.load("textRange/text"); // Load text for shapes with textFrame
              await context.sync();
  
              let text = shape.textFrame.textRange.text;
              if (text && text.trim() !== "") {
                slideText.push(text.trim()); // Collect non-empty text
              }
            } catch (error) {
              console.error(`Error reading text from shape ${j} on slide ${i + 1}:`, error);
              continue; // Skip this shape if it fails to load text
            }
          }
        }
  
        // Only add slide text if it contains valid content
        if (slideText.length > 0) {
          allText.push(`Slide ${i + 1}: ${slideText.join(" ")}`);
        }
  
        // Optional: Add some logging for progress
        if (i % 5 === 0) {
          console.log(`Processed ${i + 1} slides...`);
        }
      }
  
      // Step 3: Send extracted text to Groq API for filtering
      let filteredStickyNotes = await validateStickyNotes(allText);
  
      console.log("Extracted Sticky Notes:\n", filteredStickyNotes);
  
      // Step 4: Display in Task Pane
      document.getElementById("item-subject").innerText = filteredStickyNotes || "No sticky notes found.";
    });
  } catch (error) {
    console.error("Error extracting sticky notes:", error);
  }
  }
  
  // Function to send extracted text to Groq API for filtering
  async function validateStickyNotes(textArray) {
  const API_KEY = "API-key"; 
  const GROQ_URL = "https://api.groq.com/openai/v1/chat/completions";
  
  const prompt = `Extract and return only the instructional content from the following text.
  Instructional content includes tasks, action points, to-do lists, reminders, or directives.
  Do not return general text, just instructions as plain text.\n\nText: ${textArray.join("\n")}\n\nExtracted Instructions:`;
  
  const payload = {
    model: "llama3-70b-8192",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.2,
  };
  
  try {
    const response = await fetch(GROQ_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });
  
    if (!response.ok) {
      throw new Error("Failed to fetch from Groq API");
    }
  
    const data = await response.json();
    return data.choices[0].message.content.trim();
  } catch (error) {
    console.error("Error in Groq API:", error);
    return "Validation Failed";
  }
  }
  