Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    console.log("Office.js is ready!");

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // document.getElementById("run").onclick = modifyText;
    document.getElementById("run").onclick = alignShapes;

  }
});

// Reposition a shape on a slide
async function repositionShape() {
  try {
    await PowerPoint.run(async (context) => {
      let slide = context.presentation.slides.getItemAt(0); 
      let shape = slide.shapes.getItemAt(0); 

      // Change the position of the shape
      shape.left = 200; 
      shape.top = 150;  
      await context.sync();

      console.log("Shape repositioned successfully");
    });
  } catch (error) {
    console.error("Error repositioning shape:", error);
  }
}


// Modify text of a specific shape on a slide
async function modifyText() {
  try {
    await PowerPoint.run(async (context) => {
      let slide = context.presentation.slides.getItemAt(0);
      let shape = slide.shapes.getItemAt(0); 

      if (shape.textFrame) {
        shape.textFrame.textRange.text = "New Text"; 
        await context.sync();
        console.log("Text modified successfully");
      } else {
        console.log("No text found in this shape.");
      }
    });
  } catch (error) {
    console.error("Error modifying text:", error);
  }
}


// Align shapes relative to each other
async function alignShapes() {
  try {
    await PowerPoint.run(async (context) => {
      let slide = context.presentation.slides.getItemAt(0); 
      let shapes = slide.shapes;
      shapes.load("items"); 
      await context.sync();

      let shape1 = shapes.items[0]; 
      let shape2 = shapes.items[1]; 

      // Align shape2 to be below shape1
      shape2.top = shape1.top + shape1.height + 20; 

      await context.sync();
      console.log("Shapes aligned successfully.");
    });
  } catch (error) {
    console.error("Error aligning shapes:", error);
  }
}



// Modify text, reposition shape, and change layout in one function
async function modifySlideLayoutAndText() {
  try {
    await PowerPoint.run(async (context) => {
      let slide = context.presentation.slides.getItemAt(0); 
      let shape = slide.shapes.getItemAt(0); 

      if (shape.textFrame) {
        shape.textFrame.textRange.text = "Modified Text!";
      }
      shape.left = 300;
      shape.top = 200;

      slide.layout = PowerPoint.SlideLayout.titleOnly;

      await context.sync();
      console.log("Slide modified successfully with new text, position, and layout");
    });
  } catch (error) {
    console.error("Error modifying slide:", error);
  }
}