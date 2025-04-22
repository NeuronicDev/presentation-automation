async function extractAllSlideShapes() {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      if (!slides.items || slides.items.length === 0) {
        console.warn("Presentation contains no slides.");
        return [];
      }
      console.log(`Found ${slides.items.length} slides. Processing metadata...`);
      const allMetadata = [];
      function isOverlapping(a, b) {
        return !(
          a.left + a.width < b.left ||
          b.left + b.width < a.left ||
          a.top + a.height < b.top ||
          b.top + b.height < a.top
        );
      }

      async function extractShapeDetails(shape, context, slideIndex, zIndex, parentGroupId = null) {
        try {
          shape.load("id, name, left, top, width, height, type, altTextDescription");
          await context.sync();
  
          let shapeText = "";
          let typeName = shape.type || "Unknown";
          let fontName = null;
          let fontSize = null;
          let fontBold = false;
          let fontItalic = false;
          let textAlign = null;

          try {
            if (shape.textFrame) {
              shape.textFrame.load("textRange");
              await context.sync();
  
              if (shape.textFrame.textRange) {
                const textRange = shape.textFrame.textRange;
                textRange.load("text, font/name, font/size, font/bold, font/italic, textAlign");
                await context.sync();
                shapeText = textRange.text ? textRange.text.trim() : "";
                fontName = textRange.font.name;
                fontSize = textRange.font.size;
                fontBold = textRange.font.bold;
                fontItalic = textRange.font.italic;
                textAlign = textRange.textAlign;
              }
            }
          } catch (_) {
          }

          const isLikelyIcon = typeName === "Graphic" || (shape.name && shape.name.toLowerCase().includes("icon"));

          const shapeData = {
            id: shape.id,
            name: shape.name || "",
            parentGroupId: parentGroupId,
            top: shape.top,
            left: shape.left,
            width: shape.width,
            height: shape.height,
            text: shapeText,
            type: typeName,
            altText: shape.altTextDescription || "",
            zIndex: zIndex,
            isLikelyIcon: isLikelyIcon,
            slideIndex: slideIndex + 1,
            overlapsWith: [],
            font: {
              name: fontName,
              size: fontSize,
              bold: fontBold,
              italic: fontItalic
            },
            textAlign: textAlign
          };
  
          allMetadata.push(shapeData);
  
          if (typeName === "Group" && shape.groupItems) {
            shape.groupItems.load("items");
            await context.sync();
  
            for (let i = 0; i < shape.groupItems.items.length; i++) {
              await extractShapeDetails(shape.groupItems.items[i], context, slideIndex, `${zIndex}.${i}`, shape.id);
            }
          }
        } catch (err) {
          console.warn("Skipped shape due to error:", err);
        }
      }

      async function processSlide(slide, index) {
        try {
          const shapes = slide.shapes;
          shapes.load("items");
          await context.sync();
  
          for (let i = 0; i < shapes.items.length; i++) {
            await extractShapeDetails(shapes.items[i], context, index, i);
          }
        } catch (err) {
          console.error(`Error processing slide ${index}:`, err);
        }
      }

    // Process each slide
    for (let s = 0; s < slides.items.length; s++) {
        await processSlide(slides.items[s], s);
      }
  
      // Calculate overlaps (only within same slide)
      for (let i = 0; i < allMetadata.length; i++) {
        for (let j = 0; j < allMetadata.length; j++) {
          if (
            i !== j &&
            allMetadata[i].slideIndex === allMetadata[j].slideIndex &&
            isOverlapping(allMetadata[i], allMetadata[j])
          ) {
            allMetadata[i].overlapsWith.push(allMetadata[j].id);
          }
        }
      }

    // Print to console
    const metadataJSON = JSON.stringify(allMetadata, null, 2);
    console.log("Extracted Slide Shape Metadata with Text Properties:\n", metadataJSON);

    // Send to backend
    const filename = `metadata.json`;
    try {
      const response = await fetch("http://localhost:8000/upload-metadata", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          filename: filename,
          path: "slide_images",
          data: allMetadata
        })
      });
      const result = await response.json();
      console.log("Metadata uploaded to backend:", result);
    } catch (err) {
      console.error("Failed to send metadata to backend:", err);
    }
    return allMetadata;
  });
}