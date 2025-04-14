async function extractAllSlideShapes() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const allMetadata = [];

    function isOverlapping(a, b) {
      return !(
        a.left + a.width < b.left ||
        b.left + b.width < a.left ||
        a.top + a.height < b.top ||
        b.top + b.height < a.top
      );
    }

    async function extractShapeDetails(shape, slideIndex, zIndex, parentGroupId = null) {
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
        } catch (e) {
          // Shape has no text or textFrame
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
          slideIndex: slideIndex,
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

        // Handle groups recursively
        if (typeName === "Group" && shape.groupItems) {
          shape.groupItems.load("items");
          await context.sync();

          for (let i = 0; i < shape.groupItems.items.length; i++) {
            await extractShapeDetails(shape.groupItems.items[i], slideIndex, `${zIndex}.${i}`, shape.id);
          }
        }

      } catch (e) {
        console.warn("Skipped shape due to error:", e);
      }
    }

    // Iterate through all slides
    for (let s = 0; s < slides.items.length; s++) {
      const slide = slides.items[s];
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      for (let i = 0; i < shapes.items.length; i++) {
        await extractShapeDetails(shapes.items[i], s, i);
      }
    }

    // Check overlaps
    for (let i = 0; i < allMetadata.length; i++) {
      for (let j = 0; j < allMetadata.length; j++) {
        if (i !== j && isOverlapping(allMetadata[i], allMetadata[j])) {
          allMetadata[i].overlapsWith.push(allMetadata[j].id);
        }
      }
    }

    // Print to console
    const metadataJSON = JSON.stringify(allMetadata, null, 2);
    console.log("Extracted Slide Shape Metadata with Text Properties:\n", metadataJSON);

    // Send to backend
    fetch("http://localhost:8000/upload-metadata", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        filename: "metadata.json",
        path: "slide_images",
        data: allMetadata
      })
    })
    .then(res => res.json())
    .then(data => console.log("Metadata saved to backend:", data))
    .catch(err => console.error("Failed to send metadata to backend:", err));

    return allMetadata;
  });
}