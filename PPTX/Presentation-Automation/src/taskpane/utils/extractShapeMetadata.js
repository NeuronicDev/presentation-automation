async function extractAllSlideShapes() {
  console.log("[Sync Metadata] Starting full metadata extraction...");
  let overallSuccess = true;

  try {
      await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items/id");
          await context.sync();

          if (!slides.items || slides.items.length === 0) {
              console.warn("[Sync Metadata] Presentation contains no slides.");
              return;
          }
          const slideCount = slides.items.length;
          console.log(`[Sync Metadata] Found ${slideCount} slides. Processing sequentially...`);

          for (let s = 0; s < slideCount; s++) {
              console.log(`[Sync Metadata] ---> Processing Slide Index ${s}...`);
              let currentSlideMetadata = [];
              let slideSuccess = false;

              try {
                  const slide = context.presentation.slides.getItemAt(s);
                  const shapes = slide.shapes;
                  shapes.load("items/id, items/name, items/left, items/top, items/width, items/height, items/type, items/altTextDescription");
                  await context.sync();

                  const shapesItems = shapes.items;
                  const shapesToGetTextFrom = [];
                  const basicMetadataMap = new Map();

                  for (let i = 0; i < shapesItems.length; i++) {
                      const shape = shapesItems[i];
                      const standardizedName = `Google Shape;${shape.id};p${s + 1}`;
                      const shapeData = {
                          id: shape.id,
                          name: standardizedName,
                          parentGroupId: null,
                          top: shape.top, left: shape.left,
                          width: shape.width, height: shape.height,
                          text: "",
                          type: shape.type || "Unknown",
                          altText: shape.altTextDescription || "",
                          zIndex: i,
                          isLikelyIcon: shape.type === "Graphic" || (shape.name?.toLowerCase().includes("icon")),
                          slideIndex: s,
                          overlapsWith: [],
                          font: { name: null, size: null, bold: false, italic: false },
                          textAlign: null
                      };
                      basicMetadataMap.set(shape.id, shapeData);

                      if (shape.type === 'TextBox' || shape.type === 'Placeholder' || shape.type.includes('Text')) {
                          shapesToGetTextFrom.push(shape);
                      }
                  }

                  for (let i = 0; i < shapesItems.length; i++) {
                      const shapeA = shapesItems[i];
                      for (let j = i + 1; j < shapesItems.length; j++) {
                          const shapeB = shapesItems[j];
                          if (isOverlapping(shapeA, shapeB)) {
                              basicMetadataMap.get(shapeA.id).overlapsWith.push(shapeB.id);
                              basicMetadataMap.get(shapeB.id).overlapsWith.push(shapeA.id);
                          }
                      }
                  }

                  if (shapesToGetTextFrom.length > 0) {
                      console.log(`[Sync Metadata] Slide ${s}: Loading text for ${shapesToGetTextFrom.length} shapes...`);
                      for (const shape of shapesToGetTextFrom) {
                          shape.load("textFrame/textRange/text, textFrame/textRange/font/name, textFrame/textRange/font/size, textFrame/textRange/font/bold, textFrame/textRange/font/italic, textFrame/textRange/textAlign");
                      }
                      await context.sync();

                      for (const shape of shapesToGetTextFrom) {
                          const entry = basicMetadataMap.get(shape.id);
                          if (!entry) continue;

                          try {
                              const tr = shape.textFrame?.textRange;
                              if (tr) {
                                  entry.text = tr.text?.trim() || "";
                                  entry.font.name = tr.font.name;
                                  entry.font.size = tr.font.size;
                                  entry.font.bold = tr.font.bold;
                                  entry.font.italic = tr.font.italic;
                                  entry.textAlign = tr.textAlign;
                              }
                          } catch (textError) {
                              if (!textError.message?.includes("RichApi.Error")) {
                                  console.warn(`[Sync Metadata] Slide ${s}: Could not load text for shape ID ${shape.id}: ${textError.message}`);
                              }
                          }
                      }
                      console.log(`[Sync Metadata] Slide ${s}: Text loading complete.`);
                  }
                  currentSlideMetadata = Array.from(basicMetadataMap.values());
                  console.log(currentSlideMetadata) ////////////////////////////////////////////////////////////////
                  slideSuccess = true;
              } catch (slideError) {
                  console.error(`[Sync Metadata] Failed to process slide index ${s}:`, slideError);
                  overallSuccess = false;
              }

              if (slideSuccess && currentSlideMetadata.length > 0) {
                  console.log(`[Sync Metadata] Slide ${s} processed. Uploading metadata_${s}.json...`);
                  try {
                      const response = await fetch("http://localhost:8000/upload-metadata", {
                          method: "POST",
                          headers: { "Content-Type": "application/json" },
                          body: JSON.stringify({
                              filename: `metadata_${s}.json`,
                              path: "slide_images/metadata",
                              data: currentSlideMetadata
                          })
                      });

                      if (!response.ok) {
                          throw new Error(`Upload failed with status ${response.status}`);
                      }
                      const result = await response.json();
                      console.log(`[Sync Metadata] Slide ${s} metadata uploaded:`, result);
                  } catch (uploadError) {
                      console.error(`[Sync Metadata] Upload failed for Slide ${s}:`, uploadError);
                      overallSuccess = false;
                  }
              } else if (slideSuccess) {
                  console.log(`[Sync Metadata] Slide ${s} had no shapes to upload.`);
              }
          }
      });
      console.log("[Sync Metadata] Finished processing all slides.");
      return overallSuccess;
  } catch (error) {
      console.error("[Sync Metadata] Unexpected error during overall process:", error);
      return false;
  }
}

function isOverlapping(shapeA, shapeB) {
  return !(
      shapeA.left + shapeA.width < shapeB.left ||
      shapeA.left > shapeB.left + shapeB.width ||
      shapeA.top + shapeA.height < shapeB.top ||
      shapeA.top > shapeB.top + shapeB.height
  );
}