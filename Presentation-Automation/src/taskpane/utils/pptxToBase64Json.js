async function sendPptxAsBase64ToBackend() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, async (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get file:", result.error.message);
                return reject(result.error);
            }
            const file = result.value;
            const slices = [];
            let received = 0;
            const collectSlices = (index) => {
                file.getSliceAsync(index, (sliceResult) => {
                    if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                        console.error("Failed to get slice:", sliceResult.error.message);
                        return reject(sliceResult.error);
                    }
                    slices[index] = sliceResult.value.data;
                    received++;
                    if (received === file.sliceCount) {
                        const totalLength = slices.reduce((sum, slice) => sum + slice.length, 0);
                        const combined = new Uint8Array(totalLength);
                        slices.reduce((offset, slice) => {
                            combined.set(slice, offset);
                            return offset + slice.length;
                        }, 0);
                        const blob = new Blob([combined.buffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
                        const reader = new FileReader();
                        reader.onloadend = () => {
                            const base64 = reader.result.split(',')[1];
                            const payload = {
                                filename: "presentation.pptx",
                                filetype: blob.type,
                                createdAt: new Date().toISOString(),
                                base64: base64
                            };
                            fetch("http://localhost:8000/upload-pptx", {
                                method: "POST",
                                headers: { "Content-Type": "application/json" },
                                body: JSON.stringify(payload)
                            })
                            .then(response => response.json())
                            .then(data => {
                                console.log("Backend response:", data);
                                resolve(data);
                            })
                            .catch(err => {
                                console.error("Error sending to backend:", err);
                                reject(err);
                            });
                        };
                        reader.readAsDataURL(blob);
                        file.closeAsync();
                    } else {
                        collectSlices(received);
                    }
                });
            };
            collectSlices(0);
        });
    });
}