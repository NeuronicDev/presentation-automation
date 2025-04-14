from fastapi import APIRouter, Request
import os, json

router = APIRouter()

# --- Metadata Upload Handler ---
@router.post("")
async def upload_metadata(request: Request):
    try:
        body = await request.json()

        filename = body.get("filename", "metadata.json")
        path = body.get("path", "./slide_images")
        data = body.get("data")

        os.makedirs(path, exist_ok=True)
        save_path = os.path.join(path, filename)

        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

        return {"message": "Metadata saved", "path": save_path}
    except Exception as e:
        return {"error": str(e)}