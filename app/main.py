import os
import uuid
import shutil
from fastapi import FastAPI, UploadFile
from dotenv import load_dotenv

from app.supabase_client import supabase
from app.ppt_processor import extract_text_slidewise, convert_ppt_to_images

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
BUCKET_PPTS = os.getenv("BUCKET_PPTS")
BUCKET_IMAGES = os.getenv("BUCKET_IMAGES")

app = FastAPI()


@app.get("/")
def health():
    return {"status": "PPT service running"}


@app.post("/upload-ppt/")
async def upload_ppt(file: UploadFile):
    user_id = str(uuid.uuid4())
    presentation_id = str(uuid.uuid4())

    # Windows + Cloud safe working dir
    work_dir = os.path.join("workdir", user_id, presentation_id)
    os.makedirs(work_dir, exist_ok=True)

    ppt_path = os.path.join(work_dir, file.filename)

    try:
        # 1️⃣ Save PPT
        with open(ppt_path, "wb") as f:
            f.write(await file.read())

        # 2️⃣ Upload PPT to Supabase
        ppt_storage_path = f"{user_id}/{presentation_id}.pptx"
        supabase.storage.from_(BUCKET_PPTS).upload(
            ppt_storage_path,
            ppt_path
        )

        # 3️⃣ Convert PPT → images (ORDER GUARANTEED)
        image_files = convert_ppt_to_images(ppt_path, work_dir)

        # 4️⃣ Extract text (ORDER GUARANTEED)
        slides_text = extract_text_slidewise(ppt_path)

        # SAFETY CHECK
        if len(image_files) != len(slides_text):
            raise Exception("Slide count and image count mismatch")

        # 5️⃣ Insert presentation
        supabase.table("presentations").insert({
            "id": presentation_id,
            "user_id": user_id,
            "title": file.filename,
            "ppt_path": ppt_storage_path,
            "total_slides": len(slides_text)
        }).execute()

        # 6️⃣ PERFECT 1-to-1 MAPPING (INDEX BASED)
        for index in range(len(slides_text)):
            slide = slides_text[index]
            image_file = image_files[index]
            slide_no = slide["slide_number"]

            image_storage_path = f"{user_id}/{presentation_id}/slide_{slide_no}.jpg"

            # Upload image
            supabase.storage.from_(BUCKET_IMAGES).upload(
                image_storage_path,
                image_file
            )

            # Public URL
            image_url = (
                f"{SUPABASE_URL}/storage/v1/object/public/"
                f"{BUCKET_IMAGES}/"
                f"{image_storage_path}"
            )

            # Insert slide row
            supabase.table("slides").insert({
                "presentation_id": presentation_id,
                "user_id": user_id,
                "slide_number": slide_no,
                "image_url": image_url,
                "extracted_text": slide["text"]
            }).execute()

        return {
            "status": "success",
            "presentation_id": presentation_id,
            "slides": len(slides_text)
        }

    finally:
        # 7️⃣ Cleanup
        shutil.rmtree(work_dir, ignore_errors=True)
