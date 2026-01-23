import layoutparser as lp
import numpy as np
from pdf2image import convert_from_path
from PIL import Image
import io


def extract_layout_image_text_blocks(pdf_path):
    print("[INFO] Starting PDF to image conversion...")
    pages = convert_from_path(pdf_path, dpi=150, thread_count=4)
    print(f"[INFO] Total pages converted: {len(pages)}")

    # Load model
    model_path = r"C:\SmartCatalog\SmartCatalog\src\smartcatalog\model\model_final.pth"
    config_path = r"C:\SmartCatalog\SmartCatalog\src\smartcatalog\model\config.yml"

    print("[INFO] Loading layout model...")
    model = lp.Detectron2LayoutModel(
        config_path=config_path,
        model_path=model_path,
        label_map={0: "Text", 1: "Title", 2: "List", 3: "Table", 4: "Figure"},
        extra_config=["MODEL.ROI_HEADS.SCORE_THRESH_TEST", 0.8],
        device="cpu"
    )

    ocr_agent = lp.TesseractAgent(languages="eng+vie")
    results = []

    for page_idx, image in enumerate(pages):
        print(f"[INFO] Processing page {page_idx + 1}...")
        layout = model.detect(image)

        image_blocks = []
        text_blocks = []

        for block in layout:
            try:
                bbox = [int(x) for x in block.coordinates]
                segment_image = block.crop_image(np.array(image))
            except Exception as e:
                print(f"[ERROR] Cropping failed: {e}")
                continue

            if block.type == "Figure":
                width = bbox[2] - bbox[0]
                height = bbox[3] - bbox[1]
                aspect_ratio = height / (width + 1e-5)

                # Skip multi-product or tiny blocks
                if aspect_ratio < 1.5 or height < 200:
                    continue

                buf = io.BytesIO()
                Image.fromarray(segment_image).save(buf, format='PNG')
                image_blocks.append({
                    "page": page_idx + 1,
                    "bbox": bbox,
                    "image_bytes": buf.getvalue(),
                })

            elif block.type in ["Text", "Title"]:
                try:
                    text = ocr_agent.detect(segment_image, return_response=False)
                except Exception as e:
                    print(f"[ERROR] OCR failed: {e}")
                    text = ""
                text_blocks.append({
                    "text": text.strip(),
                    "bbox": bbox,
                })

        results.append({
            "page": page_idx + 1,
            "images": image_blocks,
            "texts": text_blocks,
        })

    print("[INFO] Extraction complete.")
    return results
