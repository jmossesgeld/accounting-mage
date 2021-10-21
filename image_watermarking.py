from PIL import Image

def apply_watermark(image_file_path, watermark_file_path):
    with Image.open(image_file_path) as im:
        with Image.open(watermark_file_path) as wm:
            im.paste(wm)
            im.save(f"{image_file_path}")

