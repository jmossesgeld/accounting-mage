from PIL import Image
import os
import zipfile
import io
import pathlib
from shutil import rmtree


def apply_watermark(image_file_path, watermark_file_path):
    with Image.open(image_file_path) as im:
        with Image.open(watermark_file_path) as wm:
            im.paste(wm)
            im.save(f"{image_file_path}")


def convert_images(files):
    try:
        rmtree('UPLOADED_FILES')
    except:
        print("Folder not found.")
    try:
        os.makedirs('UPLOADED_FILES')
    except:
        print("Folder already exists.")
    for file in files:
        path = os.path.join('UPLOADED_FILES/', file.filename)
        file.save(path)
        apply_watermark(path, os.path.join(os.getcwd(), 'projects/image_watermarking/watermark.png'))

    base_path = pathlib.Path('UPLOADED_FILES')
    data = io.BytesIO()
    with zipfile.ZipFile(data, mode='w') as z:
        for f_name in base_path.iterdir():
            z.write(f_name)
    data.seek(0)
    return data
