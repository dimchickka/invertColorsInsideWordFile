import os
import zipfile
import io
from PIL import Image
from tempfile import mkdtemp

input_docx = "input.docx"  # Исходный файл
output_docx = "output_inverted.docx"  # Выходной файл

def invert_image_colors(image_bytes):
    """Инвертирует цвета изображения."""
    img = Image.open(io.BytesIO(image_bytes))
    if img.mode == 'RGBA':
        r, g, b, a = img.split()
        rgb_img = Image.merge('RGB', (r, g, b))
        inverted_img = Image.eval(rgb_img, lambda x: 255 - x)
        r, g, b = inverted_img.split()
        inverted_img = Image.merge('RGBA', (r, g, b, a))
    else:
        inverted_img = Image.eval(img.convert('RGB'), lambda x: 255 - x)

    # Сохраняем в байты (PNG)
    img_byte_arr = io.BytesIO()
    inverted_img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()


def process_docx(input_path, output_path):
    """Инвертирует изображения в .docx, работая с ним как с ZIP-архивом."""
    # Создаём временную папку
    temp_dir = mkdtemp()

    try:
        # Распаковываем .docx
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Ищем все изображения в папке word/media/
        media_dir = os.path.join(temp_dir, 'word', 'media')
        if os.path.exists(media_dir):
            for img_file in os.listdir(media_dir):
                img_path = os.path.join(media_dir, img_file)

                # Читаем, инвертируем и перезаписываем изображение
                with open(img_path, 'rb') as f:
                    img_bytes = f.read()

                inverted_img_bytes = invert_image_colors(img_bytes)

                with open(img_path, 'wb') as f:
                    f.write(inverted_img_bytes)

        # Запаковываем обратно в новый .docx
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_out.write(file_path, arcname)

    finally:
        # Удаляем временную папку
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(temp_dir)


if __name__ == "__main__":
    process_docx(input_docx, output_docx)
    print(f"✅ Готово!")