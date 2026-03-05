import zipfile
import os
import sys
import shutil
from pathlib import Path


def extract_docx_flat(docx_path, output_dir):
    """
    Извлекает все файлы из DOCX в одну папку (без подпапок)
    и переименовывает .xml в .txt
    """
    # Создаём временную папку для распаковки
    temp_dir = Path("temp_extract")
    temp_dir.mkdir(exist_ok=True)

    # Создаём выходную папку
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # Извлекаем всё во временную папку
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(temp_dir)

    # Копируем все файлы в выходную папку с новыми именами
    for file_path in temp_dir.rglob("*"):
        if file_path.is_file():
            # Формируем новое имя: папка_имя.расширение
            rel_path = file_path.relative_to(temp_dir)
            new_name = str(rel_path).replace("\\", "_").replace("/", "_")

            # Если это XML, меняем расширение на .txt
            if new_name.endswith('.xml'):
                new_name = new_name[:-4] + '.txt'

            dest_path = out_path / new_name
            shutil.copy2(file_path, dest_path)
            print(f"Скопирован: {new_name}")

    # Удаляем временную папку
    shutil.rmtree(temp_dir)

    print(f"\n✅ Все файлы скопированы в одну папку: {out_path}")
    print(f"📁 Отправьте всю папку {out_path} или заархивируйте её")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Использование: python extract_flat.py <путь_к_docx> <выходная_папка>")
        sys.exit(1)

    extract_docx_flat(sys.argv[1], sys.argv[2])