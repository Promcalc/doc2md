from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import os
import tempfile
import shutil
from app.converter import convert_file

app = FastAPI(title="DOCX/DOC to Markdown Converter")

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    # Проверяем расширение
    filename = file.filename.lower()
    if not (filename.endswith('.doc') or filename.endswith('.docx')):
        raise HTTPException(status_code=400, detail="Only .doc and .docx files are allowed")
    # Сохраняем во временный файл
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name
    try:
        # Конвертируем
        output_path = convert_file(tmp_path)
        # Читаем результат и возвращаем
        with open(output_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        # Удаляем временные файлы
        os.unlink(tmp_path)
        os.unlink(output_path)
        # Возвращаем как файл для скачивания
        response = FileResponse(
            path=output_path,  # но мы уже удалили, поэтому лучше вернуть текст
            media_type='text/markdown',
            filename=os.path.splitext(filename)[0] + '.md'
        )
        # Но FileResponse требует существующий файл, поэтому создадим временный
        with tempfile.NamedTemporaryFile(delete=False, suffix='.md', mode='w', encoding='utf-8') as out:
            out.write(md_content)
            out_path = out.name
        response = FileResponse(out_path, media_type='text/markdown', filename=os.path.splitext(filename)[0] + '.md')
        # Удалим после отправки (можно через dependency)
        return response
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))