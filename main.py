from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from typing import List
import os
import re
import zipfile
import subprocess
from datetime import datetime
from pathlib import Path

app = FastAPI(
    title="华东院文件格式转换工具",
    description="文档格式转换API，支持Office文档和PDF转换",
    version="2.0.0"
)

# Configuration
UPLOAD_FOLDER = 'uploads/'
DOWNLOAD_FILES_FOLDER = 'downloads/files/'
DOWNLOAD_ZIPS_FOLDER = 'downloads/zips/'
LIBREOFFICE_PATH = 'D:/LibreOfficePortable/App/libreoffice/program/soffice.exe'

# Create necessary directories
for folder in [UPLOAD_FOLDER, DOWNLOAD_FILES_FOLDER, DOWNLOAD_ZIPS_FOLDER]:
    os.makedirs(folder, exist_ok=True)


def secure_filename(filename: str) -> str:
    """
    生成安全的文件名，支持中文字符
    """
    if isinstance(filename, str):
        from unicodedata import normalize
        filename = normalize('NFKD', filename).encode('utf-8', 'ignore')
        filename = filename.decode('utf-8')
    
    for sep in os.path.sep, os.path.altsep:
        if sep:
            filename = filename.replace(sep, ' ')
    
    # 保留中文字符
    _filename_ascii_add_strip_re = re.compile(r'[^A-Za-z0-9_\u4E00-\u9FBF.-]')
    filename = str(_filename_ascii_add_strip_re.sub('', '_'.join(filename.split()))).strip('._')
    
    return filename


def convert_with_libreoffice(input_path: str, output_dir: str, target_format: str) -> bool:
    """
    使用LibreOffice进行文件格式转换
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录
        target_format: 目标格式 (docx, xlsx, pptx, pdf)
    
    Returns:
        bool: 转换是否成功
    """
    try:
        subprocess.run(
            [LIBREOFFICE_PATH, '--headless', '--convert-to', target_format, input_path, '--outdir', output_dir],
            check=True,
            capture_output=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"转换失败 {input_path}: {e}")
        return False


def convert_to_office_format(file_path: str, filename: str, output_dir: str) -> str:
    """
    将旧版Office文档转换为新版格式
    
    Args:
        file_path: 原始文件路径
        filename: 文件名
        output_dir: 输出目录
    
    Returns:
        str: 转换后的文件路径，如果不需要转换则返回原文件复制后的路径
    """
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]
    
    # 旧版格式需要转换
    if ext == '.doc':
        target_format = 'docx'
        new_filename = f"{base_name}.docx"
    elif ext == '.xls':
        target_format = 'xlsx'
        new_filename = f"{base_name}.xlsx"
    elif ext == '.ppt':
        target_format = 'pptx'
        new_filename = f"{base_name}.pptx"
    # 新版格式直接复制
    elif ext in ['.docx', '.xlsx', '.pptx']:
        import shutil
        new_path = os.path.join(output_dir, filename)
        shutil.copy2(file_path, new_path)
        return new_path
    else:
        # 不支持的格式，也复制过去
        import shutil
        new_path = os.path.join(output_dir, filename)
        shutil.copy2(file_path, new_path)
        return new_path
    
    # 执行转换
    if convert_with_libreoffice(file_path, output_dir, target_format):
        return os.path.join(output_dir, new_filename)
    else:
        raise Exception(f"转换失败: {filename}")


def convert_to_pdf(file_path: str, filename: str, output_dir: str) -> str:
    """
    将文档转换为PDF格式
    
    Args:
        file_path: 原始文件路径
        filename: 文件名
        output_dir: 输出目录
    
    Returns:
        str: 转换后的PDF文件路径
    """
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]
    
    # 支持的文档格式
    if ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
        new_filename = f"{base_name}.pdf"
        if convert_with_libreoffice(file_path, output_dir, 'pdf'):
            return os.path.join(output_dir, new_filename)
        else:
            raise Exception(f"PDF转换失败: {filename}")
    else:
        # 不支持的格式，复制原文件
        import shutil
        new_path = os.path.join(output_dir, filename)
        shutil.copy2(file_path, new_path)
        return new_path


@app.get("/", response_class=HTMLResponse)
async def index():
    """
    返回HTML前端页面
    """
    html_path = os.path.join("templates", "index.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()


@app.post("/api/convert")
async def convert_files(
    files: List[UploadFile] = File(...),
    target_format: str = Form(...)
):
    """
    文件转换API
    
    参数:
    - files: 上传的文件列表
    - target_format: 目标格式 ('office' 或 'pdf')
    
    返回:
    - 单个文件: 直接返回转换后的文件
    - 多个文件: 返回包含所有转换文件的ZIP压缩包
    """
    if not files:
        raise HTTPException(status_code=400, detail="未选择文件")
    
    if target_format not in ['office', 'pdf']:
        raise HTTPException(status_code=400, detail="目标格式必须是 'office' 或 'pdf'")
    
    converted_files = []
    
    try:
        # 处理每个上传的文件
        for file in files:
            filename = secure_filename(file.filename)
            
            # 保存上传的文件
            upload_path = os.path.join(UPLOAD_FOLDER, filename)
            with open(upload_path, "wb") as buffer:
                content = await file.read()
                buffer.write(content)
            
            # 根据目标格式进行转换
            if target_format == 'office':
                converted_path = convert_to_office_format(upload_path, filename, DOWNLOAD_FILES_FOLDER)
            else:  # pdf
                converted_path = convert_to_pdf(upload_path, filename, DOWNLOAD_FILES_FOLDER)
            
            converted_files.append(converted_path)
            
            # 清理上传的临时文件
            os.remove(upload_path)
        
        # 如果只有一个文件，直接返回该文件
        if len(converted_files) == 1:
            file_path = converted_files[0]
            filename = os.path.basename(file_path)
            from urllib.parse import quote
            encoded_filename = quote(filename)
            return FileResponse(
                path=file_path,
                filename=filename,
                media_type='application/octet-stream',
                headers={"Content-Disposition": f"attachment; filename*=utf-8''{encoded_filename}"}
            )
        
        # 多个文件，创建ZIP压缩包
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        zip_filename = f'converted_files_{timestamp}.zip'
        zip_path = os.path.join(DOWNLOAD_ZIPS_FOLDER, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in converted_files:
                zipf.write(file_path, os.path.basename(file_path))
        
        return FileResponse(
            path=zip_path,
            filename=zip_filename,
            media_type='application/zip'
        )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")


@app.get("/api/health")
async def health_check():
    """
    健康检查接口
    """
    return {
        "status": "ok",
        "libreoffice_available": os.path.exists(LIBREOFFICE_PATH)
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=7788)
