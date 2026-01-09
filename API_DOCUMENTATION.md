# 华东院文件格式转换工具 API 文档

## 概述

本API提供文档格式转换服务，支持将旧版Office文档转换为新版格式，或将文档转换为PDF格式。

**基础URL**: `http://localhost:7788`

**API版本**: 2.0.0

## 交互式文档

本API使用FastAPI框架，提供了自动生成的交互式API文档：

- **Swagger UI**: [http://localhost:7788/docs](http://localhost:7788/docs)
- **ReDoc**: [http://localhost:7788/redoc](http://localhost:7788/redoc)

## 端点列表

### 1. 文件转换 API

转换上传的文档文件。

**端点**: `POST /api/convert`

**请求参数**:

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| files | File[] | 是 | 要转换的文件列表（支持多文件上传） |
| target_format | string | 是 | 目标格式：`office` 或 `pdf` |

**target_format 选项说明**:

- `office`: 将旧版Office文档转换为新版格式
  - `.doc` → `.docx`
  - `.xls` → `.xlsx`
  - `.ppt` → `.pptx`
  - 新版格式文件（`.docx`, `.xlsx`, `.pptx`）保持不变
  
- `pdf`: 将所有文档转换为PDF格式
  - 支持的输入格式：`.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`
  - 输出格式：`.pdf`

**响应**:

- **单文件上传**: 直接返回转换后的文件
- **多文件上传**: 返回包含所有转换文件的ZIP压缩包

**响应头**:

```
Content-Type: application/octet-stream (单文件) 或 application/zip (多文件)
Content-Disposition: attachment; filename="<文件名>"
```

**状态码**:

- `200`: 转换成功，返回文件
- `400`: 请求参数错误（未选择文件或格式不正确）
- `500`: 服务器错误（转换失败）

---

### 2. 健康检查 API

检查API服务状态和LibreOffice可用性。

**端点**: `GET /api/health`

**响应示例**:

```json
{
  "status": "ok",
  "libreoffice_available": true
}
```

---

### 3. Web界面

访问Web用户界面。

**端点**: `GET /`

**响应**: HTML页面

---

## 使用示例

### cURL 示例

#### 转换单个文件为Office格式

```bash
curl -X POST "http://localhost:7788/api/convert" \
  -F "files=@document.doc" \
  -F "target_format=office" \
  --output "document.docx"
```

#### 转换多个文件为PDF格式

```bash
curl -X POST "http://localhost:7788/api/convert" \
  -F "files=@document1.doc" \
  -F "files=@spreadsheet.xls" \
  -F "files=@presentation.ppt" \
  -F "target_format=pdf" \
  --output "converted_files.zip"
```

#### 健康检查

```bash
curl -X GET "http://localhost:7788/api/health"
```

---

### Python 示例

#### 使用 requests 库

```python
import requests

# 单文件转换
url = "http://localhost:7788/api/convert"

files = {
    'files': open('document.doc', 'rb')
}
data = {
    'target_format': 'office'
}

response = requests.post(url, files=files, data=data)

if response.status_code == 200:
    with open('document.docx', 'wb') as f:
        f.write(response.content)
    print("转换成功！")
else:
    print(f"转换失败: {response.json()}")
```

#### 多文件转换

```python
import requests

url = "http://localhost:7788/api/convert"

files = [
    ('files', open('document1.doc', 'rb')),
    ('files', open('document2.xls', 'rb')),
    ('files', open('document3.ppt', 'rb'))
]
data = {
    'target_format': 'pdf'
}

response = requests.post(url, files=files, data=data)

if response.status_code == 200:
    with open('converted_files.zip', 'wb') as f:
        f.write(response.content)
    print("转换成功！ZIP文件已保存")
else:
    print(f"转换失败: {response.json()}")
```

#### 健康检查

```python
import requests

response = requests.get("http://localhost:7788/api/health")
print(response.json())
```

---

### JavaScript 示例

#### 使用 Fetch API

```javascript
// 单文件转换
async function convertFile(file, targetFormat) {
    const formData = new FormData();
    formData.append('files', file);
    formData.append('target_format', targetFormat);

    try {
        const response = await fetch('http://localhost:7788/api/convert', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail);
        }

        const blob = await response.blob();
        
        // 下载文件
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'converted_file';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        console.log('转换成功！');
    } catch (error) {
        console.error('转换失败:', error.message);
    }
}

// 使用示例
const fileInput = document.getElementById('file-input');
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    convertFile(file, 'office');
});
```

---

## 错误处理

API返回的错误响应格式：

```json
{
  "detail": "错误描述信息"
}
```

**常见错误**:

- `未选择文件`: 请求中没有包含文件
- `目标格式必须是 'office' 或 'pdf'`: target_format参数值不正确
- `转换失败: <详细信息>`: 文件转换过程中发生错误

---

## 文件存储

转换后的文件存储在以下位置：

- **单个转换文件**: `downloads/files/`
- **ZIP压缩包**: `downloads/zips/`
- **上传临时文件**: `uploads/` (转换后自动删除)

---

## 注意事项

1. **支持的文件格式**:
   - Office格式转换：`.doc`, `.xls`, `.ppt`
   - PDF转换：`.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`

2. **LibreOffice依赖**:
   - 本API依赖LibreOffice进行文件转换
   - 确保LibreOffice已正确安装并配置路径

3. **文件大小限制**:
   - 根据服务器配置，可能存在文件大小限制
   - 建议单个文件不超过100MB

4. **并发处理**:
   - API支持并发请求
   - 对于大量文件转换，建议分批处理

---

## 启动服务

### 安装依赖

```bash
pip install -r requirements.txt
```

### 启动服务器

```bash
python main.py
```

或使用uvicorn：

```bash
uvicorn main:app --host 0.0.0.0 --port 7788 --reload
```

服务启动后，访问：
- Web界面: http://localhost:7788
- Swagger文档: http://localhost:7788/docs
- ReDoc文档: http://localhost:7788/redoc

---

## 技术栈

- **Web框架**: FastAPI 0.109.0
- **文档转换**: LibreOffice (headless mode)
- **Python依赖**:
  - uvicorn (ASGI服务器)
  - python-multipart (文件上传)
  - openpyxl (Excel处理)
  - python-docx (Word处理)

---

## 版本历史

### v2.0.0 (当前版本)
- 迁移到FastAPI框架
- 简化转换选项（office/pdf）
- 自动识别文件格式
- 单文件直接返回，多文件返回ZIP
- 分离文件和ZIP存储目录
- 提供交互式API文档

### v1.0.0
- 基于Flask的初始版本
- 支持doc→docx, xls→xlsx转换
