# 📄 Nutrient DWS API Integration Guide

---

## 🔐 Authentication

### API Keys Required
- **Processor API Key**: For `/build` endpoint operations
- **Viewer API Key**: For `/viewer` endpoint operations
- **API Base URL**: `https://api.nutrient.io` (read-only)

### Header Format
```javascript
const headers = {
  'Authorization': `Bearer ${processorApiKey}`,
  'Content-Type': 'application/json'
};
```

### Authentication Modal Implementation
- **Fields**: Processor API Key, Viewer API Key, API Base URL (read-only)
- **Storage**: localStorage with key `nutrientAuthCredentials`
- **Validation**: Credentials validated during actual API calls

---

## 🏗️ Build API (`/build`)

### Convert DOCX to PDF

#### cURL Example (Tested & Verified)
```bash
curl -X POST https://api.nutrient.io/build \
  -H "Authorization: Bearer pdf_live_VZpbfS8lRYvhKIcA8GWgzqxvl861eKQ54QRVC4ti5Wl" \
  -F file=@input.docx \
  -F instructions='{
      "parts": [
        {
          "file": "file"
        }
      ],
      "output": {
        "type": "pdf",
        "quality": "medium"
      }
    }' \
  -o result.pdf
```

#### Test Results ✅
- **HTTP Status**: 200 OK
- **Response Size**: 36,536 bytes
- **Content-Type**: application/pdf
- **File Created**: Valid PDF document, version 1.5, 1 page

#### JavaScript Implementation
```javascript
async function convertWordToPdf(documentContent: string, quality: string): Promise<string> {
    const instructions = {
        parts: [
            {
                file: "file"
            }
        ],
        output: {
            type: "pdf",
            quality: quality
        }
    };
    
    const formData = new FormData();
    const documentBlob = new Blob([documentContent], { 
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
    });
    formData.append('file', documentBlob, 'input.docx');
    formData.append('instructions', JSON.stringify(instructions));
    
    const headers = getAuthHeaders();
    delete headers['Content-Type']; // Let browser set boundary
    
    const response = await fetch(`${authCredentials.apiBaseUrl}/build`, {
        method: 'POST',
        headers: headers,
        body: formData
    });
    
    if (response.ok) {
        const pdfBlob = await response.blob();
        // Store blob for download
        (window as any).currentPdfBlob = pdfBlob;
        
        // Convert to base64 for compatibility
        const arrayBuffer = await pdfBlob.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const base64 = btoa(String.fromCharCode.apply(null, Array.from(uint8Array)));
        
        return base64;
    } else {
        throw new Error(`API request failed: ${response.status} ${response.statusText}`);
    }
}
```

### Convert PDF to DOCX

#### cURL Example
```bash
curl -X POST https://api.nutrient.io/build \
  -H "Authorization: Bearer pdf_live_VZpbfS8lRYvhKIcA8GWgzqxvl861eKQ54QRVC4ti5Wl" \
  -F file=@input.pdf \
  -F instructions='{
      "parts": [
        {
          "file": "file"
        }
      ],
      "ocr": true,
      "output": {
        "type": "docx"
      }
    }' \
  -o result.docx
```

#### JavaScript Implementation
```javascript
async function convertPdfToDocx(file: File, ocrEnabled: boolean, language: string): Promise<Blob> {
    const instructions = {
        parts: [
            {
                file: "file"
            }
        ],
        ocr: ocrEnabled,
        language: language,
        output: {
            type: "docx"
        }
    };
    
    const formData = new FormData();
    formData.append('file', file);
    formData.append('instructions', JSON.stringify(instructions));
    
    const headers = getAuthHeaders();
    delete headers['Content-Type'];
    
    const response = await fetch(`${authCredentials.apiBaseUrl}/build`, {
        method: 'POST',
        headers: headers,
        body: formData
    });
    
    if (response.ok) {
        return await response.blob();
    } else {
        throw new Error(`API request failed: ${response.status} ${response.statusText}`);
    }
}
```

---

## 👁️ Viewer API (`/viewer`)

### Upload PDF for Viewing

#### cURL Example
```bash
curl -X POST https://api.nutrient.io/viewer/documents \
  -H "Authorization: Bearer <VIEWER_API_KEY>" \
  -H "Content-Type: application/pdf" \
  -F file=@document.pdf
```

#### Expected Response
```json
{
  "document_id": "abc123"
}
```

### Embed Viewer in HTML

#### iframe Implementation
```html
<iframe 
  src="https://viewer.nutrient.io/viewer/embed?documentId=abc123" 
  width="100%" 
  height="600" 
  frameborder="0" 
  allowfullscreen>
</iframe>
```

#### Optional URL Parameters
- `readOnly=true`
- `theme=dark`
- `initialPage=2`

---

## 🛠️ Implementation Notes

### FormData Structure
- **File field**: Always use `file` (not `document`)
- **Instructions**: JSON string with proper structure
- **Content-Type**: Let browser set multipart boundary automatically

### Error Handling
```javascript
// Common HTTP Status Codes
401: "Unauthorized - Check your API key"
403: "Forbidden - Insufficient permissions"
404: "Not found - Invalid endpoint or file"
500: "Server error - Try again later"
```

### Debug Information
The Word add-in includes comprehensive debug panels showing:
- **Request Details**: URL, method, headers, payload
- **Response Details**: Status, headers, body
- **File Information**: Size, type, processing time

### File Type MIME Types
```javascript
// Word documents
'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

// PDF files
'application/pdf'

// Plain text (fallback)
'text/plain'
```

---

## 🔧 Testing Commands

### Test DOCX to PDF Conversion
```bash
curl -X POST https://api.nutrient.io/build \
  -H "Authorization: Bearer YOUR_PROCESSOR_API_KEY" \
  -F file=@assets/Invoice.docx \
  -F instructions='{"parts":[{"file":"file"}],"output":{"type":"pdf","quality":"medium"}}' \
  -o test_output.pdf \
  --write-out "HTTP Status: %{http_code}\nResponse Size: %{size_download} bytes\nContent-Type: %{content_type}\n"
```

### Test PDF to DOCX Conversion
```bash
curl -X POST https://api.nutrient.io/build \
  -H "Authorization: Bearer YOUR_PROCESSOR_API_KEY" \
  -F file=@input.pdf \
  -F instructions='{"parts":[{"file":"file"}],"ocr":true,"output":{"type":"docx"}}' \
  -o test_output.docx
```

---

## 📚 References

- **Processor API Docs**: https://www.nutrient.io/api/reference/public/
- **Viewer API Docs**: https://www.nutrient.io/api/reference/viewer/public/
- **DOCX to PDF API**: https://www.nutrient.io/api/docx-to-pdf-api/
- **Processor MCP Server**: https://github.com/PSPDFKit/nutrient-dws-mcp-server

---

**Last Updated**: 2025-08-04  
**Tested**: ✅ DOCX to PDF conversion verified with curl  
**Status**: Production ready 