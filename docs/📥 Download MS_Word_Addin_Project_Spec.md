
# 📄 Project Specification: MS Word Add-in for Nutrient.io DWS APIs

---

## Project Title
**MS Word Add-in: PDF Conversion, Redaction, OCR & Accessibility Tools via Nutrient.io DWS APIs**

**Owner:** Product Management – Nutrient.io  
**Target Audience:** Engineers responsible for Office Add-in development and API integration

---

## Objective

To build a production-ready Microsoft Word Add-in that integrates with Nutrient.io's **Build API** and **Viewer API**, delivering:

- PDF conversion
- Redaction
- Accessibility compliance (PDF/UA)
- OCR processing
- Collaboration features (SP/Teams/OneDrive)

> ✅ Initial Proof of Concept (PoC): **Import from PDF to DOCX**

---

## Primary Sidebar UX Structure

1. 📥 **Import from PDF to DOCX** *(PoC focus)*
2. 🛡️ **Redact Document**
3. 📤 **Export to PDF/A or PDF/UA**
4. 📡 **Send PDF to SharePoint / Teams / OneDrive**

---

## PoC Feature: Import from PDF to DOCX

### User Tools
- 🗂️ Drag/drop or file picker
- ✅ OCR toggle
- 🌐 Optional language selector
- ▶️ "Convert & Insert" button

### API Call
```
POST https://api.nutrient.io/build
```

### Payload Example
```json
{
  "parts": [{ "file": "document" }],
  "ocr": true,
  "output": { "type": "docx" }
}
```

---

## PoC Engineering Tasks

### 🧩 Setup
- [ ] Scaffold Word Add-in (OfficeJS + React)
- [ ] Sidebar panel for “Import from PDF”

### 📂 File Handling
- [ ] PDF drag/drop UI
- [ ] Validate file type/size

### 🔌 API Integration
- [ ] FormData with file + instructions
- [ ] Auth headers
- [ ] Error handling

### 📥 Output Handling
- [ ] Receive converted `.docx`
- [ ] Insert via OfficeJS
- [ ] Toast notifications

### ✅ Testing
- [ ] OCR and text-based PDF tests
- [ ] Edge case handling

### 🚀 Demo
- [ ] Sample files
- [ ] Short demo recording

---

## Future Feature Modules (Post-PoC)

### 🛡️ Redact Document
- Redact terms
- Strip metadata
- Preview redactions

### 📤 Export to PDF/A or PDF/UA
- Select format
- Preview with Viewer
- Advanced export options

### 📡 Send to SP/Teams/OneDrive
- OAuth login
- Destination picker
- Comment + attach source file

---

## API Integration Notes

- Endpoint: `POST https://api.nutrient.io/build`
- Auth: `Authorization: Bearer <API_KEY>`
- Processor Docs: https://www.nutrient.io/api/reference/public/
- Processor MCP Server https://github.com/PSPDFKit/nutrient-dws-mcp-server
- Viewer Docs: https://www.nutrient.io/api/reference/viewer/public/ & 

**‼️ RULE:** Never assume — always verify endpoints with official docs.

---

## Tech Stack

- OfficeJS
- React + TypeScript
- Nutrient.io Build API
- Microsoft Graph API (future)

---

## Handoff Checklist

- [ ] API key
- [ ] Sample PDFs (scanned + searchable)
- [ ] `.env` config example
- [ ] OAuth App Reg (for SharePoint)
- [ ] Access to Viewer embed config
- [ ] Link to [https://www.nutrient.io/api/docx-to-pdf-api/](https://www.nutrient.io/api/docx-to-pdf-api/)

---

**Prepared by:** Product Team @ Nutrient.io  
**Date:** 2025-08-04
