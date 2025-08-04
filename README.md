# Nutrient PDF Tools - Microsoft Word Add-in

A Microsoft Word Add-in that integrates with Nutrient.io's Document Workflow Services (DWS) APIs to provide PDF conversion, redaction, OCR processing, and accessibility compliance features.

## 🚀 Features

### Current (PoC)
- **📥 Import from PDF to DOCX** - Convert PDF files to Word documents with OCR support
- Drag & drop file upload
- OCR processing toggle
- Language selection for OCR
- Real-time progress tracking
- Error handling and status messages

### Coming Soon
- 🛡️ **Document Redaction** - Redact sensitive information
- 📤 **Export to PDF/A or PDF/UA** - Accessibility-compliant PDF export
- 📡 **SharePoint/Teams Integration** - Send documents to Microsoft 365 services

## 🛠️ Technology Stack

- **Office.js** - Microsoft Office Add-in framework
- **TypeScript** - Type-safe JavaScript
- **React** - UI framework (planned)
- **Nutrient.io Build API** - PDF processing backend
- **Nutrient.io Viewer API** - Document preview (future)

## 📋 Prerequisites

- Node.js (v16 or higher)
- Microsoft Word (desktop or web)
- Nutrient.io API keys

## 🚀 Getting Started

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure API Keys

The add-in is pre-configured with the provided API keys:
- **Processor API Key**: `pdf_live_VZpbfS8lRYvhKIcA8GWgzqxvl861eKQ54QRVC4ti5Wl`
- **Viewer API Key**: `pdf_live_7FqeZD6pwwso0fj7nZacIerdIPvejYTy0AdDjePd90S`

### 3. Start Development Server

```bash
npm start
```

This will:
- Start the local web server on `https://localhost:3000`
- Open Word and sideload the add-in
- Enable hot reloading for development

### 4. Use the Add-in

1. Open Microsoft Word
2. Go to the **Home** tab
3. Click the **PDF Tools** button in the ribbon
4. The taskpane will open with the PDF import interface

## 📖 Usage

### Import PDF to DOCX

1. **Select a PDF file**:
   - Drag and drop a PDF file onto the upload area, or
   - Click the upload area to browse and select a file

2. **Configure options**:
   - ✅ Enable/disable OCR processing
   - 🌐 Select language for OCR (English, Spanish, French, German, or Auto-detect)

3. **Convert**:
   - Click "Convert & Insert" to process the PDF
   - The converted DOCX will be inserted at your cursor position in Word

## 🔧 Development

### Project Structure

```
word-nutrient-dwsapi/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main UI
│   │   ├── taskpane.css       # Styles
│   │   └── taskpane.ts        # TypeScript logic
│   └── commands/              # Ribbon commands (if needed)
├── assets/                    # Icons and images
├── manifest.xml              # Add-in manifest
└── package.json              # Dependencies and scripts
```

### Available Scripts

- `npm start` - Start development server and sideload add-in
- `npm run build` - Build for production
- `npm run dev-server` - Start dev server only
- `npm run validate` - Validate manifest
- `npm run sideload` - Sideload add-in to Word

### API Integration

The add-in integrates with Nutrient.io's Build API:

```typescript
// Example API call
const response = await fetch('https://api.nutrient.io/build', {
    method: 'POST',
    headers: {
        'Authorization': `Bearer ${PROCESSOR_API_KEY}`
    },
    body: formData
});
```

**API Documentation**: https://www.nutrient.io/api/reference/public/

## 🧪 Testing

### Test Files
- Use both scanned PDFs (images) and searchable PDFs (text)
- Test various file sizes (up to 50MB limit)
- Test different languages for OCR

### Manual Testing Checklist
- [ ] PDF file selection via drag & drop
- [ ] PDF file selection via file picker
- [ ] File validation (type and size)
- [ ] OCR processing with different languages
- [ ] Error handling for invalid files
- [ ] Progress tracking during conversion
- [ ] Document insertion into Word
- [ ] Status message display

## 🚀 Deployment

### Production Build

```bash
npm run build
```

### Sideloading to Production

1. Build the project
2. Host the files on a web server with HTTPS
3. Update the manifest.xml with production URLs
4. Sideload the manifest to Word

## 📚 API References

- **Office.js**: https://docs.microsoft.com/office/dev/add-ins/
- **Nutrient.io Build API**: https://www.nutrient.io/api/reference/public/
- **Nutrient.io Viewer API**: https://www.nutrient.io/api/reference/viewer/public/

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For support and questions:
- **Documentation**: https://www.nutrient.io/api/docx-to-pdf-api/
- **API Support**: Contact Nutrient.io support team
- **Add-in Issues**: Create an issue in this repository

---

**Built with ❤️ by the Nutrient.io team** 