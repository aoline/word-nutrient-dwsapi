# Nutrient PDF Tools - Microsoft Word Add-in

A powerful Microsoft Word add-in that provides advanced PDF processing capabilities using the Nutrient DWS API. Convert PDFs to DOCX, create accessible PDFs, and more with enterprise-grade processing.

## 🚀 Features

### Core Functionality
- **PDF to DOCX Conversion** - Convert PDF files to editable Word documents with OCR support
- **Word to PDF Conversion** - Convert Word documents to high-quality PDFs
- **PDF/A & PDF/UA Export** - Create accessible and compliant PDFs
- **Document Redaction** - Remove sensitive information and metadata (Coming Soon)
- **SharePoint/Teams Integration** - Share documents to Microsoft 365 (Coming Soon)

### Advanced Features
- **OCR Processing** - Extract text from scanned documents with multiple language support
- **Accessibility Compliance** - Generate PDF/UA compliant documents
- **High-Quality Output** - Multiple quality settings for different use cases
- **Secure Processing** - Enterprise-grade security with API key authentication

## 📋 Prerequisites

- Microsoft Word (Desktop or Online)
- Node.js 16+ and npm
- Nutrient DWS API credentials

## 🛠️ Installation

### Development Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/nutrient-io/word-nutrient-dwsapi.git
   cd word-nutrient-dwsapi
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Configure API credentials**
   - Open the add-in in Word
   - Go to Settings → API Settings
   - Enter your Nutrient DWS API credentials:
     - Processor API Key (for PDF processing)
     - Viewer API Key (for PDF preview)

4. **Start development server**
   ```bash
   npm run dev-server
   ```

5. **Load the add-in in Word**
   ```bash
   npm start
   ```

### Production Build

1. **Build the add-in**
   ```bash
   npm run build
   ```

2. **Validate the manifest**
   ```bash
   npm run validate
   ```

## 🔧 Configuration

### API Credentials

The add-in requires two API keys from Nutrient.io:

1. **Processor API Key** - Used for PDF processing, conversion, and OCR
2. **Viewer API Key** - Used for PDF preview and viewing

### Environment Variables

- `NUTRIENT_API_BASE` - API base URL (default: `https://api.nutrient.io`)
- `DEV_SERVER_PORT` - Development server port (default: `3000`)

## 📖 Usage

### Converting PDF to DOCX

1. Open the Nutrient PDF Tools add-in in Word
2. Click on "Import from PDF to DOCX"
3. Drag and drop a PDF file or click to browse
4. Configure OCR options (if needed)
5. Click "Convert & Insert"
6. The converted document will be inserted into your Word document

### Converting Word to PDF

1. Open the Nutrient PDF Tools add-in in Word
2. Click on "Convert to PDF"
3. Configure quality settings
4. Click "Convert to PDF"
5. Download the generated PDF

### Creating Accessible PDFs

1. Open the Nutrient PDF Tools add-in in Word
2. Click on "Export to PDF/A or PDF/UA"
3. Choose your preferred format (PDF/A or PDF/UA)
4. Configure accessibility options
5. Click "Export"
6. Preview and download the accessible PDF

## 🏗️ Project Structure

```
word-nutrient-dwsapi/
├── assets/                 # Static assets (icons, images)
├── dist/                   # Built files (generated)
├── docs/                   # Documentation
├── src/
│   ├── commands/          # Office commands
│   └── taskpane/          # Main add-in interface
├── manifest.xml           # Office add-in manifest
├── package.json           # Dependencies and scripts
├── webpack.config.js      # Build configuration
└── server.js              # Development server
```

## 🔌 API Integration

The add-in integrates with the Nutrient DWS API for:

- **Document Processing** - Convert between PDF and DOCX formats
- **OCR Processing** - Extract text from scanned documents
- **Accessibility** - Generate compliant PDFs
- **Document Viewing** - Preview PDFs in the browser

### API Endpoints Used

- `POST /build` - Document processing and conversion
- `POST /viewer/documents` - Document upload for viewing
- `GET /viewer/embed` - Document preview

## 🚀 Development

### Available Scripts

- `npm run dev-server` - Start development server
- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm start` - Start Office add-in debugging
- `npm run validate` - Validate manifest
- `npm run lint` - Run linting
- `npm run serve` - Start custom development server

### Development Workflow

1. Start the development server: `npm run dev-server`
2. Start Office debugging: `npm start`
3. Make changes to the code
4. The add-in will automatically reload in Word

### Debugging

- Use browser developer tools for debugging
- Check the browser console for detailed logs
- Use the debug section in the add-in for API request/response details

## 🔒 Security

- API keys are stored locally in browser storage
- All API requests use HTTPS
- No sensitive data is logged or transmitted unnecessarily
- CORS is properly configured for secure cross-origin requests

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- **Documentation**: [https://www.nutrient.io/docs](https://www.nutrient.io/docs)
- **Support**: [https://www.nutrient.io/support](https://www.nutrient.io/support)
- **Issues**: [GitHub Issues](https://github.com/nutrient-io/word-nutrient-dwsapi/issues)

## 🗺️ Roadmap

- [ ] Document redaction features
- [ ] SharePoint/Teams integration
- [ ] Batch processing
- [ ] Advanced OCR options
- [ ] Custom templates
- [ ] Multi-language support

## 📊 Version History

### v1.0.0 (Current)
- Initial release with core PDF processing features
- PDF to DOCX conversion with OCR
- Word to PDF conversion
- PDF/A and PDF/UA export
- Authentication system
- Modern UI/UX

---

**Built with ❤️ by [Nutrient.io](https://www.nutrient.io)** 