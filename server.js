const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const https = require('https');

const app = express();
const PORT = process.env.PORT || 8080;

// Enable CORS for all routes
app.use(cors({
    origin: ['https://localhost:3000', 'https://localhost:3001', 'https://localhost:3002', 'https://localhost:8080'],
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

// Parse JSON bodies
app.use(express.json({ limit: '50mb' }));

// Parse URL-encoded bodies
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Serve static files from assets directory
app.use('/assets', express.static(path.join(__dirname, 'assets')));

// Mock API endpoints for testing
app.post('/api/nutrient/processor/documents', (req, res) => {
    console.log('📝 Mock API: Processing document request');
    console.log('📋 Request body:', req.body);
    
    // Simulate processing delay
    setTimeout(() => {
        res.json({
            document_id: 'mock-doc-' + Date.now(),
            status: 'processed',
            message: 'Document processed successfully (mock response)',
            timestamp: new Date().toISOString()
        });
    }, 1000);
});

app.post('/api/nutrient/viewer/documents', (req, res) => {
    console.log('👁️ Mock API: Viewer document request');
    console.log('📋 Request body:', req.body);
    
    // Simulate processing delay
    setTimeout(() => {
        res.json({
            document_id: 'mock-viewer-' + Date.now(),
            status: 'uploaded',
            message: 'Document uploaded to viewer successfully (mock response)',
            timestamp: new Date().toISOString()
        });
    }, 1000);
});

// Mock /build endpoint for PDF conversion
app.post('/api/nutrient/build', (req, res) => {
    console.log('🏗️ Mock API: Build/PDF conversion request');
    console.log('📋 Request body:', req.body);
    
    // Simulate processing delay
    setTimeout(() => {
        // Return the actual file blob instead of JSON
        const mockPdfPath = path.join(__dirname, 'assets', 'Invoice.docx');
        
        if (fs.existsSync(mockPdfPath)) {
            // Set proper headers for file download
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', `attachment; filename="converted-document.pdf"`);
            res.setHeader('Cache-Control', 'no-cache');
            res.setHeader('Pragma', 'no-cache');
            
            // Stream the file
            const fileStream = fs.createReadStream(mockPdfPath);
            fileStream.on('error', (error) => {
                console.error('❌ Error streaming file:', error);
                res.status(500).json({ error: 'File streaming error' });
            });
            
            fileStream.pipe(res);
            
            console.log('✅ Build endpoint returning file blob:', mockPdfPath);
        } else {
            console.error('❌ File not found:', mockPdfPath);
            res.status(404).json({
                error: 'Document not found',
                message: 'Mock document file not available',
                path: mockPdfPath
            });
        }
    }, 2000);
});

app.get('/api/nutrient/processor/documents/:id', (req, res) => {
    console.log('📄 Mock API: Get document status');
    console.log('📋 Document ID:', req.params.id);
    
    res.json({
        document_id: req.params.id,
        status: 'completed',
        download_url: `https://localhost:8080/api/nutrient/download/${req.params.id}`,
        message: 'Document processing completed (mock response)',
        timestamp: new Date().toISOString()
    });
});

app.get('/api/nutrient/download/:id', (req, res) => {
    console.log('⬇️ Mock API: Download document');
    console.log('📋 Document ID:', req.params.id);
    
    // Return a mock PDF file (actually the DOCX file for testing)
    const mockPdfPath = path.join(__dirname, 'assets', 'Invoice.docx');
    
    if (fs.existsSync(mockPdfPath)) {
        // Set proper headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="processed-${req.params.id}.docx"`);
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Pragma', 'no-cache');
        
        // Stream the file
        const fileStream = fs.createReadStream(mockPdfPath);
        fileStream.on('error', (error) => {
            console.error('❌ Error streaming file:', error);
            res.status(500).json({ error: 'File streaming error' });
        });
        
        fileStream.pipe(res);
        
        console.log('✅ File download started:', mockPdfPath);
    } else {
        console.error('❌ File not found:', mockPdfPath);
        res.status(404).json({
            error: 'Document not found',
            message: 'Mock document file not available',
            path: mockPdfPath
        });
    }
});

// Test download endpoint for debugging
app.get('/test-download', (req, res) => {
    console.log('🧪 Test download endpoint called');
    
    const testFilePath = path.join(__dirname, 'assets', 'Invoice.docx');
    
    if (fs.existsSync(testFilePath)) {
        const stats = fs.statSync(testFilePath);
        console.log('📊 File stats:', {
            path: testFilePath,
            size: stats.size,
            exists: true
        });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="test-download.docx"');
        res.setHeader('Content-Length', stats.size);
        
        const fileStream = fs.createReadStream(testFilePath);
        fileStream.pipe(res);
        
        console.log('✅ Test download started');
    } else {
        console.error('❌ Test file not found:', testFilePath);
        res.status(404).json({
            error: 'Test file not found',
            path: testFilePath
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({
        status: 'healthy',
        timestamp: new Date().toISOString(),
        proxy: 'enabled',
        cors: 'enabled',
        mock_api: 'enabled'
    });
});

// Mock health endpoint for Nutrient API
app.get('/api/nutrient/health', (req, res) => {
    res.json({
        status: 'healthy',
        service: 'nutrient-api-proxy',
        timestamp: new Date().toISOString(),
        message: 'Nutrient API proxy is working correctly (mock response)'
    });
});

// Default route - serve test viewer
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'test-viewer.html'));
});

// Serve test-viewer.html
app.get('/test-viewer.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'test-viewer.html'));
});

// Serve test_cors.html
app.get('/test_cors.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'test_cors.html'));
});

// Serve test-download.html
app.get('/test-download.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'test-download.html'));
});

// Download page with direct links to files
app.get('/downloads', (req, res) => {
    const html = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Downloads - Nutrient API Test</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 40px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }
        .header p {
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 1.1em;
        }
        .content {
            padding: 40px;
        }
        .download-section {
            margin-bottom: 30px;
            padding: 25px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            background: #fafafa;
        }
        .download-section h2 {
            margin: 0 0 15px 0;
            color: #333;
            font-size: 1.4em;
        }
        .download-section p {
            margin: 0 0 20px 0;
            color: #666;
            line-height: 1.6;
        }
        .download-link {
            display: inline-block;
            background: #667eea;
            color: white;
            text-decoration: none;
            padding: 12px 24px;
            border-radius: 6px;
            font-weight: 500;
            transition: background-color 0.3s;
        }
        .download-link:hover {
            background: #5a6fd8;
        }
        .file-info {
            background: #e8f5e8;
            border: 1px solid #28a745;
            border-radius: 4px;
            padding: 15px;
            margin: 15px 0;
        }
        .file-info strong {
            color: #155724;
        }
        .status {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.9em;
            font-weight: 500;
        }
        .status.available {
            background: #d4edda;
            color: #155724;
        }
        .status.mock {
            background: #fff3cd;
            color: #856404;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📁 File Downloads</h1>
            <p>Direct download links for testing the Nutrient API</p>
        </div>
        
        <div class="content">
            <div class="download-section">
                <h2>📄 Original DOCX File</h2>
                <p>This is the original Invoice.docx file that the add-in uses for testing. This file contains the actual content that gets processed by the API.</p>
                <div class="file-info">
                    <strong>File:</strong> Invoice.docx<br>
                    <strong>Type:</strong> Microsoft Word Document<br>
                    <strong>Source:</strong> assets/Invoice.docx<br>
                    <strong>Status:</strong> <span class="status available">Available</span>
                </div>
                <a href="/assets/Invoice.docx" class="download-link" download="Invoice.docx">
                    ⬇️ Download Invoice.docx
                </a>
            </div>
            
            <div class="download-section">
                <h2>📊 Mock PDF File</h2>
                <p>This is the mock PDF file that gets returned when you use the "Convert to PDF" tool. It's actually the same DOCX file but served with PDF headers for testing.</p>
                <div class="file-info">
                    <strong>File:</strong> processed-document.pdf<br>
                    <strong>Type:</strong> PDF Document (Mock)<br>
                    <strong>Source:</strong> assets/Invoice.docx (converted)<br>
                    <strong>Status:</strong> <span class="status mock">Mock Response</span>
                </div>
                <a href="/api/nutrient/download/mock-test" class="download-link" download="processed-document.pdf">
                    ⬇️ Download Mock PDF
                </a>
            </div>
            
            <div class="download-section">
                <h2>🔗 API Endpoints</h2>
                <p>Test the API endpoints directly:</p>
                <ul>
                    <li><strong>Health Check:</strong> <a href="/health" target="_blank">/health</a></li>
                    <li><strong>API Status:</strong> <a href="/api/nutrient/status" target="_blank">/api/nutrient/status</a></li>
                    <li><strong>CORS Test:</strong> <a href="/test_cors.html" target="_blank">/test_cors.html</a></li>
                    <li><strong>Test Viewer:</strong> <a href="/test-viewer.html" target="_blank">/test-viewer.html</a></li>
                    <li><strong>Download Test:</strong> <a href="/test-download.html" target="_blank">/test-download.html</a></li>
                </ul>
            </div>
        </div>
    </div>
</body>
</html>`;
    res.send(html);
});

// Serve auth_test.html
app.get('/auth_test.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'assets/auth_test.html'));
});

// API endpoint to test CORS
app.get('/api/test', (req, res) => {
    res.json({
        message: 'CORS test successful',
        timestamp: new Date().toISOString(),
        headers: req.headers
    });
});

// Mock API status endpoint
app.get('/api/nutrient/status', (req, res) => {
    res.json({
        service: 'nutrient-api-proxy',
        version: '1.0.0',
        status: 'operational',
        endpoints: {
            'POST /api/nutrient/processor/documents': 'Process documents',
            'POST /api/nutrient/viewer/documents': 'Upload to viewer',
            'POST /api/nutrient/build': 'Convert to PDF',
            'GET /api/nutrient/processor/documents/:id': 'Get document status',
            'GET /api/nutrient/download/:id': 'Download processed document'
        },
        timestamp: new Date().toISOString()
    });
});

// Start HTTPS server
const httpsOptions = {
    key: fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs', 'localhost.key')),
    cert: fs.readFileSync(path.join(process.env.HOME || process.env.USERPROFILE, '.office-addin-dev-certs', 'localhost.crt'))
};

https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`🚀 Server running on https://localhost:${PORT}`);
    console.log(`🔗 Test viewer: https://localhost:${PORT}/test-viewer.html`);
    console.log(`🔗 CORS test: https://localhost:${PORT}/test_cors.html`);
    console.log(`🔗 Health check: https://localhost:${PORT}/health`);
    console.log(`🔗 API status: https://localhost:${PORT}/api/nutrient/status`);
    console.log(`📁 Assets served from: https://localhost:${PORT}/assets/`);
    console.log(`📥 Downloads page: https://localhost:${PORT}/downloads`);
    console.log(`✅ Mock API endpoints enabled for testing`);
}); 