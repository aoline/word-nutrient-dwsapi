const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 3001;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('dist'));

// Log file path
const logFile = path.join(__dirname, 'api-debug.log');

// Function to write to log file
function writeToLog(message) {
    const timestamp = new Date().toISOString();
    const logEntry = `[${timestamp}] ${message}\n`;
    
    try {
        fs.appendFileSync(logFile, logEntry);
        console.log(`Logged: ${message}`);
    } catch (error) {
        console.error('Failed to write to log file:', error);
    }
}

// Logging endpoint
app.post('/api/log', (req, res) => {
    const { message } = req.body;
    writeToLog(message);
    res.json({ success: true });
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Start server
app.listen(PORT, () => {
    console.log(`Logging server running on http://localhost:${PORT}`);
    writeToLog('=== LOGGING SERVER STARTED ===');
});

// Handle graceful shutdown
process.on('SIGINT', () => {
    writeToLog('=== LOGGING SERVER STOPPED ===');
    process.exit(0);
}); 