#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Create a simple HTML page that extracts localStorage logs
const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Log Extractor</title>
</head>
<body>
    <h1>API Debug Logs</h1>
    <div id="logs"></div>
    <script>
        const logs = localStorage.getItem('api-debug-logs') || 'No logs found';
        document.getElementById('logs').innerHTML = '<pre>' + logs + '</pre>';
        
        // Also log to console for easy copying
        console.log('=== API DEBUG LOGS ===');
        console.log(logs);
    </script>
</body>
</html>
`;

const logFile = path.join(__dirname, 'api-debug.html');
fs.writeFileSync(logFile, htmlContent);

console.log(`Log extractor created at: ${logFile}`);
console.log('Open this file in a browser to see the logs from the add-in'); 