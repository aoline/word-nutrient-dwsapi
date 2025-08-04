#!/usr/bin/env node

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const logFile = path.join(__dirname, 'api-debug.log');

console.log('🚀 Starting Nutrient PDF Tools with logging...\n');

// Start the logging server
console.log('📝 Starting logging server on port 3001...');
const loggingServer = spawn('node', ['server.js'], {
    stdio: 'inherit',
    cwd: __dirname
});

// Wait a moment for logging server to start
setTimeout(() => {
    // Start the webpack dev server
    console.log('🌐 Starting webpack dev server on port 3000...');
    const webpackServer = spawn('npm', ['run', 'build:dev'], {
        stdio: 'inherit',
        cwd: __dirname
    });
    
    // Wait for webpack to build, then start the add-in
    setTimeout(() => {
        console.log('🔧 Starting Office add-in...');
        const addinServer = spawn('npm', ['start'], {
            stdio: 'inherit',
            cwd: __dirname
        });
        
        // Monitor log file for errors
        console.log('👀 Monitoring log file for errors...\n');
        monitorLogFile();
        
        // Handle process termination
        process.on('SIGINT', () => {
            console.log('\n🛑 Shutting down servers...');
            loggingServer.kill();
            webpackServer.kill();
            addinServer.kill();
            process.exit(0);
        });
    }, 5000);
}, 2000);

function monitorLogFile() {
    let lastSize = 0;
    
    setInterval(() => {
        try {
            if (fs.existsSync(logFile)) {
                const stats = fs.statSync(logFile);
                if (stats.size > lastSize) {
                    // New content added to log file
                    const content = fs.readFileSync(logFile, 'utf8');
                    const newContent = content.substring(lastSize);
                    
                    // Check for errors
                    if (newContent.includes('ERROR') || newContent.includes('Error') || newContent.includes('error')) {
                        console.log('\n🚨 ERROR DETECTED IN LOGS:');
                        console.log('='.repeat(50));
                        console.log(newContent);
                        console.log('='.repeat(50));
                        console.log('📁 Full log file: api-debug.log\n');
                    }
                    
                    lastSize = stats.size;
                }
            }
        } catch (error) {
            console.error('Error monitoring log file:', error);
        }
    }, 1000);
}

console.log('📋 Instructions:');
console.log('1. Upload a PDF file in the Word add-in');
console.log('2. Click "Convert & Insert"');
console.log('3. Watch for errors in the console above');
console.log('4. Check api-debug.log for detailed logs\n'); 