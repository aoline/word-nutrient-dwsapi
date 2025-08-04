#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

const logFile = path.join(__dirname, 'api-debug.log');

console.log('👀 Monitoring api-debug.log for errors...\n');
console.log('Press Ctrl+C to stop monitoring\n');

let lastSize = 0;

function checkLogFile() {
    try {
        if (fs.existsSync(logFile)) {
            const stats = fs.statSync(logFile);
            
            if (stats.size > lastSize) {
                // New content added to log file
                const content = fs.readFileSync(logFile, 'utf8');
                const newContent = content.substring(lastSize);
                
                // Check for errors or important messages
                if (newContent.includes('ERROR') || 
                    newContent.includes('Error') || 
                    newContent.includes('error') ||
                    newContent.includes('failed') ||
                    newContent.includes('Failed') ||
                    newContent.includes('API Error') ||
                    newContent.includes('Network Error')) {
                    
                    console.log('\n🚨 ERROR DETECTED:');
                    console.log('='.repeat(60));
                    console.log(newContent);
                    console.log('='.repeat(60));
                    console.log('📁 Full log file: api-debug.log\n');
                } else if (newContent.includes('API REQUEST') || newContent.includes('API RESPONSE')) {
                    console.log('\n📡 API Activity:');
                    console.log('-'.repeat(40));
                    console.log(newContent);
                }
                
                lastSize = stats.size;
            }
        }
    } catch (error) {
        console.error('Error monitoring log file:', error);
    }
}

// Check every 500ms
setInterval(checkLogFile, 500);

// Initial check
checkLogFile(); 