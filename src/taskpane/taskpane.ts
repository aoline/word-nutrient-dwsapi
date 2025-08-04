/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// API Configuration
const NUTRIENT_API_BASE = 'https://api.nutrient.io';
const PROCESSOR_API_KEY = 'pdf_live_VZpbfS8lRYvhKIcA8GWgzqxvl861eKQ54QRVC4ti5Wl';

// Interface for file data
interface FileData {
    file: File;
    name: string;
    size: string;
}

// Global state
let selectedFile: FileData | null = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        initializeApp();
    }
});

function initializeApp() {
    // Initialize drag and drop functionality
    initializeDragAndDrop();
    
    // Initialize file input
    initializeFileInput();
    
    // Initialize convert button
    initializeConvertButton();
    
    // Initialize options
    initializeOptions();
}

function initializeDragAndDrop() {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    const fileInput = document.getElementById('file-input') as HTMLInputElement;

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop zone when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dragDropZone.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    dragDropZone.addEventListener('drop', handleDrop, false);
    
    // Handle click to browse
    dragDropZone.addEventListener('click', () => fileInput.click());
}

function initializeFileInput() {
    const fileInput = document.getElementById('file-input') as HTMLInputElement;
    fileInput.addEventListener('change', handleFileSelect);
}

function initializeConvertButton() {
    const convertBtn = document.getElementById('convert-btn') as HTMLButtonElement;
    convertBtn.addEventListener('click', handleConvert);
}

function initializeOptions() {
    // OCR toggle and language select are already handled by HTML
    // Additional initialization can be added here if needed
}

function preventDefaults(e: Event) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e: Event) {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    dragDropZone.classList.add('dragover');
}

function unhighlight(e: Event) {
    const dragDropZone = document.getElementById('drag-drop-zone') as HTMLElement;
    dragDropZone.classList.remove('dragover');
}

function handleDrop(e: DragEvent) {
    const dt = e.dataTransfer;
    const files = dt?.files;
    
    if (files && files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e: Event) {
    const target = e.target as HTMLInputElement;
    if (target.files && target.files.length > 0) {
        handleFile(target.files[0]);
    }
}

function handleFile(file: File) {
    // Validate file type
    if (file.type !== 'application/pdf') {
        showStatus('Please select a PDF file.', 'error');
        return;
    }

    // Validate file size (max 50MB)
    if (file.size > 50 * 1024 * 1024) {
        showStatus('File size must be less than 50MB.', 'error');
        return;
    }

    // Store file data
    selectedFile = {
        file: file,
        name: file.name,
        size: formatFileSize(file.size)
    };

    // Update UI
    updateFileInfo();
    enableConvertButton();
    showStatus('PDF file selected successfully.', 'success');
}

function updateFileInfo() {
    const fileInfo = document.getElementById('file-info') as HTMLElement;
    const fileName = document.getElementById('file-name') as HTMLElement;
    const fileSize = document.getElementById('file-size') as HTMLElement;

    if (selectedFile) {
        fileName.textContent = `File: ${selectedFile.name}`;
        fileSize.textContent = `Size: ${selectedFile.size}`;
        fileInfo.style.display = 'block';
    } else {
        fileInfo.style.display = 'none';
    }
}

function enableConvertButton() {
    const convertBtn = document.getElementById('convert-btn') as HTMLButtonElement;
    convertBtn.disabled = !selectedFile;
}

function formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function handleConvert() {
    if (!selectedFile) {
        showStatus('Please select a PDF file first.', 'error');
        return;
    }

    try {
        showProgress(true);
        updateProgress(10, 'Preparing file for conversion...');

        // Get options
        const ocrEnabled = (document.getElementById('ocr-toggle') as HTMLInputElement).checked;
        const language = (document.getElementById('language-select') as HTMLSelectElement).value;

        updateProgress(30, 'Uploading to Nutrient.io...');

        // Convert PDF to DOCX using Nutrient.io API
        const docxBlob = await convertPdfToDocx(selectedFile.file, ocrEnabled, language);

        updateProgress(70, 'Inserting into Word document...');

        // Insert the converted document into Word
        await insertDocxIntoWord(docxBlob);

        updateProgress(100, 'Conversion completed!');
        
        setTimeout(() => {
            showProgress(false);
            showStatus('PDF successfully converted and inserted into Word document!', 'success');
            resetFileSelection();
        }, 1000);

    } catch (error) {
        showProgress(false);
        console.error('Conversion error:', error);
        showStatus(`Conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`, 'error');
    }
}

async function convertPdfToDocx(file: File, ocrEnabled: boolean, language: string): Promise<Blob> {
    const formData = new FormData();
    formData.append('file', file);

    // Prepare the API payload according to Nutrient.io Build API specification
    const payload = {
        parts: [{ file: 'document' }],
        ocr: ocrEnabled,
        output: { type: 'docx' }
    };

    // Add language if specified and not auto-detect
    if (language !== 'auto') {
        payload.ocr_language = language;
    }

    formData.append('instructions', JSON.stringify(payload));

    const response = await fetch(`${NUTRIENT_API_BASE}/build`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${PROCESSOR_API_KEY}`
        },
        body: formData
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
    }

    return await response.blob();
}

async function insertDocxIntoWord(docxBlob: Blob) {
    return Word.run(async (context) => {
        // Convert blob to base64
        const arrayBuffer = await docxBlob.arrayBuffer();
        const base64 = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)));

        // Insert the document at the current cursor position
        const range = context.document.getSelection();
        range.insertFileFromBase64(base64, Word.InsertLocation.replace);

        await context.sync();
    });
}

function showProgress(show: boolean) {
    const progressSection = document.getElementById('progress-section') as HTMLElement;
    progressSection.style.display = show ? 'block' : 'none';
}

function updateProgress(percentage: number, text: string) {
    const progressBar = document.getElementById('progress-bar') as HTMLElement;
    const progressText = document.getElementById('progress-text') as HTMLElement;

    progressBar.style.width = `${percentage}%`;
    progressText.textContent = text;
}

function showStatus(message: string, type: 'success' | 'error' | 'info') {
    const statusMessages = document.getElementById('status-messages') as HTMLElement;
    
    // Remove existing messages
    const existingMessages = statusMessages.querySelectorAll('.status-message');
    existingMessages.forEach(msg => msg.remove());

    // Create new message
    const messageElement = document.createElement('div');
    messageElement.className = `status-message ${type}`;
    messageElement.textContent = message;
    
    statusMessages.appendChild(messageElement);

    // Auto-remove success/info messages after 5 seconds
    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            if (messageElement.parentNode) {
                messageElement.remove();
            }
        }, 5000);
    }
}

function resetFileSelection() {
    selectedFile = null;
    updateFileInfo();
    enableConvertButton();
    
    // Reset file input
    const fileInput = document.getElementById('file-input') as HTMLInputElement;
    fileInput.value = '';
}
