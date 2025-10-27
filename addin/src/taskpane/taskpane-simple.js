/* global Office */

console.log('Simple Taskpane.js loaded');

// Global variables
let authToken = null;
let currentUser = null;

// API configuration
const API_BASE_HTTPS = 'https://localhost:8000/api';
const API_BASE_HTTP = 'http://localhost:8000/api';

// Simple API base selector
async function getApiBase() {
    try {
        // Try HTTPS first (for desktop)
        const response = await fetch(`${API_BASE_HTTPS}/health/`, { method: 'GET' });
        if (response.ok) {
            console.log('Using HTTPS API');
            return API_BASE_HTTPS;
        }
    } catch (error) {
        console.log('HTTPS failed, trying HTTP:', error.message);
    }
    
    try {
        // Fallback to HTTP
        const response = await fetch(`${API_BASE_HTTP}/health/`, { method: 'GET' });
        if (response.ok) {
            console.log('Using HTTP API');
            return API_BASE_HTTP;
        }
    } catch (error) {
        console.log('HTTP also failed:', error.message);
    }
    
    console.log('No API available, using HTTPS as default');
    return API_BASE_HTTPS;
}

// Simple status display
function showStatus(message, isError = false) {
    const statusDiv = document.getElementById('status') || document.getElementById('question-status');
    if (statusDiv) {
        statusDiv.textContent = message;
        statusDiv.style.color = isError ? 'red' : 'green';
    } else {
        console.log('Status:', message);
    }
}

// Simple login function
async function handleLogin(event) {
    event.preventDefault();
    console.log('Login attempt started');
    
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    
    if (!username || !password) {
        showStatus('Please enter username and password', true);
        return;
    }
    
    try {
        const apiBase = await getApiBase();
        const response = await fetch(`${apiBase}/auth/login/`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ username, password })
        });
        
        console.log('Login response status:', response.status);
        
        if (response.ok) {
            const data = await response.json();
            authToken = data.access;
            currentUser = data.user || { username, name: username };
            
            console.log('Login successful');
            showStatus('Login successful!');
            
            // Hide login form and show main interface
            document.getElementById('login-section').style.display = 'none';
            document.getElementById('main-interface').style.display = 'block';
            
        } else {
            const errorData = await response.json().catch(() => ({}));
            console.log('Login failed:', response.status, errorData);
            showStatus(`Login failed: ${errorData.detail || 'Unknown error'}`, true);
        }
    } catch (error) {
        console.error('Login error:', error);
        showStatus(`Login error: ${error.message}`, true);
    }
}

// Simple upload function
async function uploadAndInsert() {
    console.log('Upload attempt started');
    
    if (!authToken) {
        showStatus('Please login first', true);
        return;
    }
    
    const fileInput = document.getElementById('question-image');
    if (!fileInput || !fileInput.files.length) {
        showStatus('Please select an image first', true);
        return;
    }
    
    const file = fileInput.files[0];
    const formData = new FormData();
    formData.append('image', file);
    formData.append('question_text', 'Image uploaded from PowerPoint add-in');
    formData.append('class_id', '1');
    
    try {
        const apiBase = await getApiBase();
        const response = await fetch(`${apiBase}/questions/`, {
            method: 'POST',
            body: formData,
            headers: {
                'Authorization': `Bearer ${authToken}`
            }
        });
        
        console.log('Upload response status:', response.status);
        
        if (response.ok) {
            const data = await response.json();
            const imageUrl = data.image_url;
            
            console.log('Upload successful, inserting image:', imageUrl);
            
            // Insert image into PowerPoint
            await Office.onReady();
            await PowerPoint.run(async (context) => {
                const slide = context.presentation.getSelectedSlide();
                slide.shapes.addImage(imageUrl);
                await context.sync();
            });
            
            showStatus('Image uploaded and inserted successfully!');
        } else {
            const errorData = await response.json().catch(() => ({}));
            console.log('Upload failed:', response.status, errorData);
            showStatus(`Upload failed: ${errorData.detail || 'Unknown error'}`, true);
        }
    } catch (error) {
        console.error('Upload error:', error);
        showStatus(`Upload error: ${error.message}`, true);
    }
}

// Simple test function
function testConnection() {
    console.log('Testing connection...');
    showStatus('Testing connection...');
    
    getApiBase().then(apiBase => {
        fetch(`${apiBase}/health/`)
            .then(response => response.json())
            .then(data => {
                console.log('Health check result:', data);
                showStatus(`Connection successful: ${data.status}`);
            })
            .catch(error => {
                console.error('Health check failed:', error);
                showStatus(`Connection failed: ${error.message}`, true);
            });
    });
}

// Initialize the add-in
function initializeAddIn() {
    console.log('Initializing add-in...');
    
    try {
        // Set up event listeners
        const loginForm = document.getElementById('login-form');
        if (loginForm) {
            loginForm.addEventListener('submit', handleLogin);
            console.log('Login form listener added');
        }
        
        const uploadBtn = document.getElementById('uploadBtn');
        if (uploadBtn) {
            uploadBtn.addEventListener('click', uploadAndInsert);
            console.log('Upload button listener added');
        }
        
        const testBtn = document.getElementById('test-connection');
        if (testBtn) {
            testBtn.addEventListener('click', testConnection);
            console.log('Test button listener added');
        }
        
        // Show main interface if already logged in
        if (authToken) {
            document.getElementById('login-section').style.display = 'none';
            document.getElementById('main-interface').style.display = 'block';
        }
        
        console.log('Add-in initialized successfully');
        showStatus('Add-in ready');
        
    } catch (error) {
        console.error('Initialization error:', error);
        showStatus(`Initialization error: ${error.message}`, true);
    }
}

// Office.js ready handler
if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        console.log('Office.js ready:', info);
        initializeAddIn();
    });
} else {
    console.log('Office.js not available, initializing anyway');
    // Fallback for when Office.js is not available
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeAddIn);
    } else {
        initializeAddIn();
    }
}
