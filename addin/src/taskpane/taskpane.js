/* global Office */

console.log('Taskpane.js loaded successfully');

// Global state management
let authToken = null;
let currentUser = null;
let userType = null;
let classes = [];
let selectedClassId = null;
let studentClasses = [];
let selectedStudentClassId = null;
let availableQuestions = [];
let selectedQuestion = null;
let myAnswers = [];
let reviewClasses = [];
let selectedReviewClassId = null;
let reviewQuestions = [];
let selectedReviewQuestion = null;
let reviewAnswers = [];
let analyticsData = {};
let powerpointClasses = [];
let powerpointQuestions = [];
let selectedPowerPointClassId = null;
let selectedPowerPointQuestionId = null;
let isAutoRefreshEnabled = false;
let autoRefreshInterval = null;
let currentSlideInfo = {};
let studentAutoRefreshInterval = null; // Auto-refresh for student answers
const API_BASE = 'http://localhost:8000/api';

// JWT token decoder
function decodeJWT(token) {
    try {
        const base64Url = token.split('.')[1];
        const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
        return JSON.parse(jsonPayload);
    } catch (error) {
        console.error('Error decoding JWT:', error);
        return null;
    }
}

// Determine user type based on username
function determineUserType(username) {
    // Simple rule: if username starts with 'student', it's a student
    // Otherwise, it's an instructor
    if (username.toLowerCase().startsWith('student')) {
        return 'student';
    }
    return 'instructor';
}

// Initialize the Office.js add-in
Office.onReady((info) => {
    console.log('Office.js initialized successfully');
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);
    
    if (info.host === Office.HostType.PowerPoint) {
        console.log('PowerPoint detected, initializing add-in...');
        initializeAddIn();
    } else {
        console.log('Not PowerPoint, but initializing anyway for testing...');
        initializeAddIn();
    }
});

function initializeAddIn() {
    console.log('initializeAddIn called');
    // Initialize tab state first
    initializeTabState();
    
    // Check if already logged in
    if (authToken) {
        showLoggedInInterface();
    } else {
        showLoginInterface();
    }
    
    // Initialize event listeners
    initializeEventListeners();
}

function initializeTabState() {
    // Hide all tab content initially
    const tabContents = document.querySelectorAll('.tab-content');
    console.log('Found tab contents:', tabContents.length);
    tabContents.forEach(content => {
        content.classList.add('hidden');
        console.log('Hiding:', content.id);
    });
    
    // Remove active class from all tabs
    const tabButtons = document.querySelectorAll('.tab-button');
    console.log('Found tab buttons:', tabButtons.length);
    tabButtons.forEach(tab => {
        tab.classList.remove('active');
        console.log('Removing active from:', tab.id);
    });
}

function initializeEventListeners() {
    console.log('initializeEventListeners called');
    // Login form
    const loginForm = document.getElementById('login-form');
    if (loginForm) {
        console.log('Login form found, adding event listener');
        loginForm.addEventListener('submit', handleLogin);
    } else {
        console.error('Login form not found!');
    }
    
    // Login button (backup)
    const loginBtn = document.getElementById('btn-login');
    if (loginBtn) {
        console.log('Login button found, adding click listener');
        loginBtn.addEventListener('click', handleLogin);
    } else {
        console.error('Login button not found!');
    }
    
    // Logout button
    const logoutBtn = document.getElementById('btn-logout');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', handleLogout);
    }
    
    // Tab navigation
    const instructorTab = document.getElementById('tab-instructor');
    const studentTab = document.getElementById('tab-student');
    const reviewTab = document.getElementById('tab-review');
    
    if (instructorTab) {
        instructorTab.addEventListener('click', () => setActiveTab('instructor'));
    }
    if (studentTab) {
        studentTab.addEventListener('click', () => setActiveTab('student'));
    }
    if (reviewTab) {
        reviewTab.addEventListener('click', () => {
            setActiveTab('review');
            if (reviewClasses.length === 0) {
                loadReviewClasses();
            }
        });
    }
    
    // Question form
    const questionForm = document.getElementById('form-question');
    if (questionForm) {
        questionForm.addEventListener('submit', handleQuestionSubmit);
    }
    
    // Class selector
    const classSelect = document.getElementById('select-class');
    if (classSelect) {
        // Remove existing listener to prevent duplicates
        classSelect.removeEventListener('change', handleClassChange);
        classSelect.addEventListener('change', handleClassChange);
    }
    
    // Image upload
    const imageInput = document.getElementById('question-image');
    if (imageInput) {
        imageInput.addEventListener('change', handleImageSelect);
    }
    
    // Remove image button
    const removeImageBtn = document.getElementById('remove-image');
    if (removeImageBtn) {
        removeImageBtn.addEventListener('click', handleRemoveImage);
    }
    
    // Test API button
    const testApiButton = document.getElementById('test-api');
    if (testApiButton) {
        testApiButton.addEventListener('click', testApiConnection);
    }
    
    // Student functionality
    const studentClassSelect = document.getElementById('student-class-select');
    if (studentClassSelect) {
        studentClassSelect.addEventListener('change', handleStudentClassChange);
    }
    
    const answerForm = document.getElementById('form-answer');
    if (answerForm) {
        answerForm.addEventListener('submit', handleAnswerSubmit);
    }
    
    const cancelAnswerBtn = document.getElementById('btn-cancel-answer');
    if (cancelAnswerBtn) {
        cancelAnswerBtn.addEventListener('click', handleCancelAnswer);
    }
    
    const answerImageInput = document.getElementById('answer-image');
    if (answerImageInput) {
        answerImageInput.addEventListener('change', handleAnswerImageSelect);
    }
    
    const removeAnswerImageBtn = document.getElementById('remove-answer-image');
    if (removeAnswerImageBtn) {
        removeAnswerImageBtn.addEventListener('click', handleRemoveAnswerImage);
    }
    
    // Review functionality
    const reviewClassSelect = document.getElementById('review-class-select');
    if (reviewClassSelect) {
        reviewClassSelect.addEventListener('change', handleReviewClassChange);
    }
    
    // PowerPoint functionality
    const powerpointClassSelect = document.getElementById('powerpoint-class-select');
    if (powerpointClassSelect) {
        powerpointClassSelect.addEventListener('change', handlePowerPointClassChange);
    }
    
    const powerpointQuestionSelect = document.getElementById('powerpoint-question-select');
    if (powerpointQuestionSelect) {
        powerpointQuestionSelect.addEventListener('change', handlePowerPointQuestionChange);
    }
    
    const insertQuestionBtn = document.getElementById('btn-insert-question');
    if (insertQuestionBtn) {
        insertQuestionBtn.addEventListener('click', handleInsertQuestion);
    }
    
    const enableLiveAnswersBtn = document.getElementById('btn-enable-live-answers');
    if (enableLiveAnswersBtn) {
        enableLiveAnswersBtn.addEventListener('click', handleEnableLiveAnswers);
    }
    
    const showAnalyticsBtn = document.getElementById('btn-show-analytics');
    if (showAnalyticsBtn) {
        showAnalyticsBtn.addEventListener('click', handleShowAnalytics);
    }
    
    const quickReviewBtn = document.getElementById('btn-quick-review');
    if (quickReviewBtn) {
        quickReviewBtn.addEventListener('click', handleQuickReview);
    }
    
    const toggleAutoRefreshBtn = document.getElementById('btn-toggle-auto-refresh');
    if (toggleAutoRefreshBtn) {
        toggleAutoRefreshBtn.addEventListener('click', handleToggleAutoRefresh);
    }
}

function showLoginInterface() {
    document.getElementById('login-section').classList.remove('hidden');
    document.getElementById('logged-in-section').classList.add('hidden');
}

function showLoggedInInterface() {
    document.getElementById('login-section').classList.add('hidden');
    document.getElementById('logged-in-section').classList.remove('hidden');
    
    // Update user info
    document.getElementById('logged-in-user').textContent = `Logged in as ${currentUser}`;
    
    // Show/hide tabs based on user type
    showTabsForUserType();
    
    // Load data based on user type
    if (userType === 'instructor') {
        loadClasses();
        loadReviewClasses();
        loadPowerPointClasses();
        initializePowerPoint();
        // Set the initial active tab
        setActiveTab('instructor');
    } else if (userType === 'student') {
        loadStudentClasses();
        // Set the initial active tab
        setActiveTab('student');
        // Start auto-refresh for student answers
        startStudentAutoRefresh();
    }
}

function showTabsForUserType() {
    console.log('showTabsForUserType called with userType:', userType);
    
    const instructorTab = document.getElementById('tab-instructor');
    const studentTab = document.getElementById('tab-student');
    const reviewTab = document.getElementById('tab-review');
    const powerpointTab = document.getElementById('tab-powerpoint');
    
    console.log('Found tabs:', {
        instructor: instructorTab,
        student: studentTab,
        review: reviewTab,
        powerpoint: powerpointTab
    });
    
    // Hide all tabs first
    instructorTab.classList.add('hidden');
    studentTab.classList.add('hidden');
    reviewTab.classList.add('hidden');
    powerpointTab.classList.add('hidden');
    
    // Show tabs based on user type
    if (userType === 'instructor') {
        console.log('Showing instructor tabs');
        instructorTab.classList.remove('hidden');
        reviewTab.classList.remove('hidden');
        powerpointTab.classList.remove('hidden');
    } else if (userType === 'student') {
        console.log('Showing student tab');
        studentTab.classList.remove('hidden');
    } else {
        console.log('Unknown user type:', userType);
    }
}

async function handleLogin(event) {
    console.log('handleLogin called');
    event.preventDefault();
    
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    console.log('Login attempt for:', username);
    const loginBtn = document.getElementById('btn-login');
    const errorDiv = document.getElementById('login-error');
    
    // Show loading state
    loginBtn.textContent = 'Logging in...';
    loginBtn.disabled = true;
    errorDiv.classList.add('hidden');
    
    try {
        const response = await fetch(`${API_BASE}/auth/login/`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ username, password })
        });
        
        if (response.ok) {
            const data = await response.json();
            authToken = data.access;
            currentUser = username;
            
            // Decode JWT token to get user type
            const tokenData = decodeJWT(authToken);
            console.log('Token data:', tokenData);
            console.log('Username:', username);
            
            if (tokenData) {
                userType = tokenData.user_type || determineUserType(username);
                console.log('User type from token:', tokenData.user_type);
                console.log('User type determined:', determineUserType(username));
                console.log('Final user type:', userType);
            } else {
                userType = determineUserType(username);
                console.log('No token data, user type determined:', userType);
            }
            
            console.log('Login successful');
            showLoggedInInterface();
        } else {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Login failed');
        }
    } catch (error) {
        console.error('Login error:', error);
        errorDiv.textContent = error.message;
        errorDiv.classList.remove('hidden');
    } finally {
        loginBtn.textContent = 'Login';
        loginBtn.disabled = false;
    }
}

function handleLogout() {
    authToken = null;
    currentUser = null;
    userType = null;
    classes = [];
    selectedClassId = null;
    
    // Stop auto-refresh intervals
    stopStudentAutoRefresh();
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
        autoRefreshInterval = null;
    }
    
    // Clear forms
    document.getElementById('login-form').reset();
    document.getElementById('form-question').reset();
    document.getElementById('recent-questions').innerHTML = '<p class="no-questions">No questions yet. Create your first question above!</p>';
    
    showLoginInterface();
}

async function fetchWithAuth(url, options = {}) {
    const headers = {
        ...options.headers
    };
    
    // Only set Content-Type if not already set (for file uploads)
    if (!headers['Content-Type'] && !options.body instanceof FormData) {
        headers['Content-Type'] = 'application/json';
    }
    
    if (authToken) {
        headers['Authorization'] = `Bearer ${authToken}`;
    }
    
    const response = await fetch(url, {
        ...options,
        headers
    });
    
    if (response.status === 401) {
        handleLogout();
        throw new Error('Session expired. Please login again.');
    }
    
    return response;
}

async function loadClasses() {
    try {
        const response = await fetchWithAuth(`${API_BASE}/classes/`);
        
        if (response.ok) {
            classes = await response.json();
            populateClassSelector();
            
            if (classes.length > 0) {
                selectedClassId = classes[0].id;
                const classSelect = document.getElementById('select-class');
                
                // Temporarily remove event listener to prevent duplicate calls
                classSelect.removeEventListener('change', handleClassChange);
                classSelect.value = selectedClassId;
                classSelect.addEventListener('change', handleClassChange);
                
                // Load questions
                loadQuestions(selectedClassId);
            }
        } else {
            throw new Error('Failed to load classes');
        }
    } catch (error) {
        console.error('Error loading classes:', error);
        showStatus('Error loading classes: ' + error.message, 'error');
    }
}

function populateClassSelector() {
    const select = document.getElementById('select-class');
    select.innerHTML = '';
    
    if (classes.length === 0) {
        select.innerHTML = '<option value="">No classes available</option>';
        return;
    }
    
    classes.forEach(classItem => {
        const option = document.createElement('option');
        option.value = classItem.id;
        option.textContent = classItem.class_name;
        select.appendChild(option);
    });
}

function handleClassChange(event) {
    console.log('handleClassChange called with value:', event.target.value);
    selectedClassId = parseInt(event.target.value);
    if (selectedClassId) {
        loadQuestions(selectedClassId);
    } else {
        document.getElementById('recent-questions').innerHTML = '<p class="no-questions">Select a class to view questions.</p>';
    }
}

let isLoadingQuestions = false;
let lastLoadedClassId = null;

async function loadQuestions(classId) {
    console.log('loadQuestions called with classId:', classId);
    
    // Prevent duplicate calls for the same class
    if (isLoadingQuestions && lastLoadedClassId === classId) {
        console.log('Already loading questions for class', classId, 'skipping duplicate call');
        return;
    }
    
    isLoadingQuestions = true;
    lastLoadedClassId = classId;
    
    try {
        const response = await fetchWithAuth(`${API_BASE}/questions/?class_id=${classId}`);
        
        if (response.ok) {
            const questions = await response.json();
            console.log('Questions loaded:', questions.length);
            displayQuestions(questions);
        } else {
            throw new Error('Failed to load questions');
        }
    } catch (error) {
        console.error('Error loading questions:', error);
        showStatus('Error loading questions: ' + error.message, 'error');
    } finally {
        isLoadingQuestions = false;
    }
}

let lastDisplayedQuestions = null;

function displayQuestions(questions) {
    const container = document.getElementById('recent-questions');
    
    // Check if we're displaying the same questions (prevent duplicates)
    const questionsString = JSON.stringify(questions);
    if (lastDisplayedQuestions === questionsString) {
        console.log('Skipping duplicate display of same questions');
        return;
    }
    
    console.log('Displaying questions:', questions.length);
    lastDisplayedQuestions = questionsString;
    
    // Clear the container first
    container.innerHTML = '';
    
    if (questions.length === 0) {
        container.innerHTML = '<p class="no-questions">No questions yet. Create your first question above!</p>';
        return;
    }
    
    container.innerHTML = questions.map(question => `
        <div class="question-item">
            <div class="question-text">${question.question_text}</div>
            ${question.image_url ? `<div class="question-image"><img src="${question.image_url}" alt="Question image" class="question-thumbnail"></div>` : ''}
            <div class="question-meta">
                <span class="question-date">${new Date(question.created_at).toLocaleString()}</span>
            </div>
        </div>
    `).join('');
}

function handleImageSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // Validate file size (8MB max)
    const maxSize = 8 * 1024 * 1024; // 8MB
    if (file.size > maxSize) {
        showStatus('File size exceeds 8MB limit. Please choose a smaller image.', 'error');
        event.target.value = '';
        return;
    }
    
    // Show preview
    const reader = new FileReader();
    reader.onload = (e) => {
        const previewImg = document.getElementById('preview-img');
        const imagePreview = document.getElementById('image-preview');
        
        previewImg.src = e.target.result;
        imagePreview.classList.remove('hidden');
    };
    reader.readAsDataURL(file);
}

function handleRemoveImage() {
    document.getElementById('question-image').value = '';
    document.getElementById('image-preview').classList.add('hidden');
}

async function handleQuestionSubmit(event) {
    event.preventDefault();
    
    if (!checkAuth()) return;
    
    const formData = new FormData();
    const questionText = document.getElementById('question-text').value;
    const imageFile = document.getElementById('question-image').files[0];
    
    if (!questionText.trim()) {
        showStatus('Please enter a question text.', 'error');
        return;
    }
    
    if (!selectedClassId) {
        showStatus('Please select a class.', 'error');
        return;
    }
    
    formData.append('question_text', questionText);
    formData.append('class_id', selectedClassId);
    
    if (imageFile) {
        formData.append('image', imageFile);
    }
    
    const submitBtn = document.getElementById('btn-submit-question');
    const originalText = submitBtn.textContent;
    
    // Show loading state
    submitBtn.textContent = 'Submitting...';
    submitBtn.disabled = true;
    
    try {
        const response = await fetchWithAuth(`${API_BASE}/questions/`, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const question = await response.json();
            showStatus('Question created successfully!', 'success');
            
            // Clear form
            document.getElementById('form-question').reset();
            document.getElementById('image-preview').classList.add('hidden');
            
            // Reset cache and reload questions
            lastDisplayedQuestions = null;
            loadQuestions(selectedClassId);
        } else {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Failed to create question');
        }
    } catch (error) {
        console.error('Error creating question:', error);
        showStatus('Error creating question: ' + error.message, 'error');
    } finally {
        submitBtn.textContent = originalText;
        submitBtn.disabled = false;
    }
}

function setActiveTab(tabName) {
    console.log(`Switching to ${tabName} tab`);
    
    // Remove active class from all tabs
    const allTabs = document.querySelectorAll('.tab-button');
    console.log('Found tabs:', allTabs.length);
    allTabs.forEach(tab => {
        tab.classList.remove('active');
        console.log('Removing active from:', tab.id);
    });
    
    // Hide all tab content
    const allContents = document.querySelectorAll('.tab-content');
    console.log('Found contents:', allContents.length);
    allContents.forEach(content => {
        content.classList.add('hidden');
        console.log('Hiding:', content.id);
    });
    
    // Add active class to selected tab
    const tabButton = document.getElementById(`tab-${tabName}`);
    if (tabButton) {
        tabButton.classList.add('active');
        console.log('Activated tab:', tabButton.id);
    } else {
        console.error('Tab button not found:', `tab-${tabName}`);
    }
    
    // Show corresponding content
    const tabContent = document.getElementById(`${tabName}-content`);
    if (tabContent) {
        tabContent.classList.remove('hidden');
        console.log('Showing content:', tabContent.id);
    } else {
        console.error('Tab content not found:', `${tabName}-content`);
    }
    
    console.log(`Successfully switched to ${tabName} tab`);
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('question-status');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
    
    // Clear status after 5 seconds
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status-message';
    }, 5000);
}

function checkAuth() {
    if (!authToken) {
        showLoginInterface();
        return false;
    }
    return true;
}

async function testApiConnection() {
    const apiResult = document.getElementById('api-result');
    const testButton = document.getElementById('test-api');
    
    // Show loading state
    testButton.textContent = 'Testing...';
    testButton.disabled = true;
    apiResult.innerHTML = '<p>Connecting to Django API...</p>';
    
    try {
        // Test the Django health endpoint
        const response = await fetch(`${API_BASE}/health/`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            apiResult.innerHTML = `
                <div class="success-message">
                    <h4>✅ API Connection Successful!</h4>
                    <p><strong>Response:</strong> ${JSON.stringify(data, null, 2)}</p>
                    <p><strong>Status:</strong> ${response.status} ${response.statusText}</p>
                </div>
            `;
        } else {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
    } catch (error) {
        console.error('API Test Error:', error);
        apiResult.innerHTML = `
            <div class="error-message">
                <h4>❌ API Connection Failed</h4>
                <p><strong>Error:</strong> ${error.message}</p>
                <p><strong>Possible causes:</strong></p>
                <ul>
                    <li>Django server is not running (start with: cd backend && python manage.py runserver)</li>
                    <li>CORS configuration issue</li>
                    <li>Network connectivity problem</li>
                </ul>
            </div>
        `;
    } finally {
        // Reset button state
        testButton.textContent = 'Test API Connection';
        testButton.disabled = false;
    }
}

// Student functionality
async function loadStudentClasses() {
    try {
        const response = await fetchWithAuth(`${API_BASE}/student/classes/`);
        
        if (response.ok) {
            studentClasses = await response.json();
            console.log('Student classes loaded:', studentClasses);
            populateStudentClassSelector();
        } else {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Failed to load classes');
        }
    } catch (error) {
        console.error('Error loading student classes:', error);
        showStudentStatus('Error loading classes: ' + error.message, 'error');
    }
}

function populateStudentClassSelector() {
    const select = document.getElementById('student-class-select');
    select.innerHTML = '';
    
    if (studentClasses.length === 0) {
        select.innerHTML = '<option value="">No classes available</option>';
        return;
    }
    
    studentClasses.forEach(classItem => {
        const option = document.createElement('option');
        option.value = classItem.id;
        option.textContent = classItem.class_name;
        select.appendChild(option);
    });
    
    // Automatically select the first class and load its questions
    if (studentClasses.length > 0) {
        selectedStudentClassId = studentClasses[0].id;
        select.value = selectedStudentClassId;
        console.log('Auto-selecting class:', selectedStudentClassId);
        loadAvailableQuestions(selectedStudentClassId);
        loadMyAnswers(selectedStudentClassId);
    }
}

function handleStudentClassChange(event) {
    selectedStudentClassId = parseInt(event.target.value);
    if (selectedStudentClassId) {
        loadAvailableQuestions(selectedStudentClassId);
        loadMyAnswers(selectedStudentClassId);
    } else {
        document.getElementById('available-questions').innerHTML = '<p class="no-questions">Select a class to view questions.</p>';
        document.getElementById('my-answers').innerHTML = '<p class="no-answers">No answers submitted yet.</p>';
    }
}

async function loadAvailableQuestions(classId) {
    console.log('loadAvailableQuestions called with classId:', classId);
    try {
        const response = await fetchWithAuth(`${API_BASE}/student/questions/?class_id=${classId}`);
        
        if (response.ok) {
            availableQuestions = await response.json();
            console.log('Available questions loaded:', availableQuestions);
            displayAvailableQuestions(availableQuestions);
        } else {
            const errorData = await response.json();
            throw new Error(errorData.error || errorData.detail || 'Failed to load questions');
        }
    } catch (error) {
        console.error('Error loading available questions:', error);
        showStudentStatus('Error loading questions: ' + error.message, 'error');
    }
}

function displayAvailableQuestions(questions) {
    console.log('displayAvailableQuestions called with:', questions.length, 'questions');
    const container = document.getElementById('available-questions');
    
    if (questions.length === 0) {
        console.log('No questions available, showing message');
        container.innerHTML = '<p class="no-questions">No questions available for this class.</p>';
        return;
    }
    
    console.log('Displaying questions:', questions);
    container.innerHTML = questions.map(question => `
        <div class="question-item clickable" data-question-id="${question.id}">
            <div class="question-text">${question.question_text}</div>
            ${question.image_url ? `<div class="question-image"><img src="${question.image_url}" alt="Question image" class="question-thumbnail"></div>` : ''}
            <div class="question-meta">
                <span class="question-date">${new Date(question.created_at).toLocaleString()}</span>
                <button class="btn-answer" data-question-id="${question.id}">Answer This Question</button>
            </div>
        </div>
    `).join('');
    
    // Add click listeners to answer buttons
    container.querySelectorAll('.btn-answer').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const questionId = parseInt(btn.dataset.questionId);
            selectQuestion(questionId);
        });
    });
}

function selectQuestion(questionId) {
    selectedQuestion = availableQuestions.find(q => q.id === questionId);
    if (selectedQuestion) {
        // Display selected question
        document.getElementById('selected-question-text').textContent = selectedQuestion.question_text;
        
        const imageDisplay = document.getElementById('selected-question-image');
        if (selectedQuestion.image_url) {
            imageDisplay.innerHTML = `<img src="${selectedQuestion.image_url}" alt="Question image" class="question-display-image">`;
        } else {
            imageDisplay.innerHTML = '';
        }
        
        // Show answer form
        document.getElementById('answer-form-section').classList.remove('hidden');
        
        // Scroll to form
        document.getElementById('answer-form-section').scrollIntoView({ behavior: 'smooth' });
    }
}

function handleCancelAnswer() {
    selectedQuestion = null;
    document.getElementById('form-answer').reset();
    document.getElementById('answer-image-preview').classList.add('hidden');
    document.getElementById('answer-form-section').classList.add('hidden');
}

function handleAnswerImageSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // Validate file size (8MB max)
    const maxSize = 8 * 1024 * 1024; // 8MB
    if (file.size > maxSize) {
        showStudentStatus('File size exceeds 8MB limit. Please choose a smaller image.', 'error');
        event.target.value = '';
        return;
    }
    
    // Show preview
    const reader = new FileReader();
    reader.onload = (e) => {
        const previewImg = document.getElementById('answer-preview-img');
        const imagePreview = document.getElementById('answer-image-preview');
        
        previewImg.src = e.target.result;
        imagePreview.classList.remove('hidden');
    };
    reader.readAsDataURL(file);
}

function handleRemoveAnswerImage() {
    document.getElementById('answer-image').value = '';
    document.getElementById('answer-image-preview').classList.add('hidden');
}

async function handleAnswerSubmit(event) {
    event.preventDefault();
    
    if (!checkAuth()) return;
    
    if (!selectedQuestion) {
        showStudentStatus('Please select a question first.', 'error');
        return;
    }
    
    const formData = new FormData();
    const imageFile = document.getElementById('answer-image').files[0];
    
    if (!imageFile) {
        showStudentStatus('Please select an image for your answer.', 'error');
        return;
    }
    
    formData.append('question_id', selectedQuestion.id);
    formData.append('image', imageFile);
    
    const submitBtn = document.getElementById('btn-submit-answer');
    const originalText = submitBtn.textContent;
    
    // Show loading state
    submitBtn.textContent = 'Submitting...';
    submitBtn.disabled = true;
    
    try {
        const response = await fetchWithAuth(`${API_BASE}/student/submit-answer/`, {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const answer = await response.json();
            showStudentStatus('Answer submitted successfully!', 'success');
            
            // Clear form and hide it
            document.getElementById('form-answer').reset();
            document.getElementById('answer-image-preview').classList.add('hidden');
            document.getElementById('answer-form-section').classList.add('hidden');
            selectedQuestion = null;
            
            // Reload my answers
            loadMyAnswers(selectedStudentClassId);
        } else {
            const errorData = await response.json();
            throw new Error(errorData.detail || 'Failed to submit answer');
        }
    } catch (error) {
        console.error('Error submitting answer:', error);
        showStudentStatus('Error submitting answer: ' + error.message, 'error');
    } finally {
        submitBtn.textContent = originalText;
        submitBtn.disabled = false;
    }
}

async function loadMyAnswers(classId) {
    console.log('loadMyAnswers called with classId:', classId);
    try {
        const response = await fetchWithAuth(`${API_BASE}/student/answers/?class_id=${classId}`);
        
        if (response.ok) {
            const newAnswers = await response.json();
            console.log('My answers loaded:', newAnswers);
            
            // Check if any answers have changed (for debugging)
            const hasChanges = JSON.stringify(newAnswers) !== JSON.stringify(myAnswers);
            if (hasChanges) {
                console.log('Answers have changed, updating display');
            }
            
            myAnswers = newAnswers;
            displayMyAnswers(myAnswers);
        } else {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Failed to load answers');
        }
    } catch (error) {
        console.error('Error loading my answers:', error);
        showStudentStatus('Error loading answers: ' + error.message, 'error');
    }
}

function displayMyAnswers(answers) {
    const container = document.getElementById('my-answers');
    
    if (answers.length === 0) {
        container.innerHTML = '<p class="no-answers">No answers submitted yet.</p>';
        return;
    }
    
    container.innerHTML = answers.map(answer => `
        <div class="answer-item">
            <div class="answer-meta">
                <span class="answer-date">Submitted: ${new Date(answer.created_at).toLocaleString()}</span>
                <span class="answer-status ${answer.liked ? 'liked' : 'not-liked'}">
                    ${answer.liked ? '❤️ Liked by instructor' : '⏳ Pending review'}
                </span>
            </div>
            ${answer.image_url ? `<div class="answer-image"><img src="${answer.image_url}" alt="Answer image" class="answer-thumbnail"></div>` : ''}
        </div>
    `).join('');
}

function showStudentStatus(message, type) {
    const statusDiv = document.getElementById('answer-status');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
    
    // Clear status after 5 seconds
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status-message';
    }, 5000);
}

// Auto-refresh functionality for students
function startStudentAutoRefresh() {
    console.log('Starting student auto-refresh');
    // Clear any existing interval
    if (studentAutoRefreshInterval) {
        clearInterval(studentAutoRefreshInterval);
    }
    
    // Refresh every 5 seconds
    studentAutoRefreshInterval = setInterval(() => {
        if (userType === 'student' && selectedStudentClassId) {
            console.log('Auto-refreshing student answers');
            loadMyAnswers(selectedStudentClassId);
        }
    }, 5000);
}

function stopStudentAutoRefresh() {
    console.log('Stopping student auto-refresh');
    if (studentAutoRefreshInterval) {
        clearInterval(studentAutoRefreshInterval);
        studentAutoRefreshInterval = null;
    }
}

// Review functionality
async function loadReviewClasses() {
    console.log('loadReviewClasses called');
    try {
        const response = await fetchWithAuth(`${API_BASE}/classes/`);
        
        if (response.ok) {
            reviewClasses = await response.json();
            console.log('Review classes loaded:', reviewClasses.length);
            populateReviewClassSelector();
        } else {
            throw new Error('Failed to load classes');
        }
    } catch (error) {
        console.error('Error loading review classes:', error);
        showReviewStatus('Error loading classes: ' + error.message, 'error');
    }
}

function populateReviewClassSelector() {
    console.log('populateReviewClassSelector called with', reviewClasses.length, 'classes');
    const select = document.getElementById('review-class-select');
    select.innerHTML = '';
    
    if (reviewClasses.length === 0) {
        select.innerHTML = '<option value="">No classes available</option>';
        return;
    }
    
    reviewClasses.forEach(classItem => {
        const option = document.createElement('option');
        option.value = classItem.id;
        option.textContent = classItem.class_name;
        select.appendChild(option);
    });
    console.log('Review class selector populated with', reviewClasses.length, 'options');
    
    // Automatically select the first class and trigger change event
    if (reviewClasses.length > 0) {
        selectedReviewClassId = reviewClasses[0].id;
        select.value = selectedReviewClassId;
        console.log('Auto-selecting review class:', selectedReviewClassId);
        // Trigger the change event to load analytics and questions
        loadReviewQuestions(selectedReviewClassId);
        loadAnalytics(selectedReviewClassId);
    }
}

function handleReviewClassChange(event) {
    selectedReviewClassId = parseInt(event.target.value);
    if (selectedReviewClassId) {
        loadReviewQuestions(selectedReviewClassId);
        loadAnalytics(selectedReviewClassId);
    } else {
        document.getElementById('review-question-selection').classList.add('hidden');
        document.getElementById('review-answers-section').classList.add('hidden');
        document.getElementById('review-questions').innerHTML = '<p class="no-questions">Select a class to view questions.</p>';
    }
}

async function loadReviewQuestions(classId) {
    console.log('loadReviewQuestions called with classId:', classId);
    try {
        const response = await fetchWithAuth(`${API_BASE}/questions/?class_id=${classId}`);
        
        if (response.ok) {
            reviewQuestions = await response.json();
            console.log('Review questions loaded:', reviewQuestions.length);
            displayReviewQuestions(reviewQuestions);
            document.getElementById('review-question-selection').classList.remove('hidden');
        } else {
            throw new Error('Failed to load questions');
        }
    } catch (error) {
        console.error('Error loading review questions:', error);
        showReviewStatus('Error loading questions: ' + error.message, 'error');
    }
}

function displayReviewQuestions(questions) {
    console.log('displayReviewQuestions called with', questions.length, 'questions');
    const container = document.getElementById('review-questions');
    
    if (questions.length === 0) {
        container.innerHTML = '<p class="no-questions">No questions available for this class.</p>';
        return;
    }
    
    container.innerHTML = questions.map(question => `
        <div class="question-item clickable" data-question-id="${question.id}">
            <div class="question-text">${question.question_text}</div>
            ${question.image_url ? `<div class="question-image"><img src="${question.image_url}" alt="Question image" class="question-thumbnail"></div>` : ''}
            <div class="question-meta">
                <span class="question-date">${new Date(question.created_at).toLocaleString()}</span>
                <button class="btn-review" data-question-id="${question.id}">Review Answers</button>
            </div>
        </div>
    `).join('');
    
    console.log('Review questions displayed:', questions.length);
    
    // Add click listeners to review buttons
    container.querySelectorAll('.btn-review').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const questionId = parseInt(btn.dataset.questionId);
            selectReviewQuestion(questionId);
        });
    });
}

function selectReviewQuestion(questionId) {
    selectedReviewQuestion = reviewQuestions.find(q => q.id === questionId);
    if (selectedReviewQuestion) {
        // Update question title
        document.getElementById('review-question-title').textContent = selectedReviewQuestion.question_text;
        
        // Load answers for this question
        loadReviewAnswers(questionId);
        
        // Show answers section
        document.getElementById('review-answers-section').classList.remove('hidden');
        
        // Scroll to answers
        document.getElementById('review-answers-section').scrollIntoView({ behavior: 'smooth' });
    }
}

async function loadReviewAnswers(questionId) {
    console.log('loadReviewAnswers called with questionId:', questionId);
    try {
        // Use the instructor class answers endpoint and filter by question
        const response = await fetchWithAuth(`${API_BASE}/instructor/class-answers/?class_id=${selectedReviewClassId}`);
        
        if (response.ok) {
            const allAnswers = await response.json();
            // Filter answers for the specific question
            reviewAnswers = allAnswers.filter(answer => answer.question === questionId);
            console.log('Review answers loaded:', reviewAnswers.length);
            displayReviewAnswers(reviewAnswers);
            updateReviewStats(reviewAnswers);
        } else {
            throw new Error('Failed to load answers');
        }
    } catch (error) {
        console.error('Error loading review answers:', error);
        showReviewStatus('Error loading answers: ' + error.message, 'error');
    }
}

function displayReviewAnswers(answers) {
    const container = document.getElementById('student-answers');
    
    if (answers.length === 0) {
        container.innerHTML = '<p class="no-answers">No answers submitted for this question yet.</p>';
        return;
    }
    
    container.innerHTML = answers.map(answer => `
        <div class="answer-review-item" data-answer-id="${answer.id}">
            <div class="answer-header">
                <div class="student-info">
                    <span class="student-name">${answer.student_name || 'Unknown Student'}</span>
                    <span class="answer-date">${new Date(answer.created_at).toLocaleString()}</span>
                </div>
                <div class="answer-actions">
                    <button class="btn-like ${answer.liked ? 'liked' : 'not-liked'}" data-answer-id="${answer.id}" data-liked="${answer.liked}">
                        ${answer.liked ? '❤️ Liked' : '🤍 Like'}
                    </button>
                </div>
            </div>
            ${answer.image_url ? `
                <div class="answer-image-container">
                    <img src="${answer.image_url}" alt="Student answer" class="answer-review-image" data-image-url="${answer.image_url}">
                </div>
            ` : ''}
        </div>
    `).join('');
    
    // Add click listeners to like buttons
    container.querySelectorAll('.btn-like').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const answerId = parseInt(btn.dataset.answerId);
            const isLiked = btn.dataset.liked === 'true';
            toggleAnswerLike(answerId, !isLiked);
        });
    });
    
    // Add click listeners to answer images
    container.querySelectorAll('.answer-review-image').forEach(img => {
        img.addEventListener('click', (e) => {
            e.stopPropagation();
            const imageUrl = img.dataset.imageUrl;
            openImageModal(imageUrl);
        });
    });
}

async function toggleAnswerLike(answerId, liked) {
    try {
        const response = await fetchWithAuth(`${API_BASE}/answers/${answerId}/`, {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${authToken}`
            },
            body: JSON.stringify({ liked: liked })
        });
        
        if (response.ok) {
            const updatedAnswer = await response.json();
            
            // Update the button state
            const btn = document.querySelector(`[data-answer-id="${answerId}"] .btn-like`);
            if (btn) {
                btn.dataset.liked = liked;
                btn.textContent = liked ? '❤️ Liked' : '🤍 Like';
                btn.classList.toggle('liked', liked);
                btn.classList.toggle('not-liked', !liked);
            }
            
            // Update stats
            updateReviewStats(reviewAnswers);
            updateAnalytics();
            
            showReviewStatus(liked ? 'Answer liked!' : 'Answer unliked!', 'success');
        } else {
            throw new Error('Failed to update answer');
        }
    } catch (error) {
        console.error('Error toggling answer like:', error);
        showReviewStatus('Error updating answer: ' + error.message, 'error');
    }
}

function updateReviewStats(answers) {
    const totalAnswers = answers.length;
    const likedAnswers = answers.filter(a => a.liked).length;
    
    document.getElementById('total-answers').textContent = `${totalAnswers} answers`;
    document.getElementById('liked-answers').textContent = `${likedAnswers} liked`;
}

async function loadAnalytics(classId) {
    console.log('loadAnalytics called with classId:', classId);
    try {
        // Load questions for analytics
        const questionsResponse = await fetchWithAuth(`${API_BASE}/questions/?class_id=${classId}`);
        const questions = questionsResponse.ok ? await questionsResponse.json() : [];
        console.log('Questions for analytics:', questions.length);
        
        // Load all answers for analytics using instructor endpoint
        const answersResponse = await fetchWithAuth(`${API_BASE}/instructor/class-answers/?class_id=${classId}`);
        let allAnswers = [];
        if (answersResponse.ok) {
            allAnswers = await answersResponse.json();
            console.log('Answers for analytics:', allAnswers.length);
        } else {
            console.error('Failed to load answers for analytics');
        }
        
        // Calculate analytics
        const totalQuestions = questions.length;
        const totalAnswers = allAnswers.length;
        const likedAnswers = allAnswers.filter(a => a.liked).length;
        const participationRate = totalQuestions > 0 ? Math.round((allAnswers.length / (totalQuestions * 5)) * 100) : 0; // Assuming 5 students per class
        
        analyticsData = {
            totalQuestions,
            totalAnswers,
            likedAnswers,
            participationRate
        };
        
        console.log('Analytics data:', analyticsData);
        updateAnalytics();
    } catch (error) {
        console.error('Error loading analytics:', error);
    }
}

function updateAnalytics() {
    document.getElementById('total-questions-count').textContent = analyticsData.totalQuestions || 0;
    document.getElementById('total-answers-count').textContent = analyticsData.totalAnswers || 0;
    document.getElementById('liked-answers-count').textContent = analyticsData.likedAnswers || 0;
    document.getElementById('participation-rate').textContent = `${analyticsData.participationRate || 0}%`;
}

function openImageModal(imageUrl) {
    // Simple image modal - in a real implementation, you'd use a proper modal library
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.8);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 1000;
        cursor: pointer;
    `;
    
    const img = document.createElement('img');
    img.src = imageUrl;
    img.style.cssText = `
        max-width: 90%;
        max-height: 90%;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.5);
    `;
    
    modal.appendChild(img);
    document.body.appendChild(modal);
    
    modal.addEventListener('click', () => {
        document.body.removeChild(modal);
    });
}

// Make openImageModal available globally
window.openImageModal = openImageModal;

function showReviewStatus(message, type) {
    // Use the same status system as other tabs
    showStatus(message, type);
}

// PowerPoint Integration functionality
async function loadPowerPointClasses() {
    try {
        const response = await fetchWithAuth(`${API_BASE}/classes/`);
        
        if (response.ok) {
            powerpointClasses = await response.json();
            populatePowerPointClassSelector();
        } else {
            throw new Error('Failed to load classes');
        }
    } catch (error) {
        console.error('Error loading PowerPoint classes:', error);
        showPowerPointStatus('Error loading classes: ' + error.message, 'error');
    }
}

function populatePowerPointClassSelector() {
    const select = document.getElementById('powerpoint-class-select');
    select.innerHTML = '';
    
    if (powerpointClasses.length === 0) {
        select.innerHTML = '<option value="">No classes available</option>';
        return;
    }
    
    powerpointClasses.forEach(classItem => {
        const option = document.createElement('option');
        option.value = classItem.id;
        option.textContent = classItem.class_name;
        select.appendChild(option);
    });
}

function handlePowerPointClassChange(event) {
    selectedPowerPointClassId = parseInt(event.target.value);
    if (selectedPowerPointClassId) {
        loadPowerPointQuestions(selectedPowerPointClassId);
    } else {
        document.getElementById('powerpoint-question-select').innerHTML = '<option value="">Select a class first</option>';
    }
}

async function loadPowerPointQuestions(classId) {
    try {
        const response = await fetchWithAuth(`${API_BASE}/questions/?class_id=${classId}`);
        
        if (response.ok) {
            powerpointQuestions = await response.json();
            populatePowerPointQuestionSelector();
        } else {
            throw new Error('Failed to load questions');
        }
    } catch (error) {
        console.error('Error loading PowerPoint questions:', error);
        showPowerPointStatus('Error loading questions: ' + error.message, 'error');
    }
}

function populatePowerPointQuestionSelector() {
    const select = document.getElementById('powerpoint-question-select');
    select.innerHTML = '';
    
    if (powerpointQuestions.length === 0) {
        select.innerHTML = '<option value="">No questions available</option>';
        return;
    }
    
    powerpointQuestions.forEach(question => {
        const option = document.createElement('option');
        option.value = question.id;
        option.textContent = question.question_text.length > 50 
            ? question.question_text.substring(0, 50) + '...' 
            : question.question_text;
        select.appendChild(option);
    });
}

function handlePowerPointQuestionChange(event) {
    selectedPowerPointQuestionId = parseInt(event.target.value);
}

async function handleInsertQuestion() {
    if (!checkAuth()) return;
    
    if (!selectedPowerPointClassId || !selectedPowerPointQuestionId) {
        showPowerPointStatus('Please select both a class and a question.', 'error');
        return;
    }
    
    const selectedQuestion = powerpointQuestions.find(q => q.id === selectedPowerPointQuestionId);
    if (!selectedQuestion) {
        showPowerPointStatus('Selected question not found.', 'error');
        return;
    }
    
    const includeImage = document.getElementById('include-image').checked;
    const addAnswerSpace = document.getElementById('add-answer-space').checked;
    
    try {
        // Get current slide information
        await getCurrentSlideInfo();
        
        // Insert question into slide
        await insertQuestionIntoSlide(selectedQuestion, includeImage, addAnswerSpace);
        
        showPowerPointStatus('Question inserted successfully!', 'success');
    } catch (error) {
        console.error('Error inserting question:', error);
        showPowerPointStatus('Error inserting question: ' + error.message, 'error');
    }
}

async function getCurrentSlideInfo() {
    return new Promise((resolve, reject) => {
        // Check if we're in PowerPoint
        if (!Office.context || !Office.context.document) {
            console.log('Not in PowerPoint - using test mode');
            currentSlideInfo = {
                title: 'Test Presentation',
                slideNumber: 1,
                totalSlides: 1
            };
            updateSlideInfo();
            resolve(currentSlideInfo);
            return;
        }
        
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Get presentation info
                Office.context.document.getFilePropertiesAsync((propsResult) => {
                    if (propsResult.status === Office.AsyncResultStatus.Succeeded) {
                        currentSlideInfo = {
                            title: propsResult.value.title || 'Untitled Presentation',
                            slideNumber: 1, // This would need to be determined from the current slide
                            totalSlides: 1 // This would need to be calculated
                        };
                        updateSlideInfo();
                        resolve(currentSlideInfo);
                    } else {
                        reject(new Error('Failed to get presentation properties'));
                    }
                });
            } else {
                reject(new Error('Failed to get slide information'));
            }
        });
    });
}

function updateSlideInfo() {
    document.getElementById('current-slide-number').textContent = currentSlideInfo.slideNumber || '-';
    document.getElementById('total-slides').textContent = currentSlideInfo.totalSlides || '-';
    document.getElementById('presentation-title').textContent = currentSlideInfo.title || '-';
}

async function insertQuestionIntoSlide(question, includeImage, addAnswerSpace) {
    return new Promise((resolve, reject) => {
        // Create the content to insert
        let content = `QUESTION: ${question.question_text}\n\n`;
        
        if (includeImage && question.image_url) {
            content += `[Image: ${question.image_url}]\n\n`;
        }
        
        if (addAnswerSpace) {
            content += `STUDENT ANSWERS:\n`;
            content += `________________\n`;
            content += `________________\n`;
            content += `________________\n\n`;
            content += `Submit your answer at: [Student Portal URL]`;
        }
        
        // Insert the content into the current slide
        Office.context.document.setSelectedDataAsync(content, {
            coercionType: Office.CoercionType.Text
        }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                reject(new Error('Failed to insert content into slide'));
            }
        });
    });
}

function handleEnableLiveAnswers() {
    const btn = document.getElementById('btn-enable-live-answers');
    const isEnabled = btn.textContent.includes('Enable');
    
    if (isEnabled) {
        btn.textContent = 'Disable Live Answers';
        btn.classList.add('active');
        showPowerPointStatus('Live answer collection enabled!', 'success');
        
        // Start polling for new answers
        startAnswerPolling();
    } else {
        btn.textContent = 'Enable Live Answers';
        btn.classList.remove('active');
        showPowerPointStatus('Live answer collection disabled.', 'info');
        
        // Stop polling
        stopAnswerPolling();
    }
}

function startAnswerPolling() {
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
    }
    
    autoRefreshInterval = setInterval(async () => {
        if (selectedPowerPointQuestionId) {
            await loadLiveAnswers(selectedPowerPointQuestionId);
        }
    }, 5000); // Poll every 5 seconds
}

function stopAnswerPolling() {
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
        autoRefreshInterval = null;
    }
}

async function loadLiveAnswers(questionId) {
    try {
        const response = await fetchWithAuth(`${API_BASE}/answers/?question_id=${questionId}`);
        
        if (response.ok) {
            const answers = await response.json();
            // Update slide with new answers
            await updateSlideWithAnswers(answers);
        }
    } catch (error) {
        console.error('Error loading live answers:', error);
    }
}

async function updateSlideWithAnswers(answers) {
    return new Promise((resolve, reject) => {
        let content = `LIVE ANSWERS (${answers.length} total):\n\n`;
        
        answers.forEach((answer, index) => {
            content += `${index + 1}. ${answer.student_name || 'Student'} - `;
            content += `${answer.liked ? '❤️ Liked' : '⏳ Pending'}\n`;
            if (answer.image_url) {
                content += `   [Image: ${answer.image_url}]\n`;
            }
            content += `   Submitted: ${new Date(answer.created_at).toLocaleString()}\n\n`;
        });
        
        Office.context.document.setSelectedDataAsync(content, {
            coercionType: Office.CoercionType.Text
        }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            } else {
                reject(new Error('Failed to update slide with answers'));
            }
        });
    });
}

function handleShowAnalytics() {
    if (!selectedPowerPointQuestionId) {
        showPowerPointStatus('Please select a question first.', 'error');
        return;
    }
    
    // Switch to review tab and show analytics
    setActiveTab('review');
    showPowerPointStatus('Switched to analytics view.', 'success');
}

function handleQuickReview() {
    if (!selectedPowerPointQuestionId) {
        showPowerPointStatus('Please select a question first.', 'error');
        return;
    }
    
    // Switch to review tab and load the selected question
    setActiveTab('review');
    if (selectedReviewClassId) {
        selectReviewQuestion(selectedPowerPointQuestionId);
    }
    showPowerPointStatus('Switched to quick review.', 'success');
}

function handleToggleAutoRefresh() {
    const btn = document.getElementById('btn-toggle-auto-refresh');
    isAutoRefreshEnabled = !isAutoRefreshEnabled;
    
    if (isAutoRefreshEnabled) {
        btn.textContent = 'Disable Auto-refresh';
        btn.classList.add('active');
        startAnswerPolling();
        showPowerPointStatus('Auto-refresh enabled!', 'success');
    } else {
        btn.textContent = 'Enable Auto-refresh';
        btn.classList.remove('active');
        stopAnswerPolling();
        showPowerPointStatus('Auto-refresh disabled.', 'info');
    }
}

function initializePowerPoint() {
    // Initialize PowerPoint-specific functionality
    updatePowerPointStatus('Connected to PowerPoint', 'success');
    
    // Load current slide information
    getCurrentSlideInfo().catch(error => {
        console.error('Error getting slide info:', error);
        updatePowerPointStatus('Connected to PowerPoint (Slide info unavailable)', 'warning');
    });
}

function updatePowerPointStatus(message, type) {
    const statusIcon = document.getElementById('powerpoint-status-icon');
    const statusText = document.getElementById('powerpoint-status-text');
    
    statusText.textContent = message;
    
    // Update icon based on status
    switch (type) {
        case 'success':
            statusIcon.textContent = '✅';
            break;
        case 'warning':
            statusIcon.textContent = '⚠️';
            break;
        case 'error':
            statusIcon.textContent = '❌';
            break;
        default:
            statusIcon.textContent = '🔗';
    }
}

function showPowerPointStatus(message, type) {
    const statusDiv = document.getElementById('insert-status');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
    
    // Clear status after 5 seconds
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status-message';
    }, 5000);
}

// Error handling for Office.js
Office.onError = (error) => {
    console.error('Office.js Error:', error);
    showStatus('Office.js Error: ' + error.message, 'error');
};
