# PowerPoint Add-in Setup Guide

## Prerequisites

1. **Microsoft PowerPoint** (Windows or Mac)
2. **Development servers running:**
   - Django backend: `http://localhost:8000`
   - Add-in frontend: `https://localhost:3000`

## Installation Steps

### 1. Start the Development Servers

**Terminal 1 - Django Backend:**
```bash
cd backend
source venv/bin/activate  # On Windows: venv\Scripts\activate
python manage.py runserver
```

**Terminal 2 - Add-in Frontend:**
```bash
cd addin
npm run dev
```

### 2. Trust the Development Certificate

Since the add-in uses HTTPS on localhost, you need to trust the self-signed certificate:

**On Mac:**
1. Open Chrome or Safari and navigate to `https://localhost:3000`
2. Click "Advanced" and "Proceed to localhost (unsafe)"
3. Accept the certificate warning

**On Windows:**
1. Open Microsoft Edge and navigate to `https://localhost:3000`
2. Click "Advanced" and "Continue to localhost (unsafe)"
3. Accept the certificate warning

### 3. Install the Add-in in PowerPoint

#### Method 1: Automatic Installation (Mac) - RECOMMENDED ✅

Run the installation script:

```bash
cd addin
./install-mac.sh
```

This will automatically copy the manifest to PowerPoint's wef folder.

#### Method 2: Manual Installation (Mac)

1. **Create the wef folder** (if it doesn't exist):
```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

2. **Copy the manifest file**:
```bash
cp /Users/ali/Desktop/imageUpload/addin/dist/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

3. **Restart PowerPoint** (if it's already open)

#### Method 3: Upload from File (Mac Alternative)

1. Open PowerPoint
2. Go to **File** > **Get Add-ins** > **My Add-ins**
3. Click **"Upload My Add-in"**
4. Navigate to `/Users/ali/Desktop/imageUpload/addin/dist/manifest.xml`
5. Click **"Upload"**

#### Method 4: Windows Installation

1. Open PowerPoint
2. Go to **Insert** > **Get Add-ins** (or **My Add-ins**)
3. Select **"Upload My Add-in"**
4. Browse to `C:\Users\...\imageUpload\addin\dist\manifest.xml`
5. Click **"Upload"**

### 4. Using the Add-in in PowerPoint

1. **Open PowerPoint**
2. **Go to the Home tab**
3. **Look for "PPT Image Tool" group**
4. **Click "Show Taskpane"** button
5. **The add-in taskpane will open on the right side**

## Expected Behavior

### Login Flow
1. When you first open the add-in, you'll see the login screen
2. Enter credentials:
   - **Instructor**: `instructor1` / `password123`
   - **Student**: `student1` / `password123`

### User-Specific Interface

**For Instructors (instructor1):**
- **Instructor Tab**: Create questions with optional images
- **Review Tab**: Review and like student answers
- **PowerPoint Tab**: Insert questions into slides

**For Students (student1):**
- **Student Tab Only**: View questions and submit answers with images

### Features
- ✅ Login with username/password
- ✅ User type detection (instructor vs student)
- ✅ Dynamic tab display based on user role
- ✅ Create questions with images
- ✅ Submit answers with images
- ✅ Real-time answer status updates (auto-refresh every 5 seconds)
- ✅ Like/unlike student answers
- ✅ Analytics dashboard for instructors

## Troubleshooting

### Issue: Add-in button doesn't appear in Home tab
**Solution:** 
- Restart PowerPoint
- Re-install the add-in
- Check that both servers are running

### Issue: SSL Certificate Error
**Solution:**
- Trust the certificate in your browser first
- Visit `https://localhost:3000` and accept the security warning

### Issue: "Cannot connect to server"
**Solution:**
- Ensure Django backend is running on `http://localhost:8000`
- Ensure webpack dev server is running on `https://localhost:3000`
- Check that PostgreSQL database is running

### Issue: Login fails
**Solution:**
- Verify Django backend is accessible: `http://localhost:8000/api/health/`
- Check browser console for errors (F12)
- Verify credentials are correct

### Issue: Tabs not showing correctly
**Solution:**
- Logout and login again
- Check browser console for JWT token errors
- Verify user type is being set correctly

## Development Tips

1. **Debugging:**
   - Right-click in the taskpane → **Inspect** (or press F12)
   - Console logs will show in the browser developer tools

2. **Hot Reload:**
   - Changes to `taskpane.js`, `taskpane.html`, or `taskpane.css` will auto-reload
   - You may need to close and reopen the taskpane for manifest changes

3. **Testing:**
   - Test with both instructor and student accounts
   - Test creating questions and submitting answers
   - Verify real-time updates work (like functionality)

## Next Steps

After successful installation:
1. Login as `instructor1`
2. Create a question in the Instructor tab
3. Login as `student1` (in a separate PowerPoint window or after logout)
4. Submit an answer to the question
5. Login back as `instructor1`
6. Review and like the answer in the Review tab
7. Watch the student interface auto-update to show "Liked by instructor"

## Production Deployment

For production deployment:
1. Update the manifest URLs from `localhost` to your production domain
2. Use a valid SSL certificate (not self-signed)
3. Update `API_BASE` in `taskpane.js` to your production API URL
4. Build the add-in: `npm run build`
5. Deploy to Office Store or use centralized deployment in Microsoft 365 Admin Center

