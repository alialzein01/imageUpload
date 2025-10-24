# PPT Image Tool Add-in

PowerPoint add-in for creating image-based questions and collecting student answers.

## Setup Instructions

### 1. Install Dependencies
```bash
cd addin
npm install
```

### 2. Start Development Server
```bash
npm run dev
```
This starts the webpack dev server at `https://localhost:3000`

### 3. Trust Self-Signed Certificate
Since Office add-ins require HTTPS, you'll need to trust the self-signed certificate:

1. Open `https://localhost:3000` in your browser
2. Click "Advanced" → "Proceed to localhost (unsafe)"
3. This tells your system to trust the certificate

### 4. Sideload the Add-in in PowerPoint

#### Method 1: Shared Folder (Recommended)
```bash
# Create trusted manifests folder
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef

# Copy manifest to trusted location
cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

In PowerPoint:
1. Open PowerPoint
2. Go to Insert → Get Add-ins → My Add-ins → Shared Folder
3. Select your `manifest.xml`
4. Click Add

#### Method 2: Direct Sideload
```bash
npm run start  # Starts PowerPoint with add-in
npm run stop   # Stops the add-in
```

### 5. Test the Add-in

1. **Verify Office.js loads**: Check browser dev tools console for "Office initialized"
2. **Test tab switching**: Click Instructor, Student, Review tabs
3. **Test API connection**: 
   - Make sure Django server is running: `cd backend && python manage.py runserver`
   - Click "Test API" button in taskpane
   - Should display: `{"status": "ok"}`

## Troubleshooting

### Certificate Errors
- Visit `https://localhost:3000` in browser first
- Accept the certificate warning
- This tells your system to trust the self-signed cert

### Add-in Doesn't Appear
```bash
# Clear Office cache
rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/*
# Restart PowerPoint
# Re-sideload manifest
```

### CORS Errors
- Verify Django CORS settings include `https://localhost:3000`
- Check Django console for OPTIONS requests
- Ensure both Django and webpack dev server are running

## Project Structure

```
addin/
├── manifest.xml          # Office add-in manifest
├── package.json          # Dependencies and scripts
├── webpack.config.js     # Webpack configuration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html # Main HTML file
│   │   ├── taskpane.js   # JavaScript logic
│   │   └── taskpane.css  # Styling
│   └── assets/           # Icon images
└── dist/                 # Generated files
```

## Available Scripts

- `npm run dev` - Start development server
- `npm run build` - Build for production
- `npm run start` - Sideload add-in in PowerPoint
- `npm run stop` - Stop add-in

## Features

- **Tab Navigation**: Instructor, Student, Review modes
- **API Integration**: Test connection to Django backend
- **Professional UI**: Clean, responsive design
- **Office.js Integration**: Full Office API support

