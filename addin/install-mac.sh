#!/bin/bash

# PowerPoint Add-in Installation Script for Mac
# This script installs the add-in by copying the manifest to PowerPoint's wef folder

echo "📦 Installing PPT Image Tool Add-in for Mac..."
echo ""

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Define paths
WEF_DIR=~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
MANIFEST_SOURCE=/Users/ali/Desktop/imageUpload/addin/dist/manifest.xml

# Check if manifest file exists
if [ ! -f "$MANIFEST_SOURCE" ]; then
    echo -e "${RED}❌ Error: manifest.xml not found at $MANIFEST_SOURCE${NC}"
    echo "Please run 'npm run dev' in the addin directory first to build the files."
    exit 1
fi

# Create wef directory if it doesn't exist
if [ ! -d "$WEF_DIR" ]; then
    echo "📁 Creating wef directory..."
    mkdir -p "$WEF_DIR"
fi

# Copy manifest file
echo "📋 Copying manifest.xml to PowerPoint wef folder..."
cp "$MANIFEST_SOURCE" "$WEF_DIR/"

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✅ Manifest file copied successfully!${NC}"
    echo ""
    echo "📝 Next steps:"
    echo "1. ${YELLOW}Make sure both servers are running:${NC}"
    echo "   - Django backend: http://localhost:8000"
    echo "   - Webpack dev server: https://localhost:3000"
    echo ""
    echo "2. ${YELLOW}Open or restart PowerPoint${NC}"
    echo ""
    echo "3. ${YELLOW}Look for the add-in in the Home tab:${NC}"
    echo "   - Go to the ${GREEN}Home${NC} tab in PowerPoint"
    echo "   - Look for ${GREEN}'PPT Image Tool'${NC} group"
    echo "   - Click ${GREEN}'Show Taskpane'${NC} button"
    echo ""
    echo "4. ${YELLOW}If you don't see the add-in:${NC}"
    echo "   - Go to: Insert → Add-ins → My Add-ins"
    echo "   - Look under 'Developer Add-ins'"
    echo "   - You should see 'PPT Image Tool'"
    echo ""
    echo "🎉 Installation complete!"
else
    echo -e "${RED}❌ Error: Failed to copy manifest file${NC}"
    exit 1
fi

