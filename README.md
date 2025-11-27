# Wood App Form

A Progressive Web App (PWA) for conducting wood appliance inspections and generating professional PDF reports. This app works offline and can be installed on any device.

## Features

### 📋 Comprehensive Inspection Forms
- **Multiple Appliance Types**: Supports Stoves, Kitchen Wood Ranges, Inserts, Furnaces, Boilers, Factory Built Fireplaces, Pellet Stoves, Hearth Mounts, Outdoor Wood Boilers, and Masonry Fireplaces
- **Smart Form Adaptation**: Form fields automatically show/hide based on selected appliance type
- **Multiple Appliances**: Add and manage multiple appliances in a single inspection
- **Auto-Complete Manufacturers**: Type-ahead search for common wood appliance manufacturers

### 🔧 Advanced Clearance Tracking
- **Automatic Clearance Requirements**: Default clearance values populate based on appliance type
- **Flue Pipe Types**: Select Single Wall (18" clearance) or Double Wall (6" clearance) flue pipe
- **Shielding Support**: Per-row shielding checkboxes that automatically reduce clearances by 50%
- **Visual Feedback**: 
  - 🟢 **Green** - Actual clearance meets or exceeds requirements
  - 🔴 **Red** - Actual clearance is insufficient (within 2" of required)
  - 🟡 **Yellow** - Actual clearance has minimal buffer (exactly required or +1")

### 📄 PDF Generation
- **Professional Reports**: Generate inspection reports with all appliance details and clearances
- **Preview Mode**: Preview PDFs before saving
- **Save & Share**: Download PDFs or share directly on mobile devices
- **Chimney Code Legend**: Automatically includes chimney code descriptions

### 🎨 User Interface
- **Dark/Light Mode**: Toggle between themes for comfortable viewing in any environment
- **Mobile Responsive**: Optimized layouts for phones, tablets, and desktop
- **Offline Support**: Works without internet connection after first load
- **Form Persistence**: Data is automatically saved as you type

### 🔄 Data Management
- **Add Multiple Appliances**: Inspect multiple units in one session
- **Delete Appliances**: Remove individual appliances from multi-appliance inspections
- **Reset Form**: Clear all data with confirmation dialog
- **Auto-Save**: Your work is saved locally in your browser

## Installation Instructions

### 📱 Android Installation

1. **Open in Chrome**:
   - Open Chrome browser on your Android device
   - Navigate to your Wood App URL

2. **Install the App**:
   - Tap the menu icon (⋮) in the top-right corner
   - Select **"Add to Home screen"** or **"Install app"**
   - Confirm by tapping **"Add"** or **"Install"**

3. **Launch**:
   - Find the Wood App icon on your home screen
   - Tap to open like any native app

**Alternative Method**:
- Look for the install banner that appears at the bottom of the screen
- Tap **"Install"** on the banner

### 📱 iPhone/iPad Installation

1. **Open in Safari**:
   - Open Safari browser (must use Safari, not Chrome)
   - Navigate to your Wood App URL

2. **Add to Home Screen**:
   - Tap the Share button (box with arrow pointing up) at the bottom
   - Scroll down and tap **"Add to Home Screen"**
   - Edit the name if desired
   - Tap **"Add"** in the top-right corner

3. **Launch**:
   - Find the Wood App icon on your home screen
   - Tap to open like any native app

### 💻 Windows PC Installation

#### Method 1: Chrome/Edge (Recommended)
1. **Open in Chrome or Edge**:
   - Open Chrome or Microsoft Edge browser
   - Navigate to your Wood App URL

2. **Install as App**:
   - **Chrome**: Click the install icon (⊕) in the address bar, or go to menu (⋮) → "Install Wood App Form"
   - **Edge**: Click the install icon (⊕) in the address bar, or go to menu (⋯) → "Apps" → "Install this site as an app"
   - Click **"Install"** in the dialog

3. **Launch**:
   - The app opens in its own window
   - Find it in your Start Menu or Desktop
   - Pin to taskbar for quick access

#### Method 2: Browser Bookmark
- Simply bookmark the URL in any browser
- Use it as a web application

### 💻 Mac Installation

#### Method 1: Chrome/Safari
1. **Using Chrome**:
   - Open Chrome and navigate to the Wood App URL
   - Click menu (⋯) → "Install Wood App Form"
   - The app will open in its own window

2. **Using Safari**:
   - Add the URL to your bookmarks
   - Use as a web app in Safari

## How to Use

### Basic Workflow

1. **Start an Inspection**:
   - Enter Policy # or Name
   - Set Survey Date
   - Select who completed the inspection

2. **Enter Appliance Details**:
   - Select appliance **Type** (triggers relevant form fields)
   - Enter **Manufacturer** and **Model**
   - Select **Chimney Code** (major type and location)
   - Set **Chimney Condition**, **Shielding**, **Label**, and **Location**
   - Choose **Flue Pipe Type** if applicable (SW/DW/N/A)

3. **Record Clearances**:
   - Enter **Required** clearances (auto-filled for some types)
   - Measure and enter **Actual** clearances
   - Check **Shielded** boxes if shielding is present (reduces requirement by 50%)
   - Watch for color-coded feedback (green/yellow/red)

4. **Add Floor Pad Measurements**:
   - Default requirements are pre-filled (18" front, 8" sides/rear)
   - Enter actual measurements
   - Relevant rows hide/show based on appliance type

5. **Add Notes**:
   - Use the Notes section for additional observations
   - Notes appear in the generated PDF

6. **Add More Appliances** (optional):
   - Click **"Add Another Appliance"** to inspect multiple units
   - Each appliance gets its own section in the PDF

7. **Generate PDF**:
   - **Preview PDF**: Review before saving
   - **Save**: Download or share the PDF report
   - **Reset**: Clear form to start new inspection

### Tips & Tricks

- **Auto-Fill**: Selecting certain appliance types will auto-fill some clearance requirements
- **Flue Pipe Types**: 
  - SW (Single Wall) = 18" clearance requirement
  - DW (Double Wall) = 6" clearance requirement
  - N/A = For appliances without flue pipes
- **Shielding**: Check the shielded box for any row to automatically reduce required clearance by 50%
- **Dark Mode**: Toggle theme for comfortable viewing in low-light conditions
- **Offline Mode**: After first load, the app works without internet
- **Multi-Device**: Install on all your devices for flexibility

## Technical Details

- **Progressive Web App (PWA)**: Installable on all platforms
- **Offline Support**: Service worker caches app for offline use
- **Local Storage**: Data persists in browser localStorage
- **PDF Generation**: Client-side PDF creation with jsPDF
- **Responsive Design**: Adapts to phone, tablet, and desktop screens

## Troubleshooting

### App Won't Install
- **Android**: Ensure you're using Chrome or Edge browser
- **iPhone**: Must use Safari browser (not Chrome)
- **PC**: Use Chrome, Edge, or other Chromium-based browsers

### Data Not Saving
- Ensure cookies/localStorage is enabled in your browser
- Check browser privacy settings

### PDF Generation Issues
- Ensure JavaScript is enabled
- Check that pop-ups are not blocked
- On mobile, grant file access permissions when prompted

### Clearances Not Turning Green
- Ensure both Required and Actual fields have values
- Check that Actual value is equal to or greater than Required

### Reset Not Working
- Confirm the reset action in the dialog
- If data persists, clear browser cache

## Updates

The app automatically updates when you reload with an internet connection. The version number is displayed at the top of the form.

## Support

For issues or questions, contact the development team or your system administrator.

---

**Version**: Check the version display at the top of the app  
**Last Updated**: November 2025
