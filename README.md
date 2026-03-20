💎 Tanima Jewels - Cloud Shipping Portal
A high-efficiency tool for managing shipping labels and customer data for Tanima Jewels. This app parses addresses, generates labels, saves a local history, and syncs data to Google Sheets.

🚀 1. How to Use the App (Daily Workflow)
Step 1: Inputting Customer Data
Manual Entry: Fill in the fields (Name, House, PO, etc.). The app automatically converts your typing into Title Case (e.g., "KERALA" becomes "Kerala").

Auto-Fill (Fastest): Copy a messy address from WhatsApp/Instagram and paste it into the "Paste Scrambled Address" box. Click ✨ Auto-Fill & Capitalize. The app will intelligently extract the Name, Pincode, and Mobile number for you.

Step 2: Saving & Syncing
Add & Save Local: This button "locks in" the customer. They will appear in the History Table at the bottom and in the Print Queue.

Sync to Cloud: Once a label is added, click this button to send the details to your Google Sheet.

Note: If the status turns green, it's saved to the cloud forever.

Step 3: Printing
Print Labels: Click this to open the print window.

Printer Settings: To ensure the labels fit your sticker paper:

Margins: Set to "None" or "Minimum."

Scale: Set to "100%."

Background Graphics: Check this box.

🎨 2. How to Tweak & Customize (CSS Styling)
You can edit the "look" by changing the code in the <style> section of index.html.

Changing the Logo
Logo Size: Find .logo-container. Change width: 100px; and height: 100px;. If the logo looks too small on the label, increase these numbers.

Logo File: Your logo must be named logo.png and be in the same GitHub folder as your index.html. If you upload a new logo, make sure the name is exactly the same (case-sensitive).

Changing Colors & Fonts
Brand Gold: Search for #f1c40f and replace it with your preferred hex color code.

Fonts: Search for font-family. You can change it to 'Georgia' for a classic look or 'Montserrat' for something modern.

Adjusting Label Size
Label Height: Find .label and look for height: 68mm;. If your sticker paper has 4 labels per page, this number is usually correct. If you have more or fewer labels, adjust this height.

☁️ 3. How to Fix the Cloud Sync (Google Sheets)
If the "Sync to Cloud" button stops working, follow these steps:

Get the Script URL: Open your Google Sheet, go to Extensions > Apps Script.

Deploy: Click Deploy > Manage Deployments. Copy the Web App URL.

Update Code: In your index.html, find this line:
const WEB_APP_URL = "YOUR_URL_HERE";
Replace the old URL with the new one.

Permissions: When deploying, ensure "Who has access" is set to "Anyone". This is the most important step for GitHub to talk to Google.

🛠️ 4. Maintenance & Troubleshooting
Logo Not Showing: Ensure you didn't name the file Logo.png (capital L) while the code says logo.png (small l). GitHub is very strict about capital letters.

Old Version Showing: If you updated the code on GitHub but the website looks the same, press Ctrl + F5 to force the browser to refresh its memory.

Data Backup: Data in the "History" table is saved to your browser's memory. If you clear your browser cache, it might disappear. Always sync to the Cloud or click 📊 Export CSV once a month to keep an Excel backup.

Mobile Use: Open the link on your phone and select "Add to Home Screen". It will work just like a mobile app for on-the-go entries.