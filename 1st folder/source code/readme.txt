
🚀 SETUP INSTRUCTIONS
========================
STEP 1: Install Required Libraries

First, you need to install all the necessary Python libraries.

Open Command Prompt or Terminal
Navigate to your project folder:
cd "path/to/your/project/folder"
Run the requirements installer:
python requirements.py
This will automatically install all required libraries including:

PyAudio (for audio processing)
Vosk (for speech recognition)
Tkinter (for GUI)
PyWin32 (for Windows integration)
MySQL connector
And other dependencies



STEP 2: Setup MySQL Database

You need to import the database using XAMPP MySQL.

Download and install XAMPP from: https://www.apachefriends.org/
Start XAMPP Control Panel
Start both "Apache" and "MySQL" services
Open your web browser and go to: http://localhost/phpmyadmin/
Create a new database:

Click "New" on the left sidebar
Enter database name (e.g., "gesture_control")
Click "Create"


Import the SQL file:

Select your newly created database
Click "Import" tab
Click "Choose File" and select "gesture_db.sql"
Click "Go" to import



STEP 3: Configure Application

Make sure your system has:

A working microphone for speech recognition
PowerPoint installed (for subtitle integration)
Vosk model files in the correct path:
C:\Users[YourUsername]\Documents\VOS\model\vosk-model-en-us-0.22\

Make sure you have internet connection
STEP 4: Run the Application

Once everything is set up:

Navigate to your project folder
Run the main application:
python main_control_window.py
The application window should open and you can start using the features!

 FEATURES
==============

Presentation control using hand gestures and voice command
Real time subtitles
Real time translation
Customizable gesture

🔧 TROUBLESHOOTING
====================
Common Issues and Solutions:
Issue: "Module not found" error
Solution: Make sure you ran requirements.py successfully. Try running:
pip install [module_name]
Issue: Database connection error
Solution:

Ensure XAMPP MySQL is running
Check database name and credentials
Verify gesture_db.sql was imported correctly

Issue: Microphone not working
Solution:

Check Windows microphone permissions
Test microphone in other applications
Run the application as administrator

Issue: Vosk model not found
Solution:

Download the Vosk model from: https://alphacephei.com/vosk/models
Extract to: C:\Users[YourUsername]\Documents\VOS\model\
Ensure the folder name matches exactly: vosk-model-en-us-0.22

Issue: PowerPoint integration not working
Solution:

Ensure Microsoft PowerPoint is installed
Run the application as administrator
Check that PowerPoint is not blocked by antivirus

📋 SYSTEM REQUIREMENTS
=========================

Windows 10 or later
Python 3.10 or higher
Microsoft PowerPoint XAMPP (for MySQL database)
Working microphone
At least 4GB RAM
2GB free disk space

============================

🌐 QUICK START - DOWNLOAD READY-TO-USE VERSION
=================================================
For users who want to skip the setup process:
Visit our website: https://airswipe.netlify.app/

Download the pre-built executable (.exe) file
No installation required - just run and use!
All dependencies are included in the executable