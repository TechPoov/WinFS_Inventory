# WinFS_Inventory
WinFS_InventoryCSV is a Windows utility that scans folders, files, or both, and generates clean CSV reports for inventory, audit, and documentation purposes. The tool is fully automated using a configuration (.config) file—no pop‑ups or user prompts.

📘 WinFS_InventoryCSV
A lightweight Windows File System Inventory & Logging Utility (VBScript-based)
🚀 Overview
WinFS_InventoryCSV is a simple, script-based, zero-dependency Windows utility that scans folders and files, collects metadata, and exports the results into clean, Excel-friendly CSV reports.
It is designed for:
•	System administrators
•	Developers
•	IT auditors
•	Power users managing large folder structures
•	Automating scheduled server scans
No installation required. Runs silently. Produces professional output.
________________________________________
⭐ Key Features
•	Recursive scanning of folders
•	Supports 3 modes:
o	Files – list all files
o	Folders – list all folders
o	Both – files + folders
•	INI-style multi-job config (Job1, Job2…)
•	Structured CSV logs per job
•	Run-level summary file
•	Access-denied safe (continues even if some folders fail)
•	No popups (server-safe)
•	Supports scheduling via Task Scheduler
•	Clear versioning + documentation pack
________________________________________

🛠️ How to Use
1️⃣ Place script + config in the same folder
WinFS_InventoryCSV_V1.0.vbs  
WinFS_InventoryCSV_V1.0.config
2️⃣ Edit .config
Example:
[Job1]
ScanFolder=C:\Data
OutputFolder=C:\InventoryOut
Mode=Both

[Job2]
ScanFolder=D:\Projects
OutputFolder=D:\ScanOut
Mode=Files
3️⃣ Run the tool
Double-click:
WinFS_InventoryCSV_V1.0.vbs
or via command line:
cscript WinFS_InventoryCSV_V1.0.vbs
4️⃣ Check outputs
You will get:
•	Data CSV → details of files/folders
•	Log CSV → events, warnings, scan results
•	Summary CSV → one row per job
________________________________________
📄 Output Files Explained
✔ Data File
Contains one row per file/folder with:
•	Path
•	Name
•	Extension
•	Parent folder
•	Size (for files)
•	Created date
•	Modified date
•	Attributes
✔ Job Log
Tracks:
•	Start/end
•	Access denied folders
•	Errors
•	Each CSV created
✔ Summary File
Lists all jobs in a single place.
________________________________________
🔧 Configuration Options
Key	Meaning	Required
ScanFolder	The root folder to scan	Yes
OutputFolder	Folder where CSV/logs go	No (defaults to Output\)
Mode	Files / Folders / Both	Yes
Email	Reserved for future email summary	Optional
________________________________________
🧪 Sample Use Cases
•	Inventory of shared drives
•	Periodic audit scans
•	Checking software project directories
•	Finding large or old files
•	Pre-migration assessments
•	Cleanup planning
For more, see:
📄 docs/1_StarterKit/4_Use_Cases_User_Stories.docx
________________________________________
📘 Full Documentation
All detailed documentation is available in /docs and organized by audience:
1_StarterKit → For all users
Quick start, training, user guide
2_Management → For managers
Vision, scope, release notes
3_Admin → For sysadmins
Run instructions, scheduling, permissions
4_Developer → For maintainers
FRD, HLD, LLD, Developer Guide
5_Testing → For QA
Test plan + test cases + sample outputs
________________________________________
🧭 Versioning Strategy
The project follows:
Major.Minor (X.Y)
•	Major → architecture changes or new capabilities
•	Minor → incremental features, improvements, bug fixes
See CHANGELOG.md for full history.
________________________________________
🤝 Contributing
Contributions welcome!
Submit:
•	Pull Requests
•	Issues
•	Feature ideas
•	Bug reports
GitHub Issues tab → “New Issue”
________________________________________
📝 License
MIT License
________________________________________
📬 Contact
For technical queries:
📧 techpoov+WinFS_InventoryCSV@gmail.com

