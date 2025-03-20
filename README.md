Chainsaw Forensic Timeliner
Chainsaw Forensic Timeliner is a PowerShell-based tool that automates the process of aggregating and formatting forensic artifacts from Chainsaw into a structured Master Timeline in Excel.

This tool is designed for forensic analysts who need to process event logs, MFT, RDP events, and other forensic artifacts efficiently.

Special Thanks
Huge thanks to WithSecure Countercept (FranticTyping, AlexKornitzer) for creating Chainsaw, an invaluable tool for forensic analysis.

Features
Automatically combines all Chainsaw CSV outputs into a single Excel timeline.
Normalizes timestamps into a readable format (MM/DD/YYYY HH:MM:SS).
Assigns an artifact name to each row for easy identification.
Supports color-coding for different artifacts (see color_macro.vbs for details).
Preserves important metadata like event IDs, source addresses, user information, and service details.
Sorts the final timeline by Date/Time.
Requirements
Windows:
PowerShell (Version 5.1 or later)
ImportExcel PowerShell Module (for Excel support)
powershell
Copy
Edit
Install-Module ImportExcel -Force -Scope CurrentUser
Chainsaw (Download Here)
Optional:
Excel Macro for Color Coding:
The file color_macro.vbs can be used to apply color coding to each row based on the artifact type.
Installation
Clone this repository:

sh
Copy
Edit
git clone https://github.com/your-repo/chainsaw-timeliner.git
cd chainsaw-timeliner
Run Chainsaw (if you haven't already):

sh
Copy
Edit
chainsaw hunt C:\kape\triage\C\Windows\System32\winevt\Logs --rules rules/ --mapping sigma-event-logs-all.yml --csv --output C:\KAPE_Output\Chainsaw
Run the PowerShell script:

powershell
Copy
Edit
.\chainsaw_forensic_timeline.ps1 -CsvDirectory "C:\KAPE_Output\Chainsaw" -OutputFile "C:\KAPE_Output\Master_Timeline.xlsx"
(Optional) Apply Color Coding

Open Master_Timeline.xlsx
Run color_macro.vbs to color rows based on artifact type.
Color Coding (Excel)
The following artifact types are color-coded for better visibility:

Artifact Name	Color
account_tampering	Blue
antivirus	Green
indicator_removal	Red
lateral_movement	Orange
login_attacks	Yellow
MFT - FileNameCreated0x30	Purple
microsoft_rds_events_-_user_profile_disk	Cyan
persistence	Olive
powershell_engine_state	Pink
powershell_script	Brown
rdp_events	Lime
service_installation	Teal
sigma	Dark Orchid
Output Format
The final Master Timeline.xlsx contains the following structured columns:

Column	Description
Date/Time	Formatted timestamp (MM/DD/YYYY HH:MM:SS)
Artifact Name	Source of the artifact
Event ID	Event log ID
Computer	Hostname of the system
Detections	MITRE ATT&CK detections (if applicable)
Threat Path	Key path or process involved
User	Username associated with the event
IP Address	IP address (if applicable)
Service Details	Service type, start type, and account info
Evidence Path	Original source file location
Notes & Future Improvements
Automated color coding in PowerShell is slow. For now, color-coding is handled via color_macro.vbs.
Filtering and search functions can be added for quicker analysis.
License
This project is released under the MIT License.

