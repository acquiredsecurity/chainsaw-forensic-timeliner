# Chainsaw Forensic Timeliner

Chainsaw Forensic Timeliner is a PowerShell-based tool that automates the process of aggregating and formatting forensic artifacts from [Chainsaw](https://github.com/WithSecureLabs/chainsaw) into a structured **Master Timeline** in Excel.

This tool is designed for forensic analysts who need to process event logs, MFT, RDP events, and other forensic artifacts efficiently.

### Special Thanks
Huge thanks to **WithSecure Countercept** ([FranticTyping](https://twitter.com/FranticTyping), [AlexKornitzer](https://twitter.com/AlexKornitzer)) for creating [Chainsaw](https://github.com/WithSecureLabs/chainsaw), an invaluable tool for forensic analysis.

---
sample commandline:
.\chainsaw_forensic_timeline.ps1 -CsvDirectory "C:\chainsaw" -OutputFile "C:\chainsaw\Master_Timeline.xlsx"

-CsvDirectory  - the path to your chainsaw output
-OutputFile - the path to save your timeline to

## Features
- Automatically combines all **Chainsaw CSV outputs** into a single **Excel timeline**.
- **Normalizes timestamps** into a readable format (MM/DD/YYYY HH:MM:SS).
- Assigns an **artifact name** to each row for easy identification.
- Supports **color-coding** for different artifacts (see `color_macro.vbs` for details).
- Preserves **important metadata** like event IDs, source addresses, user information, and service details.
- Sorts the final timeline by **Date/Time**.

---
## ToDo
- I am missing some of the rule fields most likely
- Integrate Kape output

## Requirements
### Windows:
1. **PowerShell** (Version 5.1 or later)
2. **ImportExcel PowerShell Module** (for Excel support)
   ```powershell
   Install-Module ImportExcel -Force -Scope CurrentUser
3. Chainsaw (https://github.com/WithSecureLabs/chainsaw)
Optional:
Excel Macro for Color Coding:
The file color_macro.vbs can be used to apply color coding to each row based on the artifact type.

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
The final Master_Timeline.xlsx contains the following structured columns:

Column	Description
Date/Time	- Formatted timestamp (MM/DD/YYYY HH:MM:SS)
Artifact Name	- Source of the artifact
Event ID	- Event log ID
Computer	- Hostname of the system
Detections	MITRE ATT&CK detections (if applicable)
Threat Path	- Key path, file or process involved
User	- Username associated with the event
IP Address	- IP address (if applicable)
AND more! I am probably missing some important mappings. Let me know!
