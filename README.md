# HP_Firmware_Get

This PowerShell script will help you to download the vast majority of the official HP Printer Firmware directly from HP. 
It figures out the official links used by HP. If HP removes these links, the script will no longer work.

## How it works:
It guesses what format the printer uses for firmware and stores it locally in printers.csv if it is not already stored there.  
After successfully finding the changelog PDF for the desired firmware, it will search the PDF using iTextSharp for available firmware versions.  
Once it finds something that matches your search parameters, it will display them for you to select which exact firmware you want.  
It will then generate a direct download link from HP and download the printer firmware!  

## How to download:
- On the right side, there is a section that says releases.
- Click the latest version, and under Assets, download Source Code (zip).

## How to use:
- Extract the zip file.
- Right-click the HP_Firmware_get_dev_autoread.ps1 file and select Run with PowerShell.
- You will be prompted for a printer model number.
  This is expecting a format like: E55650, M630, or PW586.
- Next, it will prompt you for a firmware version.
  This is expecting a format like: 5, 5.5, 5.5.0.3, etc.
- The script will attempt to find firmware versions that match your request. If this succeeds, it will give you a list of options.
- Select a firmware from the list, by using the number associated to the line. ex. 1, 2, 3, 4, etc. Do not type in the firmware version.
- The selected firmware version will start to download into the same folder that the script exists in.


## Features:
- Printer firmware file format prediction logic.
- Store predicted logic locally for future use (In printers.csv).
- Analyze HP firmware changelogs to detect available firmware versions.


This is one of many projects I have been working on to expand my knowledge as an IT professional.  
This script is a platform for me to experiment and learn from.  
Over time I will clean this script and add more redundancy to it, but for the time being please understand this is a WORK IN PROGRESS.  

If you have an issue, please create an issue in the tabs above and I will gladly look into it.  
PRs welcome.  

## Dependencies:
[iTextSharp 5.5.10](https://github.com/itext/itextsharp/)  
      I am under the impression that itextsharp version 5 can legally be distributed in open-source software.  
      The included iTextSharp5.5.10.dll file is from https://github.com/itext/itextsharp/releases/download/5.5.10/itextsharp-all-5.5.10.zip  
      This file is unmodified.  
      It is licensed as [AGPL](https://www.gnu.org/licenses/#AGPL) software.  
      https://github.com/itext/itextsharp/blob/develop/LICENSE.md  
      Please verify if your use case is not a violation of the iTextSharp license.
