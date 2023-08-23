# HP_Firmware_Get

This script will allow you to download the vast majority of HP printer firmware directly from HP.
It guesses what format the printer uses for firmware, and store it locally in printers.csv if it is not already stored in there.
After it successfully finds the changelog PDF for your desired firmware, it will search the PDF using iTextSharp for various firmware versiona and codes.
Once it finds something that matches with your search parameters, it will display them for you to select which exact firmware you want.
It will then generate a direct download link from HP and download the printer firmware!


This is one of many projects I have been working on to expand my knowledge as an IT professional.
This script is a platform for me to experiment and learn from.
Over time I will clean this script and add more redundancy to it, but for the time being please understand this is a WORK IN PROGRESS.

If you have an issue, please create an issue in the tabs above and I will gladly look into it.
PRs welcome.
