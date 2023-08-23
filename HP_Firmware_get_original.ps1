#Created by Andrew Lawson


#Todo

#Potentially download pdf and automatically scrape it for the information needed. Could also provide scraped versions in a list.
#Use logic to determine the model_p1 and model_p2 versus being dependant on printers.csv
#Could insert firmware revision into a file for future usage/reference.



#For testing:
#$Printermodel = "M527"
#$Firmwareversion = "4.11.2.1"
#$Firmwarecode = "2411177_063753"
write-Host "`nThis script will prompt you for Printer model, Firmware Version, and Firmware Code number. `nThis information be aquired from Webjet, or from a PDF that this script will open. `n" -ForegroundColor White
write-Host ">>>>Check documentation prior to using!!<<<<" -BackgroundColor red
start-sleep 2
$Printermodel = Read-Host -Prompt 'Insert printer model'
$Firmwareversion = Read-Host -Prompt 'Insert firmware version'
$Firmware_short = $Firmwareversion -split "\." | Select-Object -First 1
#write-Host $Firmware_short

#Start getting info
try {
	$csv = Import-Csv -Path ./printers.csv -header Model,Model_p1,Model_p2,Part_num
}
catch
{
"`n>>Map the VIC HP Printer Dept folder to your OneDrive!<< `n`nGo to Teams>Warehouse on the top `nClick VIC WHSE Home then click the button ""Add shortcut to OneDrive"""
"`n`nClick CTRL C to exit this script"
#We should create the file, but this is a legacy script.
Write-Host "You are missing the printers.csv file"

start-sleep 500
Exit
}
}
if (($csv.Model).contains($Printermodel)) {
$model = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model
$model_p1 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p1
$model_p2 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p2
	} else {
    write-output "Printer model not found in printer.csv `nPlease verify your model selection `nIf accurate, please update printers.csv"
	#If its a E series, assume that it is a unique printer and that everything is the same?
	start-sleep 20
	Exit
    }



$get_changelog = Read-Host -Prompt 'Do you know the firmware revision code? type y if so.'
if ($get_changelog -ne "y") {
write-Host "Find a section in this PDF that contains the firmware version you want, and a firmware revision. `nWhen you find it, come back and click the Enter key."
start-sleep 5
$readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
start $readme_url
#start "http://ftp.hp.com/pub/softlib/software13/printers/LJM607_608_609/readme_ljM607_608_609_fs5.pdf"
pause
}


$Firmwarecode = Read-Host -Prompt 'Insert firmware revision. (Example: 2506116_037317)'


$url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/' + $model_p2 +'_fs' + $Firmwareversion + '_fw_' + $Firmwarecode + '.zip'
$file = $model_p2 +'_fs' + $Firmwareversion + '_fw_' + $Firmwarecode + '.zip'
try {
	write-Host "Downloading...."
	$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -URI $url -OutFile $file
	write-Host "Downloaded. Enjoy!"
	pause
	}
catch { 
"Error occured. Verify Firmware code and Firmware version matches `nIf you are certain everything is right, HP may not have this firmware exposed. `nContact Renee or the current HP rep"
start-sleep 20
Exit
}

