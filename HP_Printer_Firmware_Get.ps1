#Created by Andrew Lawson


#Todo:

#If PDF exists, but requested version can not be found, prompt for opening PDF and allow manual user entry.
#Add 4 wide method, as E877 is 4 wide.
#Remove verbose from M series.
#Delete everything but the first number, then try, if fail, decrease number until version 3 has been attempted.
#M405 and M405 has no posted firmware on its page, and only a firmware update utility.
	#We could try and detect this if we were motivated.
	#Website contains 002_2322C
	#https://ftp.hp.com/pub/softlib/software13/FW_CPE_Commercial/M404-M405_8A/HP_LaserJet_Pro_M304_M305_M404_M405_series_FW_002_2322C.exe
	#Release notes link might not be easy to figure out though.
	#https://support.hp.com/soar-attachment/698/col91657-release_notes_2322c.html

#Store firmware links for each printer so users do not need to run the script.


#HP "New" "Managed" printer list.
#What can we figure out, and what can't we?
#What is possible to access, and what isn't?

####Not E series....
#4301-4303 uses different format. https://ftp.hp.com/pub/softlib/software13/FW_Commercial/4301_4303/HP_LaserJet_Pro_MFP_4301-4303_Firmware_release_6.12.1.12-202306030312.bdl
#Unsucessful: https://ftp.hp.com/pub/softlib/software13/FW_Commercial/4301_4303/HP_LaserJet_Pro_MFP_4301-4303_Firmware_release_6.12.0.2002-202304281920.bdl
#Release Notes: https://support.hp.com/soar-attachment/902/col118163-releasenotesfiles_6.12.1.12-202306030312.html
#Firmware Update Utility: https://ftp.hp.com/pub/softlib/software13/printers/4301-4303/HP_LaserJet_Pro_MFP_4301-4303_Firmware_6.12.1.12-202306030312.exe
#While release notes contains 6.12.0 and 6.12.1, we may not be able to figure out old firmware.

#E785** is E78523 and E78528.
#Max current firmware is 5.6.0.2 (May 21) using typical format.

#E87740-70
#Example firmware: https://ftp.hp.com/pub/softlib/software13/printers/E87740_50_60_70/E87740_50_60_70_fs5.6.0.2_fw_2506649_040426.zip
#Release Notes: http://ftp.hp.com/pub/softlib/software13/printers/E87740_50_60_70/readme_E87740_50_60_70_fs5.pdf

#M478-479
#No publicly posted firmware
#Release notes: https://support.hp.com/soar-attachment/433/col91648-releasenotesfiles_002_2322c.html (Does contain good amount of history)
#Firmware Update Utility: https://ftp.hp.com/pub/softlib/software13/FW_CPE_Commercial/M478-M479_MA/HP_Color_LaserJet_Pro_M478_M479_series_FW_002_2322C.exe

#E826** E82650 E82660 E82670
#Example firmware: https://ftp.hp.com/pub/softlib/software13/printers/E82650_60_70/E82650_60_70_fs5.6.0.2_fw_2506649_040423.zip
#Changelog: http://ftp.hp.com/pub/softlib/software13/printers/E82650_60_70/readme_E82650_60_70_fs5.pdf

#E731** E73130_E73135_E73140
#Example Firmware https://ftp.hp.com/pub/softlib/software13/printers/E73130_35_40/E73130_35_40_fs5.6.0.2_fw_2506649_040417.zip

#E786**
#Example FIrmware: https://ftp.hp.com/pub/softlib/software13/printers/E78625_30_35/E78625_30_35_fs5.6.0.2_fw_2506649_040420.zip
#Has both Firmware Readme: http://ftp.hp.com/pub/softlib/software13/printers/E78625_30_35/readme_E78625_30_35_fs5.pdf
#and release notes: https://support.hp.com/soar-attachment/81/col113680-readmefs5.html


#####Bugs:
#If unsuccessful at finding PDF, we should decrease and increase firmware number to see if the firmware just doesn't exist for that printer. (M506 max is 3.9.12)
	#If unsuccessful, error out saying we can't figure out printer firmware format, or printer model is not valid. (ex. E82640 does not exist)
	#If PDF works, but firmware version can't be found, show other available versions.
	#CSV file format "could" be wrong, but very unlikely

#If the firmware version and reivision are on seperate pages, it breaks.
	#Haven't experienced this in a few months, but it was an issue.(Need to find printer model that still does this)

#PW556 shows valid 5.5 firmware in PDF. Says "Success!!" but does not attempt to download firmware. (Is not in printers.csv)
#PW586 works fine. (is in printers.csv)

#####Notes:
#M680 4.11.2.3 does not exist on repo anymore, but 4.12.0.1 does.

#functions

function Check-for-iTextSharp {
    param (
        [string]$iTextSharpPath
    )

    if (Test-Path -Path $iTextSharpPath) {
        # The DLL is present
        #Write-Output "The DLL is found."

        # Load the iTextSharp assembly from the DLL file
        Add-Type -Path $iTextSharpPath

        return $true
    }
    else {
        # The iTextSharp module and DLL are not found
        Write-Output "Provide the path to the iTextSharp DLL."
		#From what I understand, this file can be distributed with open source software.
		#It has been added to the github
		#Write-Output "Do you want to download iTextSharp?"
		#https://github.com/itext/itextsharp/releases/download/5.5.11/itextsharp-all-5.5.11.zip
		#Inside file has other zip files. We want itextsharp-dll-core.zip
		#Core contains itextsharp.dll and iTextSharp.xml
		#Invoke-WebRequest -URI https://github.com/itext/itextsharp/releases/download/5.5.11/itextsharp-all-5.5.11.zip -OutFile ./itextsharp/itextsharp-all-5.5.11.zip
		#Expand-Archive -LiteralPath ./itextsharp/itextsharp-dll-core.zip" -DestinationPath "./itextsharp/"
		#Expand-Archive -LiteralPath ./itextsharp/itextsharp-all-5.5.11.zip" -DestinationPath "./itextsharp/"
		#Add-Type -Path ./itextsharp/itextsharp.dll
        return $false
    }
}

function Find-DesiredVersion {
    param (
        [string]$PdfFilePath,
        [string]$DesiredVersion
    )

    # Add the iTextSharp DLL
#    Add-Type -Path "C:\Users\andre\Documents\PowerShell_test\VIHA\itextsharp.dll"

    # Create a reader object to extract text from the PDF
    $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $PdfFilePath

    # Initialize variables
    $versions = @()
    $revisions = @()
    $previousLine = ""
    $currentLine = ""
	
	try {
    # Iterate through each page in the PDF and extract the text
    for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
        $strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
        $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy)

        $lines = $currentText -split "`r?`n"

        foreach ($line in $lines) {
            $currentLine = $line

            if ($currentLine -match "HP Fleet Bundle Version Designation: FutureSmart Bundle Version") {
                $revisions += $previousLine -replace "Firmware Revision:\s+", ""
                $versions += $currentLine -replace "HP Fleet Bundle Version Designation: FutureSmart Bundle Version\s+", ""
            }

            $previousLine = $currentLine
        }
    }

    # Close the reader object
    $reader.Close()
	} catch {
		#If itextsharp is local, but windows is blocking it. (Had it happen on my Windows 10 desktop)
		Write-Host "Its possible that your itextsharp dll file is being blocked by windows. \n If you right click it, and go to properties, at the very bottom there is a unblock button."
		Write-Host "Please do this to use this functionality, or delete the dll file to use without it."
		#Running without it may or may not work. Will need to verify.
		Start-Sleep -seconds 15
		Exit
		
	}
    # Remove duplicates from versions and revisions
    $versions = $versions | Select-Object -Unique
    $revisions = $revisions | Select-Object -Unique

    # Store the unique versions in a separate variable
    $uniqueVersions = $versions | Select-Object -Unique

    # Store the unique revisions in a separate variable
    $uniqueRevisions = $revisions | Select-Object -Unique

    # Find the versions that match the desired version using regular expressions
    $desiredVersions = $uniqueVersions | Where-Object { $_ -match [regex]::Escape($DesiredVersion) }

    # Check if any desired versions are found
    if ($desiredVersions.Count -gt 0) {
        # Prompt the user to select the desired version if there are multiple matches
        if ($desiredVersions.Count -gt 1) {
            Write-Host "Multiple FS $DesiredVersion versions found:"
            for ($i = 0; $i -lt $desiredVersions.Count; $i++) {
                Write-Host "$($i + 1)." -ForegroundColor green -NoNewline
				Write-Host " $($desiredVersions[$i])"
            }
            $selection = Read-Host "Enter the green number to select the firmware version"
            $index = $selection - 1
        } else {
            $index = 0
        }

        # Check if the selected index is within the valid range
        if ($index -ge 0 -and $index -lt $desiredVersions.Count) {
            $version = $desiredVersions[$index]
            $revision = $uniqueRevisions[[array]::IndexOf($uniqueVersions, $version)]

            # Return the desired version and its corresponding revision
            return @{
                "Revision" = $revision
				"Version" = $version
            }
        }
    }

    return "Desired Version not found."
}
function Find-DesiredVersion_test {
    param (
        [string]$PdfFilePath,
        [string]$DesiredVersion
    )
	# Create a reader object to extract text from the PDF
	$reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $PdfFilePath

	# Initialize variables
	$combinedText = ""
	$versions = @()
	$revisions = @()

	# Iterate through each page in the PDF and extract text
	for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
		$strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
		$currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy)

		# Add each line as a new line to the combined text
		$combinedText += $currentText
		
		#The output format is as it is in the PDF.
		#Each important thing has its own line already.
		#But probably not in the format needed?
	}

	# Close the reader object
	$reader.Close()

    # Process each line
    foreach ($line in $combinedText) {
		#I made this $combinedText from $lines.
		#Line wasnt defined anywhere before, pretty sure.
		
        # if ($line -match "Firmware Revision:\s+(.+)") {
            # $revisions += $Matches[1]
        # }
		#Look at this next?
		
		
        if ($line -match "HP Fleet Bundle Version Designation: FutureSmart Bundle Version") {
            $versions += $line -replace "HP Fleet Bundle Version Designation: FutureSmart Bundle Version\s+", ""
        }
    }

    # Remove duplicates from versions and revisions
    #$versions = $versions | Select-Object -Unique
    #$revisions = $revisions | Select-Object -Unique

    # Find the versions that match the desired version using regular expressions
    $desiredVersions = $versions | Where-Object { $_ -match [regex]::Escape($DesiredVersion) }

    # Check if any desired versions are found
    if ($desiredVersions.Count -gt 0) {
        # Prompt the user to select the desired version if there are multiple matches
        if ($desiredVersions.Count -gt 1) {
            Write-Host "Multiple versions matching '$DesiredVersion' found:"
            for ($i = 0; $i -lt $desiredVersions.Count; $i++) {
                Write-Host "$($i + 1). $($desiredVersions[$i])"
            }
            $selection = Read-Host "Enter the number corresponding to the desired version (NOT the version number)"
            $index = $selection - 1
        } else {
            $index = 0
        }

        # Check if the selected index is within the valid range
        if ($index -ge 0 -and $index -lt $desiredVersions.Count) {
			
            $version = $desiredVersions[$index]
			Write-Host "revisions test output:"
			Write-Host $revisions
            $revision = $revisions[[array]::IndexOf($versions, $version)]
			Write-Host "revision test output:"
			Write-Host $revision
			Write-Host "versions test output:"
			Write-Host $versions
			Write-Host "version test output:"
			Write-Host $version
			#$revision
			#echo "revisions"
			#$revisions
			#echo "versions"
			#echo $desiredVersions
            # Return the desired version and its corresponding revision
            return @{
                "Revision" = $revision
                "Version" = $version
            }
        }
    }

    return "Desired Version not found."
}

function Download-Firmware-PDF {
    param (
        [string]$model_p1,
        [string]$model_p2,
		[string]$Firmware_short
    )
	$Firmware_short_orig = $Firmware_short
		try {
		Start-Sleep -milliseconds 500 #Reduce HP spam
        # Try provided firmware number
        $readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
		$response = Invoke-WebRequest -Uri $readme_url -Method Head -ErrorAction SilentlyContinue
		write-host good
		write-host $firmware_short
		} catch {
		try {
		Start-Sleep -milliseconds 500 #Reduce HP spam
		# Try one firmware version lower if provided fails.
		$Firmware_short_minus = [int]$Firmware_short_orig
		$Firmware_short_minus--
		$readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short_minus + '.pdf'
		$response = Invoke-WebRequest -Uri $readme_url -Method Head -ErrorAction SilentlyContinue
		$Firmware_short = [String]$Firmware_short_minus
		$Firmwareversion = $Firmware_short
		write-host "`r`nThis printer does not have any detected firmware for FS $Firmware_short_orig `r`nhowever it does for FS $Firmware_short. Select a version below or hold ctrl+c to exit.`r`n" -ForegroundColor yellow
		} catch {
		try {
		Start-Sleep -milliseconds 500 #Reduce HP spam
		#try one firmware above if provided fails.
		$Firmware_short_plus = [int]$Firmware_short
		$Firmware_short_plus++
		$readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short_plus + '.pdf'
		$response = Invoke-WebRequest -Uri $readme_url -Method Head -ErrorAction SilentlyContinue
		$fixed_firmware_short = $Firmware_short
		$Firmware_short = [String]$Firmware_short_plus
		$Firmwareversion = $Firmware_short
		write-host "`r`nThis printer does not have any detected firmware for FS $Firmware_short_orig `r`nhowever it does for FS $Firmware_short. Select a version below or hold ctrl+c to exit.`r`n" -ForegroundColor yellow
		} catch {}
		}
		}
		
    if ($response.StatusCode -eq 200) {
        #Write-Host "Got 200"
		#Open in browser
		$results = Check-for-iTextSharp -iTextSharpPath $iTextSharpPath
		#$results
		if (Check-for-iTextSharp -iTextSharpPath $iTextSharpPath) {
			#echo "Check-for-iTextSharp success"
			#$ProgressPreference = 'SilentlyContinue'; 
			#$file = 'readme_' + $model + '_fs' + $Firmware_short + '.pdf'
			$OutFile = $PSScriptRoot + '\readme_' + $model + '_fs' + $Firmware_short + '.pdf'
			Invoke-WebRequest -Uri $readme_url -OutFile $OutFile
			#This is being tested:
			#if ($fixed_firmware_short -eq 0) {
			$result = Find-DesiredVersion -PdfFilePath $OutFile -DesiredVersion $Firmwareversion
			#} else {
			#$result = Find-DesiredVersion -PdfFilePath $OutFile -DesiredVersion $Firmware_short
			#}
			#Write-Host "Spitting out stuff"
			#echo "out of function"
			# Extract the version and revision from the output
			$Firmwareversion = $result.Version
			$FirmwareRevision = $result.Revision
			return @{
                "Revision" = $result.Revision
				"Version" = $result.Version
            }
		} else {
		$known_revision = Read-Host -Prompt 'Do you know the firmware revision code? type n if no, y if you do.'
		if ($known_revision -eq "n" ) {
		Write-Host "Opening a PDF file for you to find."
		start $readme_url
		}
		}
	} else {
	#Write-Host "After response"
		#Try reducing Firmware_short from default, then try increasing Firmware_short.
		#Firmware_short will just be, 3, 4, 5
		
		#$Firmware_short_minus = [int]$Firmware_short		
		#$Firmware_short_minus--
		#Write-Host $Firmware_short_minus
		#If successful, $Firmware_short = [String]$Firmware_short_minus
		
		#$Firmware_short_plus = [int]$Firmware_short
		#$Firmware_short_plus_type = $Firmware_short_plus.GetType()
		#Write-Host "The variable type is: $Firmware_short_plus_type"
		
		#$Firmware_short_plus++
		#Write-Host $Firmware_short_plus
		#If successful, $Firmware_short = [String]$Firmware_short_plus
		#} catch {
		return $false
	}
Write-Host "End of function"
}

Function Manualy-Aquire-Info {

write-Host "Find a section in this PDF that contains the firmware version you want, and a firmware revision. `nWhen you find it, come back and click the Enter key."
start-sleep 5
$readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
start $readme_url
#start "http://ftp.hp.com/pub/softlib/software13/printers/LJM607_608_609/readme_ljM607_608_609_fs5.pdf"
}

function Download-Firmware {
	param (
        [string]$model_p1,
        [string]$model_p2,
		[string]$FirmwareRevision,
		[string]$Firmwareverision,
		[string]$existing_csv
    )
$url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/' + $model_p2 +'_fs' + $Firmwareversion + '_fw_' + $FirmwareRevision + '.zip'


try {
	write-Host "Trying to download from this link"
	write-Host $url
	$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -URI $url -OutFile $file
	write-Host "Downloaded. Enjoy!"
	#Why is there a "Please Enter to continue...:" between these two?
	#This is placed here so if the download fails, it will skip it, right?
	if ($exists_within_csv -eq 'False') {
	Save-to-CSV -model $model -model_p1 $model_p1 -model_p2 $model_p2 -printercsv $printercsv
	}
	start-sleep 15
	break
	}
catch { 
"Error occured. Verify Firmware code and Firmware version matches `nIf you are certain everything is right, HP may not have this firmware exposed. `nContact Renee or the current HP rep"
start-sleep 20
Exit
}
}

function Save-to-CSV {
	param (
        [string]$model_p1,
        [string]$model_p2,
		[string]$model,
		[string]$printercsv
    )
	echo "Insering missing entry into the CSV"
	$insert_to_csv = [PSCustomObject]@{ Model = $model; Model_p1 = $model_p1; Model_p2 = $model_p2; Part_num = "" }
	$insert_to_csv | Export-Csv -Path $printerscsv -Append -NoTypeInformation -Encoding UTF8 -Delimiter ','
	pause
}
##################################################################################################################################
##################################################################################################################################




#Todo

#Use method in VIHA form populator to download printers.csv from sharepoint.
#Potentially download pdf and automatically scrape it for the information needed. Could also provide scraped versions in a list.
#Use logic to determine the model_p1 and model_p2 versus being dependant on printers.csv
#Could insert firmware revision into a file for future usage/reference.

$iTextSharpPath = ".\itextsharp5.5.11.0.dll"
$printerscsv = ".\printers.csv"
#For testing:
#$Printermodel = "M527"
#$Firmwareversion = "4.11.2.1"
#$FirmwareRevision = "2411177_063753"
echo "This is Andrew Lawson's HP Printer Firmware downloader script"
echo "Enjoy!"

#write-Host "`nThis script will prompt you for Printer model, Firmware Version, and Firmware Code number. `nThis information be aquired from Webjet, or from a PDF that this script will open. `n" -ForegroundColor White
#write-Host ">>>>Check documentation prior to using!!<<<<" -BackgroundColor red
start-sleep 2
$Printermodel = Read-Host -Prompt 'Insert printer model'
$Printermodel = $Printermodel.ToUpper()
$Firmwareversion = Read-Host -Prompt 'Insert firmware version'
$Firmware_short = $Firmwareversion -split "\." | Select-Object -First 1
#write-Host $Firmware_short

#Trim any spaces user may have added.
$Printermodel = $Printermodel.Trim()
$Firmwareversion = $Firmwareversion.Trim()

#Start getting info
try {
	$csv = Import-Csv -Path $printerscsv -header Model,Model_p1,Model_p2,Part_num
	$existing_csv = 'True'
} catch {
	Write-Host "Unable to find printers.csv file. I will download the newest version from Github."
	$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -URI "https://raw.githubusercontent.com/Boby360/HP_Printer_Firmware_Get/main/printers.csv" -OutFile .\printers.csv
		try {
		$csv = Import-Csv -Path $printerscsv -header Model,Model_p1,Model_p2,Part_num
		$existing_csv = 'True'
		} catch {
		Write-Host "Attempted to download new printers.csv file but still unable to import."
		$existing_csv = 'False'
		}
	
}

if (($csv.Model).contains($Printermodel)) {
$model = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model
$model_p1 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p1
$model_p2 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p2

$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p2 -Firmware_short $Firmware_short -existing_csv $existing_csv
#echo "after function"
#$firmware_pdf
if ($firmware_pdf -eq $false ) {
	echo "The CSV file has the wrong formatting, or the printer only supports significantly higher/lower firmware version."
	break
	#Could make this delete the line.
}
# $readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/' + 'readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
# $response = Invoke-WebRequest -Uri $readme_url -Method Head -ErrorAction SilentlyContinue

    # Check if the first attempt succeeded
    # if ($response.StatusCode -eq 200) {
        # Write-Host "First attempt succeeded!"
		# Open in browser
		# $results = Check-for-iTextSharp -iTextSharpPath $iTextSharpPath
		# $results
		# if (Check-for-iTextSharp -iTextSharpPath $iTextSharpPath) {
			# echo "Check-for-iTextSharp success"
			# $ProgressPreference = 'SilentlyContinue'; 
			# $file = 'readme_' + $model + '_fs' + $Firmware_short + '.pdf'
			# $OutFile = 'C:\Users\andre\Downloads\' + $file
			# Invoke-WebRequest -Uri $readme_url -OutFile $OutFile
			# This is being tested:
			# $result = Find-DesiredVersion -PdfFilePath $OutFile -DesiredVersion $Firmwareversion
			# echo "out of function"
			# Extract the version and revision from the output
			# $Firmwareversion = $result.Version
			# $FirmwareRevision = $result.Revision
		# } else {
		# $known_revision = Read-Host -Prompt 'Do you know the firmware revision code? type n if no, y if you do.'
		# if ($known_revision -eq "n" ) {
		# echo "Opening a PDF file for you to find."
		# start $readme_url
		# }
		# }
	# }
#Add logic so we try exactly what we need, given we already know the other information.
#Once we figure it out, should we store it locally??



	 } else {
    write-output "Printer model not found in printers.csv `n Lets try to figure it out."
	#Lets try downloading a updated version.
	#https://raw.githubusercontent.com/Boby360/HP_Printer_Firmware_Get/main/printers.csv
	#Should we combine user aquired data, if so, compare between what is local and what is remote.
	#Could use a hash to verify if it has changed, or download and compare file lengths, or do an actual comparison.
	#If we do a hash, if we force user to restart script after downloading, we won't be in an endless loop, if the comparison is in this location.
	#Could compare hash and length, but no guarantee.
	
	$exists_within_csv = 'False'
	#this includes all of the searching stuff. Will be changed when we add more logic.


##################################################################################################################################
##################################################################################################################################
$model = $Printermodel
[int]$five = 5
[int]$ten = 10
$continue = "1"
if ($model -like "E*") {
	while ($continue -eq "1") {
	echo "Is E Model"
	
	######Single printer uses firmware
	try {
	Start-Sleep -Milliseconds 200
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model -model_p2 $model -Firmware_short $Firmware_short
	#echo "I made it out!"
	if ($firmware_pdf -ne $false ) {
		#echo "Success!!"
		$model_p1 = $model
		break
		#echo "Did i break?"
	} 
	#else {
	#	#echo "No success!"
	#}
	} catch {}
	
	######Two printers share same firmware
	try {
	echo "Looking for 2 wide firmware"
	# Extract the last two digits of the model
    $lastTwoDigits = $model.Substring($model.Length - 2)
	#echo "last two digits" + $lastTwoDigits
	
	#Cut off the last two digits
	$model_short = $model.Substring(0,$model.Length-2)
	
	Start-Sleep -Milliseconds 200
		
	if($lastTwoDigits -eq 50 -or $lastTwoDigits -eq 60 ) {
	$model_p1 = $model_short + "50_60"
	
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
	#When this failed, it did't spit out anything.
	if ($firmware_pdf -ne $false ) {
		echo "Success!!"
		break
	} else {
		#echo "No success!"
	}
	}
	} catch {}

	
	#Add 5 
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + $lastTwoDigits + "_" + [string]([int]$lastTwoDigits + $five)
	
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
	if ($firmware_pdf -ne $false ) {
		echo "Success!!"
		break
	} else {
		#echo "No success!"
	}
	} catch {}
	
	#Add 10
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + $lastTwoDigits + "_" + [string]([int]$lastTwoDigits + $ten)

	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
	if ($firmware_pdf -ne $false ) {
		echo "Success!!"
		break
	} else {
		#echo "No success!"
	}
	} catch {}
	
	#Minus 5
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $five) + "_" +  $lastTwoDigits

	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
	if ($firmware_pdf -ne $false ) {
		echo "Success!!"
		break
	} else {
		#echo "No success!"
	}
	} catch {}
	
	#Minus 10
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $ten) + "_" +  $lastTwoDigits
	
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
	if ($firmware_pdf -ne $false ) {
		echo "Success!!"
		break
	} else {
		#echo "No success!"
	}
	} catch {}
	

#####################################Three printers share same firmware
	try {
	Start-Sleep -Milliseconds 200

		echo "Looking for 3 wide Firmware"
		#Common Logic:
		if($lastTwoDigits -eq 25) {
		$model_p1 = $model_short + "25_30_35"
		}
		if($lastTwoDigits -eq 30) {
		$model_p1 = $model_short + "25_30_35"
		}
		if($lastTwoDigits -eq 35) {
		echo $model
		$model_p1 = $model_short + "25_30_35"
		}
		#30_35_40 now exists.
		#E731**
		if($lastTwoDigits -eq 40) {
		$model_p1 = $model_short + "40_50_60"
		}
		if($lastTwoDigits -eq 55) {
		$model_p1 = $model_short + "55_65_75"
		}
		if($lastTwoDigits -eq 65) {
		$model_p1 = $model_short + "55_65_75"
		}
		if($lastTwoDigits -eq 70) {
		$model_p1 = $model_short + "50_60_70"
		}
		if($lastTwoDigits -eq 75) {
		$model_p1 = $model_short + "55_65_75"
		}

		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
		if ($firmware_pdf -ne $false ) {
			echo "Success!!"
			break
			echo "Did i break?"
		} else {
			#echo "No success!"
		}
		
		if($lastTwoDigits -eq 50) {
			echo "Last two are 50"
			try {
			Start-Sleep -Milliseconds 200
			$model_p1 = $model + "50_60_70"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
			if ($firmware_pdf -ne $false ) {
				echo "Success!!"
				break
				echo "Did i break?"
			} else {
				#echo "No success!"
			}
			} catch {
			$model_p1 = $model_short + "40_50_60"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
						if ($firmware_pdf -ne $false ) {
							echo "Success!!"
							break
							echo "Did i break?"
						} else {
							#echo "No success!"
						}
						}
		}
		

		if($lastTwoDigits -eq 60) {
			try {
			Start-Sleep -Milliseconds 200
			$model_p1 = $model_short + "50_60_70"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
			if ($firmware_pdf -ne $false ) {
				echo "Success!!"
				break
				echo "Did i break?"
			} else {
				#echo "No success!"
			}
			} catch {
			$model_p1 = $model_short + "40_50_60"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
			if ($firmware_pdf -ne $false ) {
				echo "Success!!"
				break
				echo "Did i break?"
			} else {
				#echo "No success!"
			}
			}
		}
		
        } catch {}
	#Try to detect unexpected format:
	#Add 5 
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + $lastTwoDigits + "_" + [string]([int]$lastTwoDigits + $five) + "_" + [string]([int]$lastTwoDigits + $ten)
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
			if ($firmware_pdf -ne $false ) {
				echo "Success!!"
				break
				echo "Did i break?"
			} else {
				#echo "No success!"
			}
	
	} catch {}
	
	#Add 10
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + $lastTwoDigits + "_" + [string]([int]$lastTwoDigits + $ten) + "_" + [string]([int]$lastTwoDigits + $ten + $ten)
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
	
	#Minus 5
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $ten) + "_" + [string]([int]$lastTwoDigits - $five) + "_" +  $lastTwoDigits
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
	
	#Minus 10
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $ten - $ten) + "_" + [string]([int]$lastTwoDigits - $ten) + "_" +  $lastTwoDigits
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		
	#Middle 10
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $ten) + "_" + $lastTwoDigits + "_" + [string]([int]$lastTwoDigits + $ten)
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
	#Middle 5
	try {
	Start-Sleep -Milliseconds 200
	$model_p1 = $model_short + [string]([int]$lastTwoDigits - $ten - $ten) + "_"  + [string]([int]$lastTwoDigits - $ten) + "_" +  $lastTwoDigits
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}


#####################################four printers share same firmware
	try {
	Start-Sleep -Milliseconds 200

		echo "Looking for 4 wide firmware"
		#Common Logic:
		#Right now only know of the E877 with quad 
		#write-host $Printermodel
		if($Printermodel -like "E877*") {
		$model_p1 = $model_short + "40_50_60_70"
		}

		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
		if ($firmware_pdf -ne $false ) {
			echo "Success!!"
			break
			echo "Did i break?"
		} else {
			#echo "No success!"
		}
		
		
	Write-Host "Unsuccessful"
	echo "End of Continue"
	$bruteforce_fail = "y"
	break

	} catch {}
	}	#end of continue
	$model_p2 = $model_p1
	echo "End of E"
}#end of  E



if ($model -like "PW*" -Or $model -like "pw*") {
	echo "PW model"
	######Single printer uses firmware
	#$model_lower = $model.Substring(0,1).ToLower()
	$model_p1 = $model.ToUpper()
	$model_p2 = $model.ToLower() #<-------------- Easier Method!!
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p2 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}



}

if ($model -like "M*" -Or $model -like "m*") {
while ($continue -eq 1) {
	echo "M model"
	######Single printer uses firmware
    #M series has a captical and non captial format.
	#If M is colour, its also CLJM not LJM
	
	#Force capitcal
	$model = $model.ToUpper()
	
	[int]$model_no_letter = $model.substring(1)
	$model_no_letter_inc = $model_no_letter
	$model_no_letter_dec = $model_no_letter
	
	#Increment
	
	echo "model2:"
	$model_no_letter_inc++
	$model_2_no_letter=$model_no_letter_inc
	$model_2_no_letter
	
	echo "model3:"
	$model_no_letter_inc++
	$model_3_no_letter=$model_no_letter_inc
	$model_3_no_letter
	
	#Decrement
	$model_no_letter_dec--
	$model_0_no_letter = $model_no_letter_dec
	echo "model 0:"
	$model_0_no_letter
	$model_no_letter_dec--
	$model_neg1_no_letter = $model_no_letter_dec
	echo "model -1"
	$model_neg1_no_letter

	$model_no_last_digit = $model.substring(0,$model.Length-1)
	
	echo "test"
	$model_no_letter
	$model_no_last_digit
	$model_colour_short = "CLJ"
	$model_colour_short_ncap = $model_colour_short.ToLower()
	$model_colour_cap = $model_colour_short + $model
	$model_colour_nocap = $model_colour_short_ncap + $model
	
	
	$model_black_short = "LJ"
	$model_black_short_nocap = $model_black_short.ToLower()
	$model_black_cap = $model_black_short + $model_no_letter
	$model_black_nocap = $model_black_short_nocap + $model
	
	#Try colour
	try {
	Start-Sleep -Milliseconds 200
	$model_colour_cap
	$model_colour_nocap
	$Firmware_short
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_colour_cap -model_p2 $model_colour_nocap -Firmware_short $Firmware_short
	#$firmware_pdf
	#write-host $firmware_pdf
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
	#$firmware_pdf
	#write-host $firmware_pdf
		
	try {
	Start-Sleep -Milliseconds 200
	#Try black
	$firmware_pdf = Download-Firmware-PDF -model_p1 $model_black_cap -model_p2 $model_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
	
########################Two printers share same firmware
		echo "Two printers"
		
        # Extract the last digit of the model
        $lastDigit = $model.Substring($model.Length-1)
		echo $lastDigit
		
		#Try increment method.
		
		$model_no_letter
		#$array = 1,2
		echo $model_no_last_digit
		$model_2 = $model + "_" + [string]$model_2_no_letter
		$model_3 = $model + "_" + [string]$model_2_no_letter + "_" + [string]$model_3_no_letter
		echo $model_2
		
		
	#$model_colour_short = "CLJ"
	#$model_colour_short_ncap = $model_colour_short.Substring(0,3).ToLower()
	$model_2_colour_cap = $model_colour_short + $model_2
	$model_2_colour_nocap = $model_colour_short_ncap + $model_2
	$model_3_colour_cap = $model_colour_short + $model_3
	$model_3_colour_nocap = $model_colour_short_ncap + $model_3
	
	#$model_black_short = "LJ"
	#$model_black_short_nocap = $model_black_short.Substring(0,2).ToLower()
	$model_2_black_cap = $model_black_short + $model_2
	$model_2_black_nocap = $model_black_short_nocap + $model_2
	$model_3_black_cap = $model_black_short + $model_3
	$model_3_black_nocap = $model_black_short_nocap + $model_3
		
		try {
		Start-Sleep -Milliseconds 200
        # Generate URL and attempt to curl HTTP code 200 again
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_2_colour_cap -model_p2 $model_2_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
        # Generate URL and attempt to curl HTTP code 200 again
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_2_black_cap -model_p2 $model_2_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
        # Generate URL and attempt to curl HTTP code 200 again
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_black_cap -model_p2 $model_3_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		
		
		echo "Try Decrement Method"
		#Try Decremnt Method 
		$model_2 = "M" + [string]$model_0_no_letter + "_" + $model_no_letter
		$model_3 = "M" + [string]$model_neg1_no_letter + [string]$model_0_no_letter + "_" + $model_no_letter
		
		echo $model_2
		$model_2_colour_cap = $model_colour_short + $model_2
		$model_2_colour_nocap = $model_colour_short_ncap + $model_2
		$model_2_black_cap = $model_black_short + $model_2
		$model_2_black_nocap = $model_black_short_nocap + $model_2

		$model_3_colour_cap = $model_colour_short + $model_3
		$model_3_colour_nocap = $model_colour_short_ncap + $model_3
		$model_3_black_cap = $model_black_short + $model_3
		$model_3_black_nocap = $model_black_short_nocap + $model_3
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_2_colour_cap -model_p2 $model_2_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_2_black_cap -model_p2 $model_2_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_black_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_colour_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		##########Try Middle method
		echo "try middle method"
		$model_3 = "M" + [string]$model_0_no_letter + "_" + $model_no_letter + "_" + [string]$model_2_no_letter
		$model_3_colour_cap = $model_colour_short + $model_3
		$model_3_colour_nocap = $model_colour_short_ncap + $model_3
		$model_3_black_cap = $model_black_short + $model_3
		$model_3_black_nocap = $model_black_short_nocap + $model_3
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_black_cap -model_p2 $model_3_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}

		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_colour_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}
		
		
		try {
		Start-Sleep -Milliseconds 200
		#try PW, just incase
			#$model_no_letter = $model.substring(1)
			#$model_no_letter
			$model_p1 = "PW" + $model_no_letter
			$model_p2 = $model.ToLower() #<-------------- Easier Method!!
			
			#$model_3 = "PW" + [string]$model_neg1_no_letter + [string]$model_0_no_letter + "_" + $model_no_letter
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p2 -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					#echo "No success!"
				}
		} catch {}

		break
}




 }

		#echo "Insering missing entry into the CSV"
		#$insert_to_csv = [PSCustomObject]@{ Model = $model; Model_p1 = $model; Model_p2 = $model; Part_num = "" }
		#$insert_to_csv | Export-Csv -Path $printerscsv -Append -NoTypeInformation -Encoding UTF8 -Delimiter ','
 }#End of else for CSV check
##################################################################################################################################
##################################################################################################################################

if ($bruteforce_fail -eq "y" -Or $known_revision -eq "n") {
Function Manualy-Aquire-Info{

write-Host "Find a section in this PDF that contains the firmware version you want, and a firmware revision. `nWhen you find it, come back and click the Enter key."
start-sleep 5
$readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
start $readme_url
#start "http://ftp.hp.com/pub/softlib/software13/printers/LJM607_608_609/readme_ljM607_608_609_fs5.pdf"
}
}
#Converting to string as TrimEnd did not work.
$FirmwareRevision = [String]$firmware_pdf.Revision
$Firmwareversion = [String]$firmware_pdf.Version

if ($FirmwareRevision -eq $null ){
$FirmwareRevision = Read-Host -Prompt 'Insert firmware revision. (Example: 2506116_037317)'
}

if ($model_colour_cap  -like "C*" ){
			#echo "Is colour!"
			$model_p1 = $model_colour_cap 
			$model_p2 = $model_colour_nocap
}

#Just incase
#Trim just removes extra spaces.
$Firmwareversion = $Firmwareversion.Trim()
$FirmwareRevision = $FirmwareRevision.Trim()
$file = $model_p2 +'_fs' + $Firmwareversion + '_fw_' + $FirmwareRevision + '.zip'
$Fileexists = Test-Path -Path ./$file -PathType Leaf
if ($Fileexists -eq 'True' ) {
	$download = Read-Host -Prompt 'File already exists locally. Do you want to redownload? type y if yes.'
	if ($download -eq "y") {
		$firmware_pdf = Download-Firmware -model_p1 $model_p1 -model_p2 $model_p2 -FirmwareRevision $FirmwareRevision -Firmwareversion $Firmwareversion
	} else {
	Write-Host "File Already Exists. Not downloading."
	start-sleep 10
	break
	}
	} else {
	
$firmware_pdf = Download-Firmware -model_p1 $model_p1 -model_p2 $model_p2 -FirmwareRevision $FirmwareRevision -Firmwareversion $Firmwareversion
}
