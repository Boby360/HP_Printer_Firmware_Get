#Created by Andrew Lawson Cousineau


#todo:
#If the firmware version and reivision are on seperate pages, it breaks.
#Make manual PDF opening work. Download-Firmware-PDF isn't triggered if info is unknown.
#If unsuccessful at finding PDF, we should decrease firmware number to see if the firmware just doesn't exist for that printer.
#If PDF works, but firmware version can't be found, show other available versions.
#	Delete everything but the first number, then try, if fail, decrease number until version 3 has been attempted.
#Download iTextsharp and extract.

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
            Write-Host "Multiple versions matching '$DesiredVersion' found:"
            for ($i = 0; $i -lt $desiredVersions.Count; $i++) {
                Write-Host "$($i + 1). $($desiredVersions[$i])"
            }
            $selection = Read-Host "Enter the number corresponding to the desired version(NOT the version number)"
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
	
        # Generate URL and attempt to curl HTTP code 200 again
        $readme_url = 'https://ftp.hp.com/pub/softlib/software13/printers/' + $model_p1 + '/readme_' + $model_p2 + '_fs' + $Firmware_short + '.pdf'
	#try {
		$response = Invoke-WebRequest -Uri $readme_url -Method Head -ErrorAction SilentlyContinue
		
    if ($response.StatusCode -eq 200) {
        Write-Host "Got 200"
		#Open in browser
		$results = Check-for-iTextSharp -iTextSharpPath $iTextSharpPath
		#$results
		if (Check-for-iTextSharp -iTextSharpPath $iTextSharpPath) {
			#echo "Check-for-iTextSharp success"
			#$ProgressPreference = 'SilentlyContinue'; 
			$file = 'readme_' + $model + '_fs' + $Firmware_short + '.pdf'
			$OutFile = 'C:\Users\andre\Downloads\' + $file
			Invoke-WebRequest -Uri $readme_url -OutFile $OutFile
			#This is being tested:
			$result = Find-DesiredVersion -PdfFilePath $OutFile -DesiredVersion $Firmwareversion
			Write-Host "test"
			Write-Host $result
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
	Write-Host "After response"
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
$url

try {
	write-Host "Downloading...."
	$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -URI $url -OutFile $file
	write-Host "Downloaded. Enjoy!"
	if ($existing_csv -eq 'False') {
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
echo "If you have, or permit itextsharp.dll to be downloaded, it will completely automate the process."
echo "It is very common and well known open source file. I would deem it safe"

#write-Host "`nThis script will prompt you for Printer model, Firmware Version, and Firmware Code number. `nThis information be aquired from Webjet, or from a PDF that this script will open. `n" -ForegroundColor White
#write-Host ">>>>Check documentation prior to using!!<<<<" -BackgroundColor red
start-sleep 2
$Printermodel = Read-Host -Prompt 'Insert printer model'
$Firmwareversion = Read-Host -Prompt 'Insert firmware version'
$Firmware_short = $Firmwareversion -split "\." | Select-Object -First 1
#write-Host $Firmware_short

#Start getting info
try {
	$csv = Import-Csv -Path $printerscsv -header Model,Model_p1,Model_p2,Part_num
	$existing_csv = 'True'
}
catch
{
	$existing_csv = 'False'
	Write-Host "Unable to find printers.csv file. I will create one!"
}

if (($csv.Model).contains($Printermodel)) {
$model = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model
$model_p1 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p1
$model_p2 = ($csv.Where({[string]$_.Model -like "*$Printermodel"})).Model_p2

$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p2 -Firmware_short $Firmware_short -existing_csv $existing_csv
echo "after function"
#$firmware_pdf
if ($firmware_pdf -eq $false ) {
	echo "The CSV file has the wrong formatting"
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
    write-output "Printer model not found in printer.csv `n Lets try to figure it out."
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
	#	echo "No success!"
	#}
	} catch {}
	
	######Two printers share same firmware
	try {
	echo "Trying double wide"
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
		echo "No success!"
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
		echo "No success!"
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
		echo "No success!"
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
		echo "No success!"
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
		echo "No success!"
	}
	} catch {}
	

#####################################Three printers share same firmware
	try {
	Start-Sleep -Milliseconds 200

		echo $model + "3 wide beginning"
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
			echo "No success!"
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
				echo "No success!"
			}
			} catch {
			$model_p1 = $model_short + "40_50_60"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
						if ($firmware_pdf -ne $false ) {
							echo "Success!!"
							break
							echo "Did i break?"
						} else {
							echo "No success!"
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
				echo "No success!"
			}
			} catch {
			$model_p1 = $model_short + "40_50_60"
			$firmware_pdf = Download-Firmware-PDF -model_p1 $model_p1 -model_p2 $model_p1 -Firmware_short $Firmware_short
			if ($firmware_pdf -ne $false ) {
				echo "Success!!"
				break
				echo "Did i break?"
			} else {
				echo "No success!"
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
				echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
				}
		} catch {}

	Write-Host "Unsuccessful"
	echo "End of Continue"
	$bruteforce_fail = "y"
	break

	}#end of continue
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
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
					echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_2_black_cap -model_p2 $model_2_black_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_black_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					echo "No success!"
				}
		} catch {}
		
		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_colour_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					echo "No success!"
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
					echo "No success!"
				}
		} catch {}

		try {
		Start-Sleep -Milliseconds 200
		$firmware_pdf = Download-Firmware-PDF -model_p1 $model_3_colour_cap -model_p2 $model_3_colour_nocap -Firmware_short $Firmware_short
				if ($firmware_pdf -ne $false ) {
					echo "Success!!"
					break
				} else {
					echo "No success!"
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
					echo "No success!"
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
$FirmwareRevision = $firmware_pdf.Revision
$Firmwareversion = $firmware_pdf.Version

if ($FirmwareRevision -eq $null ){
$FirmwareRevision = Read-Host -Prompt 'Insert firmware revision. (Example: 2506116_037317)'
}

if ($model_colour_cap  -like "C*" ){
			#echo "Is colour!"
			$model_p1 = $model_colour_cap 
			$model_p2 = $model_colour_nocap
}

#Just incase

$Firmwareversion = $Firmwareversion.TrimEnd()
$FirmwareRevision = $FirmwareRevision.TrimEnd()
$file = $model_p2 +'_fs' + $Firmwareversion + '_fw_' + $FirmwareRevision + '.zip'
$Fileexists = Test-Path -Path ./$file -PathType Leaf
if ($Fileexists -eq 'True' ) {
	$download = Read-Host -Prompt 'File already exists locally. Do you want to redownload? type y if yes.'
	if ($download -eq "y") {
		write-host "yes"
		$firmware_pdf = Download-Firmware -model_p1 $model_p1 -model_p2 $model_p2 -FirmwareRevision $FirmwareRevision -Firmwareversion $Firmwareversion
	} else {
	Write-Host "File Already Exists. Not downloading."
	start-sleep 10
	break
	}
	} else {
	
$firmware_pdf = Download-Firmware -model_p1 $model_p1 -model_p2 $model_p2 -FirmwareRevision $FirmwareRevision -Firmwareversion $Firmwareversion
}
