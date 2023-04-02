######################
# LOAD PREREQUISITES #
######################

Write-Host "Loading Prerequisites..." -ForegroundColor Green

try {
    # Load the System.Windows.Forms assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Install ImportExcel in CurrentUser scope to not trip UAC flags 
    # Check if the ImportExcel module is already installed
    if (-not (Get-InstalledModule -Name ImportExcel -ErrorAction SilentlyContinue)) {
        # Install the ImportExcel module if it's not already installed
        Install-Module -Name ImportExcel -Scope CurrentUser
    }
}
catch {
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    Write-Host "Error at Powershell line: $lineNumber" -ForegroundColor Red
    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}



###############
# File Picker #
###############

# Create a file picker dialog box
$filePicker = New-Object System.Windows.Forms.OpenFileDialog
$filePicker.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
$filePicker.Multiselect = $false

# Display the file picker dialog box and get the selected file path
if ($filePicker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePath = $filePicker.FileName
}
else {
    Write-Host "File selection canceled."
    return
}

# Import the Excel document using the selected file path
try {
    # Filter out blank rows using the Where-Object cmdlet
    $excelFile = Import-Excel -Path $filePath
    #$csvFile = $excelFile #| Where-Object { $_.PSObject.Properties.Value -notcontains "" }
    $csvFile = $excelFile | Where-Object { ($_.PSObject.Properties.Value | ForEach-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0 }
}
catch {
    Write-Host "Error importing Excel file: $_" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}


##############################
# Date Convertsion Functions #
##############################

function convert-Date {
    param($inDate)
    
    if ($inDate -like "Jan*") { return "01" }
    if ($inDate -like "Feb*") { return "02" }
    if ($inDate -like "Mar*") { return "03" }
    if ($inDate -like "Apr*") { return "04" }
    if ($inDate -like "May") { return "05" }
    if ($inDate -like "Jun*") { return "06" }
    if ($inDate -like "Jul*") { return "07" }
    if ($inDate -like "Aug*") { return "08" }
    if ($inDate -like "Sep*") { return "09" }
    if ($inDate -like "Oct*") { return "10" }
    if ($inDate -like "Nov*") { return "11" }
    if ($inDate -like "Dec*") { return "12" }
    
}

function endOfDecade {
    
    $year = $args[0]
    $year = $year - 0
    $year = $year + 9 
    
    return $year
    
}

    
function codedDate($i) {
    # 1 October-December, 2001
    if ( $i -eq 'undated') {
        return "0000/0000"

    }
    elseif ($i -match "([a-zA-Z]+).?\s*-\s*([a-zA-Z]+)\s*.?\s*(\d{4})") {
        $year = $matches[3]; $month = convert-Date $matches[1]; $month2 = convert-Date $matches[2];
        return $($year + "-" + $month + "/" + $year + "-" + $month2) 
    }
    # 2 January 24, 2014 - February 24, 2018 and a few variations Done
    elseif ($i -match "([a-zA-Z]+)\s*,?\s*\b(\d{1,2})?(?:[nN][dD]|[sS][tT]|[rR][dD]|[tT][hH])?\b\s*,?\s*(\d{4})?(\s*.{1,2}\b\s*([a-zA-Z]+)\s*,?\s*\b(\d{1,2})?(?:[nN][dD]|[sS][tT]|[rR][dD]|[tT][hH])?\b\s*,?\s*(\d{4})?)" -and $i -notlike "*undated*") {
        $month = $matches[1]; $month2 = $matches[5];
        if ($month) { $month = convert-Date $month; $month = "-" + $month; } if ($matches[2]) { $day = $matches[2]; if ($day.Length -lt 2) { $day = ($day).insert(0, '0'); } $day = "-" + $day } $year = $matches[3];
        if ($month2) { $month2 = convert-Date $month2; $month2 = "-" + $month2; } if ($matches[6]) { $day2 = $matches[6]; if ($day2.Length -lt 2) { $day2 = ($day2).insert(0, '0'); }$day2 = "-" + $day2 } $year2 = $matches[7];
        if ($i -like "*Spring*" -or $i -like "*Fall*" -or $i -like "*Summer*" -or $i -like "*Winter*" ) {
            return $($year + "/" + $year2)
        }
        elseif (!$year) {
            return $($year2 + $month + $day + "/" + $year2 + $month2 + $day2)
        }
        elseif ($year2) {
            return $($year + $month + $day + "/" + $year2 + $month2 + $day2)
        }
        else {
            return $($year + $month + $day + "/" + $year + $month2 + $day2)
        }
    }
    # 3 undated
    elseif ($i -match "(\d{4})?(?:-(\d{4}))?.*(?:\s*and\s*)?undated" -and $i -ne "sfwxyswzFXSXyfqys") {
        $year = $null; $year2 = $null;
        if ($matches[1]) { $year = $matches[1] }
        if ($matches[2]) { $year2 = $matches[2] }
        #if($matches[1] -eq $null -and $matches[2] -eq $null){
        #   return $($minyear+"/"+$maxyear)
        #}
        if ($year -and $year2) {
            return $($year + "/" + $year2)
        }
        else {
                    
            return $($year)
        }
    }
    # 4 c 1790s, and 1790s
    elseif ($i -match "^(c\.?\s+)?(\d{4})s$") {
        $year = $matches[2];
        $year2 = endOfDecade $year
        return $($year + "/" + $year2) #>> $outpath
    }
    # 5 1970s-1980s
    elseif ($i -match "^\s*(\d{4})s\s*-\s*(\d{4})s\s*$") {
        $year = $matches[1];
        $year2 = $matches[2];
        $year2 = endOfDecade $year2;
        return $($year + "/" + $year2);
    }
    # 6 October, 2001
    elseif ($i -match "^[a-zA-Z]+,?\s*(\d{4})$" -and $i -notlike "Spring*" -and $i -notlike "Fall*" -and $i -notlike "Summer*" -and $i -notlike "Winter*" -and $i -inotlike "circa*") {
        if ($i -match "(^\w+)\b") { $month = $matches[1] }
        if ($i -match "(\d{4})$") { $year = $matches[1] }
        $month = convert-Date $month
        return $($year + "-" + $month) 
    }
    # 7 Spring, 2001
    elseif ($i -like "Spring*" -or $i -like "Fall*" -or $i -like "Summer*" -or $i -like "Winter*") {
        if ($i -match "(\d{4})$") { $year = $matches[1]; }
        return $($year )
    }
    # 8 October 16, 2001
    elseif ($i -match "([a-zA-Z]+)\s*(\d{1,2})(?:[nN][dD]|[sS][tT]|[rR][dD]|[tT][hH])?\s*,?\s*(\d{4})") {
        $year = $matches[3]; $day = $matches[2]; $month = convert-Date $matches[1]; 
        if ($day.Length -lt 2) { $day = ($day).insert(0, '0') }
        return $($year + "-" + $month + "-" + $day) 
    }
    # 9 October 16-18, 2001
    elseif ($i -match "([a-zA-Z]+)\s*(\d{1,2})(?:[nN][dD]|[sS][tT]|[rR][dD]|[tT][hH])?\s*(?:.{1,2})\s*\b(\d{1,2})(?:[nN][dD]|[sS][tT]|[rR][dD]|[tT][hH])?,\s*(\d{4})" -and $i -ne "hjnkejmnqwnmswdwfsvbkcfqelourpfvzsnfcgpsckwslrewhyozdhdsnafzojxez") {
        $year = $matches[4]; $day = $matches[2]; $day2 = $matches[3] ; $month = convert-Date $matches[1];
        if ($day.Length -lt 2) { $day = ($day).insert(0, '0') }
        if ($day2.Length -lt 2) { $day2 = ($day2).insert(0, '0') }
            
        return $($year + "-" + $month + "-" + $day + "/" + $year + "-" + $month + "-" + $day2) 
    }
    # 10 c. 1945-1947
    elseif ($i -match "^\s*c.\s*(\d{4})\s*-\s*(\d{4})\s*$") {
        $year = $matches[1]; $year2 = $matches[2];
        return $($year + "/" + $year2)
    }
    # 11 1945 and c. 1945
    elseif ($i -match "^\s*(?:c.|[cC][iI][Rr][cC][aA].?)?\s*(\d{4})$") {
        $year = $matches[1];
        return $($year) #>> $outpath
    }
    # 13 1942, 1045, 1945-1947
    elseif ($i -match "(\d.*\d)") {
        $str2 = $matches[1];  
        $str2 = $str2 -replace ",\s*|\s*-\s*" , ",";
        $str3 = $str2.Split(",");
        $min = $str3 | measure -Minimum -Maximum;
        $year = $min.Minimum;
        $year2 = $min.Maximum;
        $year = $year.ToString();
        $year2 = $year2.ToString();
        return $($year + "/" + $year2) ;
    }
}


# Create a new XML document
$xml = New-Object System.Xml.XmlDocument

# # Initialize the stack to store open elements
$elementStack = New-Object System.Collections.Generic.Stack[System.Xml.XmlElement]

# Create Root Element
$rootElement = $xml.CreateElement("RootElement")
$xml.AppendChild($rootElement) | Out-Null




function ConvertToXml {
    param(
        [Parameter(Mandatory = $true)] $csvFile,
        [Parameter(Mandatory = $true)] $xml
    )


    #########################################
    ######## Start XML Building Loop ########
    #########################################

    $record = 1
    $seriesID = 1
    $prevCNum
    # Start message
    Write-Host "Starting the script..." -ForegroundColor Green

    foreach ($row in $csvFile) {
              
        # Increase count of record to help identify errors.
        ++$record
        
        try {
            # Set a flag to determine if every cell is empty, blank or contains only spaces
            $allCellsEmpty = $true

            # Loop through each property (cell) in the current row
            foreach ($property in $row.PSObject.Properties) {
                # Check if the cell value is not null, not empty, and contains more than just spaces
                if (-not [string]::IsNullOrWhiteSpace($property.Value)) {
                    $allCellsEmpty = $false
                    break
                }
            }

            # If every cell is empty, blank or contains only spaces, skip the row
            if ($allCellsEmpty) {
                Write-Host "Warning: Blank row at Excel line: $record" -BackgroundColor Yellow -ForegroundColor Black
                continue
            }

            
            # Data Checks - Errors and Warnings

            # Check for required informaiton 
            if ([string]::IsNullOrEmpty($row.Attribute) -or [string]::IsNullOrEmpty($row.'c0#') -or [string]::IsNullOrEmpty($row.Title)) {
                write-host "Error: Required record information missing for record at Excel line: $record" -BackgroundColor Red  -ForegroundColor Black;
                Write-Host "Press any key to exit..."
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                exit   
            }# Checks for Missing Series ID
            # if ([string]::IsNullOrEmpty($row.'Series ID') -and $row.Attribute -eq "series") {
            #     Write-Host "Warning: Series ID missing for record at Excel line: $record" -BackgroundColor Yellow -ForegroundColor Black
            # }
            # Checks for High C# 
            if ([int]$row.'c0#' -gt 6) {
                Write-Host "Warning: High c# - You may want to check your logic. - c# = $([int]$row.'c0#') at Excel line: $record" -BackgroundColor Yellow -ForegroundColor Black
            }
            # Check for Series ID mismatch 
            if ($row.'Series ID' -or ($row.Attribute.Trim() -eq 'series')) {
                if (($row.'Series ID' -ne $null -and ($row.'Series ID'.Trim() -replace '\D', ''))  -ne $seriesID) { 
                    if (-not $row.'Series ID' -or $row.Attribute.Trim() -match "^\s*$") {
                        $currentSer = "BLANK CELL"
                    }
                    else{$currentSer = $row.'Series ID'.Trim()
                    }
                    Write-Host "Warning: Series ID mismatch for record at Excel line: $record - ID in Record: $currentSer, ID expected: ser$seriesID." -BackgroundColor Yellow -ForegroundColor Black }
                ++$seriesID
            }
            # Current C# breaks ascending pattern. 
            if ([int]$row.'c0#' -gt ($prevCNum + 1)){
                Write-Host "Warning: C# pattern broken on Excel line: $record. Previous value: $prevCnum, Expecting value: $($prevCNum + 1), actual value: $([int]$row.'c0#')."  -BackgroundColor Yellow -ForegroundColor Black 
            }
            
            
            # Starting XML Building
            
            # Get the hierarchy level and inner text from the CSV row
            $cNum = "{0:D2}" -f [int]$row.'c0#'
            
            $hierarchy = [int]$row.'c0#'
            
            # Create a new cNum element
            $newElement = $xml.CreateElement("c${cNum}")
            
            # Create the 'did' element for new element. 
            $did = $xml.CreateElement("did") 
            $newElement.AppendChild($did) 

            # Set Series ID 
            if ($row.'Series ID') {
                $newElement.SetAttribute("id", $row.'Series ID'.Trim()) 
                
            }

            # Set Level
            if ($row.Attribute) {
                $newElement.SetAttribute("level", $row.'Attribute'.Trim()) 
            }
            
            # Check if the 'Box' header exists
            if ($row.Box) {
                # Create Container Element.
                $box = $xml.CreateElement("container")
                # Add Container  Inner Text
                $box.InnerText = [int]$row.Box 
                # Add Attribute
                $box.SetAttribute("type", "box") 
                $did.AppendChild($box) 
            } # If not series or subseries populate empty value if no value given. 
            elseif ($row.Attribute.Trim() -notin 'subseries', 'series') {
                # Create Container Element.
                $box = $xml.CreateElement("container")
                # Add Container  Inner Text
                $box.InnerText = ""
                # Add Attribute
                $box.SetAttribute("type", "box")  
                $did.AppendChild($box) 
            }

            # Check if the 'File' header exists
            if ($row.File) {
                # Create Container Element.
                $file = $xml.CreateElement("container")
                # Add Container  Inner Text
                $file.InnerText = [int]$row.File
                # Add Attribute
                $file.SetAttribute("type", "folder")  
                $did.InsertAfter($file, $box) 
            } # If not series or subseries populate empty value if no value given. 
            elseif ($row.Attribute.Trim() -notin 'subseries', 'series') {
                # Create Container Element.
                $file = $xml.CreateElement("container")
                # Add Container  Inner Text
                $file.InnerText = ""
                # Add Attribute
                $file.SetAttribute("type", "folder")  
                $did.InsertAfter($file, $box) 
            }

            # Check if the 'Title' header exists
            if ($row.Title) {
                # Create the 'unittitle' child element of 'did' and set its inner text
                $unittitle = $xml.CreateElement("unittitle")
            }

            # Check if 'extref' exists
            if ($row.'Dspace URL') {
                # Create 'extref' element
                $extRef = $xml.CreateElement("extref")
                $extRef.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
                $extRef.SetAttribute("xlink:type", "simple")
                $extRef.SetAttribute("xlink:show", "new")
                $extRef.SetAttribute("xlink:actuate", "onRequest")
                $extRef.SetAttribute("xlink:href", $row.'Dspace URL'.Trim())

                # Set 'unittitle' text
                $extRef.InnerText = $row.'Title'.Trim()

                # Check if 'unitdate' exists
                if ($row.'Date') {
                    # Create 'unitdate' element
                    $unitDate = $xml.CreateElement("unitdate")
                    $unitDate.SetAttribute("era", "ce")
                    $unitDate.SetAttribute("calendar", "gregorian")
                    $unitDate.SetAttribute("normal", $(codedDate ($([string]$row.'Date').Trim())))
                    $unitDate.InnerText = ($([string]$row.'Date').Trim())

                    # Append 'unitdate' to 'extref'
                    $extRef.AppendChild($unitDate)
                }

                # Append 'extref' to 'unittitle'
                $unitTitle.AppendChild($extRef)
            }
            else {
                # Set 'unittitle' text
                $unitTitle.InnerText = $row.'Title'.Trim()

                # Check if 'unitdate' exists
                if ($row.'Date') {
                    # Create and append 'unitdate' element
                    $unitDate = $xml.CreateElement("unitdate")
                    $unitDate.SetAttribute("era", "ce")
                    $unitDate.SetAttribute("calendar", "gregorian")
                    $unitDate.SetAttribute("normal", $(codedDate ($([string]$row.'Date').Trim())))
                    $unitDate.InnerText = ($([string]$row.'Date').Trim())
                    $unitTitle.AppendChild($unitDate)
                }
            }
            $did.AppendChild($unittitle) 
            
            # Handle the hierarchy
            if ($elementStack.Count -eq 0) {
                # If the stack is empty, add the element as a child of the root
                $rootElement.AppendChild($newElement) 
            }
            elseif ($hierarchy -lt $elementStack.Peek().Name.Substring(1)) {
                # If the hierarchy is less than the current open element, close the open element and add the new element
                while ($hierarchy -lt $elementStack.Peek().Name.Substring(1)) {
                    $elementStack.Pop()
                }
                $elementStack.Peek().ParentNode.AppendChild($newElement) 
                # If the hierarchy is equal to the previous element, append to its parent. 
            }
            elseif ($hierarchy -eq $elementStack.Peek().Name.Substring(1)) {
                $elementStack.Peek().ParentNode.AppendChild($newElement) 
            }
            else {
                # Add the new element as a child of the current open element
                $elementStack.Peek().AppendChild($newElement) 
            }

            # Push the new element onto the stack
            $elementStack.Push($newElement)

            $prevCNum = [int]$row.'c0#'
        }
        catch {
            $lineNumber = $_.InvocationInfo.ScriptLineNumber
            Write-Host "Error processing record at Excel line: $record" -ForegroundColor Red
            Write-Host "Error at Powershell line: $lineNumber" -ForegroundColor Red
            Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Press any key to exit..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            exit
        }
        
    }
 
}

###############
# File Output #
###############

# Call the ConvertToXml function with the $csvFile and $xml
ConvertToXml -csvFile $csvFile -xml $xml | Out-Null

# Find full path from relative path. Some functions don't work correctly with the relative path. 

# File Name
$fileNamePrefix = "xmlOutput-RB"
# Time Based suffice (Year Month Day Hour Minute Seconds Milliseconds)
# Time Suffix updated to a human readable date time stamp. 
# You can comment out the new one and uncomment the old one if you want the epoch time. 

#$fileSuffix = $fileSuffix = (Get-Date).ToBinary().ToString().Replace("-", "")
$fileSuffix = (Get-Date).ToString("yyyy_MM_dd-HHmm_ss_fff")

$fileName = $fileNamePrefix + "-" + $fileSuffix + ".xml"
$relPath = ".\"
$filePath = "$relPath$fileNAme"
$fullPath = "$(Resolve-Path $relPath)\$fileName"
#$filePath = $fullPath

# Save the XML document to a file
try {
    $xml.Save($fullPath) 
}
catch {
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    Write-Host "Error at Powershell line: $lineNumber" -ForegroundColor Red
    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

# Read the content of the file
try {
    $content = Get-Content $fullPath
}
catch {
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    Write-Host "Error at Powershell line: $lineNumber" -ForegroundColor Red
    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

# Remove the first and last line from the content
$content = $content[1..($content.Length - 2)]

# Save the updated content to the file
try {
    Set-Content $fullPath -Value $content
}
catch {
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    Write-Host "Error at Powershell line: $lineNumber" -ForegroundColor Red
    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}



# Stop message
Write-Host "Script completed. Results written to: $fileName" -ForegroundColor Green


# Pause at the end
# Write-Host "Press any key to exit and open the output file..."
# $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

notepad.exe $fullPath
