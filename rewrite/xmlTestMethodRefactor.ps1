######################
# LOAD PREREQUISITES #
######################

# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Install ImportExcel in CurrentUser scope to not trip UAC flags 
# Check if the ImportExcel module is already installed
if (-not (Get-InstalledModule -Name ImportExcel -ErrorAction SilentlyContinue)) {
    # Install the ImportExcel module if it's not already installed
    Install-Module -Name ImportExcel -Scope CurrentUser
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

# Filter out blank rows using the Where-Object cmdlet
$excelFile = Import-Excel -Path $filePath
#$csvFile = $excelFile #| Where-Object { $_.PSObject.Properties.Value -notcontains "" }
$csvFile = $excelFile | Where-Object { ($_.PSObject.Properties.Value | ForEach-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0 }


# Create a new XML document
$xml = New-Object System.Xml.XmlDocument

# # Initialize the stack to store open elements
$elementStack = New-Object System.Collections.Generic.Stack[System.Xml.XmlElement]

# Create Root Element
$rootElement = $xml.CreateElement("RootElement")
$xml.AppendChild($rootElement) | Out-Null

$record = 2
$seriesID = 1
function ConvertToXml {
    param(
        [Parameter(Mandatory = $true)] $csvFile,
        [Parameter(Mandatory = $true)] $xml
    )


    #########################################
    ######## Start XML Building Loop ########
    #########################################

    foreach ($row in $csvFile) {

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
            continue
        }

        
        # Data Checks

        # Check for required informaiton 
        if ([string]::IsNullOrEmpty($row.Attribute) -or [string]::IsNullOrEmpty($row.'c0#') -or [string]::IsNullOrEmpty($row.Title)) {
            write-host "Error: Required record information missing for record around line $record" -BackgroundColor Red  -ForegroundColor Black;
            Read-Host 'Press enter to close' | Out-Null
            Exit   
        }# Checks for Missing Series ID
        elseif ([string]::IsNullOrEmpty($row.'Series ID') -and $row.Attribute -eq "series") {
            Write-Host "Warning: Series ID missing for record near line $record - $($row.Title)" -BackgroundColor Yellow -ForegroundColor Black
        }# Checks for High C# 
        if($row.'c0#' -gt 6){
            Write-Host "Warning: High c# - You may want to check your logic. - c# = $($row.'c0#') near line $record - $($row.Title)" -BackgroundColor Yellow -ForegroundColor Black
        }#Check for Series ID mismatch 
        if($row.'Series ID' -or $row.Attribute -eq 'series'){
            if (($row.'Series ID' -replace '\D','') -ne $seriesID){Write-Host "Warning: Series ID mismatch for record near line $record - ID in Record =  $($row.'Series ID') :: ID expected = ser$seriesID" -BackgroundColor Yellow -ForegroundColor Black}
            ++$seriesID
        }
        
        ++$record

        
        # Get the hierarchy level and inner text from the CSV row
        $cNum = "{0:D2}" -f [int]$row.'c0#'
        
        $hierarchy = [int]$row.'c0#'
        
        # Create a new cNum element
        $newElement = $xml.CreateElement("c${cNum}")
        
        # Create the 'did' element for new element. 
        $did = $xml.CreateElement("did") 
        $newElement.AppendChild($did) 

        # Check if the 'c0#' header exists
        if (!$row.'c0#') {
            # Do something if the 'c0#' header exists
            #write-output "Missing c0# in row ${rowNumber}"
        }

        # Set Series ID 
        if ($row.'Series ID') {
            $newElement.SetAttribute("id", $row.'Series ID') 
            
        }

        # Set Level
        if ($row.Attribute) {
            $newElement.SetAttribute("level", $row.'Attribute') 
        }
        
        # Check if the 'Box' header exists
        if ($row.Box) {
            # Create Container Element.
            $box = $xml.CreateElement("container")
            # Add Container  Inner Text
            $box.InnerText = $row.Box 
            # Add Attribute
            $box.SetAttribute("type", "box") 
            $did.AppendChild($box) 
        } # If not series or subseries populate empty value if no value given. 
        elseif ($row.Attribute -notin 'subseries', 'series') {
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
            $file.InnerText = $row.File
            # Add Attribute
            $file.SetAttribute("type", "folder")  
            $did.InsertAfter($file, $box) 
        } # If not series or subseries populate empty value if no value given. 
        elseif ($row.Attribute -notin 'subseries', 'series') {
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
            #$unittitle.InnerText = $row.'Title'
            
        }

        # Check if 'extref' exists
        if ($row.'Dspace URL') {
            # Create 'extref' element
            $extRef = $xml.CreateElement("extref")
            $extRef.SetAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
            $extRef.SetAttribute("xlink:type", "simple")
            $extRef.SetAttribute("xlink:show", "new")
            $extRef.SetAttribute("xlink:actuate", "onRequest")
            $extRef.SetAttribute("xlink:href", $row.'Dspace URL')

            # Set 'unittitle' text
            $extRef.InnerText = $row.'Title'

            # Check if 'unitdate' exists
            if ($row.'Date') {
                # Create 'unitdate' element
                $unitDate = $xml.CreateElement("unitdate")
                $unitDate.SetAttribute("era", "ce")
                $unitDate.SetAttribute("calendar", "gregorian")
                $unitDate.SetAttribute("normal", $row.'Date')
                $unitDate.InnerText = $row.'Date'

                # Append 'unitdate' to 'extref'
                $extRef.AppendChild($unitDate)
            }

            # Append 'extref' to 'unittitle'
            $unitTitle.AppendChild($extRef)
        }
        else {
            # Set 'unittitle' text
            $unitTitle.InnerText = $row.'Title'

            # Check if 'unitdate' exists
            if ($row.'Date') {
                # Create and append 'unitdate' element
                $unitDate = $xml.CreateElement("unitdate")
                $unitDate.SetAttribute("era", "ce")
                $unitDate.SetAttribute("calendar", "gregorian")
                $unitDate.SetAttribute("normal", $row.'Date')
                $unitDate.InnerText = $row.'Date'

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
        
    }
 
}


# Call the ConvertToXml function with the $csvFile and $xml
ConvertToXml -csvFile $csvFile -xml $xml | Out-Null

# Save the XML document to a file
$xml.Save("C:\git\EAD-XML-Conversion-Scripts\test\output.xml") 

type C:\git\EAD-XML-Conversion-Scripts\test\output.xml