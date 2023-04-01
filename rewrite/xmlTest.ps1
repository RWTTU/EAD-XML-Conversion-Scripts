# Read the CSV file
$csvFile = Import-Csv -Path "..\test\CSVForknerSHORT.csv"





# Create a new XML document
$xml = New-Object System.Xml.XmlDocument


# # Initialize the stack to store open elements
$elementStack = New-Object System.Collections.Generic.Stack[System.Xml.XmlElement]

# Create Root Element
$rootElement = $xml.CreateElement("RootElement")
$xml.AppendChild($rootElement)




foreach ($row in $csvFile) {
    #$rowNumber = $_.psobject.Properties.Value.IndexOf($row) + 1
    # Get the hierarchy level and inner text from the CSV row
    $cNum = "{0:D2}" -f [int]$row.'c0#'
    
    $hierarchy = [int]$row.'c0#'
    $innerText = $row.InnerText

    # Create a new cNum element
    $newElement = $xml.CreateElement("c${cNum}")
    
    # Create the 'did' element for new element. 
    $did = $xml.CreateElement("did")
 

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
    }

    # Check if the 'File' header exists
    if ($row.File) {
            # Create Container Element.
            $File = $xml.CreateElement("container")
            # Add Container  Inner Text
            $File.InnerText = $row.File
            # Add Attribute
            $File.SetAttribute("type", "folder")
    }

    # Check if the 'Title' header exists
    if ($row.Title) {
        # Do something if the 'Title' header exists
        #write-output "Processing row with Title $($row.Title)"
    }

    # Check if the 'Date' header exists
    if ($row.Date) {
        # Do something if the 'Date' header exists
        #write-output "Processing row with Date $($row.Date)"
    }

    
    
    # did
    # container
    # unittitle
    #unitdate
 
    

    
    # If LEVEL is SERIES
    if ($row.Attribute -eq 'series') {
    
        # Set Series ID
        
        
 
        # Create the 'unittitle' child element of 'did' and set its inner text
        $unittitle = $xml.CreateElement("unittitle")
        $unittitle.InnerText = $row.'Title'
        
        # Append Children
        
        $newElement.AppendChild($did)
        $did.AppendChild($unittitle)
    }

    # If LEVEL is SUBSERIES
    if ($row.Attribute -eq 'subseries') {
        $newElement.SetAttribute("level", $row.'Attribute')
    }

    # IF LEVEL not SERIES or SUBSERIES
    if ($row.Attribute -notin 'subseries', 'series') {
     
    }

    

    # Set the inner text if available
    # if ($innerText) {
    #     $newElement.InnerText = $innerText
    # }

    # Set attributes if available
    # foreach ($property in $row.PSObject.Properties) {
    #     $attributeName = $property.Name
    #     $attributeValue = $property.Value

    #     if ($attributeName -notin @('Hierarchy', 'InnerText') -and $attributeValue) {
    #         $newElement.SetAttribute($attributeName, $attributeValue)
    #     }
    # }

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

    #Write-Host !!!!$hierarchy
}


# Save the XML document to a file
$xml.Save("C:\git\EAD-XML-Conversion-Scripts\test\output.xml")
type C:\git\EAD-XML-Conversion-Scripts\test\output.xml