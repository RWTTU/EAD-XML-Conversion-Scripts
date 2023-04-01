#################
# DATE PATTERNS #
#################

# Regex pattern for Case 1: October-December, 2001
$regex_case1 = "(?i)([a-z]+).?\s*-\s*([a-z]+)\s*.?\s*(\d{4})"

# Regex pattern for Case 2: January 24, 2014 - February 24, 2018 and a few variations
$regex_case2 = "(?i)([a-z]+)\s*,?\s*\b(\d{1,2})?(?:nd|st|rd|th)?\b\s*,?\s*(\d{4})?(\s*.{1,2}\b\s*([a-z]+)\s*,?\s*\b(\d{1,2})?(?:nd|st|rd|th)?\b\s*,?\s*(\d{4})?)"

# Regex pattern for Case 3: Some date and undated
$regex_case3 = "(?i)(\d{4})?(?:-(\d{4}))?.*(?:\s*and\s*)?undated"

# Regex pattern for Case 4: c 1790s, and 1790s
$regex_case4 = "(?i)^(c\.?\s+)?(\d{4})s$"

# Regex pattern for Case 5: 1970s-1980s
$regex_case5 = "(?i)^\s*(\d{4})s\s*-\s*(\d{4})s\s*$"

# Regex pattern for Case 6: October, 2001
$regex_case6 = "(?i)^[a-z]+,?\s*(\d{4})$"

# Regex pattern for Case 7: Spring, 2001
$regex_case7 = "(?i)spring|fall|summer|winter"

# Regex pattern for Case 8: October 16, 2001
$regex_case8 = "(?i)([a-z]+)\s*(\d{1,2})(?:nd|st|rd|th)?\s*,?\s*(\d{4})"

# Regex pattern for Case 9: October 16-18, 2001
$regex_case9 = "(?i)([a-z]+)\s*(\d{1,2})(?:nd|st|rd|th)?\s*(?:.{1,2})\s*\b(\d{1,2})(?:nd|st|rd|th)?,\s*(\d{4})"

# Regex pattern for Case 10: c. 1945-1947
$regex_case10 = "(?i)^\s*c.\s*(\d{4})\s*-\s*(\d{4})\s*$"

# Regex pattern for Case 11: 1945 and c. 1946
$regex_case11 = "(?i)^\s*(?:c.|circa)?\s*(\d{4})$"

# Regex pattern for Case 12: 1942, 1045, 1945-1947
$regex_case12 = "(\d.*\d)"


##############
# FORMAT XML #
##############

function Format-Xml {
    <#
    .SYNOPSIS
    Formats an array of strings as the text of an XML document.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Text
    )

    begin {
        $data = New-Object System.Collections.ArrayList
    }

    process {
        $data.Add($Text -join "`n") | Out-Null
    }

    end {
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($data -join "`n")
        $sw = New-Object System.IO.StringWriter
        $writer = New-Object System.Xml.XmlTextWriter($sw)
        $writer.Formatting = [System.Xml.Formatting]::Indented
        $doc.WriteContentTo($writer)
        $sw.ToString()
    }
}



##############
# IMPORT XLS #
##############


function Import-Xls 
{ 

<# 
.SYNOPSIS 
Import an Excel file. 
 
.DESCRIPTION 
Import an excel file. Since Excel files can have multiple worksheets, you can specify the worksheet you want to import. You can specify it by number (1, 2, 3) or by name (Sheet1, Sheet2, Sheet3). Imports Worksheet 1 by default. 
 
.PARAMETER Path 
Specifies the path to the Excel file to import. You can also pipe a path to Import-Xls. 
 
.PARAMETER Worksheet 
Specifies the worksheet to import in the Excel file. You can specify it by name or by number. The default is 1. 
Note: Charts don't count as worksheets, so they don't affect the Worksheet numbers. 
 
.INPUTS 
System.String 
 
.OUTPUTS 
Object 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet 1 
Import Worksheet 1 from employees.xlsx 
 
.EXAMPLE 
".\employees.xlsx" | Import-Xls -Worksheet "Sheet2" 
Import Worksheet "Sheet2" from employees.xlsx 
 
.EXAMPLE 
".\deptA.xslx", ".\deptB.xlsx" | Import-Xls -Worksheet 3 
Import Worksheet 3 from deptA.xlsx and deptB.xlsx. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.EXAMPLE 
Get-ChildItem *.xlsx | Import-Xls -Worksheet "Employees" 
Import Worksheet "Employees" from all .xlsx files in the current directory. 
Make sure that the worksheets have the same headers, or have some headers in common, or that it works the way you expect. 
 
.LINK 
Import-Xls 
http://gallery.technet.microsoft.com/scriptcenter/17bcabe7-322a-43d3-9a27-f3f96618c74b 
Export-Xls 
http://gallery.technet.microsoft.com/scriptcenter/d41565f1-37ef-43cb-9462-a08cd5a610e2 
Import-Csv 
Export-Csv 
 
.NOTES 
Author: Francis de la Cerna 
Created: 2011-03-27 
Modified: 2011-04-09 
#Requires –Version 2.0 
#> 
 
 [CmdletBinding(SupportsShouldProcess=$true)] 
     
    Param( 
        [parameter( 
            mandatory=$true,  
            position=1,  
            ValueFromPipeline=$true,  
            ValueFromPipelineByPropertyName=$true)] 
        [String[]] 
        $Path, 
     
        [parameter(mandatory=$false)] 
        $Worksheet = 1, 
         
        [parameter(mandatory=$false)] 
        [switch] 
        $Force 
    ) 
 
    Begin 
    { 
        function GetTempFileName($extension) 
        { 
            $temp = [io.path]::GetTempFileName(); 
            $params = @{ 
                Path = $temp; 
                Destination = $temp + $extension; 
                Confirm = $false; 
                Verbose = $VerbosePreference; 
            } 
            Move-Item @params; 
            $temp += $extension; 
            return $temp; 
        } 
             
        # since an extension like .xls can have multiple formats, this 
        # will need to be changed 
        # 
        $xlFileFormats = @{ 
            # single worksheet formats 
            '.csv'  = 6;        # 6, 22, 23, 24 
            '.dbf'  = 11;       # 7, 8, 11 
            '.dif'  = 9;        #  
            '.prn'  = 36;       #  
            '.slk'  = 2;        # 2, 10 
            '.wk1'  = 31;       # 5, 30, 31 
            '.wk3'  = 32;       # 15, 32 
            '.wk4'  = 38;       #  
            '.wks'  = 4;        #  
            '.xlw'  = 35;       #  
             
            # multiple worksheet formats 
            '.xls'  = -4143;    # -4143, 1, 16, 18, 29, 33, 39, 43 
            '.xlsb' = 50;       # 
            '.xlsm' = 52;       # 
            '.xlsx' = 51;       # 
            '.xml'  = 46;       # 
            '.ods'  = 60;       # 
        } 
         
        $xl = New-Object -ComObject Excel.Application; 
        $xl.DisplayAlerts = $false; 
        $xl.Visible = $false; 
    } 
 
    Process 
    { 
        $Path | ForEach-Object { 
             
            if ($Force -or $psCmdlet.ShouldProcess($_)) { 
             
                $fileExist = Test-Path $_ 
 
                if (-not $fileExist) { 
                    Write-Error "Error: $_ does not exist" -Category ResourceUnavailable;             
                } else { 
                    # create temporary .csv file from excel file and import .csv 
                    # 
                    $_ = (Resolve-Path $_).toString(); 
                    $wb = $xl.Workbooks.Add($_); 
                    if ($?) { 
                        $csvTemp = GetTempFileName(".csv"); 
                        $ws = $wb.Worksheets.Item($Worksheet);
                        $ws 
                        $ws.SaveAs($csvTemp, $xlFileFormats[".csv"]); 
                        $wb.Close($false); 
                        Remove-Variable -Name ('ws', 'wb') -Confirm:$false; 
                        #notepad $csvTemp
                        Import-Csv $csvTemp 
                        Remove-Item $csvTemp -Confirm:$false -Verbose:$VerbosePreference; 
                    } 
                } 
            } 
        } 
    } 
    
    End 
    { 
        $xl.Quit(); 
        Remove-Variable -name xl -Confirm:$false; 
        [gc]::Collect(); 
    } 
} 


##################
# GET FILE NAMES #
##################

function Get-FileName {
    <#
    .SYNOPSIS
    Opens a dialog box that allows the user to select a file from the file system.

    .DESCRIPTION
    The Get-FileName function opens an Open File dialog box that allows the user to select a file from the file system. The selected file's name (including the full path) is returned as the output of the function.

    .PARAMETER InitialDirectory
    The initial directory to display in the Open File dialog box. If no value is specified, the current working directory is used.

    .EXAMPLE
    PS C:\> Get-FileName -InitialDirectory "C:\Users\JohnDoe\Desktop"

    This example opens the Open File dialog box with the initial directory set to "C:\Users\JohnDoe\Desktop".

    .OUTPUTS
    System.String
    The name of the selected file, including the full path.

    .NOTES
    This function uses the System.Windows.Forms .NET assembly to create an Open File dialog box.
    #>
    [CmdletBinding()]
    param (
        [string]$InitialDirectory = $PWD
    )

    Add-Type -AssemblyName System.Windows.Forms | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.Filter = "All files (*.*)| *.*"
    if ($OpenFileDialog.ShowDialog() -eq 'OK') {
        return $OpenFileDialog.FileName
    }
}


##################
# DATE FUNCTIONS #
##################


#convert date to number 
function Convert-Date {
    param($inDate)

    $monthMapping = @{
        'Jan' = '01'
        'Feb' = '02'
        'Mar' = '03'
        'Apr' = '04'
        'May' = '05'
        'Jun' = '06'
        'Jul' = '07'
        'Aug' = '08'
        'Sep' = '09'
        'Oct' = '10'
        'Nov' = '11'
        'Dec' = '12'
    }

    foreach ($month in $monthMapping.Keys) {
        if ($inDate -like "$month*") {
            return $monthMapping[$month]
        }
    }
}

function Get-EndOfDecade {
    param(
        [int]$Year
    )

    $endYear = ($Year - ($Year % 10)) + 9
    return $endYear
}

function minDate {
    param($csv1, $minrun)

    $years = @()

    foreach ($i in $csv1) {
        if ($i.Date -eq $null) { continue }

        $date = $i.Date

        $year, $year2 = $null, $null

        if ($date -match "(\d{4})") { $year = $matches[1] }

        switch -Regex ($date) {
            # Case 1: October-December, 2001
            $regex_case1 { $year = $matches[3] }
            # Case 2: January 24, 2014 - February 24, 2018 and a few variations
            $regex_case2 {
                if ($matches[3]) { $year = $matches[3] }
                if ($matches[7]) { $year2 = $matches[7] }
            }
            # Case 3: undated
            $regex_case3 {
                if ($matches[1]) { $year = $matches[1] }
                if ($matches[2]) { $year2 = $matches[2] }
            }
            # Case 4: c 1790s, and 1790s
            $regex_case4 {
                $year = $matches[2]
                $year2 = Get-EndOfDecade -Year $year
            }
            # Case 5: 1970s-1980s
            $regex_case5 {
                $year = $matches[1]
                $year2 = Get-EndOfDecade -Year $matches[2]
            }
            # Case 6: October, 2001
            $regex_case6 {
                if (-not ($date -like "Spring*" -or $date -like "Fall*" -or $date -like "Summer*" -or $date -like "Winter*" -or $date -inotlike "circa*")) {
                    $year = $matches[1]
                }
            }
            # Case 7: Spring, 2001
            $regex_case7 {
                if ($date -match "(\d{4})$") { $year = $matches[1] }
            }
            # Case 8: October 16, 2001
            $regex_case8 { $year = $matches[3] }
            # Case 9: October 16-18, 2001
            $regex_case9 { $year = $matches[4] }
            # Case 10: c. 1945-1947
            $regex_case10 {
                $year = $matches[1]
                $year2 = $matches[2]
            }
            # Case 11: 1945 and c. 1946
            $regex_case11 { $year = $matches[1] }
            # Case 12: 1942, 1045, 1945-1947
            $regex_case12 {
                $str2 = $matches[1]
                $str2 = $str2 -replace ",\s*|\s*-\s*", ","
                $str3 = $str2.Split(",")
                $min = $str3 | Measure-Object -Minimum -Maximum
                $year = $min.Minimum
                $year2 = $min.Maximum
            }
        }
    
            if ($year) { $years += $year }
            if ($year2) { $years += $year2 }
    }
    
    if ($minrun -eq 0) { return ($years | Measure-Object -Minimum).Minimum }
    else { return ($years | Measure-Object -Maximum).Maximum }
}

function codedDate{
$year = ""
$year2 = ""
$month = "" 
$month2 = ""
$day = ""
$day2 = ""
$i = $args[0]
$minyear = $args[1]
$maxyear = $args[2]
    
    # 1 October-December, 2001
    if($i -match $regex_case1 ){
        $year = $matches[3]; $month = Convert-Date $matches[1]; $month2 = Convert-Date $matches[2];
        return $($year+"-"+$month+"/"+$year+"-"+$month2) 
    }
    # 2 January 24, 2014 - February 24, 2018 and a few variations Done
    elseif($i -match $regex_case2 -and $i -notlike "*undated*"){
        $month = $matches[1]; $month2 = $matches[5];
        if($month){$month = Convert-Date $month; $month = "-"+$month;} if($matches[2]){$day = $matches[2];if($day.Length -lt 2){ $day =  ($day).insert(0,'0');} $day = "-"+$day} $year = $matches[3];
        if($month2){$month2 = Convert-Date $month2; $month2 = "-"+$month2;} if($matches[6]){$day2 = $matches[6];if($day2.Length -lt 2){ $day2 =  ($day2).insert(0,'0');}$day2 = "-"+$day2} $year2 = $matches[7];
        if($i -like "*Spring*" -or   $i -like "*Fall*" -or   $i -like "*Summer*" -or   $i -like "*Winter*" ){
            return $($year+"/"+$year2)
        }elseif(!$year){
            return $($year2+$month+$day+"/"+$year2+$month2+$day2)
        }elseif($year2){
            return $($year+$month+$day+"/"+$year2+$month2+$day2)
        }else{
            return $($year+$month+$day+"/"+$year+$month2+$day2)
        }
    }
    # 3 undated
    elseif($i -match $regex_case3 -and $i -ne "sfwxyswzFXSXyfqys"){
        $year = $null; $year2=$null;
        if($matches[1]){$year = $matches[1]}
        if($matches[2]){$year2 = $matches[2]}
        if($matches[1] -eq $null -and $matches[2] -eq $null){
            return "$minyear/$maxyear"
        }elseif($year -ne $null -and $year2 -ne $null){
            return $($year+"/"+$year2)
        }else{
                    
            return $($year)
        }
    }
     # 4 c 1790s, and 1790s
    elseif($i -match $regex_case4 ){ $year = $matches[2];
        $year2 = Get-EndOfDecade -Year $year
    return $($year+"/"+ $year2) #>> $outpath
    }
    # 5 1970s-1980s
    elseif($i -match $regex_case5 ){
        $year = $matches[1];
        $year2 = $matches[2];
        $year2 = Get-EndOfDecade -Year $year2;
        return $($year+"/"+$year2);
    }
    # 6 October, 2001
    elseif($i -match $regex_case6 -and   $i -notlike "Spring*" -and   $i -notlike "Fall*" -and   $i -notlike "Summer*" -and   $i -notlike "Winter*" -and   $i -inotlike "circa*"){
        if ($i -match "(^\w+)\b"){ $month = $matches[1]}
        if ($i -match "(\d{4})$"){ $year = $matches[1]}
        $month = Convert-Date $month
        return $($year+"-"+$month) 
    }
    # 7 Spring, 2001
    elseif($i -like "Spring*" -or   $i -like "Fall*" -or   $i -like "Summer*" -or   $i -like "Winter*"){
        if ($i -match "(\d{4})$"){ $year = $matches[1]; }
    return $($year )
    }
    # 8 October 16, 2001
    elseif($i -match  $regex_case8 ){$year = $matches[3]; $day = $matches[2]; $month = Convert-Date $matches[1]; 
        if($day.Length -lt 2){ $day =  ($day).insert(0,'0')}
        return $($year+"-"+$month+"-"+$day) 
    }
    # 9 October 16-18, 2001
    elseif($i -match $regex_case9 -and $i -ne "hjnkejmnqwnmswdwfsvbkcfqelourpfvzsnfcgpsckwslrewhyozdhdsnafzojxez"){
        $year = $matches[4]; $day = $matches[2]; $day2 = $matches[3] ;  $month = Convert-Date $matches[1];
        if($day.Length -lt 2){ $day =  ($day).insert(0,'0')}
        if($day2.Length -lt 2){ $day2 =  ($day2).insert(0,'0')}
         
        return $($year+"-"+$month+"-"+$day+"/"+$year+"-"+$month+"-"+$day2) 
    }
    # 10 c. 1945-1947
    elseif($i -match $regex_case10 ){$year = $matches[1]; $year2 = $matches[2];
        return $($year+"/"+$year2)
    }
    # 11 1945 and c. 1945
    elseif($i -match $regex_case11  ){ $year = $matches[1];
        return $($year) #>> $outpath
    }
    # 12 1942, 1045, 1945-1947
    elseif($i -match  $regex_case12 ){ $str2 = $matches[1];  
        $str2 = $str2 -replace ",\s*|\s*-\s*" , ",";
        $str3 = $str2.Split(",");
        $min = $str3 | measure -Minimum -Maximum;
        $year = $min.Minimum;
        $year2 = $min.Maximum;
        $year = $year.ToString();
        $year2 = $year2.ToString();
        return $($year+"/"+ $year2) ;
    }
   
 }
 
# **************************** Entry Point to Script **********************************
if( get-module -ListAvailable -Name ImportExcel){}else{

Install-Module ImportExcel -Force}

write-host "Processing CSV. Please wait..."
$file = Get-FileName -initialDirectory "c:fso" 
#$csv = $file | Import-Xls -Worksheet 1

$csv = Import-Excel -Path $file
 
$minyear = minDate -csv1 $csv -minrun 0
$maxyear = minDate -csv1 $csv -minrun 1

$outfile = ".\xmlOutput" + $(get-date).Tobinary() + ".xml"

$csv | foreach-object {

$_.Title = $_.Title.replace("&","&amp;")

}

$count = 1
$preSer = 0 
$a = $csv | Measure-Object
$a = $a.Count
$a = [int]$a
$progresspercent = 0
$progressposition = 1
$record = 2

foreach( $i in $csv){

    ###Progress Bar###
    Write-Progress -Activity "Working..." -PercentComplete $progresspercent -CurrentOperation  "Processing Record $progressposition / $a... " -Status "Please wait."
            
    ### Start Loop ###

$clevel = $i.'c0#'

    if($i.Attribute -eq "" -and $i.'c0#' -eq "" -and $i.Title -eq ""){continue}
        
    if($i.Attribute -eq "" -or $i.'c0#' -eq "" -or $i.Title -eq ""){
        
        write-host "Error: Required record information missing for record around line $record" -BackgroundColor Red  -ForegroundColor Black;
        write-host "Deleting partial file..."
        rm $outfile;
        
        Read-Host 'Press enter to close' | Out-Null
        Exit   
        }
        
    if($perSer -ge $clevel){
        do{
        $("</c0$perSer>") >> $outfile
        $perSer--
        }until($perSer+1 -eq $clevel) 
            
        }
        
       if($i.Attribute -in ("series", "subseries")){
          if(!$i.'Series ID' -and $i.Attribute -eq "series"){Write-Host "Warning: Series ID Missing for record on line $record - $($i.Title)" -BackgroundColor red -ForegroundColor Black }      
           
          if($i.Attribute -eq "subseries"){
           $("<c0$($clevel) level=""$($i.Attribute)""><did>") >> $outfile
           }elseif($i.Attribute -eq "series"){$("<c0$($clevel) id=""$($i.'Series ID')"" level=""$($i.Attribute)""><did>") >> $outfile}
        
            if($i."Dspace URL"){
        #Link Title
            $("<unittitle>'r'n<extref xmlns:xlink=""http://www.w3.org/1999/xlink"" xlink:type=""simple"" xlink:show=""new"" xlink:actuate=""onRequest""  
        xlink:href=""$($i."Dspace URL")"">
            $($i.Title+" ")</extref>
        </unittitle>
            </did>") >> $outfile}else{
        #No Link Title
            $("<unittitle>$($i.Title+" ")</unittitle>
            </did>") >> $outfile}
            }else

            {
           $("<c0$($clevel) level=""$($i.Attribute)""> <did>
        <container type=""box"">$($i.Box)</container>
        <container type=""folder"">$($i.File)</container>") >> $outfile

        if($i."Dspace URL"){
            $date = $i.Date
            if ($date -eq "undated") {
                $date = "0000/0000"
            } else {
                $date = codedDate $i.Date $minyear $maxyear
            }
            
            #Link Title
            $("<unittitle><extref xmlns:xlink=""http://www.w3.org/1999/xlink"" xlink:type=""simple""
            xlink:show=""new"" xlink:actuate=""onRequest""  
            xlink:href=""$($i."Dspace URL")"">$($i.Title+" ")<unitdate era=""ce""
            calendar=""gregorian"" normal=""$date"">$($i.Date)</unitdate></extref></unittitle>
            </did>") >> $outfile
        } else {
            $date = $i.Date
            if ($date -eq "undated") {
                $date = "0000/0000"
            } else {
                $date = codedDate $i.Date $minyear $maxyear
            }
                    
            #No Link Title
            $("<unittitle>$($i.Title+" ")<unitdate era=""ce""
            calendar=""gregorian"" normal=""$date"">$($i.Date)</unitdate></unittitle>
            </did>") >> $outfile
        } 
        
    }
           
$record++
$perSer = $i.'c0#'
$perSer = [int]$perSer
$progressposition++
$progresspercent = ($progressposition/$a)*100
}

 do{
        $("</c0$perSer>") >> $outfile
        $perSer--
        }until($perSer -eq 0)
        
        ###End Progress###
        Write-Progress -Activity "Working..." -Completed -Status "All done."

        #################################Format Document##########################################
    write-host "Processing Complete!"     
    
   notepad $outfile
   #pause 
   #Read-Host 'Press enter to close' | Out-Null