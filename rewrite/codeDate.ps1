# Import Data 
$csvdata = Get-Content ..\dateFormats\dates.csv

function Parse-DateString {
    param (
        [string]$inputString
    )

    $result = New-Object -TypeName PSObject -Property @{
        Year = $null
        Month = $null
        Day = $null
    }

    # Try to parse input string as a date
    $dateTime = [DateTime]::MinValue
    if ([DateTime]::TryParse($inputString, [ref]$dateTime)) {
        $result.Year = $dateTime.Year
        $result.Month = $dateTime.Month.ToString("D2")
        $result.Day = $dateTime.Day.ToString("D2")
    } else {
        # Check for year, month, and day components separately
        switch -regex ($inputString) {
            "\b\d{4}\b" {
                $result.Year = $matches[0]
            }
            "\b\d{1,3}\b" {
                $result.Day = $matches[0].ToString("D2")
            }
            "^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*" {
                $monthAbbrev = $matches[1]
                $monthNum = (Get-Culture).DateTimeFormat.AbbreviatedMonthNames.IndexOf($monthAbbrev) + 1
                if ($monthNum -gt 0) {
                    $result.Month = $monthNum.ToString("D2")
                }
            }
        }
    }

    return $result
}




function HasMultipleYears($inputString) {
    $regex = "\d{4}"
    $yearMatches = [regex]::Matches($inputString, $regex)
        $regex = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        $monthMatches = [regex]::Matches($inputString, $regex)
        $regex = "\d{1,2}"
        $dayMatches = [regex]::Matches($inputString, $regex)
     
    write-host "y:$($yearMatches.Count) m:$($monthMatches.Count) d:$($dayMatches.Count)"
    

       


    # If your string has 1 year and no months
    # if($yearCount -eq 1 -and $inputString -notmatch "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\ds)"){
    #     write-host $yearCount - $inputString => $matches.value
    # }#else($yearCount -eq 2){}
}

foreach($i in $csvdata){
    $i
    $(HasMultipleYears "$i")
    #write-host "$($i.case) $(Parse-DateString "$($i.date)")"
    # $pieces = $($i.date).Split("-")
    # $a = $pieces[0]
    # $b = $pieces[1]
    # $a = Parse-DateString "$a"
    # $a
    # if($b){
    #     $b = Parse-DateString "$b"
    # }else{Write-Host not b}
    # $b


}


