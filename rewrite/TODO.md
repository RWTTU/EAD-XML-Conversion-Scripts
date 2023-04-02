# XML 

## fix relation
- [v]
Fix container and did relationship
did > container > unit title

## Add attrib

add unit date era 
`era="ce"`


## add warning 
- [v]
Warn after 7 c

## look at adding in this warning 
- [v]
```powershell
 $clevel = $i.'c0#'
    
        if($i.Attribute -eq "" -and $i.'c0#' -eq "" -and $i.Title -eq ""){continue}
            
        if($i.Attribute -eq "" -or $i.'c0#' -eq "" -or $i.Title -eq ""){
            
            write-host "Error: Required record information missing for record around line $record" -BackgroundColor Red  -ForegroundColor Black;
            write-host "Deleting partial file..."
            rm $outfile;
            
            Read-Host 'Press enter to close' | Out-Null
            Exit   
            }
            
```

- [v]
```powershell
if(!$i.'Series ID' -and $i.Attribute -eq "series"){Write-Host "Warning: Series ID Missing for record on line $record - $($i.Title)" -BackgroundColor red -ForegroundColor Black }    
```


## Dspace urls 
-[v]
```powershell
  
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
                
                #Link Title
                $("<unittitle><extref xmlns:xlink=""http://www.w3.org/1999/xlink"" xlink:type=""simple""
                xlink:show=""new"" xlink:actuate=""onRequest""  
                xlink:href=""$($i."Dspace URL")"">$($i.Title+" ")<unitdate era=""ce""
                calendar=""gregorian"" normal=""$(codedDate $i.Date $minyear $maxyear)"">$($i.Date)</unitdate></extref></unittitle>
                </did>") >> $outfile}else{
                
                #No Link Title
                $("<unittitle>$($i.Title+" ")<unitdate era=""ce""
                calendar=""gregorian"" normal=""$(codedDate $i.Date $minyear $maxyear)"">$($i.Date)</unitdate></unittitle>
                </did>") >> $outfile} 
             }
```