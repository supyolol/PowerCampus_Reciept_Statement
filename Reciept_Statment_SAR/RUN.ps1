

## clean up

Remove-Item .\emails.csv -ErrorAction SilentlyContinue
Remove-Item .\*.pdf -ErrorAction SilentlyContinue
Remove-Item .\ids_2.csv -ErrorAction SilentlyContinue
Remove-Item .\ids.csv -ErrorAction SilentlyContinue
Remove-Item PATH\ids.csv -ErrorAction SilentlyContinue







$Space = @"
########################################################################################################################
############################################   PRODUCTION DATABASE   ###################################################
########################################################################################################################
"@


$banner = @"
======================================================================================================================
======================================================================================================================
==========================================   SAR Reciept and Statement  ==============================================
================================================================================ By: Efrain Romero Ledesma ===========
======================================================================================================================


"@



function Show-Menu
 {
      param (
            [string]$Title = 'SAR Recipet Statement'
      )
      cls
      $banner
      
      
      Write-Host "   Type '1' For Batch Number."
      Write-Host "   Type '2' For Batch Number. Recepits Only."
      Write-Host "   Type '3' For Entry Date Range and specific Student."
      Write-Host "   Type '4' FERPA Recepit."
      Write-Host "   Type 'q' to quit."
 }


 do
{
      Show-Menu
      Write-Host " "
      $input = Read-Host "Please make a selection"
      switch ($input)
      {
            '1' {

cls
$Space

Write-Host "   "
Write-Host "   "

$ConfirmBATCH = $false
while(-not $ConfirmBATCH){
$batchinput = Read-Host "Enter BatchNumber"


#Confirm Batch Num

$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select BATCH from CHARGECREDIT
where BATCH = '$($batchinput)'
GROUP BY BATCH



"


If ($SQLData2){

   Write-Host "Batch Number Found!" -ForegroundColor Green
   $ConfirmBATCH = $true

}
else{

   Write-Host "Batch Number Not Found! Try again." -ForegroundColor Red
   $ConfirmBATCH = $false

}

}

## SQL CALL for data

$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT,FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate from CHARGECREDIT
where BATCH = '$($batchinput)' and CHARGE_CREDIT_TYPE in ('R')
GROUP BY PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT,ENTRY_DATE



"

## check for reocrds on STOP for reason COLLECTIONS
$ForLooplolnoob = $SQLData2 | select @{n='PEOPLE_ID' ;e={$_.PEOPLE_ORG_ID}}

$COLLrecords = foreach($id in $ForLooplolnoob){

$SQLData3 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "

select PEOPLE_ID from STOPLIST where PEOPLE_ID = '$($id.PEOPLE_ID)' and STOP_REASON = 'COLL'




"

$SQLData3

}









## Prompt User

Write-Host "   "
Write-Host "You'll be working with the following records."
Write-Host "You can remove records. Read the prompts below."
Write-Host "Records on STOP for collections will be highligted in Red"

write-host "   "
Write-Host "Batch Number: $($batchinput)"



## Logic to turn record/row red if on stop


$SQLData2 | Format-Table @{

    # Credit https://stackoverflow.com/questions/20705102/how-to-colorise-powershell-output-of-format-table - Jason Shirk
    # Credit https://www.valtech.com/insights/setting-a-row-colour-in-powershell--format-table/



    Label = 'PEOPLE_ORG_ID'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.PEOPLE_ORG_ID)${e}[0m"
    
    
    
    }


},
@{




    Label = 'ACADEMIC_YEAR'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.ACADEMIC_YEAR)${e}[0m"
    
    
    
    }


},
@{




    Label = 'ACADEMIC_TERM'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.ACADEMIC_TERM)${e}[0m"
    
    
    
    }


},
@{


    Label = 'CHARGE_CREDIT_CODE'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.CHARGE_CREDIT_CODE)${e}[0m"
    
    
    
    }


},
@{

    Label = 'AMOUNT'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.AMOUNT)${e}[0m"
    
    
    
    }


},
@{


    Label = 'entryDate'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.entryDate)${e}[0m"
    
    
    
    }


}





## Remove Records here
## Logic to remove students from the workflow.

$Userarray = @()
do {
    Write-Host "Once you have finished entering IDs to exclude.Type 'Done'." -ForegroundColor Yellow
 $input = (Read-Host "Please enter the student Ids who you wish to exclude out of sending Email Receipts too")
 
 if ($input -ne '') {$Userarray += $input}
}
until ($input -eq 'Done')

## 
#$remove = Import-Csv .\ids_2.csv 
$remove = $SQLData2

$finaldata = $remove | where {$Userarray  -notcontains $_.PEOPLE_ORG_ID}






#------------------------------------------------------------------
#---------NEW collection only logic -------------------------------
#------------------------------------------------------------------

$collectiondata = $remove | where {$Userarray  -contains $_.PEOPLE_ORG_ID}

if($collectiondata){

   

   Write-Host "                     "

   $collreocrdsONLY =  $collectiondata | select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,entryDate | sort -Unique * | Export-Csv .\idscoll.csv -NoTypeInformation

   ## Remove " 
   $file2= ".\idscoll.csv"
   (Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii

   ## Move ids.csv to python working dir. Needed for RUN.bat and MAIN.py to run.
   Copy-Item .\idscoll.csv -Destination PATH\ -Force

   ## Logging ids.csv
   Copy-Item .\idscoll.csv -Destination .\logs\
   Rename-Item .\logs\idscoll.csv -NewName "Collection_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv"


   #Display coll for End User
   Write-Host "The following collection records will only receive Recpeits."
   $collectiondata | select PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT | ft | Out-Host

   Write-Host " "
   Write-Host " "
   
   Read-Host "Press enter to Create Receipets for Collection Records." 

   
    ## Start Batch File to python Script

    Start-Process PATH\RUN_COLL.bat -Wait -NoNewWindow | Out-Null

    ## Move PDFs to Powershell working dir.
    Move-Item PATH\*.pdf `
    -Destination PATH\Recipet_Statement_SAR



    ## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

    $idstogetdata = Import-Csv .\idscoll.csv 

    ## iterate thru ids to get more info on student. 

    foreach($id in $idstogetdata) {

    ## SQL CALL to get more info on student.

    $SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


    Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

    "
    ## Creation of Ids.csv
    $SQLData9 | Export-Csv .\ids_2_coll.csv -NoTypeInformation -Append

    }


    ## Info displayed to console while running in order to remove student from the workflow
    Write-Host "   "
    Write-Host "The following students will get emails."
    $Datatoshow = Import-Csv .\ids_2_coll.csv 
    $Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME |Out-Host
    Write-Host "   "


    ##Display students getting emails

    Write-Host "   "
    Write-Host "Type 'Send Email' to start sending emails."


    ##Sending emails Logic


    $sendemailsquestionmark = Read-Host " "

    if($sendemailsquestionmark -eq 'Send Email'){



    foreach($id in $Datatoshow){
    $sql_DATA =    Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query "



    select EmailAddress.PeopleOrgId,EmailAddress.Email,PEOPLE.FIRST_NAME,PEOPLE.LAST_NAME 
    from EmailAddress
    left join PEOPLE
    on PEOPLE.PEOPLE_ID = EmailAddress.PeopleOrgId
    where PeopleOrgId = '$($id.PEOPLE_ID)'



    "
    $sql_DATA |  select PeopleOrgId,Email,FIRST_NAME,LAST_NAME | sort -Unique * | Export-Csv .\emails.csv -NoTypeInformation -Append 

    }
 
 
    $file2= ".\emails.csv"
    (Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


    $ReadyEmailData = Import-Csv .\emails.csv
    $ReadyEmailData | Export-Csv ".\logs\emails_coll_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

    Foreach($e in $ReadyEmailData){

    
            try{
                $From = "EMAIL"
                $To = $($e.Email)
                $Subject = "Receipt"
    $Body = "Hi $($e.FIRST_NAME),
        
Attached is your Receipt!

Thanks!




    "
        
                $SMTPServer = "*SERVER*"
                $SMTPPort = "*PORT*"
                $pwd = "*PASSWORD*"
                $username = "*USERNAME*"
                $securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
                $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
                $AttachFile = ".\$($e.PeopleOrgId)_billing_Receipt_$((get-date).ToString("MM-dd-yyyy")).pdf"

                Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $AttachFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
        
                Write-Host "Success! Email was sent to Email:$($e.Email), Name:$($e.FIRST_NAME) $($e.LAST_NAME), ID:$($e.PeopleOrgId)" -ForegroundColor Green
                Start-Sleep 1
        
            }
            catch{

                Write-Host "Failure! An error occured. Email was not sent to Email$($e.Email),Name:$($e.FIRST_NAME) $($e.LAST_NAME),ID: $($e.PeopleOrgId)" -ForegroundColor Red
            }

        }

    }
    else {


    Write-Host "Blah!"



    }

    ##logging
    $PDFsFiles = Get-ChildItem .\*.pdf | select Name

    Foreach($PDF in $PDFsFiles){
        $PDFfileName = $PDF | select -ExpandProperty Name
        Copy-Item ".\$($PDFfileName)" -Destination .\herstory 
        Rename-Item "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)" `
        -NewName "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).pdf"
    }

    ##clean up
    
    Remove-Item .\emails.csv -ErrorAction SilentlyContinue
    Remove-Item .\*.pdf -ErrorAction SilentlyContinue
    Remove-Item .\ids_2_coll.csv  -ErrorAction SilentlyContinue
    Remove-Item .\idscoll.csv -ErrorAction SilentlyContinue
    Remove-Item PATH\idscoll.csv -ErrorAction SilentlyContinue

    




 }

#------------------------------------------------------------------
#---------end of collection only logic back to main logic ---------
#------------------------------------------------------------------





## Display Records after removal.
Write-Host "   "
Write-host "Collection Records Process is Complete. Back to the Main Process." -ForegroundColor Yellow
Write-Host "   "
Write-Host "You'll be working with the following records."
$finaldata | select PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT | ft | Out-Host





#------------------------------------------------------------------
#---------NEW exclude logic -------------------------------
#------------------------------------------------------------------


$Userarray1 = @()
do {
    Write-Host "THE REAL EXCLUDE Once you have finished entering IDs to exclude.Type 'Done'." -ForegroundColor Yellow
 $input1 = (Read-Host "Please enter the student Ids who you wish to exclude out of sending Email Receipts too")
 
 if ($input1 -ne '') {$Userarray1 += $input1}
}
until ($input1 -eq 'Done')

## 
#$remove = Import-Csv .\ids_2.csv 
$remove = $finaldata


$finaldata2 = $remove | where {$Userarray1  -notcontains $_.PEOPLE_ORG_ID}


#------------------------------------------------------------------
#--------- end of NEW exclude logic -------------------------------
#------------------------------------------------------------------



## Creation of Ids.csv
$finaldata2 | select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,entryDate | sort -Unique * | Export-Csv .\ids.csv -NoTypeInformation




## Remove " 
$file2= ".\ids.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii



## Move ids.csv to python working dir. Needed for RUN.bat and MAIN.py to run.
Copy-Item .\ids.csv -Destination PATH\ -Force


## Logging ids.csv
$SQLData2 | Export-Csv ".\logs\OG_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Write-Host "   "
## Prompt user to run RUN.bat/MAIN.py script
Read-Host "Press Enter Create Statements and Receipts."
Write-Host "   "

## Start Batch File to python Script

Start-Process PATH\RUN.bat -Wait -NoNewWindow | Out-Null

## Move PDFs to Powershell working dir.

Move-Item PATH\*.pdf `
-Destination PATH\Recipet_Statement_SAR

Write-Host "   "

# cls clears screen on powershell console while running.

#cls

## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\ids.csv 

## iterate thru ids to get more info on student. 

foreach($id in $idstogetdata) {

## SQL CALL to get more info on student.

$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"
## Creation of Ids.csv
$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation -Append

}


## Info displayed to console while running in order to remove student from the workflow

Write-Host "The following students will get emails."
$Datatoshow = Import-Csv .\ids_2.csv 
$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME |Out-Host
Write-Host "   "


$finaldata | Export-Csv ".\logs\After_excluded_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

##Display students getting emails

Write-Host "   "
Write-Host "Type 'Send Email' to start sending emails."


##Sending emails Logic


$sendemailsquestionmark = Read-Host " "

if($sendemailsquestionmark -eq 'Send Email'){



foreach($id in $finaldata){
$sql_DATA =    Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query "



select EmailAddress.PeopleOrgId,EmailAddress.Email,PEOPLE.FIRST_NAME,PEOPLE.LAST_NAME 
from EmailAddress
left join PEOPLE
on PEOPLE.PEOPLE_ID = EmailAddress.PeopleOrgId
where PeopleOrgId = '$($id.PEOPLE_ORG_ID)'



"
$sql_DATA |  select PeopleOrgId,Email,FIRST_NAME,LAST_NAME  | sort -Unique * | Export-Csv .\emails.csv -NoTypeInformation -Append 

}
 
 
$file2= ".\emails.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


$ReadyEmailData = Import-Csv .\emails.csv
$ReadyEmailData | Export-Csv ".\logs\emails_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Foreach($e in $ReadyEmailData){

    
        try{
            $From = "*EMAILADDRESS*"
            $To = $($e.Email) 
            $Subject = "Receipt and Statement"
$Body = "Hi $($e.FIRST_NAME),
        
Attached is your Receipt and Statement!



"
        
            $SMTPServer = "*SERVER*"
            $SMTPPort = "*PORT*"
            $pwd = "*PASSWORD*"
            $username = "*USERNAME*"
            $securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
            $AttachFile = ".\$($e.PeopleOrgId)_billing_statment_$((get-date).ToString("MM-dd-yyyy")).pdf",".\$($e.PeopleOrgId)_billing_Receipt_$((get-date).ToString("MM-dd-yyyy")).pdf"

            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $AttachFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
        
            Write-Host "Success! Email was sent to Email:$($e.Email), Name:$($e.FIRST_NAME) $($e.LAST_NAME), ID:$($e.PeopleOrgId)" -ForegroundColor Green
            Start-Sleep 2
        
        }
        catch{

            Write-Host "Failure! An error occured. Email was not sent to Email$($e.Email),Name:$($e.FIRST_NAME) $($e.LAST_NAME),ID: $($e.PeopleOrgId)" -ForegroundColor Red
        }

    }

}
else {


Write-Host "Blah!"



}

##logging

$PDFsFiles = Get-ChildItem .\*.pdf | select Name

Foreach($PDF in $PDFsFiles){
    $PDFfileName = $PDF | select -ExpandProperty Name
    Copy-Item ".\$($PDFfileName)" -Destination .\herstory 
    Rename-Item "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)" `
    -NewName "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).pdf"
}

##clean up

Remove-Item .\emails.csv -ErrorAction SilentlyContinue
Remove-Item .\*.pdf -ErrorAction SilentlyContinue
Remove-Item .\ids_2.csv -ErrorAction SilentlyContinue
Remove-Item .\ids.csv -ErrorAction SilentlyContinue
Remove-Item PATH\ids.csv -ErrorAction SilentlyContinue



Read-Host "Press Enter to return to the Main Menu."



                


            } 


            '2'{ 
            
            
cls
$Space

Write-Host "   "
Write-Host "   "

$ConfirmBATCH = $false
while(-not $ConfirmBATCH){
$batchinput = Read-Host "Enter BatchNumber"


#Confirm Batch Num

$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select BATCH from CHARGECREDIT
where BATCH = '$($batchinput)'
GROUP BY BATCH



"


If ($SQLData2){

   Write-Host "Batch Number Found!" -ForegroundColor Green
   $ConfirmBATCH = $true

}
else{

   Write-Host "Batch Number Not Found! Try again." -ForegroundColor Red
   $ConfirmBATCH = $false

}

}

## SQL CALL for data

$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT,FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate from CHARGECREDIT
where BATCH = '$($batchinput)' and CHARGE_CREDIT_TYPE in ('R')
GROUP BY PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT,ENTRY_DATE



"

## check for reocrds on STOP for reason COLLECTIONS
$ForLooplolnoob = $SQLData2 | select @{n='PEOPLE_ID' ;e={$_.PEOPLE_ORG_ID}}

$COLLrecords = foreach($id in $ForLooplolnoob){

$SQLData3 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "

select PEOPLE_ID from STOPLIST where PEOPLE_ID = '$($id.PEOPLE_ID)' and STOP_REASON = 'COLL'




"

$SQLData3

}









## Prompt User

Write-Host "   "
Write-Host "You'll be working with the following records."
Write-Host "   "
Write-Host "Records on STOP for collections will be highligted in Red"

write-host "   "
Write-Host "Batch Number: $($batchinput)"



## Logic to turn record/row red if on stop


$SQLData2 | Format-Table @{

    # Credit https://stackoverflow.com/questions/20705102/how-to-colorise-powershell-output-of-format-table - Jason Shirk
    # Credit https://www.valtech.com/insights/setting-a-row-colour-in-powershell--format-table/



    Label = 'PEOPLE_ORG_ID'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.PEOPLE_ORG_ID)${e}[0m"
    
    
    
    }


},
@{




    Label = 'ACADEMIC_YEAR'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.ACADEMIC_YEAR)${e}[0m"
    
    
    
    }


},
@{




    Label = 'ACADEMIC_TERM'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.ACADEMIC_TERM)${e}[0m"
    
    
    
    }


},
@{


    Label = 'CHARGE_CREDIT_CODE'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.CHARGE_CREDIT_CODE)${e}[0m"
    
    
    
    }


},
@{

    Label = 'AMOUNT'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.AMOUNT)${e}[0m"
    
    
    
    }


},
@{


    Label = 'entryDate'
    Expression = 
    {
    
        if($COLLrecords.PEOPLE_ID -eq $_.PEOPLE_ORG_ID){
            
            $color = "31" #red
            
        }
        else{
        
            $color = "0" #white
        
        }
        $e = [char]27
        #"$e[${color}m$($_.Name)${e}[0m"
        "$e[${color}m$($_.entryDate)${e}[0m"
    
    
    
    }


}




$finaldata = $SQLData2


## Display Records after removal.
Write-Host "   "
Write-Host "   "
Write-Host "You'll be working with the following records."
$finaldata | select PEOPLE_ORG_ID,ACADEMIC_YEAR,ACADEMIC_TERM,CHARGE_CREDIT_CODE,AMOUNT | ft | Out-Host


## Creation of Ids.csv
$finaldata | select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,entryDate | sort -Unique * | Export-Csv .\idscoll.csv -NoTypeInformation

## Remove " 
$file2= ".\idscoll.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii

## Move ids.csv to python working dir. Needed for RUN.bat and MAIN.py to run.
Copy-Item .\idscoll.csv -Destination PATH\ -Force

## Logging ids.csv
$SQLData2 | Export-Csv ".\logs\OG_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Write-Host "   "
## Prompt user to run RUN.bat/MAIN.py script
Read-Host "Press Enter Receipts."
Write-Host "   "

## Start Batch File to python Script

Start-Process PATH\RUN_COLL.bat -Wait -NoNewWindow | Out-Null

## Move PDFs to Powershell working dir.

Move-Item PATH\*.pdf `
-Destination PATH\Recipet_Statement_SAR

Write-Host "   "

# cls clears screen on powershell console while running.

#cls

## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\idscoll.csv

## iterate thru ids to get more info on student. 

foreach($id in $idstogetdata) {

## SQL CALL to get more info on student.

$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"
## Creation of Ids.csv
$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation -Append

}


## Info displayed to console while running in order to remove student from the workflow

Write-Host "The following students will get emails."
$Datatoshow = Import-Csv .\ids_2.csv 
$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME |Out-Host
Write-Host "   "


$finaldata | Export-Csv ".\logs\After_excluded_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

##Display students getting emails

Write-Host "   "
Write-Host "Type 'Send Email' to start sending emails."


##Sending emails Logic


$sendemailsquestionmark = Read-Host " "

if($sendemailsquestionmark -eq 'Send Email'){



foreach($id in $finaldata){
$sql_DATA =    Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query "



select EmailAddress.PeopleOrgId,EmailAddress.Email,PEOPLE.FIRST_NAME,PEOPLE.LAST_NAME 
from EmailAddress
left join PEOPLE
on PEOPLE.PEOPLE_ID = EmailAddress.PeopleOrgId
where PeopleOrgId = '$($id.PEOPLE_ORG_ID)'



"
$sql_DATA |  select PeopleOrgId,Email,FIRST_NAME,LAST_NAME  | sort -Unique * | Export-Csv .\emails.csv -NoTypeInformation -Append 

}
 
 
$file2= ".\emails.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


$ReadyEmailData = Import-Csv .\emails.csv
$ReadyEmailData | Export-Csv ".\logs\emails_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Foreach($e in $ReadyEmailData){

    
        try{
            $From = "*EMAILADDRESS*"
            $To = $($e.Email) 
        
            $Subject = "Receipt and Statement"
$Body = "Hi $($e.FIRST_NAME),
        
Attached is your Receipt and Statement!

Thanks!



"
        
            $SMTPServer = "*SERVER*"
            $SMTPPort = "*PORT*"
            $pwd = "*PASSWORD*"
            $username = "*USERNAME*"
            $securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
            $AttachFile = ".\$($e.PeopleOrgId)_billing_Receipt_$((get-date).ToString("MM-dd-yyyy")).pdf"

            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $AttachFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
        
            Write-Host "Success! Email was sent to Email:$($e.Email), Name:$($e.FIRST_NAME) $($e.LAST_NAME), ID:$($e.PeopleOrgId)" -ForegroundColor Green
            Start-Sleep 1
        
        }
        catch{

            Write-Host "Failure! An error occured. Email was not sent to Email$($e.Email),Name:$($e.FIRST_NAME) $($e.LAST_NAME),ID: $($e.PeopleOrgId)" -ForegroundColor Red
        }

    }

}
else {


Write-Host "Blah!"



}

##logging

$PDFsFiles = Get-ChildItem .\*.pdf | select Name

Foreach($PDF in $PDFsFiles){
    $PDFfileName = $PDF | select -ExpandProperty Name
    Copy-Item ".\$($PDFfileName)" -Destination .\herstory 
    Rename-Item "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)" `
    -NewName "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).pdf"
}

##clean up

Remove-Item .\emails.csv -ErrorAction SilentlyContinue
Remove-Item .\*.pdf -ErrorAction SilentlyContinue
Remove-Item .\ids_2.csv -ErrorAction SilentlyContinue
Remove-Item .\ids.csv -ErrorAction SilentlyContinue
Remove-Item .\idscoll.csv -ErrorAction SilentlyContinue
Remove-Item PATH\ids.csv -ErrorAction SilentlyContinue



Read-Host "Press Enter to return to the Main Menu."
            
            
            
            
            
            
            
            
            
            
            }


            '3' {
                
cls
$Space

Write-Host "   "
Write-Host "   "
# Entry Date Valiation
$Valid1 = $false
while(-not $Valid1){
Write-Host "Enter Entry Date Range ie 05/03/2021"


$EntryDate1 = Read-Host "Entry Date 1"


try{
    [ValidatePattern("(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}")]
    $TEST = [string]$EntryDate1 = $EntryDate1
    if($TEST){
    
    #Write-Host "Good Date! $($EntryDate1)" -ForegroundColor Green
    $Valid1 = $true

    }else{
    
    
    $Valid1 = $false
    }
}
catch{
    Write-Host "Incorrect Date Format. Try this format 05/03/2021" -ForegroundColor Red
    continue
}



}




# Entry Date Valiation
$Valid2 = $false
while(-not $Valid2){
Write-Host "Enter Entry Date Range ie 05/03/2021"


$EntryDate2 = Read-Host "Entry Date 2"


try{
    [ValidatePattern("(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}")]
    $TEST2 = [string]$EntryDate2 = $EntryDate2
    if($TEST2){
    
    #Write-Host "Good Date! $($EntryDate2)" -ForegroundColor Green
    $Valid2 = $true

    }else{
    
    
    $Valid2 = $false
    }
}
catch{
    Write-Host "Incorrect Date Format. Try this format 05/03/2021" -ForegroundColor Red
    continue
}



}





Write-Host "Enter Student ID"



$checkid = $false
while(-not $checkid){

$StudentID = Read-Host "Student ID"
$SQLStudentID =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "

select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($StudentID)'



"

if($SQLStudentID ){
    Write-Host "$($SQLStudentID.PEOPLE_ID) is a vaild student ID" -ForegroundColor Green
    $checkid = $true
}
else{
    Write-host "Invalid Student ID. Try again." -ForegroundColor Red
   
}



}


$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate from CHARGECREDIT
where ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)' and PEOPLE_ORG_ID = '$($SQLStudentID.PEOPLE_ID)' and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO')
GROUP BY PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,ENTRY_DATE


"

# For Output/Console Display only
$SQLDataDISPLAYONLY =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "



select 
PEOPLE_ORG_CODE_ID,
ACADEMIC_TERM,
ACADEMIC_YEAR,
AMOUNT,
FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate 
from CHARGECREDIT
where 
(ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)'
and PEOPLE_ORG_CODE_ID = 'P$($SQLStudentID.PEOPLE_ID)' 
and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO'))
or
(ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)'
and PEOPLE_ORG_CODE_ID in (select ORG_CODE_ID from ORGANIZATION where ORG_NAME_2 = 'P$($SQLStudentID.PEOPLE_ID)' ) 
and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO'))
GROUP BY PEOPLE_ORG_CODE_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,ENTRY_DATE



"


Write-Host "   "
Write-Host "Entry Date range $($EntryDate1) - $($EntryDate2)"
write-host "Student Id : $($SQLStudentID.PEOPLE_ID)"
$SQLStudentID | Out-Host
$SQLDataDISPLAYONLY | select PEOPLE_ORG_CODE_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,entryDate | ft |Out-Host

$SQLData2 | select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,entryDate | sort -Unique * | Export-Csv .\ids.csv -NoTypeInformation
$SQLData2 | Export-Csv "PATH\Recipet_Statement_SAR\logs\OG_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

$file2= ".\ids.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


Copy-Item .\ids.csv -Destination PATH -force

$datacheck = '.\ids.csv'

If ((Get-Item $datacheck).Length -eq 0kb) {
    
    Write-Host "No Data Found for the Student and Entry Date you provided." -ForegroundColor Yellow
                        

    }
else {



Write-Host "   "
## Prompt user to run py script
Read-Host "Press Enter to Create Statement and Receipt."
Write-Host "   "

## Start Batch File to python Script

Start-Process PATH\RUN.bat -Wait -NoNewWindow | Out-Null

Write-Host "   "

#cls
## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\ids.csv 

foreach($id in $idstogetdata) {
#'$($id.PEOPLE_ORG_ID)'
$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"

$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation -Append

}

## copy over pdfs

Move-Item PATH\*.pdf `
-Destination PATH\Recipet_Statement_SAR\

#cls
## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\ids.csv 

foreach($id in $idstogetdata) {
#'$($id.PEOPLE_ORG_ID)'
$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"

$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation

}




Write-Host "The following students will get emails."
$Datatoshow = Import-Csv .\ids_2.csv 

$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME | `
 Export-Csv "PATH\Recipet_Statement_SAR\logs\ids_2_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Write-Host "   "

$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME |Out-Host

Write-Host "Type 'Send Email' to start sending emails."

$sendemailsquestionmark = Read-Host " "

if($sendemailsquestionmark -eq 'Send Email'){



foreach($id in $Datatoshow){
$sql_DATA =    Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query "



select EmailAddress.PeopleOrgId,EmailAddress.Email,PEOPLE.FIRST_NAME,PEOPLE.LAST_NAME 
from EmailAddress
left join PEOPLE
on PEOPLE.PEOPLE_ID = EmailAddress.PeopleOrgId
where PeopleOrgId = '$($id.PEOPLE_ID)'
--and EmailAddress.EmailType = 'CAMP'



"
$sql_DATA |  select PeopleOrgId,Email,FIRST_NAME,LAST_NAME  |  sort -Unique * | Export-Csv .\emails.csv -NoTypeInformation -Append 

}
 
 
$file2= ".\emails.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


$ReadyEmailData = Import-Csv .\emails.csv
$ReadyEmailData | Export-Csv ".\logs\emails_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Foreach($e in $ReadyEmailData){

    
        try{
           $From = "*EMAILADDRESS*"
      
            $To = $($e.Email)
        
            $Subject = "Receipt and Statement"
$Body = "Hi $($e.FIRST_NAME),
        
Attached is your Receipt and Statement!

Thanks!




"
                $SMTPServer = "*SERVER*"
                $SMTPPort = "*PORT*"
                $pwd = "*PASSWORD*"
                $username = "*USERNAME*"

            $securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
            $AttachFile = ".\$($e.PeopleOrgId)_billing_statment_$((get-date).ToString("MM-dd-yyyy")).pdf",".\$($e.PeopleOrgId)_billing_Receipt_$((get-date).ToString("MM-dd-yyyy")).pdf"

            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $AttachFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
        
            Write-Host "Success! Email was sent to Email:$($e.Email), Name:$($e.FIRST_NAME) $($e.FIRST_NAME), ID:$($e.PeopleOrgId)" -ForegroundColor Green
            Start-Sleep 2
        
        }
        catch{

            Write-Host "Failure! An error occured. Email was not sent to Email$($e.Email),Name:$($e.FIRST_NAME) $($e.FIRST_NAME),ID: $($e.PeopleOrgId)" -ForegroundColor Red
        }

    }

}


else {


Write-Host "Blah!" -ForegroundColor Red



}

#logging
$PDFsFiles = Get-ChildItem .\*.pdf | select Name

Foreach($PDF in $PDFsFiles){
    $PDFfileName = $PDF | select -ExpandProperty Name
    Copy-Item ".\$($PDFfileName)" -Destination .\herstory 
    Rename-Item "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)" `
    -NewName "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).pdf"
}


Remove-Item .\emails.csv -ErrorAction SilentlyContinue
Remove-Item .\*.pdf -ErrorAction SilentlyContinue
Remove-Item .\ids_2.csv -ErrorAction SilentlyContinue
Remove-Item .\ids.csv -ErrorAction SilentlyContinue
Remove-Item PATH\ids.csv -ErrorAction SilentlyContinue


Read-Host "Press Enter to return to the Main Menu."
                

}
                


            } 

            '4' {

cls
$Space

Write-Host "   "
Write-Host "   "
# Entry Date Valiation
$Valid1 = $false
while(-not $Valid1){
Write-Host "Enter Entry Date Range ie 05/03/2021"


$EntryDate1 = Read-Host "Entry Date 1"


try{
    [ValidatePattern("(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}")]
    $TEST = [string]$EntryDate1 = $EntryDate1
    if($TEST){
    
    #Write-Host "Good Date! $($EntryDate1)" -ForegroundColor Green
    $Valid1 = $true

    }else{
    
    
    $Valid1 = $false
    }
}
catch{
    Write-Host "Incorrect Date Format. Try this format 05/03/2021" -ForegroundColor Red
    continue
}



}




# Entry Date Valiation
$Valid2 = $false
while(-not $Valid2){
Write-Host "Enter Entry Date Range ie 05/03/2021"


$EntryDate2 = Read-Host "Entry Date 2"


try{
    [ValidatePattern("(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}")]
    $TEST2 = [string]$EntryDate2 = $EntryDate2
    if($TEST2){
    
    #Write-Host "Good Date! $($EntryDate2)" -ForegroundColor Green
    $Valid2 = $true

    }else{
    
    
    $Valid2 = $false
    }
}
catch{
    Write-Host "Incorrect Date Format. Try this format 05/03/2021" -ForegroundColor Red
    continue
}



}

write-host " "
Write-host "What Email address do you wish to use?"
$theotherEmail = Read-Host "Email Address"

# logging for alt email
$theotherEmail | Out-file ".\logs\ALT_emails_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).txt"


Write-Host "  "

Write-Host "Enter Student ID"



$checkid = $false
while(-not $checkid){

$StudentID = Read-Host "Student ID"
$SQLStudentID =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "

select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($StudentID)'



"

if($SQLStudentID ){
    Write-Host "$($SQLStudentID.PEOPLE_ID) is a vaild student ID" -ForegroundColor Green
    $checkid = $true
}
else{
    Write-host "Invalid Student ID. Try again." -ForegroundColor Red
   
}



}


$SQLData2 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate from CHARGECREDIT
where ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)' and PEOPLE_ORG_ID = '$($SQLStudentID.PEOPLE_ID)' and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO')
GROUP BY PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,ENTRY_DATE


"


### HERE
# For Output/Console Display only
$SQLDataDISPLAYONLY =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "



select 
PEOPLE_ORG_CODE_ID,
ACADEMIC_TERM,
ACADEMIC_YEAR,
AMOUNT,
FORMAT(ENTRY_DATE, 'yyyy-MM-dd') as entryDate 
from CHARGECREDIT
where 
(ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)'
and PEOPLE_ORG_CODE_ID = 'P$($SQLStudentID.PEOPLE_ID)' 
and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO'))
or
(ENTRY_DATE between '$($EntryDate1)' AND '$($EntryDate2)'
and PEOPLE_ORG_CODE_ID in (select ORG_CODE_ID from ORGANIZATION where ORG_NAME_2 = 'P$($SQLStudentID.PEOPLE_ID)' ) 
and CHARGE_CREDIT_CODE in ('DEFOUTSCH','CRPAYMENT','CRINTLDEPO'))
GROUP BY PEOPLE_ORG_CODE_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,ENTRY_DATE



"

### END of HERE


Write-Host "   "
Write-Host "Entry Date range $($EntryDate1) - $($EntryDate2)"
write-host "Student Id : $($SQLStudentID.PEOPLE_ID)"
$SQLStudentID | Out-Host
$SQLDataDISPLAYONLY | select PEOPLE_ORG_CODE_ID,ACADEMIC_TERM,ACADEMIC_YEAR,AMOUNT,entryDate | ft |Out-Host

$SQLData2 | select PEOPLE_ORG_ID,ACADEMIC_TERM,ACADEMIC_YEAR,entryDate | sort -Unique * | Export-Csv .\ids.csv -NoTypeInformation
$SQLData2 | Export-Csv "PATH\Recipet_Statement_SAR\logs\OG_ids_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

$file2= ".\ids.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


Copy-Item .\ids.csv -Destination PATH -force

$datacheck = '.\ids.csv'

If ((Get-Item $datacheck).Length -eq 0kb) {
    
    Write-Host "No Data Found for the Student and Entry Date you provided." -ForegroundColor Yellow
                        

    }
else {




Write-Host "   "
## Prompt user to run py script
Read-Host "Press Enter to Create FERPA Receipt."
Write-Host "   "

## Start Batch File to python Script

Start-Process PATH\RUN_FERPA.bat -Wait -NoNewWindow | Out-Null

Write-Host "   "

#cls
## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\ids.csv 

foreach($id in $idstogetdata) {
#'$($id.PEOPLE_ORG_ID)'
$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"

$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation -Append

}

## copy over pdfs

Move-Item PATH\*.pdf `
-Destination PATH\Recipet_Statement_SAR\

#cls
## Get SQL data from ids.csv for PEOPLE_ID,FIRST_NAME, LAST_NAME

$idstogetdata = Import-Csv .\ids.csv 

foreach($id in $idstogetdata) {
#'$($id.PEOPLE_ORG_ID)'
$SQLData9 =  Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*"  -Query "


Select PEOPLE_ID,FIRST_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id.PEOPLE_ORG_ID)'

"

$SQLData9 | Export-Csv .\ids_2.csv -NoTypeInformation

}




Write-Host "The following students will get emails."
$Datatoshow = Import-Csv .\ids_2.csv 

$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME | `
 Export-Csv "PATH\Recipet_Statement_SAR\logs\ids_2_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Write-Host "   "

$Datatoshow | select PEOPLE_ID,FIRST_NAME,LAST_NAME |Out-Host

Write-Host "Type 'Send Email' to start sending emails."

$sendemailsquestionmark = Read-Host " "

if($sendemailsquestionmark -eq 'Send Email'){



foreach($id in $Datatoshow){
$sql_DATA =    Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query "



select PEOPLE_ID,FIRST_NAME,LAST_NAME
from PEOPLE
where PEOPLE_ID = '$($id.PEOPLE_ID)'



"
$sql_DATA |  select PEOPLE_ID,FIRST_NAME,LAST_NAME   | Export-Csv .\FERPA_option_3.csv -NoTypeInformation -Append 

}
 
 
$file2= ".\FERPA_option_3.csv"
(Get-Content $file2) | Foreach-Object {$_ -replace '"', ''}|Out-File $file2 -Encoding ascii


$ReadyEmailData = Import-Csv .\FERPA_option_3.csv
$ReadyEmailData | Export-Csv ".\logs\FERPA_option_3_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).csv" -NoTypeInformation

Foreach($e in $ReadyEmailData){

    
        try{
            $From = "*EMAILADDRESS*"
            $To = $theotherEmail
        
            $Subject = "Receipt and Statement"
$Body = "Hi $($e.FIRST_NAME),
        
Attached is your Receipt and Statement!

Thanks!


"
        
                $SMTPServer = "*SERVER*"
                $SMTPPort = "*PORT*"
                $pwd = "*PASSWORD*"
                $username = "*USERNAME*"

            $securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
            $AttachFile = ".\$($e.PEOPLE_ID)_billing_Receipt_$((get-date).ToString("MM-dd-yyyy")).pdf"

            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $AttachFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
        
            Write-Host "Success! Email was sent to Email:$($theotherEmail), Name:$($e.FIRST_NAME) $($e.LAST_NAME), ID:$($e.PEOPLE_ID)" -ForegroundColor Green
            Start-Sleep 2
        
        }
        catch{

            Write-Host "Failure! An error occured. Email was not sent to Email$($theotherEmail),Name:$($e.FIRST_NAME) $($e.LAST_NAME),ID: $($e.PEOPLE_ID)" -ForegroundColor Red
        }

    }

}


else {


Write-Host "Blah!" -ForegroundColor Red



}

#logging
$PDFsFiles = Get-ChildItem .\*.pdf | select Name

Foreach($PDF in $PDFsFiles){
    $PDFfileName = $PDF | select -ExpandProperty Name
    Copy-Item ".\$($PDFfileName)" -Destination .\herstory 
    Rename-Item "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)" `
    -NewName "PATH\Recipet_Statement_SAR\herstory\$($PDFfileName)_$((get-date).ToString("yyyy_MM_dd_hh_mm_ss")).pdf"
}


Remove-Item .\FERPA_option_3.csv -ErrorAction SilentlyContinue
Remove-Item .\*.pdf -ErrorAction SilentlyContinue
Remove-Item .\ids_2.csv -ErrorAction SilentlyContinue
Remove-Item .\ids.csv -ErrorAction SilentlyContinue
Remove-Item PATH\ids.csv -ErrorAction SilentlyContinue


Read-Host "Press Enter to return to the Main Menu."
                    
}          

            } 

            'q' {
                 return
            }
      }
      pause
 }
 until ($input -eq 'q')





