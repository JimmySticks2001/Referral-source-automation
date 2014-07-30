#
#This script will automate the task of inserting referral sources into Avatar.
#

#Before running this script make sure the newest version of the PM Data Collection Workbook is being used and that the
# Validator has been run and all rows in the Referral_Source sheet have passed. If any errors make it through to this script the whole
# import process could be de-railed.
#Next, open Avatar to Referral Source Maintenance and don't click anything. 
#Run the .bat file associated with this script to start the program. It will open an open file dialog where the user 
# can select the PM Workbook which contains the list of guarantors that need to be entered. After the user chooses the
# workbook this script will read the Referral_Source sheet in the specified workbook. If there are any problems during this operation
# such as the workbook not being available to open, or the Referral_Source sheet not being found, the script will display a meesage box
# informing the user of the error and will safely exit.
#Next the script will inform the user that Avatar needs to be open and the Referral Source Maintenance form needs to be displayed. 
# It also needs to be the only form open or else the tab order will be incorrect. When the user presses the OK button on the message 
# box the script will being the automated process of entering in all of the guarantors listed in the Guarantors sheet in the workbook. 
# If the script needs to be stopped for any reason the user can close the command prompt window that the .bat file opened. This will 
# kill the automation process.


#Function to open an OpenFileDialog. This will allow the user to find the file in the filesystem rather than typing the path
# by hand or having the path hard-coded into the script. Takes a string representing the initial path to open to. Returns a 
# string representing a filepath.
Function Get-FileName($initialDirectory)
{
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}#end Get-FileName

#This function allows for easy removing of a file. It checks to make sure the filepath is valid before attempting to remove. 
# This eliminates the error caused by attempting to remove a file which might not exist. 
Function Remove-File($fileName)
{
    if(Test-Path -path $fileName)
    {
        Remove-Item -path $fileName
    }
}#end Remove-File

#This function checks to make sure that the string we want to send is Avatar approved and that it is able to be sent via WASP's
# Send-Keys function.
Function Send-String($string)
{
    if(!$string)    #if the string is null or empty, we need to output a space due to Send-Keys erroring due to a blank string
    {
        $string = " "
    }
    if($string -match "[(]")    #all parentheses need to be escaped 
    {
        $string = $string -replace "[(]", "{(}"
    }
    if($string -match "[)]")
    {
        $string = $string -replace "[)]", "{)}"
    }
    if($string -match ".+[ ].+")    #if the string contains a space between stuff, we need to build the apporpriate send keys string
    {
        $string = $string -replace " ", "+( )"
    }
    
    Select-Window javaw | Send-Keys $string
    Start-Sleep -Milliseconds $pauseTime

}#end Send-String



#----------------------------------End function definition-----------------------------------------------------#



#Load the Windows forms library so we can make open file dialogs and message boxes.
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

#Get current user account and set some paths.
$user = $env:USERNAME
$myDocs = [environment]::getfolderpath("mydocuments")
$toLoc = "$myDocs\WindowsPowerShell\WASP"
$fromLoc = "\\ntst.com\dfs\ProjectManagement\Share\Plexus\Plexus_Tools\WASP\*"


#Create the PowerShell modules directory, place the WASP automation .dll into the new directory, and unblock the file so this script
# can use it. If the path already exists, do nothing.
"checking for WASP module"
If (!(Test-Path $toLoc))
{
    "getting module from N:\ProjectManagement\Share\Plexus\Plexus_Tools\WASP"
    New-Item $toLoc -type directory -Force | Out-Null
    Copy-Item $fromLoc ($toLoc)
    "unblocking module"
    Unblock-File $toLoc\WASP.dll
}

#Load the module which should now be in the directory we created.
"loading WASP module"
Import-Module $toLoc

#Call the Get-FileName function defined above to get the filepath of the Excel workbook.
"awaiting user input"
$filePath = Get-FileName -initialDirectory "c:\fso"

#If the filepath came back empty, tell the user and end the script.
if(!$filePath)
{
    [System.Windows.Forms.MessageBox]::Show("Could not get the filepath of the chosen file. Ending script." , "Error")
    Exit
}

#Create a filepath for a temporary CSV file. This file will be used as temporary storage for the data stored in the Referral_Source tab
# in the workbook. It is faster for PowerShell to import data from a CSV file rather than importing directly from Excel. The
# temporary file will be placed in the user's temp directory and will be removed whenever Windows feels like it or at the end of 
# this script.
"creating temp csv filepath"
$csvFile = ($env:temp + "\" + ((Get-Item -path $filePath).name).Replace(((Get-Item -path $filePath).extension),".csv"))
Remove-File $csvFile                                #remove and previous version of the temporary file

"opening PM workbook"
$excel = new-object -comobject excel.application    #start a new Excel COM object
$excel.Visible = $False                             #make Excel visible or not
$workbook = $excel.Workbooks.Open($filePath)        #open the workbook from the specified path

"finding macroReferral_Source sheet"
#Loop throught each sheet to find the macroReferral_Source sheet. This way the sheet can move around and we can still find it.
foreach($worksheetIterator in $workbook.worksheets)
{
    $temp = $worksheetIterator.name
    
    #Print the current sheet to the console, this will show the user that something is happening as it takes a while to find the sheet.
    Write-Host "`r$temp                               " -NoNewLine

    #If we found the sheet named macroGuarantors, assign it to a variable, notify the user, and exit the loop.
    if($temp -eq "macroReferral_Source")
    {
        $worksheet = $worksheetIterator
        Write-Host ""
        Write-Host "found sheet"
        break
    }
}

#If we didn't find the sheet, inform the user and exit the script.
if($worksheet.name -ne "macroReferral_Source")
{
    [System.Windows.Forms.MessageBox]::Show("Could not find the macroReferral_Source tab in the chosen workbook. Ending script." , "Error")
    Exit
}

"creating temp csv file"
$worksheet.SaveAs($csvFile, 6)    #Save the Referral_Source worksheet as a CSV file
$workbook.Saved = $True           #Mark the workbook as saved so it won't prompt if you want to save before closing
"closing PM workbook"
$workbook.Close()                 #Close the workbook
$excel.Quit()                     #Close Excel
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null    #Release the Excel COM object because we have to
[System.GC]::Collect()            #Tell garbage collection to do it's thing
[System.GC]::WaitForPendingFinalizers()

"importing temp csv file into memory"
#Import the CSV file into an array.
$referralCSV = Import-Csv -path $csvFile

"removing temp csv file"
#Remove the csv file we created, we no longer need it.
Remove-File $csvFile 

"awaiting user input"
#Tell the user that Avatar should be open to the Referral Source Maintenance form.
$response = [System.Windows.Forms.MessageBox]::Show("Please start Avatar and open the Referral Source Maintenance form." + [Environment]::NewLine +
                                                    "Make sure it is the only form open." +
                                                    [Environment]::NewLine + [Environment]::NewLine + 
                                                    "Press OK when ready." , "Status", 1)

#If the user hit the cancel button, end the script.
if($response -eq "CANCEL")
{
    Exit
}

#Set the pause time to 1.5 seconds (1500 milliseconds) between commands.
$pauseTime = 1500

"beginning automated entry"
""
"guarantor:"
#Make Avatar the active window.
Select-Window javaw | Set-WindowActive | Out-Null 
Start-Sleep -Milliseconds $pauseTime

#Loop through each row in the CSV file. 
For($row = 0; $row -lt $referralCSV.Count; $row++)
{
    #If the first column in this row is empty, we are done with the sheet. We are assuming that the Guarantors sheet is valid and 
    # that all of the required columns are filled in. If the first required column is not filled out then we must have reached the
    # end of the completed rows. Break out of the loop.
    if(!$referralCSV[$row].1)
    {
        break
    }

    Send-String " "                   #Select Add
    Start-Sleep -Milliseconds $pauseTime
    Send-String $referralCSV[$row].1  #New referral source Code - 1
    Send-String "{TAB}"               #move to Referral Source - Name
    Send-String $referralCSV[$row].2  #Referral source Name - 2
    Send-String "{TAB}"

    #If Practitioner ID# is blank in the CSV we need to alter the way we move through the form. 
    if(!$referralCSV[$row].3)
    {
        Send-String "{TAB}"
    }
    else
    {
        Send-String $referralCSV[$row].3  #Practitioner ID# - 3
        Start-Sleep -Milliseconds $pauseTime
        Send-String "{ENTER}"
        Send-String "{TAB}"
    }
    
    Send-String $referralCSV[$row].4  #Referral Source - Speciality - 4
    Send-String "{TAB}"
    Send-String $referralCSV[$row].5  #Referral Source Category - 5
    Send-String "{TAB}"
    Send-String $referralCSV[$row].6  #Referral Source - Phone - 6
    Send-String "{TAB}"
    Send-String $referralCSV[$row].7  #Referral Source - Agency - 7
    Send-String "{TAB}"
    Send-String $referralCSV[$row].8  #Referral Source - Address Street 1 - 8
    Send-String "{TAB}"
    Send-String $referralCSV[$row].9  #Referral Source - Address Street 2 - 9
    Send-String "{TAB}"
    Send-String $referralCSV[$row].10 #Referral Source - Zipcode - 10
    Send-String "{TAB}"
    Send-String $referralCSV[$row].11 #Referral Source - City - 11
    Send-String "{ENTER}"
    Send-String "{TAB}"
    Send-String "{TAB}"
    Send-String "{TAB}"
    Send-String $referralCSV[$row].12 #Referral Source - State - 12
    Send-String "{ENTER}"

    Select-Window javaw | Send-Keys "+({TAB})"
    Start-Sleep -Milliseconds $pauseTime
    Select-Window javaw | Send-Keys "+({TAB})"
    Start-Sleep -Milliseconds $pauseTime
    Send-String " "
    Start-Sleep -Milliseconds $pauseTime
    Send-String " "
    Start-Sleep -Milliseconds $pauseTime
    Send-String "{TAB}"
    Send-String "{TAB}"
    Send-String "{TAB}"
    Send-String "{TAB}"

}#end row selection loop

#Inform the user that everything went swimmingly.
[System.Windows.Forms.MessageBox]::Show("Referral source entry complete." , "Status") | Out-Null

#All done!