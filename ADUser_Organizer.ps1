#THIS SCRIPT ELEVATES THE STARTED PS SESSION TO ADMIN
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

#SETTING THE CONSOLE COLOR TO BLACK, STARTING THE TRANSCRIPT
CMD /C Color 0F
$DesktopPath = [Environment]::GetFolderPath("Desktop")
Start-Transcript -path $DesktopPath\ADConf.txt -append

#IMPORTING AD MODULE
Import-Module ActiveDirectory

#SWITCH-CASE SCENARIO
$switchcasemessage = @"
*******************************************************************************************
* 1. If this is your first time running this script, input 1 and the script will create a *
* directory and parse all Active Directory users' data into one CSV file.                 *
*                                                                                         *
*                                           |||                                           *
*                                                                                         *
* 2. If you already have an adusers.csv updated and ready in C:\by3142, input 2 and the   *
* program will start the corresponding attribute update process.                          *
*                                                                                         *
*                                           |||                                           *
*                                                                                         *
* 3. If your file is ready, but is elsewhere or has a different name, input 3 to declare  *
* the filepath.                                                                           *
*                                                                                         *
*                                           |||                                           *
*                                    [ADDITIONAL INFO]                                    *
* It is HEAVILY recommended to use either ModernCSV or LibreCalc in order to create or    *
* maintain the CSV file. Alternatively, you can select, copy and paste cells or entire    *
* tables into both those programs from Excel (including the web app).                     *
*                                                                                         *
* The CSV file itself should contain 12 columns with corresponding info.                  *
* Those columns are:                                                                      *
* 1. DisplayName - simply the name of the employee [String]                               *
* 2. JobTitle - the job title [String]                                                    *
* 3. UserPrincipalName - primary email address (aka the SMTP (login email)) [String]      *
* 4. SamAccountName - username (raw, without the domain) [String]                         *
* 5. Department - name of the department [String]                                         *
* 6. ManagerSamAccName - username of the manager (raw, without the domain) [String]       *
* 7. Company - name of the company or office [String]                                     *
* 8. Messenger - user's messenger username [String]                                       *
* 9. MobilePhone - user's mobile phone number [String]                                    *
* 10. ProxyAddress1 - additional proxy address or the smtp (format: smtp:email@email.com; *
* NOT to be confused with SMTP)                                                           *
* 11. ProxyAddress2 - additional proxy address or the smtp (format: smtp:email@email.com; *
* NOT to be confused with SMTP)                                                           *
* 12. ProxyAddress3 - additional proxy address or the smtp (format: smtp:email@email.com; *
* NOT to be confused with SMTP)                                                           *
* 13. TargetDN - target organizational unit's distinguishedName attribute (AS IS)         *
*                                                                                         *
* DisplayName, Manager, Messenger, MobilePhone, ProxyAddress1/2/3 and TargetDN fields can *
* be left empty since the script doesn't necessarily require them, at the same time       *
* collecting invalid inputs into those fields and listing them at the very end.           *
* All the other fields are mandatory.                                                     *
*******************************************************************************************
"@

Write-Host $switchcasemessage -ForegroundColor Cyan
$fileswitchcase = Read-Host -Prompt "ANSWER"
Write-Output "You've entered: $fileswitchcase"

if ($fileswitchcase -eq 1) {
    New-Item -Path "c:\" -Name "by3142" -ItemType "directory"
    Get-ADUser -Filter * -Properties CN, DisplayName, Title, Description, Department, Manager, Company, EmailAddress, MobilePhone, SamAccountName, DistinguishedName | Where { $_.Enabled -eq $True} | export-csv C:\by3142\adusers.csv
    Read-Host -Prompt "CHECK C:\BY3142 FOR THE USERS CSV. EDIT THE FIELDS AS NEEDED AND PRESS ANY KEY TO CONTINUE 1/2"
    Read-Host -Prompt "CHECK C:\BY3142 FOR THE USERS CSV. EDIT THE FIELDS AS NEEDED AND PRESS ANY KEY TO CONTINUE 2/2"
    $filepath = "C:\by3142\adusers.csv"
    Write-Output "Filepath = $filepath"
} elseif ($fileswitchcase -eq 2) {
    $filepath = "C:\by3142\adusers.csv"
    Write-Output "Filepath set to: $filepath"
} elseif ($fileswitchcase -eq 3) {
    $filepath = Read-Host -Prompt "ENTER FILE'S ABSOLUTE PATH"
    Write-Output "Filepath set to: $filepath"
} else {
    Read-Host -Prompt "Answer not found. Press Enter to exit"
    Exit
}

#DOMAIN VARIABLE
$domainvariable = @"
*************************************
* Enter your domain for Primary and *
* proxy email addresses.            *
* This will also be used to scope   *
* exclusively users of said domain  *
* when applying the script.         *
* Example: sakurada.lan             *    
*************************************
"@
Write-Host $domainvariable -ForegroundColor Cyan
$domainvariable = Read-Host -Prompt "DOMAIN"

#OBJECT MOVE TO TARGET DN PROMPT
$adobjmoveprompt = @"
**************************************
* Do you wish to active the cycle    *
* that changes AD object's DN's and  *
* moves them to target OU?           *
*                                    *
* This will only work on objects     *
* that have TargetDN variable set.   *
*                                    * 
*              [yes/no]              *
**************************************
"@
Write-Host $adobjmoveprompt -ForegroundColor Red
$adobjmoveprompt = Read-Host -Prompt "ANSWER"

#YESNO-SAFEHOUSE
$safehousemessage = @"
****************************************************
* Upon continuation, the script will run a         *
* Set-ADUser cycle for the entire AD. It is highly *
* advised that you make sure that everything is    *
* correct in the CSV file and as needed in the PS  *
* script.                                          *
*                                                  *
*       Do you wish to continue? [yes/no]          *
****************************************************
"@
Write-Host $safehousemessage -ForegroundColor Yellow
$safehouse = Read-Host -Prompt "ANSWER"
Write-Output "You've entered: $safehouse"

#DECLARING THE VARIABLES FOR ERROR MESSAGES, IMPORTING THE CSV FILE AND STARTING THE AD CONFIGURATION
$adusers = Import-Csv -Path $filepath
$proxyaddrErrormessage = "ERROR: WRONG DATA FORMAT. Correct data format: smtp:emailaddress$domainvariable"
$user404Errormessage = @"
**********************************
*        USER NOT FOUND!        *        
**********************************
*USER: 
"@
$usernotfoundlist = @"
**********************************
*        USERS NOT FOUND:        *        
**********************************
*LIST:

"@
$managernotsetlist = @"
**************************************
*        MANAGER NOT SET FOR:        *        
**************************************
*LIST:

"@
$managernotfoundlist = @"
*************************************
*        MANAGERS NOT FOUND:        *        
*************************************
*LIST:

"@ 
$incorrectsmtplist = @"
**********************************************
*        PROXY ADDRESS INCORRECT FOR:        *        
**********************************************
*LIST:

"@
$incorrectdnlist = @"
**********************************************
*     DISTINGUISHED NAMES INCORRECT FOR:     *        
**********************************************
*LIST:

"@

if ($safehouse -eq "yes") { 
    $adusers | ForEach-Object {
        if ($_.UserPrincipalName -like "*$domainvariable") {
            $error.clear()
            Try {
                $checksamaccname = Get-ADUser $_.SamAccountName
            } catch {
                #Catch function inherits from global environment only if you put $global: before the variable, and I don't want to do that here.
            }
            
            if ($error) {
                Write-Output " "
                Write-Host $user404Errormessage $_.SamAccountName -ForegroundColor Red
                Write-Output "======================="
                $usernotfoundlist = $usernotfoundlist + " $($_.SamAccountName) `n"
            } elseif (!$error) {
                Write-Output " "
                Write-Host "DisplayName | SAN:    $($_.DisplayName) | $($_.SamAccountName)"  -ForegroundColor Cyan
                Set-ADUser $_.SamAccountName -Description $_.JobTitle
                Set-ADUser $_.SamAccountName -Title $_.JobTitle
                Write-Host "Job title:            $($_.JobTitle)" -ForegroundColor Yellow
                Set-ADUser -Identity $_.SamAccountName -Add @{ProxyAddresses="SMTP:$($_.UserPrincipalName)"}
                Set-ADUser -Identity $_.SamAccountName -EmailAddress $_.UserPrincipalName
                Write-Host "Primary address:      $($_.UserPrincipalName)" -ForegroundColor DarkGray
                Set-ADUser $_.SamAccountName -Department $_.Department
                Write-Host "Department:           $($_.Department)" -ForegroundColor DarkYellow
                if (!$_.ManagerSamAccName) {
                    Write-Host "Manager:              MANAGER NOT SET" -ForegroundColor Red
                    $managernotsetlist = $managernotsetlist + " $($_.SamAccountName) `n"
                } else {
                    $error.clear()
                    Try {
                        $varmanagerinfo = Get-ADUser $_.ManagerSamAccName
                    } catch {
                        #still no desire.
                    }
                    
                    if ($error) {
                        Write-Host "Manager:              MANAGER NOT FOUND" -ForegroundColor Red
                        $managernotfoundlist = $managernotfoundlist + " $($_.ManagerSamAccName) FOR $($_.SamAccountName) `n"
                    } else {
                        $varmanagerdisplayname = $varmanagerinfo.Name
                        Set-ADUser $_.SamAccountName -Manager $_.ManagerSamAccName
                        Write-Host "Manager:              $varmanagerdisplayname |" $_.ManagerSamAccName -ForegroundColor Blue
                    }
                }
                Set-ADUser $_.SamAccountName -Company $_.Company
                Write-Host "Company:              $($_.Company)" -ForegroundColor Magenta

                if ($_.Messenger) {
                    Set-ADUser -Identity $_.SamAccountName -Mobile $_.Messenger
                    Write-Host "Messenger:            $($_.Messenger)" -ForegroundColor Gray
                } elseif (!$_.Messenger) {
                    Set-ADUser -Identity $_.SamAccountName -Mobile $null
                    Write-Host "Messenger:            NOT PRESENT; NULL'D!" -ForegroundColor DarkRed
                } else {
                    #SKIPPING
                }

                if ($_.MobilePhone) {
                    Set-ADUser -Identity $_.SamAccountName -OfficePhone "+$($_.MobilePhone)"
                    Write-Host "Work/Business phone:  +$($_.MobilePhone)" -ForegroundColor Gray
                } elseif (!$_.MobilePhone) {
                    Set-ADUser -Identity $_.SamAccountName -OfficePhone $null
                    Write-Host "Work/Business phone:  NOT PRESENT; NULL'D!" -ForegroundColor DarkRed
                } else {
                    #SKIPPING
                        #INFO: THE OFFICEPHONE ATTRIBUTE FROM AD MIRRORS TO TELEPHONENUMBER ATTRIBUTE AND THEN REPLICATES TO AAD'S BUSINESSPHONES ATTRIBUTE.
                        #HOWEVER, THIS IS NEITHER STANDARDIZED NOR DOCUMENTED, AND THEREFORE CAN BE "FIXED" ANY TIME BY MICROSOFT.
                            #ONCE AGAIN: OFFICEPHONE GOES TO TELEPHONENUMBER IN AD, TELEPHONE NUMBER REPLICATES TO BUSINESSPHONES IN AAD.
                                #YET SOMEHOW MICROSOFT IS CONSIDERED ENTERPRISE.
                                    #4 TRILLION USD MARKET CAPITALIZATION BTW.
                }

                #UNCOMMENT TO MERGE MOBILE PHONE AND MESSENGER INTO A SINGLE ATTRIBUTE IF OFFICEPHONE NO LONGER REPLICATES TO AAD
                #Set-ADUser -Identity $_.SamAccountName -OfficePhone $null #CLEARS TELEPHONENUMBER ATTRIBUTE IN AD AND BUSINESSPHONES IN AAD
                #if ($_.Messenger -and $_.MobilePhone) { 
                    #Set-ADUser -Identity $_.SamAccountName -Mobile "$($_.Messenger) | +$($_.MobilePhone)"
                    #Write-Host "Messenger:            $($_.Messenger)" -ForegroundColor DarkGray
                    #Write-Host "Phone number:         $($_.MobilePhone)" -ForegroundColor DarkGray
                #} elseif ($_.Messenger -and !$_.MobilePhone) {
                    #Set-ADUser -Identity $_.SamAccountName -Mobile $_.Messenger
                    #Write-Host "Messenger:            $($_.Messenger)" -ForegroundColor DarkGray
                #} elseif (!$_.Messenger -and $_.MobilePhone) {
                    #Set-ADUser -Identity $_.SamAccountName -Mobile $_.MobilePhone
                    #Write-Host "Phone number:         $($_.MobilePhone)" -ForegroundColor DarkGray
                #} else {
                    #SKIPPING
                #}

                if ($_.ProxyAddress1 -like "smtp:*@$domainvariable") {
                    Set-ADUser -Identity $_.SamAccountName -Add @{ProxyAddresses=$_.ProxyAddress1}
                    Write-Host "Proxy address 1:      $($_.ProxyAddress1)" -ForegroundColor DarkGray
                } elseif (!$_.ProxyAddress1){
                    #SKIPPING
                } else {
                    Write-Output " "
                    Write-Host $proxyaddrErrormessage -ForegroundColor Red
                    Write-Output "Current ProxyAddr~1:  $($_.ProxyAddress1)"
                    Write-Output " "
                    $incorrectsmtplist = $incorrectsmtplist + " $($_.ProxyAddress1) `n"
                }

                if ($_.ProxyAddress2 -like "smtp:*@$domainvariable") {
                    Set-ADUser -Identity $_.SamAccountName -Add @{ProxyAddresses=$_.ProxyAddress2}
                    Write-Host "Proxy address 2:      $($_.ProxyAddress2)" -ForegroundColor DarkGray
                } elseif (!$_.ProxyAddress2){
                    #SKIPPING
                } else {
                    Write-Output " "
                    Write-Host $proxyaddrErrormessage -ForegroundColor Red
                    Write-Output "Current ProxyAddr~2:  $($_.ProxyAddress2)"
                    Write-Output " "
                    $incorrectsmtplist = $incorrectsmtplist + " $($_.ProxyAddress2) `n"
                }

                if ($_.ProxyAddress3 -like "smtp:*@$domainvariable") {
                    Set-ADUser -Identity $_.SamAccountName -Add @{ProxyAddresses=$_.ProxyAddress3}
                    Write-Host "Proxy address 3:      $($_.ProxyAddress3)" -ForegroundColor DarkGray
                } elseif (!$_.ProxyAddress3){
                    #SKIPPING
                } else {
                    Write-Output " "
                    Write-Host $proxyaddrErrormessage -ForegroundColor Red
                    Write-Output "Current ProxyAddr~3:  $($_.ProxyAddress3)"
                    Write-Output " "
                    $incorrectsmtplist = $incorrectsmtplist + " $($_.ProxyAddress3) `n"
                }

                if ($adobjmoveprompt -eq "yes") {
                    if (!$_.TargetDN) {
                        Write-Host "TARGET DN:            EMPTY" -ForegroundColor DarkGray
                        #SKIPPING
                    } elseif ($_.TargetDN) {
                        Write-Output " "
                        Write-Host "MOVING USER TO:       $($_.TargetDN)" -ForegroundColor Blue
                        try {
                            Get-ADUser $_.SamAccountName | Move-ADObject -TargetPath $_.TargetDN -ErrorAction Stop
                            Write-Host "USER MOVED!" -ForegroundColor DarkGreen
                        } catch {
                            Write-Host "ERROR: DN INCORRECT OR NOT FOUND!" -ForegroundColor Red
                            $global:incorrectdnlist = $global:incorrectdnlist + " $global:SamAccountName `n"
                        }
                        $varuserdn = Get-ADUser $_.SamAccountName | Select DistinguishedName
                        $varuserdisplayname = Get-ADUser $_.SamAccountName | Select Name
                        if ($varuserdn[0].DistinguishedName -eq "CN=$($varuserdisplayname[0].Name),$($_.TargetDN)") {
                            Write-Host "DISTINGUISHED NAMES MATCHED!" -ForegroundColor DarkGreen
                        } else {
                            Write-Host "ERROR: DISTINGUISHED NAMES NOT MATCHED!" -ForegroundColor Red
                            Write-Host "USER'S DN:            $($varuserdn[0].DistinguishedName)" -ForegroundColor Red
                            Write-Host "TARGET DN:            CN=$($varuserdisplayname[0].Name),$($_.TargetDN)" -ForegroundColor Red
                            $incorrectdnlist = $incorrectdnlist + " $($_.SamAccountName) `n"
                        }
                    } else { 
                        Write-Host "TARGET DN ERROR. CHECK THE CSV FILE." -ForegroundColor Red
                        $incorrectdnlist = $incorrectdnlist + " $($_.SamAccountName) `n"
                    }
                } else {
                    #SKIPPING
                }

                Write-Output " "
                Write-Host "DONE!" -ForegroundColor Green
                Write-Output "======================="
            }
        }
    }
} else {
    Write-Output "Huh?"
} 

Write-Host $usernotfoundlist -ForegroundColor Red
Write-Output " "
Write-Host $managernotsetlist -ForegroundColor Red
Write-Output " "
Write-Host $managernotfoundlist -ForegroundColor Red
Write-Output " "
Write-Host $incorrectsmtplist -ForegroundColor Red
Write-Output " "
Write-Host $incorrectdnlist -ForegroundColor Red
Write-Output " "

$by3142 = @"
___.           ________  ____   _____ ________  
\_ |__ ___.__. \_____  \/_   | /  |  |\_____  \ 
 | __ <   |  |   _(__  < |   |/   |  |_/  ____/ 
 | \_\ \___  |  /       \|   /    ^   /       \ 
 |___  / ____| /______  /|___\____   |\_______ \
     \/\/             \/          |__|        \/
                 END OF SCRIPT.                
         https://github.com/Volnorez3142        
"@
Write-Host $by3142 -ForegroundColor Magenta

Write-Output " "
Write-Host "Do you want to run an AD-AAD syncrhonization cycle? [yes/no(else)]"
$synccycleprompt = Read-Host
if ($synccycleprompt -eq "yes") {
    Start-ADSyncSyncCycle -PolicyType Delta
    Read-Host -Prompt "END OF SCRIPT. THE TRANSCRIPT CAN BE FOUND ON THE DESKTOP. PRESS ENTER TO EXIT."
} else {
    #SKIPPING
    Read-Host -Prompt "END OF SCRIPT. THE TRANSCRIPT CAN BE FOUND ON THE DESKTOP. PRESS ENTER TO EXIT."
}

#STOPPING THE TRANSCRIPT
Stop-Transcript
