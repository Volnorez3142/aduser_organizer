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
***********************************************************************
* 1. If this is your first time running this script, input 1 and the  *
* script will create a directory and parse all Active Directory users'*
* data into one CSV file.                                             *
*                                                                     *
*                                 |||                                 *
*                                                                     *
* 2. If you already have an adusers.csv updated and ready in C:\by3142*
* input 2 and the program will start the corresponding attribute      *
* update process.                                                     *
*                                                                     *
*                                 |||                                 *
*                                                                     *
* 3. If your file is ready, but is elsewhere or has a different name, *
* input 3 to declare the filepath.                                    *
*                                                                     *
*                                 |||                                 *
*                          [ADDITIONAL INFO]                          *
* The CSV file should include 9 columns with corresponding info.      *
* Those columns are:                                                  *
* 1. SamAccountName - username (without the domain) [String]          *
* 2. UserPrincipalName - primary email address [String]               *
* 3. JobTitle - name of the new job title [String]                    *
* 4. Department - name of the new department [String]                 *
* 5. ManagerSamAccName - username of the manager [String]             *
* 6. Company - name of the new company [String]                       *
* 7. Mobile - user's mobile phone or messenger tag [String]           *
* 8. ProxyAddress1 - additional proxy address [String]                *
* 9. ProxyAddress2 - additional proxy address [String]                *
* 10. ProxyAddress3 - additional proxy address [String]               *
* ProxyAddress1/2/3 and manager fields can be left empty. All the     *
* other fields are mandatory.                                         *
***********************************************************************
Input: 
"@

Write-Host $switchcasemessage -ForegroundColor Cyan
$fileswitchcase = Read-Host
Write-Output "You've entered: $fileswitchcase"

if ($fileswitchcase -eq 1) {
    New-Item -Path "c:\" -Name "by3142" -ItemType "directory"
    Get-ADUser -Filter * -Properties CN, DisplayName, Title, Description, Department, Manager, Company, EmailAddress, MobilePhone, SamAccountName | Where { $_.Enabled -eq $True} | export-csv C:\by3142\adusers.csv
    Read-Host -Prompt "CHECK C:\BY3142 FOR THE USERS CSV. EDIT THE FIELDS AS NEEDED AND PRESS ANY KEY TO CONTINUE 1/2"
    Read-Host -Prompt "CHECK C:\BY3142 FOR THE USERS CSV. EDIT THE FIELDS AS NEEDED AND PRESS ANY KEY TO CONTINUE 2/2"
    $filepath = "C:\by3142\adusers.csv"
    Write-Output "Filepath = $filepath"
} elseif ($fileswitchcase -eq 2) {
    $filepath = "C:\by3142\adusers.csv"
    Write-Output "Filepath set to: $filepath"
} elseif ($fileswitchcase -eq 3) {
    Write-Output "Please enter the absolute file path below: "
    $filepath = Read-Host
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
*DOMAIN:
"@
Write-Host $domainvariable -ForegroundColor Cyan
$domainvariable = Read-Host

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
$safehouse = read-host
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

if ($safehouse -eq "yes") { 
    $adusers | ForEach-Object {
        if ($_.UserPrincipalName -like "*$domainvariable") {
            $error.clear()
            Try {
                $checksamaccname = Get-ADUser $_.SamAccountName
            } catch {
                #Catch function doesn't inherit from global environment so we gotta do an extra ifelse below
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
                        #still not inheriting.
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

                if ($_.Mobile) {
                    Set-ADUser -Identity $_.SamAccountName -Mobile $_.Mobile
                    Write-Host "Mobile:               $($_.Mobile)" -ForegroundColor DarkGray
                } elseif (!$_.Mobile){
                    #SKIPPING
                }

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

                Write-Host "DONE!" -ForegroundColor DarkGreen
                Write-Output "======================="
            }
        }
    }
} else {
    Write-Output "Huh?"
} 

Write-Host $usernotfoundlist -ForegroundColor Red
Write-Host $managernotsetlist -ForegroundColor Red
Write-Host $managernotfoundlist -ForegroundColor Red
Write-Host $incorrectsmtplist -ForegroundColor Red

#STOPPING THE TRANSCRIPT
Write-Host " "
Stop-Transcript

Read-Host -Prompt "
___.           ________  ____   _____ ________  
\_ |__ ___.__. \_____  \/_   | /  |  |\_____  \ 
 | __ <   |  |   _(__  < |   |/   |  |_/  ____/ 
 | \_\ \___  |  /       \|   /    ^   /       \ 
 |___  / ____| /______  /|___\____   |\_______ \
     \/\/             \/          |__|        \/
       END OF SCRIPT. PRESS ENTER TO EXIT.       
   THE TRANSCRIPT CAN BE FOUND ON THE DESKTOP.  
                        "
