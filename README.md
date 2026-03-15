A mediocre attempt to automate the process of Active Directory user attribute and organization structure update.
Instructions unwrap throughout the script itself.
Not remotely ideal, but way better than updating 10+ amount of users manually, and saves hours of time that can be better channeled elsewhere.

**ModernCSV or LibreCalc is recommended to build and maintain the CSV file. Alternatively, both those programs can past cells from XLSX files.**
The reason is simple: Excel, upon opening a CSV file, "lowers the accuracy" of 15+ digit numbers (a phone number, for example), effectively transforming it into a text that states 9E+-something-something.

**As any other Powershell file, you might need to unlock this before running. Is done through RMB > Properties > Unlock.**

The CSV file itself should contain 12 columns with corresponding info.
Those columns are: 
1. DisplayName - simply the name of the employee      
2. JobTitle - the job title      
3. UserPrincipalName - primary email address (aka the SMTP (login email))
4. SamAccountName - username (raw, without the domain)
5. Department - name of the department    
6. ManagerSamAccName - username of the manager (raw, without the domain; required to set up the organizational structure) 
7. Company - name of the company or office 
8. Messenger - user's messenger username
9. MobilePhone - user's mobile phone number 
10. ProxyAddress1 - additional proxy address or the smtp (format: smtp:email@email.com; NOT to be confused with SMTP) 
11. ProxyAddress2 - additional proxy address or the smtp (format: smtp:email@email.com; NOT to be confused with SMTP) 
12. ProxyAddress3 - additional proxy address or the smtp (format: smtp:email@email.com; NOT to be confused with SMTP)
13. TargetDN - target organizational unit's distinguishedName attribute (should be copied AS IS)

DisplayName, Manager, Messenger, MobilePhone, ProxyAddress1/2/3 and TargetDN fields can be left empty since the script itself doesn't necessarily require them, at the same time collecting invalid inputs into those fields and listing them at the very end.
All the other fields are mandatory. 

The script refers to SamAccountName in order to grab the user from AD and change it's attributes, therefore the user should already be created before running the script.
Doesn't trigger/crash if the proxy address' domain is set wrong, or if the user or user's manager isn't found, but rather lists at the very end both the wrong proxy addresses, wrong users and wrong managers.

In a perfect scenario, the task of filling 90% of the CSV file can be passed to non-softheaded HR (rare), and upon making sure that everything is correct, the script can be started on any scale of a company.
Hope this saves you some time.

As a matter of principle, **no AI was used upon writing this.**
