A mediocre attempt to automate the process of Active Directory user attribute and organization structure update.
Instructions unwrap throughout the script itself.
Not remotely ideal, but way-way better than updating 10+ amount of users manually.

The script refers to SamAccountName in order to grab the user from AD and change it's attributes, therefore the user should already be created before running the script.
Doesn't get triggered if the proxy address' domain is set wrong, or if user or user's manager isn't found, but rather lists at the very end both the wrong proxy addresses, wrong users and wrong managers.

Requires a CSV file with followed columns:
1. DisplayName – Not mandatory and doesn't interact with the script, but useful to the editor.
2. JobTitle – position name
3. UserPrincipalName – primary email address aka the SMTP
4. SamAccountName – the username itself
5. Department – name of the department
6. ManagerSamAccName – user's manager's SamAccountName (not mandatory; required to set up the organizational structure)
7. Company – name of the company
8. ProxyAddress1/2/3 – alias email addresses aka smtp (not to be confused with SMTP)

In a perfect scenario, the task of filling 90% of the CSV file can be passed to non-softheaded HR (rare), and upon making sure that everything is correct, the script can be started on any scale of a company.
Hope this saves you some time.
