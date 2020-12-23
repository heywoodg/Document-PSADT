# Document-PSADT

This is a PowerShell function that when passed PowerShell Application Deployment Toolkit file, will extract out the relevent information and place it in a MS Word document. So if you have a lot of PSADT files for distributing applications, you can get docuemntation created automatically.

The created Word document will include the following:
1) Details of the application to be installed
2) Pre-Installation steps, specifically any commands that have been added to the pre-installation section. 
3) Installation tasks that are required for the applcation to be installed
4) Post-Installation details, commands that are run after the application has been installed.
5) Pre-Uninstallation, specifically any commands that have been added to the pre-uninstallation section. 
6) Uninstallation tasks that are performed when the application is removed
7) Post-Uninstallation details, commands that are run after the application has been uninstalled.

The document is then saved at the end of the function.
