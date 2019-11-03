# syncfromsharepointonline
Sync a library (Office Templates) from sharepointOnline to local machine

# make sure SharePointPNP is installed and has some stored credentials
# Put this in a seperate GPO that runs once per machine
# 
# Install-Module SharePointPnPPowerShellOnline via the MSI package
# Add-PnPStoredCredential -Name https://yoursite.sharepoint.com -Username minimumaccess@tenant.dk -Password (ConvertTo-SecureString -string "insecurepassword" -AsPlainText -Force)
# make sure the office templates folder is set to $targetDIR\Skabeloner
