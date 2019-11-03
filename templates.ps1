###################################################
#                                                 #
# Pull a library from SharePoint to local machine #
#                                                 #
###################################################

function get-Templates()
{
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, HelpMessage="Local target for templates")]
        [String]
        [ValidateScript({If(Test-Path $_){$true}else{Throw "Invalid path given: $_"}})] 
        $location,
        [Parameter(Mandatory=$true,ValueFromPipeline, HelpMessage="SharePoint Online source site URL")]
        [String]
        $sourceUrl,
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [String]
        $subSiteRelativePath,
         [Parameter(Mandatory=$true,ValueFromPipeline)]
        [String]
        $DocumentLibraryName
    )
    Process{
         Write-Host "###################### Pull templates From SharePoint Online ######################" -ForegroundColor Green

        $path = $location.TrimEnd('\')

        Write-Host "Provided Source URL :"$sourceUrl -ForegroundColor Green
        Write-Host "Provided Document Library :"$DocumentLibraryName -ForegroundColor Green
        Write-Host "Provided Path :"$path -ForegroundColor Green

         try{
         
                 $Folder = $path.Trim() ;
                 
                 $traceLog = $path.Trim()+'\traceoutput.txt'

                 $sourceMappedLocation = 'SPO:\' 
                 if($subSiteRelativePath -ne "")
                 {
                    $sourceMappedLocation = 'SPO:\' + $subSiteRelativePath
                 }

             
                 Set-PnPTraceLog -Off # -LogFile $traceLog -Level Debug
                 Connect-PnPOnline -Url $sourceUrl -CreateDrive # -CurrentCredentials -Credentials $credentials

                cd $sourceMappedLocation 
                pwd

                Write-Host "Downloading documents from Document Library.." -ForegroundColor Cyan
                Copy-PnpItemProxy  -r -force $DocumentLibraryName $Folder

                Set-PnPTraceLog -Off
               # cd $location 
               # pwd

            }
            catch{
             Write-Host $_.Exception.Message -ForegroundColor Red
            }

    }
}

#####################################################
#
# CONFIG
#
######################################################

    #Target Directory for templates
    $targetDIR = $env:APPDATA + '\Microsoft\Templates'

#####################################################
#
# SCRIPT EXECUTION
#
######################################################

# make sure SharePointPNP is installed and has some stored credentials
# Put this in a seperate GPO that runs once per machine
# 
# Install-Module SharePointPnPPowerShellOnline via the MSI package
# Add-PnPStoredCredential -Name https://yoursite.sharepoint.com -Username minimumaccess@tenant.dk -Password (ConvertTo-SecureString -string "insecurepassword" -AsPlainText -Force)
# make sure the office templates folder is set to $targetDIR\Skabeloner

# Remove folder before getting the latest version
if (Test-Path $targetDIR\Skabeloner) { Get-ChildItem $targetDIR -Recurse | Remove-Item -Force;}
New-Item -ItemType Directory -Force -Path $targetDIR

# getting latest version of the Templates
get-Templates -sourceUrl "https://yoursite.sharepoint.com/" -subSiteRelativePath 'SubPath' -DocumentLibraryName "path/to/templates" -location $targetDIR
