#region local modules 
#common functions
function WriteMessage([string]$msg, [System.ConsoleColor]$forecolor) {
    $currforecolor = $Host.UI.RawUI.ForegroundColor
    if ($forecolor -ne $null){
       $Host.UI.RawUI.ForegroundColor = $forecolor
    }
    Write-Output -InputObject $msg
    #set it back
    $Host.UI.RawUI.ForegroundColor = $currforecolor
}

#Used to start the Mage utility.  Can be used for other process executionss as well
function StartProcess([Parameter(Position=0, mandatory=$true)]
                      [string]$commandfileName, 
                      [string]$commandarguments,
                      [string]$allowedExitCodes = $null )
{

    If ($allowedExitCodes -eq $null) {
        $allowedExitCodes = ""
    }

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "$commandfileName" 
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true
    $pinfo.ErrorDialog = $false
    $pinfo.Arguments = "$commandarguments"

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null

    [string]$stdout = $p.StandardOutput.ReadToEnd()
    [string]$stderr = $p.StandardError.ReadToEnd()

    $p.WaitForExit()  
        
    $exitCode = $p.ExitCode

    $p.Close()
    $p.Dispose()  

    If($exitCode -ne 0)
    {   

       if ($allowedExitCodes.Contains($exitCode) -eq $false) {   
          if ($stderr+"" -eq "") {
            $stderr = "No error information was provided"
          }
          Write-Error -Message "$stderr"
          throw "Error in executing command $commandfileName $commandarguments :: $stderr" 
       }
       else {
          if ($stderr+"" -ne "") {
             WriteMessage -msg "$stderr" -forecolor DarkYellow
          }
          if ($stdout+"" -ne "") { 
             WriteMessage -msg $stdout
          }
       }
    }
    else
    {
        WriteMessage -msg $stdout
    }
}

# Mage tool - https://docs.microsoft.com/en-us/dotnet/framework/tools/mage-exe-manifest-generation-and-editing-tool
function GetMageUtil()
{
	$microsoftSDKFolder = "Microsoft SDKs\Windows"
	$mageUtil = "mage.exe"
	$windowsProgramx86 = ${env:ProgramFiles(x86)}
	$microsoftSDKLocation = Join-Path -Path $windowsProgramx86 -Child $microsoftSDKFolder
	$magePath = $null

	if ((Test-Path -Path $microsoftSDKLocation) -ne $true){
		WriteMessage -msg "GetMageUtil: The path $microsoftSDKLocation is not found and is required to deploy to production! Production deployments will fail." -forecolor Yellow
		return $null
	}
	else 
	{
		$childObjects = (Get-Item -Path (Join-Path -Path $microsoftSDKLocation -ChildPath v*)) | Sort-Object -Descending
		if ($childObjects.Length -le 0){
			WriteMessage -msg "GetMageUtil: The path $microsoftSDKLocation is empty. No SDK version folders were found and is required to deploy to production!" -forecolor Yellow
			return $null
		}
		else
		{
			$foundBin = $false
			foreach($childObject in $childObjects){
				$magePath = (Join-Path -Path $childObject.FullName -ChildPath "bin")
				If ((Test-Path -Path $magePath) -eq $true){
					$foundBin = $true
					break;
				}
			}
			if ($foundBin -eq $false){
				WriteMessage -msg "GetMageUtil: The SDK folders in $microsoftSDKLocation do not contain a bin folder where the mage utility is found. Please investigate.  Without the mage utility, production deployment is not possible!" -forecolor Yellow
				return $null
			}
			else
			{
				$childObjects = (Get-ChildItem -Directory $magePath  -Filter "*.*") 
				if ($childObjects.Length -le 0){
					WriteMessage -msg "GetMageUtil: The SDK folder $magePath is empty. Please investigate. I am trying to locate the mage utility. Without it, production deployment is not possible!" -forecolor Yellow
					return $null
				}
				else 
				{
                    $childObjects = Sort-Object -InputObject $childObjects -Descending
					if ((Test-Path -Path (Join-Path -Path $magePath -ChildPath $mageUtil) ) -eq $true){
						$magePath = Join-Path $magePath -ChildPath $mageUtil
						return $null
					}
					else
					{
						$foundMage = $false
						foreach($childObject in $childObjects){
                            $magePath = (Join-Path -Path $childObject.FullName -ChildPath $mageUtil)
							if ((Test-Path -Path $magePath) -eq $true){
								$foundMage = $true
								break
							}
						}
						if ($foundMage -eq $false){
							WriteMessage -msg "GetMageUtil: Could not find the MAGE utility in any folders within $magePath. Without the mage utility, production deployment is not possible!" -forecolor Yellow
							return $null
						}
					}
				}
			}
		}
	}
	return $magePath
}

#Compute has for SHA1.  
function ComputeHash( [string]$filePath ) {
  
  [System.IO.FileStream]$fs = $null
  Try
  {
	 $sha1Managed = New-Object System.Security.Cryptography.SHA1Managed
     $fs = New-Object System.IO.FileStream "$filePath", "Open", "Read" 
	 
	[byte[]] $hash = $sha1Managed.ComputeHash($fs)
	return ([System.Convert]::ToBase64String( $hash, 0, $hash.Length))
  }
  Finally
  {
     If ($fs -ne $null) 
	 {
		$fs.Close()
		$fs.Dispose()
	 }
  }
}

#Compute 256 hash
function ComputeHash256( [string]$filePath) {

  [System.IO.FileStream]$fs = $null
  Try
  {
     If ((Test-Path -Path $filePath) -ne $true) {
        throw "File $filePath not found"
     }

     $cryptoType = New-Object System.Security.Cryptography.SHA256Managed
     $fs = New-Object System.IO.FileStream "$filePath", "Open", "Read"
	 
	[byte[]] $hash = $cryptoType.ComputeHash($fs)
	return ([System.Convert]::ToBase64String( $hash, 0, $hash.Length))
	   
  }
  Finally
  {
     If ($fs -ne $null) 
	 {
		$fs.Close()
		$fs.Dispose()
	 }
  }
}

#Transfer the application settings from the environment config to the deployment configuration file
function UpdateAppSettings($destConfigContent, $envAppConfigContent) {

	$destAppsettings = $destConfigContent.configuration.appSettings
	$envAppsettings = $envAppConfigContent.configuration.appSettings

    $appSettingFilter = "//configuration/appSettings"
	foreach($appSetting in $envAppsettings.SelectNodes("//add")){
        $key = $appSetting.key
        $filter = ("{0}/add[@key='{1}']" -f $appSettingFilter, $key)
        $destSettingNode = $destConfigContent.SelectSingleNode($filter)
        if ($destSettingNode -eq $null){
           $newAppSetting = $destConfigContent.CreateElement("add")
           $newAppSetting = $destAppsettings.AppendChild($newAppSetting)
           $newAppSetting.SetAttribute("key", $key)
           $newAppSetting.SetAttribute("value", $appSetting.value)
           Write-Host -Object "WARNING: Setting Key $key was not found and added in destination config - $destConfigFile" -ForegroundColor Yellow
        } else {
           $destSettingNode.value = $appSetting.value
        }
	}
	$destConfigContent.Save($destConfigFile) 
}
#Transfer service model settings 
function UpdateServiceModelSettings($destConfigContent, $envAppConfigContent) { 
	$webServiceUrl = $envAppConfigContent.configuration.'system.servicemodel'.client.endpoint.address 
	if ($webServiceUrl+"" -ne ""){
		#Expected that destination deployment configuration file already contain the client endpoint elements
		#We are now updating the endpoint to match the environment
		$destConfigContent.configuration.'system.servicemodel'.client.endpoint.address = $webServiceUrl 
	}

}
#Update manifest to insert the hash from the updated files
function UpdateManifestHash($destAppManifestFile, $destBinaryLocation) {

    [xml]$manifestFileContent = (Get-Content -Path $destAppManifestFile)
    foreach($assemblyFile in $manifestFileContent.assembly.dependency){
       $codeBase = $assemblyFile.dependentAssembly.codebase 
       If ($codeBase+"" -eq ""){
           #element not found
           continue
       }
       $fileFullPath = (Join-Path -Path $destBinaryLocation -ChildPath ("{0}.deploy" -f $codeBase))    
       $configHash = ComputeHash -filePath  $fileFullPath
       $assemblyFile.dependentAssembly.hash.DigestValue = $configHash
    }

    foreach($assemblyFile in $manifestFileContent.assembly.file){
       $fileFullPath = (Join-Path -Path $destBinaryLocation -ChildPath ("{0}.deploy" -f $assemblyFile.name))    
       $configHash = ComputeHash -filePath  $fileFullPath 
       $assemblyFile.hash.DigestValue = $configHash
    }

    $manifestFileContent.Save($destAppManifestFile);
}

#endregion 
#-----------------------------------------------------------------------------------------------------------------------------------

#main
Get-Date
$pslocation =  Split-Path $MyInvocation.MyCommand.Path -Parent
WriteMessage -msg  ("Set pslocation={0}" -f $pslocation)

Try
{    
    $deploymentUrlFormat = "http://dell-dev64/Test/{0}"
#	########### validation ############################
    $mageUtil = GetMageUtil

    If ($mageUtil -ne $null){
        If ((Test-Path -Path $mageUtil) -ne $true){
			 $mageUtil = $null
        }
    }

	if ($mageUtil -eq $null){
		throw "Main::GetMageUtil - Could not find MAGE utility"
	}
	
	# 1.  Decide on environment (i.e. virtual directory where Test,QA,PROD reside )
    $deployFromPublishPath = Join-Path -Path $pslocation -ChildPath "Publish\bin"
    
    $environment = (Read-Host -Prompt "Which environment are you deploying (i.e. Test, QA, PROD)?").ToUpper() 
    While($environment -ne 'Test' -and $environment -ne 'QA' -and $environment -ne 'PROD'){
        $environment = (Read-Host -Prompt "Please respond with either 'Test', 'QA' or 'PROD'. Enter 'X' to exit.").ToUpper() 
        if ($environment -eq 'X'){
            return
        }
    }
    
    #2. Set the deployment url used to set the click-once binaries
    $deploymentUrl = ($deploymentUrlFormat -f $environment)
	
    #3.  Set the physical path corresponding to the url..
    $deployToPath = (Read-Host -Prompt "What is the UNC deployment path, also the IIS physical path for client installation, for the client?" )
    $result = (Test-Path -Path "$deployToPath")
    WriteMessage -msg $result
    if ((Test-Path -Path "$deployToPath") -eq $false) {
        WriteMessage -msg "Cannot find deployment path, $deployToPath.  Please create path and try again.  Make sure you are running as Administrator." -forecolor Red
        return 
    }

    #4. Prepare destination 
    WriteMessage "Removing previous deployment from $deployToPath" -forecolor Yellow 
    Remove-Item -Path "$deployToPath\*" -Verbose -Force -Recurse

    WriteMessage "Copying deployment from $deployFromPublishPath to $deployToPath" -forecolor Yellow 
	Copy-Item -Path "$deployFromPublishPath\*" -Destination $deployToPath -Force -Recurse
	WriteMessage "$environment deployment starting" -forecolor Yellow 
    
    #5. Setup 
    
	#5.1  file names - these are the expected file fnames 
    $applicationManifestFileName = "ClientApplication.application"
    $certFileName = "ClientApplication_TemporaryKey.pfx" 
	$appConfigENV =  ("App.config.{0}" -f $environment)
    $appManifestName = "ClientApplication.exe.manifest"
    $appConfigName = "ClientApplication.exe.config.deploy"
	$deployAppExecNameName = "ClientApplication.exe.deploy" 
    $setupExec = "setup.exe" 
 
    #5.2 pointers
    $applicationFileLocation = (Join-Path -Path $deployToPath -ChildPath ("Application Files"))
	$applicationFilePaths = (Get-ChildItem -Path "$applicationFileLocation") | Sort-Object -Descending
    $destBinaryLocation = $applicationFilePaths[0].FullName
	$destConfigFile = (Join-Path -Path $destBinaryLocation -ChildPath $appConfigName)
    $deploymentManifestFile = (Join-Path -Path $deployToPath -ChildPath $applicationManifestFileName);
    $destDeploymentManifestFile = (Join-Path -Path $destBinaryLocation -ChildPath $applicationManifestFileName)
    $certFileLocation = (Join-Path -Path $pslocation -ChildPath $certFileName) 
	$environmentAppConfigLocation = (Join-Path -Path $pslocation -ChildPath $appConfigENV)
    $destAppManifestFile = (Join-Path -Path $destBinaryLocation -ChildPath $appManifestName) 
    $destDeployAppExecFile = (Join-Path -Path $destBinaryLocation -ChildPath $deployAppExecName)
	$setupPath = (Join-Path -Path $deployToPath -ChildPath $setupExec)
    
    #5.3 Validation 
    If (( Test-Path -Path ($setupPath)) -ne $true){
		throw "Cannot locate - $setupPath. Copy failed."
	}
	If (( Test-Path -Path ($deploymentManifestFile)) -ne $true){
		throw "Cannot locate - $deploymentManifestFile.  Copy failed."
	}
    If (( Test-Path -Path ($destDeploymentManifestFile)) -ne $true){
		throw "Cannot locate - $destDeploymentManifestFile.  Copy failed."
	}
    If (( Test-Path -Path ($certFileLocation)) -ne $true){
		throw "Cannot locate - $certFileLocation.  Copy failed."
	}
    If (( Test-Path -Path ($destAppManifestFile)) -ne $true){
		throw "Cannot locate - $destAppManifestFile.  Copy failed."
	}
    If (( Test-Path -Path ($destDeployAppExecFile)) -ne $true){
		throw "Cannot locate - $destDeployAppExecFile.  Copy failed."
	}
	If (( Test-Path -Path ($environmentAppConfigLocation)) -ne $true){
		throw "Cannot locate - $environmentAppConfigLocation.  Copy failed."
	}
	If (( Test-Path -Path ($destConfigFile)) -ne $true){
		throw "Cannot locate - $destConfigFile.  Copy failed."
	}

	#5.4 change the setup objects in the destination to point to generated environment endpoint
    StartProcess -commandfileName $setupPath  -commandarguments " -url=""$deploymentUrl""" 

    #NOTE: Switch config content to the environment
	[xml]$envAppConfigContent = (Get-Content -Path $environmentAppConfigLocation)
	[xml]$destConfigContent = (Get-Content -Path $destConfigFile)

	#Update the service model settings - this only updates a WCF endpoint currently, if available
	UpdateServiceModelSettings -destConfigContent $destAppManifestFile  -envAppConfigContent $envAppConfigContent

	#Update the configuration appsettings section
    UpdateAppSettings -destConfigContent $destConfigContent -envAppConfigContent $envAppConfigContent 

    #5.5 Update hash in all binaries
    UpdateManifestHash -destAppManifestFile $destAppManifestFile -destBinaryLocation $destBinaryLocation

    #NOTE: In order to avoid MSB3113 (file not found) errors on the next step, I removed the ".deploy" on a copy.  
    #this is only temporary
    WriteMessage -msg "Create temp files used when updating cert on application manifest"

    [string[]]$versionDeployFiles = Get-ChildItem -File -Filter "*.deploy" -Path $destBinaryLocation -Recurse -Name
    $filesToCleanup = New-Object System.Collections.ArrayList

	#Copy the existing .deploy files to a file without the ".deploy" extension.  As mentioned, this is a work around.
    foreach($deployFile in $versionDeployFiles) {
        $nonDeployFile = (Join-Path -Path $destBinaryLocation -childpath ($deployFile.Replace(".deploy","")))
        $deployFile = (Join-Path -Path $destBinaryLocation -childpath $deployFile)

        $index = $filesToCleanup.Add($nonDeployFile)
        Copy-Item -Path $deployFile -Destination $nonDeployFile
    }

    #NOTE: Set the cert again - for the application manifest, there will be MSB3113 errors (file not found). 
    #It appears that this may be due to it looking for the exact referenced names. I try to fix this error 
	#by making a copy of the file without the ".deploy" extenstion.  However, if I see the errors, I just accepted them.  
	#Seems that the utility is able to update the manifest without an issue.
    StartProcess -commandfileName $mageUtil -commandarguments " -u ""$destAppManifestFile"" -cf ""$certFileLocation""" 

    WriteMessage -msg "Delete temp files created earlier"
    
    #5.6 Clean up temporary files to satisify cert update and avoid MSB3113 errors
    foreach($fileToCleanup in $filesToCleanup){
        Remove-Item -Path $fileToCleanup
    }

    WriteMessage -msg "Update deployment manifest" 
    
    #5.7 Update deployment manifest with application manifest hash
    $configHash = ComputeHash -filePath  $destAppManifestFile
    [xml]$manifestFileContent = (Get-Content -Path $deploymentManifestFile)
    $manifestFileContent.assembly.dependency.dependentAssembly.hash.DigestValue = $configHash 
    $manifestFileContent.Save($deploymentManifestFile);

    #5.8 Update cert for deployment manifest
    StartProcess -commandfileName $mageUtil -commandarguments " -u ""$deploymentManifestFile"" -cf ""$certFileLocation""" 

    #5.9 Copy the deployment manifest back to the version folder
    Copy-Item -Path $deploymentManifestFile -Destination $destDeploymentManifestFile -Force -Recurse

	WriteMessage -msg "Deployment for $environment complete!"
}
Catch
{
    Write-Error -Message ("Message: {0}, Exception: {1}" -f $Error[0].Message, $Error[0].Exception)
    $innerexception = $Error[0].InnerException

    While($innerexception) {
        Write-Error -Message ("InnerException:: {0}" -f $innerexception)
        $innerexception = $innerexception.InnerException
    }
}
Finally
{		
    pause
}
