<#
  Author: Wayne Bellows)
  Source: https://github.com/NME/InstallM365Apps

  Enhancements:
  - Auto-detect Windows edition to choose single-session vs multi-session behavior.
  - Multi-session: SharedComputerLicensing=1, Updates=FALSE
  - Single-session: omit SharedComputerLicensing, Updates=TRUE
  - Applications param optional (empty = install full core suite)
  - Guardrails for Applications vs Type
  - Variable fix in ConfigureOfficeXML
#>

#description: Install or uninstall Microsoft Office applications (auto-detect AVD single vs multi-session)
#execution mode: Individual
#tags: Microsoft, Custom Image Template Scripts

[CmdletBinding()] Param (
    # Optional; leave empty to install the full core M365 Apps suite (no Visio/Project)
    [Parameter()]
    [ValidateSet("Word","PowerPoint","Access","Excel","OneNote","Outlook","Publisher","Visio","Project")]
    [string[]]$Applications = @(),

    [Parameter(Mandatory)]
    [ValidateSet("32", "64")]
    [string]$Version,

    [Parameter(Mandatory)]
    [ValidateSet("Add", "Remove")]
    [string]$Type
)

function AddProductsToConfigurationXML {
    [CmdletBinding()] Param (
        [Parameter(Mandatory)]
        [ValidateSet("Visio","Project")]
        [string[]]$Applications,

        [Parameter(Mandatory)]
        $xmlFile,

        [Parameter(Mandatory)]
        [string]$xmlFilePath,

        [Parameter(Mandatory)]
        [ValidateSet("32", "64")]
        [string]$Version
    )

    Begin {
        try {
            $addElement = $xmlFile.DocumentElement.Add
            if ($null -eq $addElement) { Throw "Not able to access the xml element" }
            $addElement.setAttribute("OfficeClientEdition", $Version)

            $VisioProductID   = "VisioProRetail"
            $ProjectProductID = "ProjectProRetail"
        } catch { $PSCmdlet.ThrowTerminatingError($PSitem) }
    }

    Process {
        try {
            foreach ($app in $Applications) {
                Write-Host " AVD AIB Customization Office apps: Request to add $app"

                $productElement  = $xmlFile.CreateElement("Product")
                $languageElement = $xmlFile.CreateElement("Language")
                $languageElement.setAttribute("ID", "MatchOS")
                $productElement.AppendChild($languageElement)

                if ($app -eq "Visio")   { $productElement.setAttribute("ID", $VisioProductID) }
                if ($app -eq "Project") { $productElement.setAttribute("ID", $ProjectProductID) }

                $addElement.AppendChild($productElement)
            }
        } catch { $PSCmdlet.ThrowTerminatingError($PSitem) }
    }

    End { $xmlFile.Save($xmlFilePath) }
}

function RemoveProductsFromConfigurationXML {
    [CmdletBinding()] Param (
        # Allow empty: no exclusions => full core suite
        [Parameter()]
        [AllowEmptyCollection()]
        [ValidateSet("Word","PowerPoint","Access","Excel","OneNote","Outlook","Publisher")]
        [string[]]$Applications = @(),

        [Parameter(Mandatory)]
        $xmlFile,

        [Parameter(Mandatory)]
        [string]$xmlFilePath,

        [Parameter(Mandatory)]
        [ValidateSet("32", "64")]
        [string]$Version
    )

    Begin {
        try {
            $addElement = $xmlFile.DocumentElement.Add
            $addElement.setAttribute("OfficeClientEdition", $Version)

            $productElement = $xmlFile.CreateElement("Product")
            $productElement.setAttribute("ID", "O365ProPlusRetail")

            $languageElement = $xmlFile.CreateElement("Language")
            $languageElement.setAttribute("ID", "MatchOS")
            $productElement.AppendChild($languageElement)
        } catch { $PSCmdlet.ThrowTerminatingError($PSitem) }
    }

    Process {
        try {
            foreach ($app in $Applications) {
                Write-Host " AVD AIB Customization Office apps: Request to remove $app"
                $excludeElement = $xmlFile.CreateElement("ExcludeApp")
                $excludeElement.setAttribute("ID", $app)
                $productElement.AppendChild($excludeElement)
            }
            $xmlFile.DocumentElement.Add.AppendChild($productElement) | Out-Null
        } catch { $PSCmdlet.ThrowTerminatingError($PSitem) }
    }

    End { $xmlFile.Save($xmlFilePath) }
}

function ConfigureOfficeXML($Applications, $xmlFile, $xmlFilePath, $Version, $Type) {
    if ($Type -eq "Add") {
        Write-Host " AVD AIB Customization Office apps: Adding office applications"
        AddProductsToConfigurationXML -Applications $Applications -xmlFile $xmlFile -xmlFilePath $xmlFilePath -Version $Version
    } else {
        Write-Host " AVD AIB Customization Office apps: Removing office applications"
        RemoveProductsFromConfigurationXML -Applications $Applications -xmlFile $xmlFile -xmlFilePath $xmlFilePath -Version $Version
    }
}

function Get-IsMultiSession {
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $pn = $cv.ProductName
        return ($pn -match '(?i)multi-?\s*session')
    } catch {
        # Fallback: assume single-session if we can't read
        return $false
    }
}

function New-ConfigXmlText {
    param(
        [Parameter(Mandatory)][bool]$IsMultiSession
    )
    # Choose update toggle and shared licensing property based on edition
    $updatesLine = if ($IsMultiSession) { '<Updates Enabled="FALSE" />' } else { '<Updates Enabled="TRUE" />' }
    $sharedLine  = if ($IsMultiSession) { '<Property Name="SharedComputerLicensing" Value="1" />' } else { '' }

@"
<Configuration>
  <Add Channel="MonthlyEnterprise">
  </Add>
  <RemoveMSI />
  $updatesLine
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  $sharedLine
</Configuration>
"@
}

function installOfficeUsingODT($Applications, $Version, $Type) {

    Begin {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Host "Starting AVD AIB Customization : Office Apps : $((Get-Date).ToUniversalTime())"

        # Guardrails for parameter misuse
        if ($Type -eq 'Remove' -and ($Applications -contains 'Visio' -or $Applications -contains 'Project')) {
            throw "When Type='Remove', Applications can only include core apps (Word, Excel, Outlook, PowerPoint, OneNote, Access, Publisher). Use Type='Add' to install Visio/Project."
        }
        if ($Type -eq 'Add' -and ($Applications | Where-Object {$_ -notin @('Visio','Project')})) {
            throw "When Type='Add', Applications can only include 'Visio' and/or 'Project'."
        }

        $isMulti = Get-IsMultiSession
        Write-Host ("Detected Windows edition: {0}" -f ($(if ($isMulti) { "AVD Multi-session" } else { "AVD Single-session" })))

        $configXML = New-ConfigXmlText -IsMultiSession:$isMulti

        $ODTDownloadLinkRegex = '/officedeploymenttool[a-z0-9_-]*\.exe$'
        $guid = [guid]::NewGuid().Guid
        $tempFolder = (Join-Path -Path "C:\temp\" -ChildPath $guid)
        $ODTDownloadUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
        $templateFilePathFolder = "C:\AVDImage"

        if (!(Test-Path -Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory | Out-Null
        }

        Write-Host "AVD AIB Customization Office Apps : Created temp folder $tempFolder"
    }

    Process {
        try {
            $HttpContent = Invoke-WebRequest -Uri $ODTDownloadUrl -UseBasicParsing

            if ($HttpContent.StatusCode -ne 200) { 
                throw "Office Installation script failed to find Office deployment tool link -- Response $($HttpContent.StatusCode) ($($HttpContent.StatusDescription))"
            }

            $ODTDownloadLinks = $HttpContent.Links | Where-Object { $_.href -Match $ODTDownloadLinkRegex }
            $ODTToolLink = $ODTDownloadLinks[0].href
            Write-Host "AVD AIB Customization Office Apps : Office deployment tool link is $ODTToolLink"

            $ODTexePath = Join-Path -Path $tempFolder -ChildPath "officedeploymenttool.exe"

            Write-Host "AVD AIB Customization Office Apps : Downloading ODT tool into folder $ODTexePath"
            $ODTResponse = Invoke-WebRequest -Uri "$ODTToolLink" -UseBasicParsing -UseDefaultCredentials -OutFile $ODTexePath -PassThru
            if ($ODTResponse.StatusCode -ne 200) { 
                throw "Office Installation script failed to download Office deployment tool -- Response $($ODTResponse.StatusCode) ($($ODTResponse.StatusDescription))"
            }

            Write-Host "AVD AIB Customization Office Apps : Extracting setup.exe into $tempFolder"
            Start-Process -FilePath $ODTexePath -ArgumentList "/extract:`"$($tempFolder)`" /quiet" -PassThru -Wait -NoNewWindow | Out-Null

            $setupExePath = Join-Path -Path $tempFolder -ChildPath 'setup.exe'
            
            # Construct XML config file for Office Deployment Tool setup.exe
            $xmlFilePath = Join-Path -Path $tempFolder -ChildPath 'installOffice.xml'

            Write-Host "AVD AIB Customization Office Apps : Saving xml content into xml file : $xmlFilePath"
            $configXML | Out-File -FilePath $xmlFilePath -Force -Encoding ascii
            
            [XML]$xmlDoc = Get-Content $xmlFilePath
            ConfigureOfficeXML -Applications $Applications -xmlFile $xmlDoc -xmlFilePath $xmlFilePath -Version $Version -Type $Type
            
            Write-Host "AVD AIB Customization Office Apps : Running setup.exe to download Office"
            $ODTRunSetupExe = Start-Process -FilePath $setupExePath -ArgumentList "/download $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden
            if (!$ODTRunSetupExe) { Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to download Office" }
            if ($ODTRunSetupExe.ExitCode) { Throw "AVD AIB Customization Office Apps : Exit code $($ODTRunSetupExe.ExitCode) returned from `"$setupExePath`" to download Office" }

            Write-Host "AVD AIB Customization Office Apps : Running setup.exe to Install Office"
            $InstallOffice = Start-Process -FilePath $setupExePath -ArgumentList "/configure $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $tempFolder -WindowStyle Hidden
            if (!$InstallOffice) { Throw "AVD AIB Customization Office Apps : Failed to run `"$setupExePath`" to install Office" }
            if ($InstallOffice.ExitCode) { Throw "AVD AIB Customization Office Apps : Exit code $($InstallOffice.ExitCode) returned from `"$setupExePath`" to install Office" }
        } catch {
            $PSCmdlet.ThrowTerminatingError($PSitem)
        }
    }

    End {
        if (Test-Path -Path $tempFolder -ErrorAction SilentlyContinue) {
            Remove-Item -Path $tempFolder -Force -Recurse -ErrorAction Continue
        }
        if (Test-Path -Path $templateFilePathFolder -ErrorAction SilentlyContinue) {
            Remove-Item -Path $templateFilePathFolder -Force -Recurse -ErrorAction Continue
        }

        $stopwatch.Stop()
        $elapsedTime = $stopwatch.Elapsed
        Write-Host "Ending AVD AIB Customization : Office Apps - Time taken: $elapsedTime"
    }
}

# Entry point
installOfficeUsingODT -Applications $Applications -Version $Version -Type $Type
