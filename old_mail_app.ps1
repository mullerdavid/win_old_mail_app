# https://apps.microsoft.com/detail/9WZDNCRFHVQM?hl=hu-hu&gl=HU
# 9WZDNCRFHVQM
# microsoft.windowscommunicationsapps
# https://store.rg-adguard.net/
# without Outlook popup is microsoft.windowscommunicationsapps_16005.14326.21422.0_neutral_~_8wekyb3d8bbwe.appxbundle 


$bundlepath = Join-Path $PSScriptRoot "appx\microsoft.windowscommunicationsapps_16005.14326.21422.0_neutral_~_8wekyb3d8bbwe.appxbundle"
$MakePri = Join-Path $PSScriptRoot "bin\makepri.exe"

# Self elevate
if (!
    (New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )) {
    Start-Process -FilePath "powershell" -ArgumentList (
            "-File", $MyInvocation.MyCommand.Source, $args | %{ $_ }
        ) -Verb RunAs
    exit
}

# Helpers
function Expand-Archive-Custom {
	param(
        [Parameter(Mandatory=$true)][string]$archivePath,
        [Parameter(Mandatory=$true)][string]$destinationDir,
        [Parameter(Mandatory=$false)][string]$fileToExtract
    )
	Add-Type -Assembly System.IO.Compression.FileSystem
	
	$resolvedArchivePath = Convert-Path -LiteralPath $archivePath
	$resolvedDestinationDir = Convert-Path -LiteralPath $destinationDir
	$archive = [IO.Compression.ZipFile]::OpenRead( $resolvedArchivePath )
	try {
		if(!($fileToExtract)){
			[IO.Compression.ZipFileExtensions]::ExtractToDirectory( $archive, $resolvedDestinationDir )
		} else {
			if( $foundFile = $archive.Entries.Where({ $_.FullName -eq $fileToExtract }, 'First') ) {
				$destinationFile = Join-Path $resolvedDestinationDir $foundFile.Name
				[IO.Compression.ZipFileExtensions]::ExtractToFile( $foundFile[ 0 ], $destinationFile )
			}
			else {
				Write-Error "File not found in ZIP: $fileToExtract"
			}
		}
	}
	finally {
		if( $archive ) { $archive.Dispose() }
	}

}

function Force-Resolve-Path {
	param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string] $FileName
    )
    $FileName = Resolve-Path $FileName -ErrorAction SilentlyContinue -ErrorVariable _frperror
    if (-not($FileName)) {
        $FileName = $_frperror[0].TargetObject
    }
    return $FileName
}

function Escape-Arg {
	param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string] $Arg
    )
    return "`"$Arg`""
}
	
# Enable developer mode
$RegistryKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
if (-not(Test-Path -Path $RegistryKeyPath)) {
    New-Item -Path $RegistryKeyPath -ItemType Directory -Force | Out-Null
}
New-ItemProperty -Path $RegistryKeyPath -Name AllowDevelopmentWithoutDevLicense -PropertyType DWORD -Value 1 -erroraction SilentlyContinue | Out-Null

# remove packages
Get-AppxPackage Microsoft.windowscommunicationsapps -AllUsers | Remove-AppxPackage -AllUsers
Get-AppxPackage Microsoft.OutlookForWindows -AllUsers | Remove-AppxPackage -AllUsers


# install older version
# Add-AppxPackage -ForceApplicationShutdown -ForceUpdateFromAnyVersion -Path $bundlepath 

# install older version to path

$defaultInstallLocation = "c:\\Program Files\\WindowsCommunicationsApps"
$installLocation = Read-Host "Install folder [$($defaultInstallLocation)]"
$installLocation = ($defaultInstallLocation,$installLocation)[[bool]$installLocation]

Write-Output "Extracting..."

New-Item -ItemType Directory -Force -Path $installLocation -ErrorAction SilentlyContinue | Out-Null
New-Item -ItemType Directory -Force -Path $installLocation -Name "_appx" -ErrorAction SilentlyContinue | Out-Null
$appxpath = Join-Path $installLocation "_appx"
New-Item -ItemType Directory -Force -Path $installLocation -Name "_pri" -ErrorAction SilentlyContinue | Out-Null
$pripath = Join-Path $installLocation "_pri"
New-Item -ItemType Directory -Force -Path $installLocation -Name "_xml" -ErrorAction SilentlyContinue | Out-Null
$xmlpath = Join-Path $installLocation "_xml"
$manifestpath = Join-Path $installLocation "Appxmanifest.xml"

$bin = "outlookim_x64.appx"
Expand-Archive-Custom $bundlepath $appxpath
$binpath = Join-Path $appxpath $bin
Expand-Archive-Custom $binpath $installLocation
Remove-Item $installLocation -Recurse -Include @("AppxBlockMap.xml", "AppxSignature.p7x", "[[]Content_Types[]].xml", "AppxMetadata")

$(Get-Item "$($appxpath)\*" -Include "outlookim.language-*.appx") | ForEach-Object {
	$lang = [regex]::match($_.Name,"language-(.*).appx").Groups[1].Value
	New-Item -ItemType Directory -Force -Path $installLocation -Name "_tmp" -ErrorAction SilentlyContinue | Out-Null
	$tmppath = Join-Path $installLocation "_tmp"
	Expand-Archive-Custom $_.Fullname $tmppath
	Move-Item -Path "$($tmppath)\$($lang)" -Destination "$($installLocation)\$($lang)"
	Move-Item -Path "$($tmppath)\resources.pri" -Destination "$($pripath)\resources-$($lang).pri"
	Move-Item -Path "$($tmppath)\Appxmanifest.xml" -Destination "$($xmlpath)\Appxmanifest-$($lang).xml"
	Remove-Item $tmppath -Recurse -Force
}

Remove-Item $appxpath -Recurse -Force

Write-Output "Dumping resources..."
$priconfig=@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<resources targetOsVersion="10.0.0" majorVersion="1">
	<index root="\" startIndexAt="\">
		<default>
			<qualifier name="Language" value="en-US" />
		</default>
		<indexer-config type="folder" foldernameAsQualifier="true" filenameAsQualifier="true"
			qualifierDelimiter="." />
		<indexer-config type="PRI" />
		<indexer-config type="priinfo" />
	</index>
</resources>
"@
$priconfig | Out-File -FilePath "$($pripath)\priconfig.xml" -Encoding ascii

Copy-Item "$($installLocation)\resources.pri" -Destination "$($pripath)\resources.pri" | Out-Null
$MakePri = $($MakePri | Resolve-Path)
Copy-Item $MakePri -Destination "$($installLocation)\MakePri.pri" | Out-Null

$PriItem = Get-Item "$($pripath)\*" -Include "*.pri"
$procs = $(
	foreach ($Item in $PriItem) {
		$a = @(
			"dump",
			"/if",("$($pripath)\$($Item.Name)"|Force-Resolve-Path|Escape-Arg),
			"/o",
			"/es",("$($pripath)\resources.pri"|Resolve-Path|Escape-Arg),
			"/of",("$($pripath)\$($Item.Name).xml"|Force-Resolve-Path|Escape-Arg),
			"/dt","detailed"
		)
		Start-Process -NoNewWindow -RedirectStandardOutput ".\NUL" -PassThru $MakePri -ArgumentList $a
	}
)
$procs | Wait-Process

Write-Output "Creating pri from dumps...."
$a = @(
	"new",
	"/pr",($pripath|Resolve-Path|Escape-Arg),
	"/cf",("$($pripath)\priconfig.xml"|Resolve-Path|Escape-Arg),
	"/of",("$($installLocation)\resources.pri"|Force-Resolve-Path|Escape-Arg),
	"/mn",($manifestpath|Resolve-Path|Escape-Arg),
	"/o"
)
$procs = $(Start-Process -NoNewWindow -RedirectStandardOutput ".\NUL" -PassThru $MakePri -ArgumentList $a )
$procs | Wait-Process

$ProjectXmlFile = Get-Item $manifestpath
$ProjectXml = [xml](Get-Content $ProjectXmlFile)
$ProjectResources = $ProjectXml.Package.Resources;
$(Get-Item "$($xmlpath)\*" -Include "*.xml") | ForEach-Object {
    $($([xml](Get-Content $_)).Package.Resources.Resource) | ForEach-Object {
        $ProjectResources.AppendChild($($ProjectXml.ImportNode($_, $true))) | Out-Null
    }
}
$ProjectXml.Save($ProjectXmlFile.Fullname)


Write-Output "Installing/Registering...."
$ProjectXml.Package.Dependencies.PackageDependency | ForEach-Object {
	$dep = $_.Name
	$minver = [System.Version]$_.MinVersion
	$found = $false
	Get-AppxPackage $dep | Where-Object { $_.Architecture -eq "x64" } | ForEach-Object {
		if ([System.Version]$_.Version -ge $minver) {
			$found = $true
		}
	}
	if (-not $found) {
		Get-Item "$($bundlepath | Split-Path)\*" -Include "$($dep)*_x64_*.appx" | ForEach-Object {
			Add-AppxPackage $_.FullName
		}
	}
}

Add-AppxPackage -register $manifestpath

pause