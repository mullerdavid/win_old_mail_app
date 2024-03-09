
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