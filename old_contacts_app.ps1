$dependencies = (
    "Microsoft.NET.Native.Runtime.1.7",
	"Microsoft.NET.Native.Framework.1.7",
	"Microsoft.NET.Native.Runtime.2.2",
	"Microsoft.NET.Native.Framework.2.2",
	"Microsoft.VCLibs.140.00"
)
foreach ($dep in $dependencies)
{
	Get-ChildItem ".\appx" -Filter "$dep*.Appx" | Add-AppxPackage -ForceApplicationShutdown
}
Add-AppxPackage ".\appx\Microsoft.People_2018.1130.2327.0_neutral_~_8wekyb3d8bbwe.AppxBundle"
