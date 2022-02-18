$packagesPath = "K:\AosService\PackagesLocalDirectory\"
$desktopPath = [System.Environment]::GetFolderPath('Desktop');
$version = "10-0-23"
$securityOutputPath = -join($desktopPath, "\SecurityFiles_", $version);
$roleOutputPath = -join($securityOutputPath, "\AxSecurityRole");
$dutyOutputPath = -join($securityOutputPath, "\AxSecurityDuty");
$privOutputPath = -join($securityOutputPath, "\AxSecurityPrivilege");

if(-Not (Test-Path -Path $securityOutputPath)){
    [System.IO.Directory]::CreateDirectory($securityOutputPath);
}
if(-Not (Test-Path -Path $roleOutputPath)){
    [System.IO.Directory]::CreateDirectory($roleOutputPath);
}
if(-Not (Test-Path -Path $dutyOutputPath)){
    [System.IO.Directory]::CreateDirectory($dutyOutputPath);
}
if(-Not (Test-Path -Path $privOutputPath)){
    [System.IO.Directory]::CreateDirectory($privOutputPath);
}
    
Get-ChildItem -Path $packagesPath -Directory | 
ForEach-Object{
    $packageName = $_.Name
    $modulePath = -join($packagesPath, $packageName)
    Get-ChildItem -Path $modulePath -Directory |
    ForEach-Object{
        $subFolderName = $_.Name
        if($subFolderName -ne "bin" -and $subFolderName -ne "Descriptor" -and $subFolderName -ne "Resources" -and $subFolderName -ne "XppMetadata" -and $subFolderName -ne ".pkgrefgen")
        {
            $rolePath = -join($modulePath, "\", $subFolderName, "\AxSecurityRole")
            $dutyPath = -join($modulePath, "\", $subFolderName, "\AxSecurityDuty")
            $privPath = -join($modulePath, "\", $subFolderName, "\AxSecurityPrivilege")
            if(Test-Path -Path $rolePath){
                Get-ChildItem -Path $rolePath -File |
                ForEach-Object{
                    Copy-Item $_.FullName -Destination $roleOutputPath
                }
            }
            if(Test-Path -Path $dutyPath){
                Get-ChildItem -Path $dutyPath -File |
                ForEach-Object{
                    Copy-Item $_.FullName -Destination $dutyOutputPath
                }
            }
            if(Test-Path -Path $privPath){
                Get-ChildItem -Path $privPath -File |
                ForEach-Object{
                    Copy-Item $_.FullName -Destination $privOutputPath
                }
            }
        }
    }
    
}
read-host “End of Script... Press Enter to Continue”
