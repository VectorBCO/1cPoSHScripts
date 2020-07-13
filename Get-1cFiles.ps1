$CPath = 'AppData\Roaming\1C\1CEStart\ibases.v8i'
$NotPingableComputers = "C:\Tmp\not_pingable_computers.txt"

Function Parse-Config {
    param (
        $ConfigFilePath
    )
    $dbfn = ''
    $srv = ''
    $db = ''
    $ConfigContent = Get-Content $ConfigFilePath -Encoding UTF8
    $DBs = $ConfigContent | ForEach-Object { 
        if ($_ -match '(?''dbFriendlyName''^\[.+\]$)'){
            $dbfn = $Matches['dbFriendlyName']
        }
        if ($_ -match 'Connect=Srvr="(?''Srv''[^"]+)";Ref="(?''db''[^"]+)";'){
            $srv = $Matches['srv']
            $db = $Matches['DB']
        }
        if ($db -and $dbfn -and $srv){
            [pscustomobject]@{
                DBFriendlyName = $dbfn
                ServerName = $srv
                DBName = $db
            }
            $dbfn = ''
            $srv = ''
            $db = ''
        }
    }
    return $DBs
}

$Result = Get-ADComputer -Filter * | Foreach-Object {
    $PCName = $_.Name
    If (Test-Connection $PCName -Count 2 -Quiet){
        # Computer Pingable
        Get-ChildItem "\\$PCName\c$\Users\" -ErrorAction SilentlyContinue | Foreach-Object {
            $User = $_.Name
            # Check that 1C config exists
            if (Test-Path "\\$PCName\c$\Users\$User\$CPath"){
                $Databases = Parse-Config -ConfigFilePath "\\$PCName\c$\Users\$User\$CPath"
                $Databases | Select-Object *, @{n='UserName';e={$User}}
            } else {
                # Config file not found
            }
        }
    } Else {
        # Write down computer name to the "unavailable computers" list
        $PCName | Out-File $NotPingableComputers -Append
    }
}


$Errors = @()
$ErrorActionPreference = 'Stop'
$Result | Group-Object DBName | ForEach-Object {
    $GroupName = $_.Name
    $Members = $_.Group

    try{
        try {
            $ADGroup = Get-ADGroup $GroupName
            if (! [string]::IsNullOrEmpty($ADGroup)){
                $GroupExists = $true
                Write-Host "ADGroup with identifier '$GroupName' found"
                $ADGroup | Out-Host
            }
        } catch{
            if (! ($_ -match 'Cannot find an object')){
                $Errors += "Get-ADGroup failed with error: $_"
                Write-Host $Errors[-1] -ForegroundColor Red
            } else {
                $GroupExists = $False
            }
        }
        if (! $GroupExists){
            Write-Host "Creating '$GroupName'"
            New-ADGroup -Name $GroupName -GroupScope Universal
            $ADGroup = Get-ADGroup $GroupName
            if (! [string]::IsNullOrEmpty($ADGroup)){
                Write-Host "ADGroup successfully created '$GroupName'"
                $ADGroup | Out-Host
            }
            Write-Host "Adding members $(($Members.UserName -join ', ') -replace ', $') to the group '$GroupName'"
            Add-ADGroupMember -Identity $ADGroup -Members $Members.UserName
        }
    } catch {
        $Errors += "Iteration failed with error: $_"
        Write-Host $Errors[-1] -ForegroundColor Red
    }
}

if ($Errors.Count -ge 1){
    Write-Host "Errors during execution:" -ForegroundColor Red
    $Errors | ForEach-Object {
        Write-Host "  - $_" -ForegroundColor Red
    }
}