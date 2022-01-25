<#Import-FileSironKYC.ps1
.SYNOPSIS
Utility for loading Accuity sanction and pep list to Siron KYC
  - extracting "slice" zip-archives (eg 1006.zip, 1010.zip, etc) from main (eg MLUPIDGWL.ZIP)
  - loop over extracted zip archives
  - run Accuity watch list loader for every "slice" zip
  - verify loader log file 
  - load files from theWall (in*.txt)

.DESCRIPTION
Usage:
Import-FileSironKYC.ps1 [-Test] [-ShowConfig] [-V] [<CommonParameters>]

Parameters:
    [-Test<SwitchParameter>] - run script without delete content of the source file;
	[-ShowConfig<SwitchParameter>] - show configuration parameters from json file;
	[-Help<SwitchParameter>] - print out this help.

.EXAMPLE


.NOTES
	Author:	Viesturs Skila
	Version: 1.5.2
#>
[CmdletBinding()] 
param (
    [switch]$Test,
    [switch]$ShowConfig,
	[switch]$Help
)
begin {
    #mail integrācija
    $SmtpServer = 'your mail server host'
    $mailTo = @('your@email','another@email')
    <#------------------------------------------------------------------------------------------------------
    # ZEMĀK NEKO NEMAINĪT!!!
    ------------------------------------------------------------------------------------------------------#>
    #Skripta tehniskie mainīgie
	$CurVersion = "1.5.2"
    $ScriptWatch = [System.Diagnostics.Stopwatch]::startNew()
    #skripta update vajadzībām
    $__ScriptName = $MyInvocation.MyCommand
    $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
    #Skritpa konfigurācijas datne
    $jsonFileCfg = "$__ScriptPath\config.json"
    #logošanai
    $LogFileDir = "$__ScriptPath\log"
	$LogFile = "$((Get-ChildItem $__ScriptName).BaseName)-$(Get-Date -Format "yyyyMMdd")"
    $ScriptRandomID = Get-Random -Minimum 100000 -Maximum 999999
	$ScriptUser = Invoke-Command -ScriptBlock { whoami }
    $Script:forMailReport = @()
    #$mailTo = @('Viesturs.Skila@expobank.eu','Ilana.Einmane@expobank.eu','Valdis.Peisenieks@expobank.eu')
    #scripta darbībai nepeiciešamie mainīgie
    $ScriptTempDir = "$__ScriptPath\tmp"
    $AccuityListsDir = "$__ScriptPath\list"
    $Script:AccHashCookie = "imported.dat"
    $ImportAMLFileNames = "AMLFileNames.txt"
    $__LockFile = "$__ScriptPath\lock-kyc.dat"
    #Konfigurācijas JSON datnes parauga objekts
    $CfgTmpl = [PSCustomObject]@{
        'RemoteAccuityServerPath' = $null
        'RemoteDirPEP' = $null
        'RemoteDirGWL' = $null
        'RemoteDirEDD' = $null
        'RemoteFileName' = $null
        'SironAccImpCommand' = $null
        'SironAccImpArgument' = $null
        'SironKYCscoringCommand' = $null
        'SironKYCscoringArgument' = $null
        'SironKYCHomePath' = $null
        'SironAMLHomePath' = $null
    }#endOfobject

    #Parādam ekrānā versiju un beidzam darbu
	if ( $V ) {
		Write-Host "`n$CurVersion`n"
		Exit 0
	}#endif

    #parādam ekrānā Help un beidzam darbu
	if ($Help) {
		Write-Host "`nVersion:[$CurVersion]`n"
		#$text = (Get-Command "$__ScriptPath\$__ScriptName" ).ParameterSets | Select-Object -Property @{n = 'Parameters'; e = { $_.ToString() } }
		$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
		$text | ForEach-Object { Write-Host $($_) }
		Write-Host "For more info write `'Get-Help `.`\$__ScriptName -Examples`'"
		Exit
	}#endif

	if ( -not ( Test-Path -Path $LogFileDir ) ) { $null = New-Item -ItemType "Directory" -Path $LogFileDir }
    if ( -not ( Test-Path -Path $ScriptTempDir ) ) { $null = New-Item -ItemType "Directory" -Path $ScriptTempDir }


    Function Get-IsRunningScript {
        if ( Test-Path -Path $__LockFile -PathType Leaf) {
            return $true
        }#endif
        return $False
    }#endOffunction

    <# ----------------
	Declare write-log function
	Function's purpose: write text to screen and/or log file
	Syntax: wrlog [-log] [-bug] -text <string>
	#>
	Function Write-msg { 
		[CmdletBinding(DefaultParameterSetName="default")]
		[alias("wrlog")]
		Param(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[string]$text,
			[switch]$log,
			[switch]$bug
		)#param
		try {
			if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO'}
			$timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
			if ( $log -and $bug ) {
                Write-Warning "[$flag] $text"	
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
			}#endif
            elseif ( $log ) {
                Write-Verbose $flag"`t"$text
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |`
                Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text"
			}#endelseif
            else {
                Write-Verbose $flag"`t"$text
			}#else
		}#endtry
		catch {
            Write-Warning "[Write-msg] $($_.Exception.Message)"
		}#endtry
	}#endOffunction
    
    Function Write-ErrorMsg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )#param
    Write-msg -log -bug -text "[$name] Error: $($InputObject.Exception.Message)"
    $string_err = $InputObject | Out-String
    Write-msg -log -bug -text "$string_err"
    }#endOffunction

    <# ----------------
	Formējam un sūtam e-pastus
	#>
    Function Send-Mail {
        foreach ( $a in $Script:forMailReport) {
            if ( $a.Contains("[ERROR]") ) { $logError = $true }
        }
        $server = $env:computername.ToLower()
        $mailParam = @{
            SmtpServer = $SmtpServer
            To = $mailTo
            From = "no-reply@$server.ltb.lan"
            Subject = "[$($env:computername)]$(if ($logError) {":[ERROR]"} else {":[SUCCESS]"}) report from [$__ScriptName]"
        }
        $mailReportBody = @"
        <!DOCTYPE html>
        <html>
        <head>
        <style>
        h2 {
          color: blue;
          font-family: tahoma;
          font-size: 100%;
        }
        p {
          font-family: tahoma;
          color: blue;
          font-size: 80%;
          margin: 0;
        }
        table {
          font-family: tahoma, sans-serif;
          border-collapse: collapse;
          width: 100%;
          font-size: 90%;
        }
        
        td, th {
          border: 1px solid #dddddd;
          text-align: left;
          padding: 8px;
          font-size: 75%;
        }
        
        tr:nth-child(even) {
          background-color: #dddddd;
        }
        </style>
        </head>
        <body>
        <h2>$(Get-Date -Format "yyyy.MM.dd HH:mm") events from log file [$LogFile.log]</h2>
        <br>
        <table><tbody>
            $( ForEach ( $line in $Script:forMailReport ) { 
                if ($line.Contains("[ERROR]") -or $line.Contains(" error") ) {
                    "<tr><td><p style=font-size:100%;color:red;>$line</p></td></tr>"
                } elseif ($line.Contains("SUCCESS:") -or $line.Contains(" success") -or $line.Contains("[SUCCESS]")) {
                    "<tr><td><p style=font-size:135%;color:green;>$line</p></td></tr>"
                } else {
                    "<tr><td>$line</td></tr>"
                } 
            } )
        </tbody></table>
        <br><br>
        <p>Powered by Powershell</p>
        <p style="font-size: 60%;color:gray;">[$__ScriptName] version $CurVersion</p>
        </body>
        </html>
"@
        try {
            Send-MailMessage @mailParam -Body $mailReportBody -BodyAsHtml -ErrorAction Stop
        }#endtry 
        catch {
            Write-ErrorMsg -Name 'smtpErr' -InputObject $_
        }#endcatch

    }#endOffunction

    <# ----------------
        Declare Get-jsonData function
        Function's purpose: to get the data from the json file, to parse and to valdiate against the template ojects' values by name
        Syntax: gjData [object] [jsonFileName]
    #>
    Function Get-JSONData {
        [cmdletbinding(DefaultParameterSetName="default")]
        [alias("gjData")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [object]$template,
            [Parameter(Position = 1, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$jsonFile
        )
        #Write-msg -text "[gjData] jsonFile : [$jsonFile]"
        $jsonData = [PSCustomObject]{}
        try {
            $jsonData = Get-Content -Path $jsonFile -Raw -ErrorAction STOP | ConvertFrom-Json
        } 
        catch {
            # ja JSON datne nav nolasāma vai neeksistē, paziņojam un beidzam darbu
            Write-msg -log -bug -text "[gjData] Fatal error - file [$jsonFile] corrupted. Exit."
            Write-ErrorMsg -Name 'gjData' -InputObject $_
            Stop-Watch -Timer $ScriptWatch -Name Script
            Send-Mail
            Exit 1
        } #endcatch
        if ( ($jsonData.count) -gt 0 ) {
            if ($Test) { Write-msg -text "[gjData] [$jsonFile].count : $($jsonData.count)" }
            # let's compare properties of $Cfg and $jsonData 
            $aa = $template[0].psobject.properties.name | Sort-Object
            $bb = $jsonData[0].psobject.properties.name | Sort-Object
            foreach ($name in $bb) {
                # ja JSON konfig struktūra neatbilst cfg template, paziņojam un beidzam darbu
                if ( $aa -eq $name ) {
                    # Write-msg -text "`t template[$aa] = jsonData[$name)]"
                }#endif
                else {
                    Write-msg -log -bug -text "[gjData] Fatal error - Unknown variable name [$name] in import file [$jsonFile]. Exit."
                    Send-Mail
                    Stop-Watch -Timer $ScriptWatch -Name Script
                    Exit 1
                } #endelse
            } #endforeach
        } #endendif
            return $jsonData
        } #endOffunction
        
    <# ----------------
    Funkcija Repair-JSONfile
    izveido no parauga objekta jaunu JSON datni
    #>
    Function Repair-JSONFile {
        [cmdletbinding(DefaultParameterSetName="default")]
        [alias("repJFile")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.Object]$object,
            [Parameter(Position = 1, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$jsonFile
        )
        try {
            $object | ConvertTo-Json | Out-File $jsonFile
        }#endtry
        catch {
            Write-ErrorMsg -Name 'repJFile' -InputObject $_
        }#endcatch
    }#endOffunction

    Function Stop-Watch {
        [CmdletBinding()] 
        param (
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [object]$Timer,
            [Parameter(Mandatory = $True)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        $Timer.Stop()
        if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)"} else { $bMin = "$($Timer.Elapsed.Minutes)" }
        if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)"} else { $bSec = "$($Timer.Elapsed.Seconds)" }
        Write-msg -log -text "[$Name] finished in $(
            if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
            elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
            else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
            )"
    }#endOffunction

    Function Set-Housekeeping {
        [cmdletbinding(DefaultParameterSetName="default")]
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$CheckPath
        )
        #šeit iestatam datņu dzēšanas periodu
        #tagad iestatīts, kad tiek saarhivēts viss, kas vecāks par 30 dienām
        $maxAge = ([datetime]::Today).addDays(-30)

        $filesByMonth = Get-ChildItem -Path $CheckPath -File |
                Where-Object -Property LastWriteTime -lt $maxAge |
                Group-Object { $_.LastWriteTime.ToString("yyyy\\MM") }
        if ( $filesByMonth.count -gt 0 ) {
            try {
                $filesByMonth.Group | Remove-Item -Recurse -Force -ErrorAction Stop
                Write-msg -log -text "[Housekeeping] Succesfully cleaned files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
            }#endtry
            catch {
                Write-ErrorMsg -Name 'Housekeeping' -InputObject $_
            }#endcatch
        }#endif
        else {
            Write-msg -log -text "[Housekeeping] There's no files older than [$($maxAge.ToShortDateString())] in [$CheckPath]"
        }#endelse
    }#endOfFunctions

    <# ----------------------------------------------
    DEFINĒJAM DARBA FUNKCIJAS
    ---------------------------------------------- #>
    <# ----------------
    Funkcija Set-FileHash 
    cepumiņā saglabā informāciju, par ielādētā saraksta arhīva datni
    #>
    Function Set-HashCookie {
        Param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$gHash,
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Path
        )
        $gHash | Out-File "$Path\$Script:AccHashCookie" -Force
		Write-msg -log -text "[SetHash] cookie changed."											   
    }#endOffunction
    
    <# ----------------
    Funkcija Get-FileHash 
    atgriež saglabā informāciju, par ielādētā saraksta arhīva datni
    #>
    Function Get-HashCookie {
        Param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        $HashTime = [System.Diagnostics.Stopwatch]::startNew()
		Write-msg -log -text "[GetHash] Checking hash of [$Name]..."											
        $Hash = Get-FileHash $Name

        Stop-Watch -Timer $HashTime -Name GetHash								 
        return $Hash.Hash
    }#endOffunction

    <# ----------------
    Funkcija Set-AccListPropertiesForFile 
    meklē norādītajā direktorijā extractsummary.txt 
    un atrod tajā mainīgo listids, listnames un listtypes vērtības
    #>
    Function Set-AccListPropertiesForFile {
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$SourceDir
        )
        $AccList= [PSCustomObject]@{}
        try{
            if ( Test-Path "$SourceDir\extractsummary.txt" -PathType Leaf ) {
                $_listids = Select-String "listids=(.*)" "$SourceDir\extractsummary.txt" -ErrorAction Stop
                $listids = $_listids.Matches.Groups[1].Value
                $_listnames = Select-String "listnames=(.*)" "$SourceDir\extractsummary.txt" -ErrorAction Stop
                $listnames = $_listnames.matches.groups[1].Value
                $_listtypes = Select-String "listtypes=(.*)" "$SourceDir\extractsummary.txt" -ErrorAction Stop
                $listtypes = $_listtypes.matches.groups[1].Value

                if ( $listtypes -match "Interdiction" ) { $listtypes = "embargo" } else { $listtypes = "pep" }

                $AccList | Add-Member -MemberType NoteProperty -Name 'Listids' -Value $listids -Force
                $AccList | Add-Member -MemberType NoteProperty -Name 'Listtypes' -Value $listtypes -Force
                $AccList | Add-Member -MemberType NoteProperty -Name 'Listnames' -Value $listnames -Force

                <#------------------------------
                Pievienojam konstatntes, kuras jāņem no
                accuity.properties.template datnes un jāiekļauj properties datnes formēšanā
                ------------------------------#>
                $AccList | Add-Member -MemberType NoteProperty -Name 'Description' -Value "PIDs, Entity Consolidation, Standardized Data, and Native Script Names" -Force
                $AccList | Add-Member -MemberType NoteProperty -Name 'LoadOriginalScriptNames' -Value 'true' -Force
            }#endif
            else {
                Write-msg -log -text "[AccListSiron] There's no extractsummary.txt file in [$SourceDir]"
            }#endelse
        }#endtry
        catch {
            Write-ErrorMsg -Name 'AccListSiron' -InputObject $_
        }#endcatch
        return $AccList
    }#endOfFunction
    
    <# ----------------
    Funkcija Set-AccPropertiesFile 
    izveido accuity{$Listids}.properties datni, ievieto tajā mainīgo listids, listnames un listtypes vērtības
    funkcija atgriež properties faila atrašānās vietu
    #>
    Function Set-AccPropertiesFile {
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [object]$AccProp
        )
        $Content = @"

Listname=$($AccProp.Listids)
Listtype=$($AccProp.Listtypes)
Providertype=$($AccProp.Listnames)
Description=$($AccProp.Description)

LoadOriginalScriptNames=$($AccProp.LoadOriginalScriptNames)
"@
        $AccPropFileName = "$($Cfg.SironAccPropretiesPath)\accuity$($AccProp.Listids).properties"
        #$Content | Out-File $AccPropFileName -Force
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
        [System.IO.File]::WriteAllLines($AccPropFileName, $Content, $Utf8NoBomEncoding)

        return $AccPropFileName
    }#endOffunction

    <# ----------------
    Funkcija Start-SironBatch 
    darbina SironKYC bat palaišanas komandas ar argumentu
    funkcija pārtver cmd output un saglabā mainīgajā $outout,
    kuru varam apstrādāt uz kļūdām
    #>
    Function Start-SironBatch {
        Param(
            [Parameter(Position = 0, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Command,
            [Parameter(Position = 1, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$WorkingDir,
            [Parameter(Position = 2, Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$OutputFile
        )
        $BatchTime = [System.Diagnostics.Stopwatch]::startNew()
        Write-msg -log -text "[Runner] execute [$($Command.trim())]"
        try {
            $psi = New-object System.Diagnostics.ProcessStartInfo -Property @{
                CreateNoWindow = $true
                UseShellExecute = $false
                RedirectStandardOutput = $true
                RedirectStandardError = $true
                FileName = 'cmd.exe'
                Arguments = @("`/c $($Command.trim())")
                WorkingDirectory = $WorkingDir
            }
            $process = New-Object System.Diagnostics.Process 
            $process.StartInfo = $psi 
            $process.Start()
            $output = $process.StandardOutput.ReadToEnd() 
            $process.WaitForExit() 
            $output | Out-File $OutputFile
        }#endtry
        catch {
            Write-ErrorMsg -Name 'Runner' -InputObject $_
        }#endcatch

        Stop-Watch -Timer $BatchTime -Name Runner
    }#endOffunction

    <# ----------------
    Funkcija Get-PostCheck 
    pārbauda KYC log datņu direktoriju pēc paterniem
    #>
    Function Get-PostCheck {

        try {
            $LogPath = $null
            $LogPath = [ordered]@{}
            Get-ChildItem -Path "$($Cfg.SironKYCHomePath)\client\0001\" -Filter "workcus*" -Attributes Directory -ErrorAction Stop |
                ForEach-Object {
                    #ievietojam parametru Hashtable name un value laukos
                    $LogPath.Add($_.FullName,$_.FullName)
                    Write-msg -log -text "[PostCheck] found WORKCUST directory:[$($_.FullName)]."
                }
        }#endtry
        catch {
            Write-ErrorMsg -Name 'PostCheck' -InputObject $_
        }#endcatch

        $LogFileNamesPatterns = @(
            @{name="error.txt"}
            @{name="SP*.txt"}
        )
        $ErrorPatterns = @(
            @{name=" I/O error "}
            @{name=" I/O-ERROR "}
            @{name=" ERROR_CODE "}
            @{name=" Error in"}
            @{name=" F-NAME"}
            @{name=" RECCNT"}
        )
        $Script:noError = $True
        $LogPath.GetEnumerator() | 
            #Where-Object { $_.name -match 'Path' } | 
            ForEach-Object {
                if ( Test-Path -Path $_.value ) {
                    $ErrorFileNames = $null
                    $ErrorFileNames = @()
                    #Meklējam sliktos failus
                    foreach ( $fPattern in $LogFileNamesPatterns) {
                        if ( Test-Path -Path "$($_.value)\*" -PathType leaf -Include "$($fPattern.name)" ) {
                            $ErrorFileNames = (Get-ChildItem "$($_.value)\$($fPattern.name)").name
                            Write-msg -log  -bug -text "[PostCheck] found [$ErrorFileNames] in the direcory [$($_.value)]."
                            $Script:noError = $False
                        }#endif
                    }#endforeach
                    
                    $Object = $null
                    $Object = @()
                    foreach ( $wPatt in $ErrorPatterns ) {
                        Get-ChildItem -Path "$($_.value)\*.txt" -File | 
                            Where-Object { $_.LastWriteTime -ge ( [datetime]::Today ) } |
                            ForEach-Object {
                                If (Get-Content $_.FullName | Select-String -Pattern "$($wPatt.name)") {
                                    $Object += Select-String -Path $_.FullName -Pattern "$($wPatt.name)" -CaseSensitive
                                    }#endif
                            }#endforeach
                    }#endofreach
                    
                    if ( $Object.Count -gt 0 ) {
                        $Object | Select-Object FileName, Pattern, Line | ForEach-Object {
                            Write-msg -log -text "[PostCheck] in [$($_.FileName)]: [$($_.Line.trim())]:"
                            $Script:noError = $False
                        }#endforeach
                    }#endif
    
                }#endif
                else {
                    Write-msg -log -bug -text "[PostCheck] not found [$($_.value)]. Skipped check."
                }#endelse
            }#endforeach
        if ( $Script:noError ) {
            Write-msg -log -text "[PostCheck]:[SUCCESS] nothing bad found in log directories."
        }#endif

    }#endOffunction

    <# ----------------------------------------------
    Funkciju definēšanu beidzām
    IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
    ---------------------------------------------- #>
    #Clear-Host
    Write-msg -log -text "[-----] Script started in [$(if ($Test) {"Test"}
		elseif ($ShowConfig) {"ShowConfig"} 
		else {"Default"})] mode. Used config file [$jsonFileCfg]"

    # Ielasām parametrus no JSON datnes
    # ja neatrodam, tad izveidojam JSON datnes paraugu ar noklusētām vērtībām
    if ( -not ( Test-Path -Path $jsonFileCfg -PathType Leaf) ) {
        Write-Warning "[Check] Config JSON file [$jsonFileCfg] not found."

        #No cfg template izveidojam objektu ar vērtībām
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RemoteAccuityServerPath' -Value '\\argo\CL_Lists' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RemoteDirPEP' -Value 'PEP' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RemoteDirGWL' -Value 'GWL' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RemoteDirEDD' -Value 'EDD' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'RemoteFileName' -Value 'ML*.ZIP' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironKYCHomePath' -Value 'D:\FICO\SironKYC' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironAMLHomePath' -Value 'D:\FICO\SironAML' -Force

        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironAccImpCommand' -Value 'start_update_Accuity.bat' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironAccImpArgument' -Value '0001' -Force

        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironKYCscoringCommand' -Value 'start_scoring_prs.bat' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironKYCscoringArgument' -Value '' -Force

        #Izsaucam JSON izveidošanas funkciju
        Repair-JSONFile $CfgTmpl $jsonFileCfg
        Write-Warning "[Check] Config JSON file [$jsonFileCfg] created."
        #Ielasām datus no jaunizveidotā JSON
        $_Cfg = Get-jsonData $CfgTmpl $jsonFileCfg
    }#endif
    else {
        #Ielasām datus no JSON
        $_Cfg = Get-JSONData $CfgTmpl $jsonFileCfg
    }#endelse

    $SironKYCHomePath = $env:KYC_ROOT
    if ( [String]::IsNullOrWhiteSpace($SironKYCHomePath) ) {
        Write-msg -log -text "[Cfg] cannot find KYC_ROOT variable from system environment"
        $SironKYCHomePath = $_Cfg.SironKYCHomePath
    }
    $SironAMLHomePath = $env:AML_ROOT
    if ( [String]::IsNullOrWhiteSpace($SironAMLHomePath) ) {
        Write-msg -log -text "[Cfg] cannot find AML_ROOT variable from system environment"
        $SironAMLHomePath = $_Cfg.SironAMLHomePath
    }

    #Definējam skripta konstantes
    $Cfg=[ordered]@{}
    $Cfg.Add("SironKYCHomePath",$SironKYCHomePath)
    $Cfg.Add("SironAMLHomePath",$SironAMLHomePath)
    $Cfg.Add("RemoteAccuityServerPath",$_Cfg.RemoteAccuityServerPath)
    $Cfg.Add("RemoteDirPEP",$_Cfg.RemoteDirPEP)
    $Cfg.Add("RemoteDirGWL",$_Cfg.RemoteDirGWL)
    $Cfg.Add("RemoteDirEDD",$_Cfg.RemoteDirEDD)
    $Cfg.Add("RemoteFileName",$_Cfg.RemoteFileName)
    $Cfg.Add("SironDataInputAccuityPath","$SironKYCHomePath\client\0001\data\input\accuity")
    $Cfg.Add("SironAccPropretiesPath","$SironKYCHomePath\custom\tool")
    $Cfg.Add("SironAccBatchPath","$SironKYCHomePath\system\watchlist\batch")
    $Cfg.Add("SironAccImpCommand",$_Cfg.SironAccImpCommand)
    $Cfg.Add("SironAccImpArgument",$_Cfg.SironAccImpArgument)
    $Cfg.Add("KYCdataInputPath","$SironKYCHomePath\client\0001\data\input")
    $Cfg.Add("AMLdataInputPath","$SironAMLHomePath\client\0001\data\input")
    $Cfg.Add("SironKYCbatchPath","$SironKYCHomePath\system\scoring\batch")
    $Cfg.Add("SironKYCscoringCommand",$_Cfg.SironKYCscoringCommand)
    $Cfg.Add("SironKYCscoringArgument",$_Cfg.SironKYCscoringArgument)
    $Cfg.Add("SironKYCResultFileName","$SironKYCHomePath\client\0001\workcust\retcodeStartClientPrs.txt")

    if ( -not ( Test-Path -Path "$__ScriptPath\$ImportAMLFileNames" -PathType Leaf ) ) { 
        $FilesFromAML = @("in_Customer.txt","in_Customer_Extension_Nohist.txt", "in_beneficial_owner.txt") 
        $FilesFromAML | Out-File "$__ScriptPath\$ImportAMLFileNames" -Force
    }#endif
    else {
        $FilesFromAML = Get-Content "$__ScriptPath\$ImportAMLFileNames"
    }#endelse

    #Ja norādīts parametrs, tad parādam ielasītos datus no JSON un beidzam darbu
    if ( $ShowConfig ) {
        Write-Host "Configuration param:"
        Write-Host "-----------------------------------------------------"
        Write-Host "CurVersion`t`t:[$CurVersion]"
        Write-Host "__ScriptName`t`t:[$__ScriptName]"
        Write-Host "__ScriptPath`t`t:[$__ScriptPath]"
        Write-Host "jsonFileCfg`t`t:[$jsonFileCfg]"
        Write-Host "LogFileDir`t`t:[$LogFileDir]"
        Write-Host "LogFile`t`t`t:[$LogFile]"
        Write-Host "SmtpServer`t`t:[$SmtpServer]"
        Write-Host "mailTo`t`t`t:[$mailTo]"
        Write-Host "ScriptTempDir`t`t:[$ScriptTempDir]"
        Write-Host "AccuityListsDir`t`t:[$AccuityListsDir]"
        Write-Host "`nFrom [$ImportAMLFileNames] file:"
        Write-Host "-----------------------------------------------------"
        foreach ( $item in $FilesFromAML ) {
            Write-Host "$item"
        }#endforeach
        Write-Host "`nFrom config file:"
        Write-Host "-----------------------------------------------------"
        $Cfg | Sort-Object -Property Name | Format-Table Name, Value -AutoSize
        Write-Host "-----------------------------------------------------"
        Stop-Watch -Timer $ScriptWatch -Name Script
        Exit 0
    }#endif

    if ( Get-IsRunningScript ) {
        Write-msg -log -bug -text "[Check] another instance of script is running."
        Stop-Watch -Timer $ScriptWatch -Name Script
        Exit 1
    }#endif 
    else {
        "1" | Out-File $__LockFile
    }#endelse

}#endOfbegin

process {
    <#--------------------------------------------
    Pārbaudam darbu izpildes priekšnosacījumus
    --------------------------------------------#>
    $Script:isItEnd = $false
    if ( -not ( Test-Path -Path $Cfg.RemoteAccuityServerPath ) ) { 
        Write-msg -log -bug -text "[Check] Accuity server [$($Cfg.RemoteAccuityServerPath)] is no accessible"
        $Script:isItEnd = $true
    }#endif
    else {
        $AccuityPath = @(
        @{name="PEP"; src="$($Cfg.RemoteAccuityServerPath)\$($Cfg.RemoteDirPEP)"; dst="$AccuityListsDir\acc$($Cfg.RemoteDirPEP)"; tmp="$ScriptTempDir\tmp.$($Cfg.RemoteDirPEP)"}
        @{name="GWL"; src="$($Cfg.RemoteAccuityServerPath)\$($Cfg.RemoteDirGWL)"; dst="$AccuityListsDir\acc$($Cfg.RemoteDirGWL)"; tmp="$ScriptTempDir\tmp.$($Cfg.RemoteDirGWL)"}
        @{name="EDD"; src="$($Cfg.RemoteAccuityServerPath)\$($Cfg.RemoteDirEDD)"; dst="$AccuityListsDir\acc$($Cfg.RemoteDirEDD)"; tmp="$ScriptTempDir\tmp.$($Cfg.RemoteDirEDD)"}
        )

        $WallPath = @(
        @{src="$($Cfg.AMLdataInputPath)"; dst="$($Cfg.KYCdataInputPath)"}
        )

        foreach ($line in $AccuityPath) {
            if ( -not (Test-Path -Path $line.src) ) {
                Write-msg -log -text "[$($line.src)] --> [$($line.dst)]"
                Write-msg -log -bug -text "[Check] Directory [$($line.src)] is no accessible"
                $Script:isItEnd = $true
            }#endif
            elseif ( -not (Test-Path -Path $line.dst) ) {
                $null = New-Item -ItemType "Directory" -Path $line.dst
            }#endelseif
        }#endforeach
        foreach ($line in $WallPath) {
            if ( -not (Test-Path -Path $line.src) ) {
                Write-msg -log -text "[$($line.src)] --> [$($line.dst)]"
                Write-msg -log -bug -text "[Check] Directory [$($line.src)] is no accessible"
                $Script:isItEnd = $true
            }#endif
            elseif ( -not (Test-Path -Path $line.dst) ) {
                $null = New-Item -ItemType "Directory" -Path $line.dst
                $Script:isItEnd = $true
            }#endelseif
        }#endforeach

        $cfg.GetEnumerator() | 
            Where-Object { $_.name -match 'Path' } | 
                ForEach-Object {
                    if ( -not ( Test-Path $_.value ) ) {
                        Write-msg -log -bug -text "[Check] Directory [$($_.value)] is not accessible. Exit";
                        $Script:isItEnd = $true
                    }#endif
                }#endforeach

        if ( -not ( Test-Path "$($Cfg.SironAccBatchPath)\$($Cfg.SironAccImpCommand)" ) ) {
            Write-msg -log -bug -text "[Check] command file ["$($Cfg.SironAccBatchPath)\$($Cfg.SironAccImpCommand)"] is not accessible. Exit";
            $Script:isItEnd = $true
        }#endif

        if ( -not ( Test-Path "$($Cfg.SironKYCbatchPath)\$($Cfg.SironKYCscoringCommand)" ) ) {
            Write-msg -log -bug -text "[Check] command file ["$($Cfg.SironKYCbatchPath)\$($Cfg.SironKYCscoringCommand)"] is not accessible. Exit";
            $Script:isItEnd = $true
        }#endif

    }#endelse

    if ( $Script:isItEnd ) {
        Write-msg -log -bug -text "[Check] Script terminated with error."
        Stop-Watch -Timer $ScriptWatch -Name Script
        Send-Mail
        Remove-Item $__LockFile -Force
        exit 1
    }#endif
    <#--------------------------------------------------
    Lejupielādējam no Accuity servera aktuālos sarakstus
    --------------------------------------------------#>
    $AccuityWatch = [System.Diagnostics.Stopwatch]::startNew()
    foreach ($line in $AccuityPath) {
        $GetFiles = Get-ChildItem "$($line.src)\$($Cfg.RemoteFileName)"
        if ( $GetFiles.Count -gt 0 ) {
            foreach ( $file in $GetFiles ) {

                #Lejupielādējam Accuity sarakstus
                #--------------------------------------------
                try {
                    if ( Test-Path "$($line.dst)\$($Script:AccHashCookie)" -PathType leaf ) {
                        Write-msg -log -text "[Download] Checking for [$($line.name)] list update..."
                        $HashCookie = Get-Content "$($line.dst)\$($Script:AccHashCookie)"
                        Copy-Item $file -Destination $line.dst -Force
                        Write-msg -log -text "[Download] [$($line.name)] list [$($file.FullName)] downloaded."																						 
                        #$fHash = Get-HashCookie $file
                        $fHash = Get-HashCookie -Name "$((Split-Path $file).Replace("$($line.src)","$($line.dst)"))\$(Split-Path -Path $file -Leaf -Resolve)"

                        if ( $fHash -notlike $HashCookie ) {
                            Write-msg -log -text "[Download] [$($line.name)] list update found and [$($file.FullName)], LastWriteTime:[$(($file.LastWriteTime).ToShortDateString()) $(($file.LastWriteTime).ToLongTimeString())]."
                            Set-HashCookie -gHash $fHash -Path $line.dst
                        }#endif
                        else {
                            Write-msg -log -text "[Download] the latest [$($line.name)] list already loaded."

                            #ja neesam testa režīmā, tad izlaižam ciklu
                            if ( -not $Test ) {
                                Continue
                            }#endif
                        }#endelse
                    }#endif
                    else {
                        Write-msg -log -text "[Download] [$($line.name)] list found! File [$($file.FullName)], LastWriteTime:[$(($file.LastWriteTime).ToShortDateString()) $(($file.LastWriteTime).ToLongTimeString())]."
                        Copy-Item $file -Destination $line.dst -Force
						Write-msg -log -text "[Download] [$($line.name)] list [$($file.FullName)] downloaded."																						
                        #$fHash = Get-HashCookie $file
                        $fHash = Get-HashCookie -Name "$((Split-Path $file).Replace("$($line.src)","$($line.dst)"))\$(Split-Path -Path $file -Leaf -Resolve)"
                        Set-HashCookie -gHash $fHash -Path $line.dst
                    }#endelse
                }#endtry
                catch {                    
                    Write-ErrorMsg -Name 'Download' -InputObject $_
                    Continue
                }#endcatch

                #Importējam Accuity sarakstus SironKYC
                #--------------------------------------------
                try {
                    if ( Test-Path "$($line.dst)\$($file.Name)" -PathType leaf ) {
                        #Atarhivējam galvenot saraksta mapi pagaidu mapē
                        if ($Test) { Write-msg -log -text "[Expand] [$($line.dst)\$($file.Name)] to [$($line.tmp)]" }
                        Expand-Archive -Path "$($line.dst)\$($file.Name)" -DestinationPath "$($line.tmp)" -Force -ErrorAction Stop

                        #Iegūstam Accuity sarakstu arhīva datnes, kuras apstrādāt
                        $CLArchivesList = Get-ChildItem "$($line.tmp)" -Filter '*.zip'
                        if ($Test) { $CLArchivesList | Format-Table LastWriteTime,Name,@{n='Size(MB)';e={[math]::round($_.Length / 1MB,3)}} -AutoSize }
                        $total = $CLArchivesList.Count
                        Write-msg -log -text "[Import] got [$total] [$($line.name)] lists to proceed"
                        
                        #Sagatavojam ielādei katru apakšarhīva datni
                        #--------------------------------------------
                        if ( $CLArchivesList.Count -gt 0 ) {
                            $count = 0
                            ForEach ( $arch in $CLArchivesList ) {
                                $count++
                                #atarhivējam arhīva datni Siron input mapē
                                try {
                                    if ($Test) { Write-msg -log -text "[Expand] [$count]of[$total] list [$($arch.BaseName)] to [$($Cfg.SironDataInputAccuityPath)]" }
                                    Expand-Archive -Path $arch.FullName -DestinationPath $Cfg.SironDataInputAccuityPath -Force -ErrorAction Stop
                                }#endtry
                                catch {
                                    Write-ErrorMsg -Name 'Expand' -InputObject $_
                                    continue
                                }#endcatch
                                
                                #atrodam Accuity saraksta identifikācijas parametrus
                                $AccProp = Set-AccListPropertiesForFile $Cfg.SironDataInputAccuityPath

                                #ja funkcija Set-AccListPropertieFile neko neatgriež, izlaižam
                                if ( [String]::IsNullOrWhiteSpace($AccProp.listids) -or [String]::IsNullOrWhiteSpace($AccProp.listnames) -or [String]::IsNullOrWhiteSpace($AccProp.listtypes) ) {
                                    Write-msg -log -bug -text "[Import] $($arch.BaseName) skipped."
                                    Continue
                                }

                                if ($Test) { Write-msg -log -text "[Import] listID:[$($AccProp.listids)]; listnames:[$($AccProp.listnames)]; listtypes:[$($AccProp.listtypes)]" }

                                #izveidojam Siron nepieciešamā formātā properties datni
                                $AccPropFileName = Set-AccPropertiesFile $AccProp
                                if ($Test) { Write-msg -log -text "[Import] created [$AccPropFileName]" }

                                #Ielādējam sarakstu Accuity
                                #start_update_Accuity.bat 0001 d:\FICO\SironKYC\custom\tool\accuity1098.properties
                                #palaižam funkciju
                                #--------------------------------------------
                                $RunnerLog = "$ScriptTempDir\AccImpList-$($AccProp.listids).log"
                                Write-msg -log -text "[Runner] runs [$count]of[$total] command with [$AccPropFileName]; log:[$RunnerLog]"

                                #Start-SironBatch [Command+Argument] [WorkingDir] [LogFile]
                                $null = Start-SironBatch "$($Cfg.SironAccImpCommand) $($Cfg.SironAccImpArgument) $AccPropFileName" $Cfg.SironAccBatchPath $RunnerLog
                                #Meklējam Runner logā norādi uz Accuity batch procesa log failu
                                try{
                                    if ( Test-Path $RunnerLog -PathType Leaf ) {
                                        $lOutput = Select-String "Logfile is `"(.*)`"" $RunnerLog -ErrorAction Stop
                                        $ResultFilePath = $lOutput.Matches.Groups[1].Value
                                        if ($Test) {Write-msg -log -text "[LogCheck] [$ResultFilePath]"}
                                    }#endif
                                    else {
                                        Write-msg -log -bug -text "[LogCheck] did not found the runner log file."
                                    }
                                }#endtry
                                catch {
                                    Write-ErrorMsg -Name 'LogCheck' -InputObject $_
                                }#endcatch

                                # Meklējam Accuity batch procesa log failā ielādes procesu statusu
                                if ( Test-Path $ResultFilePath -PathType Leaf ) {
                                    $ResultLog = Get-Content $ResultFilePath
                                        if ( $ResultLog.Count -gt 0 ) {
                                            ForEach ( $line in $ResultLog ) {
                                                if ( $line.Contains('Error in update') ) { Write-msg -log -bug -text "[LogCheck] from [$ResultFilePath] log: [$line]"; break }
                                                if ( $line.Contains('*  Processing successful') ) { Write-msg -log -text "[LogCheck] from [$ResultFilePath] log: [$line]"; break }
                                            }#endforeach
                                        }#endif
                                    }#endif
                                else {
                                    Write-msg -log -bug -text "[LogCheck] Accuity update log was not found."
                                }
                                
                            }#endforeach
                        }#endif
                    }#endif
                    else {
                        Write-msg -log -bug -text "[Import] [$($line.dst)\$($file.Name)] not found."
                        Continue
                    }
                }#endtry
                catch {
                    Write-ErrorMsg -Name 'Import' -InputObject $_
                    Continue
                }#endcatch

            }#endforeach
        }#endif
    }#endforeach
    Stop-Watch -Timer $AccuityWatch -Name Import
    <#--------------------------------------------------
    Lejupielādējam no Wall servera aktuālos sarakstus
    --------------------------------------------------#>
    $FilesFromAML.GetEnumerator() | ForEach-Object {
        try {
            $file = "$($WallPath.src)\$_"
            $fObject = Get-ChildItem $file -ErrorAction Stop
            Copy-Item $file -Destination $WallPath.dst -ErrorAction Stop
            Write-msg -log -text "[Imp-Wall] Successfuly copied [$file], LastWriteTime:[$(($fObject.LastWriteTime).ToShortDateString()) $(($fObject.LastWriteTime).ToLongTimeString())]"
        }#endtry
        catch {
            Write-ErrorMsg -Name 'Imp-Wall' -InputObject $_
        }#endcatch
    }#endforeach

    $timestamp = Get-Date -Format "yyyyMMdd"
    $RunnerLog = "$LogFileDir\SironKYCscoring-$timestamp.log"
    if ($Test) { Write-msg -log -text "[Runner] runs command [$($Cfg.SironKYCscoringCommand) $($Cfg.SironKYCscoringArgument)] in [$($Cfg.SironKYCbatchPath)] log to [$RunnerLog]" }
    # Start-SironBatch [komanda] [darba direktorija] [logfails]
    $null = Start-SironBatch "$($Cfg.SironKYCscoringCommand) $($Cfg.SironKYCscoringArgument)" $Cfg.SironKYCbatchPath $RunnerLog

    #$Cfg.KYCResultFilePath = "D:\TONBELLER\SironKYC\client\0001\workcust\retcodeStartClientPrs.txt"
    if ( Test-Path $Cfg.SironKYCResultFileName -PathType Leaf ) {
        $ResultLog = Get-Content $Cfg.SironKYCResultFileName
            if ( $ResultLog.Count -gt 0 ) {
                ForEach ( $line in $ResultLog ) {
                    if ( $line.Contains('Error in update') -or $line.Contains('*  Error') ) {
                        Write-msg -log -bug -text "[LogCheck] log [$($Cfg.SironKYCResultFileName)]"
                        Write-msg -log -bug -text "[LogCheck] from log: [$line]"
                        break
                    }#endif
                    if ( $line.Contains('*  Processing successful') ) {
                        Write-msg -log -text "[LogCheck] from [$($Cfg.SironKYCResultFileName)] log: [$line]"
                        break
                    }#endif
                }#endforeach
            }#endif
        }#endif
    else {
        Write-msg -log -bug -text "[LogCheck] SironKYC retcodeStartClientPrs.txt log was not found."
    }

}#endOfprocess

end {

    Get-PostCheck

    <#Dzēšam tmp un err mapju saturu, kas vecāks par 30 dienām
    #laika periodu varam koriģēt funkcijā 
    Set-Housekeeping $ScriptTempDir
    #  #>
    Remove-Item "$ScriptTempDir\*" -Recurse -Force
    Remove-Item "$($Cfg.SironDataInputAccuityPath)\*" -Recurse -Force
    Remove-Item $__LockFile -Force
    Set-Housekeeping -CheckPath $LogFileDir
    Stop-Watch -Timer $ScriptWatch -Name Script
    Send-Mail
    <#iespējojam, ja griba sūtīt e-pastus tikai kļūdu gadījumā
    foreach ( $line in $Script:forMailReport) {
        if ( $line.Contains("[ERROR]") ) { Send-Mail }
    }#endforeach
    #   #>
}#endOfend