<#
.SYNOPSIS
Skripts paņem no theWall in*.txt datnes, tās nokonvertē uz UTF-8 , iekopē norādītajā Siron datu ielādes direktorijā un izpilda komandu

.DESCRIPTION
Usage:
Import-FileSironAML.ps1 [-Test] [-ShowConfig] [-V] [<CommonParameters>]

.PARAMETER WallPath
Override wall import directory path given in config file. Set parameter NoEncoding to true.

.PARAMETER NoEncoding
Script not change input files encoding. Use this switch if you know for sure input files are encoded in pure UTF-8.

.PARAMETER Test
Run script without delete content of the source files

.PARAMETER ShowConfig
Show configuration parameters from json file

.PARAMETER Help
print out help

.EXAMPLE
Import-FileSironAML.ps1 -NoEncoding

.NOTES
	Author:	Viesturs Skila
	Version: 1.4.0
#>
[CmdletBinding()] 
param (
    [ValidateScript( {
        if ( -NOT ( $_ | Test-Path -PathType Container) ) {
            Write-Host "Directory does not exist"
            throw
        }#endif
        return $True
    } ) ]
    [System.IO.FileInfo]$WallPath,
    [switch]$NoEncoding,
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
	$CurVersion = "1.4.0"
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
    $__LockFile = "$__ScriptPath\lock-aml.dat"
    #scripta darbībai nepeiciešamie mainīgie
    $ScriptTempDirIN = "$__ScriptPath\tmp\in"
    $ScriptTempDirOUT = "$__ScriptPath\tmp\out"
    $ScriptDataArchive = "$__ScriptPath\archive"
    #Konfigurācijas JSON datnes parauga objekts
    $CfgTmpl = [PSCustomObject]@{
        'SironAMLHomePath' = $null
        'WallSharePath' = $null
        'SironAMLBatchCommand' = $null
        'SironAMLBatchCommandArgument' = $null
        'RunWithoutArgumentTill' = $null
    }#endOfobject

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
    if ( -not ( Test-Path -Path $ScriptTempDirIN ) ) { $null = New-Item -ItemType "Directory" -Path $ScriptTempDirIN }
    if ( -not ( Test-Path -Path $ScriptTempDirOUT ) ) { $null = New-Item -ItemType "Directory" -Path $ScriptTempDirOUT }
    if ( -not ( Test-Path -Path $ScriptDataArchive ) ) { $null = New-Item -ItemType "Directory" -Path $ScriptDataArchive }

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
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |
                    Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptUser] [$flag] $text"
			}#endif
            elseif ( $log ) {
                Write-Verbose $flag"`t"$text
				Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |
                    Out-File "$LogFileDir\$LogFile.log" -Append -ErrorAction Stop
                $Script:forMailReport += "$timeStamp [$ScriptUser] [$flag] $text"
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
                if ($line.Contains("[ERROR]")) {
                    "<tr><td><p style=font-size:100%;color:red;>$line</p></td></tr>"
                } elseif ($line.Contains("SUCCESS:")) {
                    "<tr><td><p style=font-size:135%;color:green;>$line</p></td></tr>"
                } else {
                    "<tr><td>$line</td></tr>"
                } 
            } )
        </tbody></table>
        <br>
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
                    Stop-Watch -Timer $ScriptWatch -Name Script
                    Send-Mail
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

        $LogPath = @(
            @{name="$SironAMLHomePath\client\0001\log\scoring"}
        )
        $ErrorPatterns = @(
            @{name=" I/O error "}
            @{name=" Error in"}
            @{name=" F-NAME"}
            @{name=" RECCNT"}
        )

        foreach ( $dir in $LogPath ) {
            if ( Test-Path -Path $dir.name ) {
                $Object = @()
                foreach ( $wPatt in $ErrorPatterns ) {
                    Get-ChildItem -Path "$($dir.name)\*.txt" -File | 
                        Where-Object { $_.LastWriteTime -ge ( [datetime]::Today ) } |
                        ForEach-Object {
                            If (Get-Content $_.FullName | Select-String -Pattern "$($wPatt.name)") {
                                $Object += Select-String -Path $_.FullName -Pattern "$($wPatt.name)" -CaseSensitive
                                }#endif
                        }#endforeach
                }#endofreach
                
                if ($Object.Count -gt 0) {
                    $Object | Select-Object FileName, Pattern, Line | ForEach-Object {
                        Write-msg -log -text "[PostCheck] [$($_.FileName)]: [$($_.Line.trim())]"
                    }#endforeach
                }#endif

            }#endif
            else {
                Write-msg -log -bug -text "[PostCheck] not found [$($dir.name)]. Skipped check."
            }#endelse
        }#endforeach

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
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'WallSharePath' -Value 'C:\TEMP' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironAMLBatchCommand' -Value 'start_scoring.bat 0001' -Force
        $CfgTmpl | Add-Member -MemberType NoteProperty -Name 'SironAMLBatchCommandArgument' -Value '' -Force

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
    
    #Definējam skripta konstantes
    $SironAMLHomePath = $env:AML_ROOT
    if ( [String]::IsNullOrWhiteSpace($SironAMLHomePath) ) {
        Write-msg -log -text "[Cfg] cannot find SironAML home path from system environment"
        $SironAMLHomePath = $_Cfg.SironAMLHomePath
    }

    $Cfg=[ordered]@{}
    $Cfg.Add("SironAMLHomePath",$SironAMLHomePath)

    if ($WallPath) {
        $Cfg.Add("WallSharePath",$WallPath)
        $NoEncoding = $true
    } else {
        $Cfg.Add("WallSharePath",$_Cfg.WallSharePath)
    }
    $Cfg.Add("SironDataInputPath","$SironAMLHomePath\client\0001\data\input")
    $Cfg.Add("SironAMLBatchPath","$SironAMLHomePath\system\scoring\batch")
    $Cfg.Add("SironImportLogFileName","$SironAMLHomePath\client\0001\log\scoring\retcode_SCORE_ALL.txt")
    $Cfg.Add("SironAMLBatchCommand",$_Cfg.SironAMLBatchCommand)
    $Cfg.Add("SironAMLBatchCommandArgument",$_Cfg.SironAMLBatchCommandArgument)
    $Cfg.Add("RunWithoutArgumentTill",$_Cfg.RunWithoutArgumentTill)

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
        Write-Host "ScriptTempDirIN`t`t:[$ScriptTempDirIN]"
        Write-Host "ScriptTempDirOUT`t:[$ScriptTempDirOUT]"
        Write-Host "ScriptDataArchive`t:[$ScriptDataArchive]"
        Write-Host "`nFrom config file:"
        Write-Host "-----------------------------------------------------"
        $Cfg | Format-Table name,value
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
    $Script:isItEnd = $false
    $fileListFromWall= @()
    $fileListFromIN=@()
    $fileListfromOUT=@()
    $timeStamp = Get-Date -Format "yyyyMMddHHss"
    #pārbaudām vai visas mapes ir pieejamas
    $cfg.GetEnumerator() | 
    Where-Object { $_.name -match 'Path' } | 
        ForEach-Object {
            if ( -not ( Test-Path $_.value ) ) {
                Write-msg -log -bug -text "[Check] Directory [$($_.value)] is not accessible. Exit";
                $Script:isItEnd = $true
            }#endif
        }#endforeach

    if ( $Script:isItEnd ) { 
        Stop-Watch -Timer $ScriptWatch -Name Script
        Send-Mail
        Remove-Item $__LockFile -Force
        Exit 1
    }#endif

    <#-----------------------------------------------------
    #iztīram skripta temp direktorijas
    #-----------------------------------------------------#>
    if ( Test-Path -Path "$ScriptTempDirIN\*.txt" -PathType Leaf ) {
        Write-msg -text "[tmpInCheck] Found temporary files in [$ScriptTempDirIN]. Delete them."
        $null = Remove-Item "$ScriptTempDirIN\*.txt"
    }#endif
    if ( Test-Path -Path "$ScriptTempDirOUT\*.txt" -PathType Leaf ) {
        Write-msg -text "[tmpOutCheck] Found temporary files in [$ScriptTempDirOUT]. Delete them."
        $null = Remove-Item "$ScriptTempDirOUT\*.txt"
    }#endif

    <#-----------------------------------------------------
    #ielasām Wall importa direktorijā esošo datņu atribūtus
    #----------------------------------------------------#>
    $fileListFromWall = Get-ChildItem -Path $Cfg.WallSharePath -File | Where-Object -Property Length -gt 0
    if ( $fileListFromWall.count -gt 0 ) {
        Write-msg -log -text "[Import] found [$($fileListFromWall.count)] files for import from [$($Cfg.WallSharePath)]:"

        foreach ( $WallFile in $fileListFromWall ) {
            Write-msg -log -text "[WALL] file [$($WallFile.Name)], LastWriteTime:[$($WallFile.LastWriteTime)]"
            if ( -not ( $WallFile.LastWriteTime.ToShortDateString() -eq ([datetime]::Today).ToShortDateString() ) ) {
                if ( -not $Test ) {
                    Write-msg -log -text "[WALL] file [$($WallFile.Name)] creation time is not today!"
                }#endif
            }#endif
        }#endforeach

        #kopējam/pārvietojam datnes no Wall importa direktorijas
        #----------------------------------------------------#>
        try {
            if ( $Test ) {
                $fileListFromWall | Copy-Item -Destination $ScriptTempDirIN -ErrorAction Stop
            }#endif
            else {
                $fileListFromWall | Move-Item -Destination $ScriptTempDirIN -ErrorAction Stop
            }#else
        }#endtry
        catch {
            Write-ErrorMsg -Name 'Import' -InputObject $_
        }#endcatch

        #kovertējam datnes UTF-8 kodējumā un dzēšam no tmpIN mapes
        #----------------------------------------------------#>
        try {
            $fileListFromIN = Get-ChildItem -Path "$ScriptTempDirIN\in*.txt" -File | Where-Object -Property Length -gt 0

            #pārbaudām vai importējamo datņu skaits sakrīt ar esošajām tmpIn mapē
            if ( $fileListFromWall.count -eq $fileListFromIN.count ) 
            {
                Write-msg -log -text "[Encoding] to UTF-8 file:[$(if($NoEncoding){"Disabled"}else{"Enabled"})]."
                if ( $NoEncoding ) 
                {
                    $fileListFromIN | Copy-Item -Destination $ScriptTempDirOUT -ErrorAction Stop
                }#endif
                else 
                {

                    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
                    foreach ( $file in $fileListFromIN)
                    {
                        Write-msg -log -text "[Encoding] to UTF-8 file [$($file.FullName)]"
                        $content = Get-Content $file.FullName 
                        $TempDirOUT = "$ScriptTempDirOUT\$($file.Name)"
                        [System.IO.File]::WriteAllLines($TempDirOUT, $content, $Utf8NoBomEncoding)
                    }#endforeach
                }#endelse

                if ($Test) {Write-msg -log -text "[Encoding] delete temp files from [$ScriptTempDirIN]"}
                $fileListFromIN | Remove-Item
            }#endif
            else 
            {
                Write-msg -log -bug -text "[Encoding] there's just [$($fileListFromIn.count)] of [$($fileListFromWall.count)] files in [$ScriptTempDirIN]"
                Stop-Watch -Timer $ScriptWatch -Name Script
                Send-Mail
                Remove-Item $__LockFile -Force
                Exit 1
            }#endelse
        }#endtry
        catch {
            #Write-msg -log -bug -text "[FileEncode] Error: $($_.Exception.Message)"
            Write-ErrorMsg -Name 'Encoding' -InputObject $_
        }#endcatch

        #kopējam datnes uz mērķa Siron direktoriju
        #----------------------------------------------------#>
        $fileListfromOUT = Get-ChildItem -Path $ScriptTempDirOUT -File 

        if ( $getTempINfileList.count -eq $filesToSiron.count ) {
            try {
                Write-msg -log -text "[Export] files to [$($Cfg.SironDataInputPath)]"
                $fileListfromOUT | Copy-Item -Destination $Cfg.SironDataInputPath -Force -ErrorAction Stop

                #arhivējam avota datnes no SironAML InputDir uz arhīva mapi
                #--------------------------------------------------------#>
                Write-msg -log -text "[Compress] files to archive [$ScriptDataArchive\in-$timeStamp.zip]"
                
                Get-ChildItem -Path "$($Cfg.SironDataInputPath)\in*.txt" |
                    Where-Object { $_.LastWriteTime -ge ( [datetime]::Today ) -and $_.Length -ne 0 } |
                    Compress-Archive  -DestinationPath "$ScriptDataArchive\in-$timeStamp" -ErrorAction Stop

                    if ($Test) {Write-msg -log -text "[Delete] temp files from [$ScriptTempDirOUT]"}
                    $fileListfromOUT | Remove-Item
            }#endtry
            catch {
                Write-ErrorMsg -Name 'Export' -InputObject $_
            }#endcatch


        }#endif
        else {
            Write-msg -log -bug -text "[Export] there's just [$($fileListfromOUT.count)] of [$($fileListFromWall.count)] files in [$ScriptTempDirOUT]"
            Stop-Watch -Timer $ScriptWatch -Name Script
            Send-Mail
            Remove-Item $__LockFile -Force
            Exit 1
        }#endelse

    }#endif
    else {
        Write-msg -log -text "[WALL] There's no import files in [$($Cfg.WallSharePath)]"
    }#endelse
    
    <#-----------------------------------------------------
    #izpildām SIRON AML datu ielādes komandu
    #----------------------------------------------------#>
    if ( -not ( Test-Path -Path $Cfg.SironAMLBatchPath ) ) {
        Write-msg -log -bug -text "[Runner] [$($Cfg.SironAMLBatchPath)] is no accessible."
        Write-msg -log -bug -text "[Runner] Execution of command [$($Cfg.SironAMLBatchCommand)] skipped!"
    }#endif
    else {

        try {
        #Start-SironBatch [Command+Argument] [WorkingDir] [LogFile]
        $timestamp = Get-Date -Format "yyyyMMddHHmm"

        if ( ( (Get-Date) -ge ([datetime]::Today) ) -and ( (Get-Date) -le ([datetime]::Today).AddHours($Cfg.RunWithoutArgumentTill) ) ) {

            Write-msg -log -text "[Runner] runs [$($Cfg.SironAMLBatchCommand)] in [$($Cfg.SironAMLBatchPath)] log to [$LogFileDir\SironAMLscoring-$timestamp.log]"
            $null = Start-SironBatch "$($Cfg.SironAMLBatchCommand)" $Cfg.SironAMLBatchPath "$LogFileDir\SironAMLscoring-$timestamp.log"
        }#endif
        else {

            Write-msg -log -text "[Runner] runs [$($Cfg.SironAMLBatchCommand) $($Cfg.SironAMLBatchCommandArgument)] in [$($Cfg.SironAMLBatchPath)] log to [$LogFileDir\SironAMLscoring-$timestamp.log]"
            $null = Start-SironBatch "$($Cfg.SironAMLBatchCommand) $($Cfg.SironAMLBatchCommandArgument)" $Cfg.SironAMLBatchPath "$LogFileDir\SironAMLscoring-$timestamp.log"
        }#endelse


        }#endtry
        catch {
            Write-ErrorMsg -Name 'Export' -InputObject $_
        }#endcatch

        try {
            $SironImportLog = Get-Content -Path $Cfg.SironImportLogFileName -ErrorAction Stop
            $SironImportFile = Get-ChildItem -Path $Cfg.SironImportLogFileName -ErrorAction Stop
            
            #pārbaudam log datnes šodienas izveides
            if ( $SironImportFile.LastWriteTime.ToShortDateString() -eq ([datetime]::Today).ToShortDateString() ) {
                #meklējam divus ierakstus
                # SUCCESS:
                # ERROR:
                foreach ( $line in $SironImportLog) {
                    if ( $line.Contains("ERROR:") ) {
                        Write-msg -log -bug -text "[SironLOG] [$($Cfg.SironImportLogFileName)]"
                        Write-msg -log -bug -text "[SironLOG] from log: $line"

                        #Ievācam pierādījumus
                        Get-PostCheck
                        break
                    }#endif
                    if ( $line.Contains("SUCCESS:") ) {
                        Write-msg -log -text "[SironLOG] from log: $line"
                        break
                    }#endif
                }#endforeach
            }#endif
            else {
                Write-msg -log -bug -text "[SironLOG] the log file's creation [$($SironImportFile.LastWriteTime.ToShortDateString())] time is not today!"
            }#endelse

        }#endtry
        catch {
            Write-ErrorMsg -Name 'SironLOG' -InputObject $_
        }#endcatch
    }#endelse

    <#-----------------------------------------------------
    #Dzēšam tmp un err mapju saturu, kas vecāks par 30 dienām
    #laika periodu varam koriģēt funkcijā 
    #----------------------------------------------------#>
    Set-Housekeeping $ScriptDataArchive
    Set-Housekeeping $LogFileDir

}#endOfprocess

end {

    Stop-Watch -Timer $ScriptWatch -Name Script
    Send-Mail
    Remove-Item $__LockFile -Force
    <#iespējojam, ja griba sūtīt e-pastus tikai kļūdu gadījumā
    foreach ( $line in $Script:forMailReport) {
        if ( $line.Contains("[ERROR]") ) { Send-Mail; break }
    }#endforeach
    #   #>
}#endOfend