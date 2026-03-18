$cAppName 			 = "Set_ComputerName_By_MACAddress_from_file"
$cLogPath 			 = "C:\Dell\"

# required to RENAME the PC
# if the PC is already joined to the domain then  DomAccount   ( Dom Admin ? )              is required
# if the PC is  NOT    joined to the domain then  LocalAccount ( member of Administrators ) is required

$cComputerRenameAccount   = "administrator"		
$cComputerRenamePWD       = "TopSecret1"

$cDomainJoinDomain        = "Deploy"
$cDomainJoinDomainOU      = "OU=testou,DC=Deploy,DC=net"               # OU string has to match the AD path, notice upper and lower-case, notice spaces
$cDomainJoinAccount       = "svcAccountDomJoin"                        # "svcAccountDomJoin"
$cDomainJoinPWD           = "TopSecret"                                # "TopSecret"

$cCSVPsDrive_Root		  = "\\Win11-c\share$"                         # \\Server\ShareName  of the CSV file 
$cCSVFilePath             = "MACADDRESS_To_HOSTNAME.CSV"
        #    TAB separated list expected in CSV file with  MACADDRESS{TAB}HOSTNAME
        #
        # 5C:26:0A:64:33:93{TAB}WN7-2z286q3
        # 5C:26:0A:64:33:92{TAB}WN7-2z286q2
        # 5C:26:0A:64:33:97{TAB}WN7-2z286q1
        # 5C:26:0A:64:33:90{TAB}WN7-2z286q0

$cCSVFilePathUnc          = "$cCSVPsDrive_Root\$cCSVFilePath"
$cCSVFileReadAccount	  =  $cComputerRenameAccount
$cCSVFileReadPWD		  =  $cComputerRenamePWD

$bComputerName_has_to_be_Changed = $false		# if a rename is required 
$sNewComputerNameFromCSV         = $Null

Function Main
    {
    cls
    $Logfile = $cLogPath + $cAppName + ".log" 
    If ( -not ( Test-Path $cLogPath ) ) { $NewFoloder = New-Item -path $cLogPath -ItemType Directory}
    If ( Test-Path $Logfile ) { Remove-Item -path $Logfile -force }

    Logging "$cAppName - Start" 
    Logging "$cAppName - Part 1 - Verify that the network share and CSV file (MACAddress <-> Computername) are available"
    Logging "$cAppName - Part 1 - a) PRE-Script-Run - analyze network connections ( Get-PS-Drive(), Get-SmbMapping() ) " 
    Logging "$cAppName - Part 1 - b) PRE-Script-Run - is $cCSVPsDrive_Root already connected or SMB-mapped" 
    Logging "$cAppName - Part 2 - a) prepare network - either New-PS-Drive() for $cCSVPsDrive_Root" 
    Logging "$cAppName - Part 2 - a) prepare network - or use existing PS-Drive/SmbMapping to $cCSVPsDrive_Root" 
    Logging "$cAppName - Part 3 - is CSV file (MAC<->Hostname) accessable over the network"
    Logging "$cAppName - Part 4 - get current NetIPConfigurations (MacAddress,IPv4Address,Interface/Adapter)"
    Logging "$cAppName - Part 6 - before JOIN / RENAME - delete existing 'NetSetup.LOG'"
    Logging "$cAppName - Part 7 - JOIN DOMAIN - 2026-03-18:  will be skipped, because not fully developped"
    Logging "$cAppName - Part 8 - RENAME - rename current hostname to '`$CSVComputerName'"
    
    Logging "$cAppName ============== Part 1 ======================================================================="
    # Query-PSDrive_Existence -SearchProperty DriveLetter    -SearchForString 'C'
    # Query-PSDrive_Existence -SearchProperty Root           -SearchForString '\\Win11-c\share$'
    $PsDriveAlreadyConnected    = Query-PSDrive_Existence    -SearchProperty Root       -SearchForString $cCSVPsDrive_Root
    $SmbMappingAlreadyConnected = Query-SmbMapping_Existence -SearchProperty RemotePath -SearchForString $cCSVPsDrive_Root

    Logging "$cAppName ============== Part 2 ======================================================================="
    if ( $PsDriveAlreadyConnected -eq $Null -and $SmbMappingAlreadyConnected -eq $Null )
        {
        Logging "$cAppName - Part 2 - b) prepare network - New-PS-Drive() for $cCSVPsDrive_Root" 
        $PsDrive_Name_DriveLetterOnly = Get-FreeDriveLetter 
        Logging "$cAppName - NO PSDrive and NO SmbMapping -> create new PSDrive with (next free) letter '$PsDrive_Name_DriveLetterOnly'"
        $NewPsDrive = NetworkShare_Connect -DriveLetter $PsDrive_Name_DriveLetterOnly -NetworkShare $cCSVPsDrive_Root -FileReadAccount $cCSVFileReadAccount -FileReadPWD $cCSVFileReadPWD
        #  $NewPsDrive | FT | Out-String | Logging
        if ( $NewPsDrive -eq $Null) 
            {
            Logging "$cAppName - New-PS-Drive() to '$cCSVPsDrive_Root' failed - soemthing went wrong"
            Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
            Return -2147467259  # = 0x80004005
            }
        Logging "$cAppName - analyze current network connections"
        $PsDriveNowConnected    = Query-PSDrive_Existence    -SearchProperty Root       -SearchForString $cCSVPsDrive_Root
        $SmbMappingNowConnected = Query-SmbMapping_Existence -SearchProperty RemotePath -SearchForString $cCSVPsDrive_Root
        }
    else
        {
        Logging "$cAppName - Part 2 - b) prepare network - re-use existing New-PS-Drive() for $cCSVPsDrive_Root" 
        $PsDrive_Name_DriveLetterOnly = $Null
        $PsDriveNowConnected = $PsDriveAlreadyConnected 
        $SmbMappingNowConnected = $SmbMappingAlreadyConnected 
        }
    
    if (  $PsDriveNowConnected )
        {
        Logging "$cAppName - PS-Drive exists"
        $PsDriveNowConnected | Select Name,Provider,Root | Out-String | Logging #  Write-Host
        }
    else
        { 
        Logging "$cAppName - PS-Drive to '$cCSVPsDrive_Root' is missing - soemthing went wrong"
        Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
        Return 5
        }


    # ###########################################################
    Logging "$cAppName ============== Part 3 ======================================================================="
    Logging "$cAppName - Part 3 - is CSV file (MAC<->Hostname) accessable over the network"
    # only for DEBUG
    # Get-SmbMapping
    if (  $SmbMappingNowConnected )
        {
        Logging "$cAppName - SmbMapping to '$cCSVPsDrive_Root' exists"
        $SmbMappingNowConnected | Out-String | Logging # Write-Host

        if (-not (Test-Path  $cCSVPsDrive_Root)){  Logging "$cAppName - ERROR:   Network share is NOT accessable - '$cCSVPsDrive_Root'"     }
        else                                    {  Logging "$cAppName - SUCCESS: Network share is accessable     - '$cCSVPsDrive_Root'"     }
    
        if (-not (Test-Path  $cCSVFilePathUnc)) {  Logging "$cAppName - ERROR:   CSV file is NOT accessable      - '$cCSVFilePathUnc'"  }
        else                                    {  Logging "$cAppName - SUCCESS: CSV file  is accessable         - '$cCSVFilePathUnc'"  }

        If ( (Test-Path $cCSVFilePathUnc) -eq $False ) 
            {
    		Logging "$FctName ERROR - file is missing '$cCSVFilePathUnc'"
	    	Logging "$FctName END - ERROR`r`n"  
            Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
            return 1
		    }
        }
    else
        { 
        Logging "$cAppName - SmbMapping to '#$cCSVPsDrive_Root' is missing - soemthing went wrong"
        Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
        Return 5
        }

    # ###########################################################
    Logging "$cAppName ============== Part 4 ======================================================================="
    Logging "$cAppName - Part 4 - get current NetIPConfigurations (MacAddress,IPv4Address,Interface/Adapter)"
    $AllNetIPConfigurations = Get-NetIPConfiguration | Where {   $_.NetAdapter.Status -ne 'Disconnected' } |   select  @{n='MacAddress'; e={$_.NetAdapter.MacAddress}}, @{n='IPv4Address';e={$_.IPv4Address[0]}}, InterfaceAlias, InterfaceDescription
    $AllNetIPConfigurations | ft | out-string |  Logging # write-host
        <#
                MacAddress        IPv4Address     InterfaceAlias                    
                ----------        -----------     --------------                    
                28-00-AF-0A-B6-C8 192.168.0.95    vEthernet (External USB GbE WD19) 
                00-15-5D-00-5F-00 169.254.143.228 vEthernet (Internal - HV and Host)
                00-15-5D-6A-64-44 172.26.80.1     vEthernet (Default Switch)     
        #>

    # ###########################################################
    Logging "$cAppName ============== Part 5 ======================================================================="
    Logging "$cAppName - Part 5 - search MAC in CSV"
    $ComputerName = $Null
    ForEach ($NetIPConfiguration In $AllNetIPConfigurations)
        {
	    $ComputerName = Get-ComputerName_From_Csv -CsvFilePath $cCSVFilePathUnc -MACADDRESS $NetIPConfiguration.MACAddress   # '5C:26:0A:64:33:93'
   		If ( $ComputerName -ne $Null )
            { 
   			Logging  "$cAppName - new Computername '$ComputerName' found in CSV file, no more MACAddress-es will be searched" 
   			break
            }
        else   
            { 
            Logging  "$cAppName - new Computername NOT found in CSV file, next MACAddress will get searched"  
            }
        }
    
    If ( $ComputerName -eq $Null ) {
        Logging  "$cAppName - new Computername is NOT found in CSV file "
        Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected  -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
        get-SmbMapping
        return 1
        }
   
    Disconnect_PsDrive__if_required -StatusBefore $PsDriveAlreadyConnected  -StatusNow $PsDriveNowConnected -DisconnectDriveLetter $PsDrive_Name_DriveLetterOnly
    if ( get-SmbMapping -RemotePath $cCSVPsDrive_Root  -ErrorAction SilentlyContinue  )
        {
        Logging  "$cAppName - Remove-SmbMapping -RemotePath $cCSVPsDrive_Root  "
        $RemoveResult = Remove-SmbMapping -RemotePath $cCSVPsDrive_Root -force -PassThru   # -ErrorAction SilentlyContinue
        $RemoveResult
        }

    # ###########################################################
    Logging "$cAppName ============== Part 6 ======================================================================="
    Logging "$cAppName - Part 6 - before JOIN / RENAME - delete existing 'NetSetup.LOG'"
    
    # elevation required to remove "c:\windows\debug\NetSetup.LOG"
    # this files save details / logging for any kind of Domain-Join
    if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
        {    
        Logging "$cAppName - Script is started:   Elevated."
        $IsElevated = $True
        If ( Test-Path "c:\windows\debug\NetSetup.LOG" ) { Remove-Item -path  "c:\windows\debug\NetSetup.LOG" -force }
        }
    else 
        {    
        Logging "$cAppName - Script is started 'Not elevated' - cannot remove 'C:\Windows\Debug\NetSetup.LOG'"
        $IsElevated = $false
        }

    # ###########################################################
    Logging "$cAppName ============== Part 7 ======================================================================="
    Logging "$cAppName - Part 7 - JOINJ DOMAIN - 2026-03-18:  will be skipped, because not fully developped"
    Logging "$cAppName"
    #  Join_Domain_Or_Workgroup   "DOMAIN"

    # ###########################################################
    Logging "$cAppName ============== Part 8 ======================================================================="
    Logging "$cAppName - Part 8 - RENAME - rename current hostname to '$ComputerName'"
    [Int]$RenameResult          = 0
    [String]$RenameResultString = $Null
    [Bool]$RenameResultBool     = $False

    if ( $IsElevated -eq $True )
        {
        #		WMIC ComputerSystem where Name=COMPUTERNAME call Rename Name=NewName
        #		WMIC ComputerSystem where Name= COMPUTER-NAME  call Rename Name=NewName
        $RenameResultString = ComputerName__Rename -NewComputerName  $ComputerName  -sLocalAdminAccount $cComputerRenameAccount -sLocalAdminPWD  $cComputerRenamePWD
       # $RenameResult | Out-String | Logging # Write-Host
        if ( $RenameResultString -match "The changes will take effect after you restart the computer") { $RenameResult = 3010}
        }
    else
        {
        Logging "$cAppName - Script is started 'Not elevated' - cannot rename current hostname to '$ComputerName'"
        }
  Logging "$cAppName - End - return = '$RenameResult'" 
  return $RenameResult
  }  

#=============================================================================
#=============================================================================
Function ComputerName__Rename {
    param ( [String]$NewComputerName, [String]$sLocalAdminAccount, [String]$sLocalAdminPWD)
    <#
        $OldComputerName = (gwmi win32_computersystem).Name.Trim()
        $ComputerObject = Get-WmiObject Win32_ComputerSystem -ComputerName $OldComputerName
    
        # Set the computer name
        Write-Host "Renaming computer from '$OldComputerName' to '$NewComputerName'"
        $passThru = Rename-Computer -NewName $newName -PassThru
    #>

    [String]$FctName = "ComputerName__Rename() -"
    [String]$FctReturn = $Null
	Logging "$FctName START - new name '$NewComputerName'"  

	$OldComputerName = (gwmi win32_computersystem).Name.Trim()
    Logging ("$FctName current ComputerName = '" + $OldComputerName + "'")
	Logging ("$FctName new     ComputerName = '" + $NewComputerName + "'")
		
	If ( $NewComputerName.ToUpper() -eq $OldComputerName.ToUpper() )  
        {
		$FctReturn = 'new name already set - nothing to do'
        Logging "$FctName $FctReturn"
        return $FctReturn
		}
    else
        {

        Try {
            $RenameReturn = Rename-Computer -NewName $NewComputerName -PassThru -ErrorAction stop
            }
        catch [System.ComponentModel.Win32Exception]    
            {
            $ErrCodeHex = '0x{0:X}' -f $_.Exception.ErrorCode  # (  -2147467259 -> 0x80004005 )
            # Error 0x80004005 is a general "unspecified error" in Windows, commonly indicating 
            # - permission issues, 
            # - blocked network access, 
            # - failed updates. 
            # Solutions include running the Windows Update troubleshooter, enabling SMBv1, checking file permissions

            Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "
            Logging "$FctName ERROR - $($_.Exception.ErrorCode) = $ErrCodeHex"  
            Logging "$FctName ERROR - $($_.Exception.Message)"   
            Logging "$FctName ERROR - Please improve the script"
            $_.Exception | select * | FL | Out-String | Logging # Write-Host
            
            }
        catch
            {
            $ErrCodeHex = '0x{0:X}' -f $_.Exception.ErrorCode  # (  -2147467259 -> 0x80004005 )
            Logging "$FctName ERROR - $($_.Exception.ErrorCode) = $ErrCodeHex" 
            Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "
            Logging "$FctName ERROR - Exception Type        : $($_.Exception.GetType().FullName)" -ForegroundColor Red
            # Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "Exception Type: $($_.Exception.GetType().Name)"     -ForegroundColor Red     # 'SessionStateException' or 'DriveNotFoundException'
            Logging "$FctName ERROR - Exception Message     : $($_.Exception.Message)"         -ForegroundColor Red
        
            <#
                 If ( $Return = 1326 ) 
                        {   
					    Logging "$FctName ERROR - general description : 'invalid credentials' - see netsetup.log"
					    Logging "$FctName ERROR - maybe 'connection to DomainController failed' - do not use different credentials for DomJoinAccount and NetAccessAccount if CSV-file is located on DC - see netsetup.log"
                        }
            #>
            }
    
        <#
           # $RenameReturn | select *
    
            HasSucceeded    : True
            NewComputerName : HV-ImageAssist
            OldComputerName : WIN-MUFOEUAEPP4
        #>
        [Bool]$RenameReturnHasSucceeded = $RenameReturn.HasSucceeded
		If ( $RenameReturnHasSucceeded -ne $True ) 
            {
			$FctReturn = "Rename failed. Error = $($Err.Number) - $($Err.Description) - for troubleshooting start  notepad.exe c:\windows\debug\NetSetup.LOG"
            }
		Else
            {
            $FctReturn = "The changes will take effect after you restart the computer $OldComputerName."
            }
        }
	Logging "$FctName END - return = '$FctReturn'" 
    Return $FctReturn
    }



#=============================================================================
#=============================================================================
Function Join_Domain_Or_Workgroup
    {
    param ( [String]$sDomain_or_Workgroup )
    <#
#The JoinDomainOrWorkgroup method joins a computer system to a domain or workgroup. This method is new for Windows  XP.!
#
#Function JoinDomainOrWorkgroup(  ByVal Name As String, ByVal Password As String,  ByVal UserName As String, [ ByVal AccountOU As String ],  ByVal FJoinOptions As Integer ) As Integer
#
#Parameters
#############
#Name 			[in] Specifies the domain or workgroup to join. Cannot be NULL. 
#Password 		[in] If the UserName parameter specifies an account name, the Password parameter must point to the password to use when connecting to the domain controller. Otherwise, this parameter must be NULL. 
#				Password must use a high authentication level, not less than RPC_C_AUTHN_LEVEL_PKT_PRIVACY when connecting to Winmgmt or SetProxyBlanket on the IWbemServices pointer. If local to Winmgmt this is not a concern.
#UserName 		[in] Pointer to a constant null-terminated character string that specifies the account name to use when connecting to the domain controller. Must specify a domain NetBIOS name and user account, for example, Domain\user. If this parameter is NULL, the caller#s context is used.
#				Windows XP, Windows .NET Server 2003 family:  Using the user principal name (UPN) in the form user@domain is also supported.
#				UserName must use a high authentication level, not less than RPC_C_AUTHN_LEVEL_PKT_PRIVACY when connecting to Winmgmt or SetProxyBlanket on the IWbemServices pointer. If local to Winmgmt this is not a concern.
#AccountOU 		[in, optional] Specifies the pointer to a constant null-terminated character string that contains the RFC 1779 format name of the organizational unit (OU) for the computer account. If you specify this parameter, the string must contain a full path, otherwise AccountOU must be NULL.
#				Example: OU=testOU, DC=domain, DC=Domain, DC=com 
#FJoinOptions 	[in] Set of bit flags defining the join options.

#				Value Meaning 
#				#############
#				0 Join Domain				Default. Joins the computer to a domain. If this value is not specified, joins the computer to a workgroup.
#				1 Acct Create				Creates the account on the domain.
#				2 Acct Delete				Delete the account when the domain is left.
#				4 Win9X Upgrade				The join operation is part of an upgrade of Windows 95/98 to Windows NT/2000.
#				5 Domain Join If Joined		Allows a join to a new domain even if the computer is already joined to a domain.
#				6 Join Unsecure				Performs an unsecured join.
#				7 Machine Password Passed	The machine, not the user, password passed. This option is only valid for unsecure joins.
#				8 Deferred SPN Set			Writing SPN and DnsHostName attribtes on the computer object should be deferred until the rename that follows the join.
#				18 Install Invocation		The APIs were invoked during install. 
#
# Return Values	The JoinDomainOrWorkgroup method returns 0 on success or if no options are involved. Any other value indicates an error.
#
#	sDomain_or_Workgroup = "DOMAIN"    or "WORKGROUP"
#   if "WORKGROUP" => 	OU = NULL
#>
	Logging "Join_Domain_Or_Workgroup() - START - '$sDomain_or_Workgroup'"

	$JOIN_WORKGROUP = 0
	$JOIN_DOMAIN = 1
	$ACCT_CREATE = 2
	$ACCT_DELETE = 4
	$WIN9X_UPGRADE = 16
	$DOMAIN_JOIN_IF_JOINED = 32
	$JOIN_UNSECURE = 64
	$MACHINE_PASSWORD_PASSED = 128
	$DEFERRED_SPN_SET = 256
	$INSTALL_INVOCATION = 262144


#	$cDomainJoinDomain   = "Deploy"
#	$cDomainJoinDomainOU = "OU=testOU, DC=domain, DC=Deploy, DC=com"
#	$cDomainJoinAccount  = "svcAccountDomJoin"
#	$cDomainJoinPWD      = "TopSecret"
	
	
	$objNetwork = CreateObject("WScript.Network")
	$strComputer = objNetwork.ComputerName

	$objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\$strComputer\root\cimv2:Win32_ComputerSystem.Name='$strComputer'")
	If ( $sDomain_or_Workgroup -eq "DOMAIN" ) 
        {
		Logging "Join_Domain_Or_Workgroup() - Joining the Domain       = '$cDomainJoinDomain'"
		Logging "Join_Domain_Or_Workgroup() - Joining the OU           = '$cDomainJoinDomainOU'"
		Logging "Join_Domain_Or_Workgroup() - Joining with JoinAccount = '$cDomainJoinDomain\$cDomainJoinAccount'"
        ReturnValue = $objComputer.JoinDomainOrWorkGroup($cDomainJoinDomain, $cDomainJoinPWD, "$cDomainJoinDomain\$cDomainJoinAccount", $cDomainJoinDomainOU, $JOIN_DOMAIN + $ACCT_CREATE)
		Logging "Join_Domain_Or_Workgroup() - Reboot for Domain-Join to go into effect"
        }
	ElseIf ( sDomain_or_Workgroup -eq "WORKGROUP" ) 
        {
		Logging "Join_Domain_Or_Workgroup() - Joining the Workgroup '$cDomainJoinDomain'"  
		ReturnValue = objComputer.JoinDomainOrWorkGroup($cDomainJoinDomain,$NULL ,$NULL , $NULL, $JOIN_WORKGROUP + $ACCT_CREATE + $DEFERRED_SPN_SET)
		Logging "Join_Domain_Or_Workgroup() - Reboot for Workgroup-Join to go into effect"
        }
	Else
        {
		Logging "Join_Domain_Or_Workgroup() - '$sDomain_or_Workgroup' is unknown or misstyped" 
		ReturnValue = 9999
	    }

	If ( Returnvalue -eq 234 -or Returnvalue -eq 2 ) 
        {   
		Logging "Join_Domain_Or_Workgroup() - ERROR '$cDomainJoinDomainOU' =   specified OU is not supported" 
	    }
 
	If ( Returnvalue -eq 2691 ) 
        {   
		Logging "Join_Domain_Or_Workgroup() - ERROR  the specified machine is already joined to '$cDomainJoinDomain'" 
	    }

	If ( Returnvalue -eq 2224 ) 
        {   
		Logging "Join_Domain_Or_Workgroup() - ERROR  the specified machine account already exists in domain '$cDomainJoinDomain'" 
		Logging "Join_Domain_Or_Workgroup() - ERROR  desired OU string '$cDomainJoinDomainOU'"
		Logging "Join_Domain_Or_Workgroup() - ERROR  OU string has to match exactly the AD path, notice upper and lower-case, notice spaces" 
	    }

	If ( $Returnvalue -ne 0 ) 
        { 
		Logging "`r`n`r`n" +  "Join_Domain_Or_Workgroup() - for troubleshooting start  notepad.exe c:\windows\debug\NetSetup.LOG" 
		Logging "`r`nJoin_Domain_Or_Workgroup() - ERROR`r`n" 
	    }

	Logging "Join_Domain_Or_Workgroup() - END - return = '$ReturnValue'" 
    return $ReturnValue 
	
}


#=============================================================================
#=============================================================================
Function Disconnect_PsDrive__if_required 
    {
    Param ($StatusBefore , $StatusNow, [String]$DisconnectDriveLetter )
    $FctName          = "Disconnect_PsDrive__if_required() -"
    Logging "$FctName PSDrive 'StatusBefore'          = '$StatusBefore'"
    Logging "$FctName PSDrive 'StatusNow'             = '$StatusNow'"
    Logging "$FctName PSDrive 'DisconnectDriveLetter' = '$DisconnectDriveLetter'"
    if ( $StatusBefore -eq $Null -and $StatusNow -ne $Null )
        {
        # Remove PS-Drive if connected by this script
        If ( $DisconnectDriveLetter  -ne $Null )
            {
            Logging "$FctName PSDrive '$DisconnectDriveLetter' was created by this script - will get removed"
            # get-PSDrive    -Name "A"
            # Remove-PSDrive -Name "A"
            Remove-PSDrive -Name $DisconnectDriveLetter
            }
        }
    }   


#=============================================================================
#=============================================================================
Function Get-FreeDriveLetter 
    {
    $FreeDriveLetter = $Null
    $drvlist=(Get-PSDrive -PSProvider filesystem).Name
    Foreach ($drvletter in "ABDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()) 
        {
        If ($drvlist -notcontains $drvletter) {
            $FreeDriveLetter = $drvletter
            break
            }
        }
    return $FreeDriveLetter
    }


# =============================================================================
# =============================================================================
Function Get-ComputerName_From_Csv
    {
    param ([String]$CsvFilePath, [String]$MACADDRESS , [Switch]$MyDebug)

<#
    # EXAMPLE
    #
    # CSV file is TAB separated    (    5C:26:0A:64:33:93	WN7-2z286q3 )
    #
    # MACADDRESS =  '5C:26:0A:64:33:93'

    $ComputerName = Get-ComputerName_From_Csv -CsvFilePath "C:\__B\Rename_and_Join_by_CSV_v3\MACADDRESS_To_HOSTNAME.CSV" -MACADDRESS '5C:26:0A:64:33:93'
     
    #
    # => Computername = 'WN10-2z286q3"
    #
#>

    $FctName          = "Get-ComputerName_From_Csv() -"
    $NewComputerName = $Null

    if ( $MyDebug ) { Logging "$FctName START"  }
	If ( (Test-Path $CsvFilePath) -eq $False ) 
        {
		Logging "$FctName ERROR - file is missing '$CsvFilePath'"
		if ( $MyDebug ) { Logging "$FctName END - ERROR`r`n"  }
		Return $sNewComputerName
        }

    
    # https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7.5#-delimiter
    $Delimter = "`t"    # The default is a comma (,).
                        # Enter a character, such as a colon (:). 
                        # To specify a semicolon (;), enclose it in single quotation marks. 
                        # To specify escaped special characters such as tab (`t), enclose it in double quotation marks.

    $Header = 'MAC', 'Computername'      
    Logging "$FctName Openening CSV file (TAB separated list) '$CsvFilePath' - query for row containing '$MACADDRESS'"                
    $NewComputerName= (Import-Csv -Delimiter $Delimter   -Path $CsvFilePath -Header $Header  | Where-Object -Property MAC -eq $MACADDRESS).ComputerName
    If ( $NewComputerName -eq $Null) { 
        $MACADDRESS = $MACADDRESS.replace("-",":")
        Logging "$FctName Openening CSV file (TAB separated list) '$CsvFilePath' - query for row containing '$MACADDRESS'" 
        $NewComputerName= (Import-Csv -Delimiter $Delimter   -Path $CsvFilePath -Header $Header  | Where-Object -Property MAC -eq $MACADDRESS).ComputerName
        }


    If ( $NewComputerName -eq $Null) { Logging "$FctName NO new computername found!"  } else {  Logging "$FctName new computername found '$NewComputerName'"  }
	if ( $MyDebug ) { Logging "$FctName END" }
    Return $NewComputerName
	}


# =============================================================================
# =============================================================================
Function NetworkShare_Connect
    {
    param ( [String]$DriveLetter, [String]$NetworkShare, [String]$FileReadAccount, [String]$FileReadPWD, [Switch]$MyDebug)

    # example
    #  NetworkShare_Connect	-DriveLetter $PsDrive_Name_DriveLetterOnly  -NetworkShare $cCSVPsDrive_Root -FileReadAccount $cCSVFileReadAccount -FileReadPWD  $cCSVFileReadPWD

    $FctName = "NetworkShare_Connect() -"
    Logging "$FctName START - connect '$NetworkShare' as drive '$DriveLetter' - using Domain\Account '$FileReadAccount'"
	
    $SecureStringPwd      = $FileReadPWD | ConvertTo-SecureString -AsPlainText -Force
    $Creds = New-Object System.Management.Automation.PSCredential -ArgumentList ($FileReadAccount, $SecureStringPwd)

    $AllPsDrivesFileSystem = Get-PSDrive -PSProvider FileSystem   | select Name, Root, Description,DisplayRoot
    $NewPsDrive  = $AllPsDrivesFileSystem | Where { $_.Root -eq  $NetworkShare }
    if ( $NewPsDrive  -ne $Null )
        {
        # already connected  -> return already used drive letter
        $NewPsDrive | FT | Out-String | Logging # Write-Host
        }
    else
        {
        Try {
            #  -Scope is required, else drive will get disconnected at the end of this function
	        $NewPsDrive = New-PSDrive -Name $DriveLetter -Root $NetworkShare -PSProvider "FileSystem" -Credential $Creds -Scope Script -ErrorAction stop
            <#
                Name           Used (GB)     Free (GB) Provider      Root                                                                                                                                                                                                                                                                    CurrentLocation
                ----           ---------     --------- --------      ----                                                                                                                                                                                                                                                                    ---------------
                R                                      FileSystem    \\Win11-c\share$           
            #>
            }
        catch [System.Management.Automation.SessionStateException]   # [DriveAlreadyExists],[ResourceExists]
            {
            Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "
            if ( $($_.Exception.Message) -eq "A drive with the name '$DriveLetter' already exists.")
                {
                Logging "$FctName ERROR - Please improve the script: drive letter '$DriveLetter' is already in use"
                }
            else
                {
                Logging "$FctName ERROR - please disconnect all network drives to '$NetworkShare' before you restart the script"
                }
            }
        catch [System.ComponentModel.Win32Exception]   # 
            {
            $ErrCodeHex = '0x{0:X}' -f $_.Exception.ErrorCode  # (  -2147467259 -> 0x80004005 )
            # Error 0x80004005 is a general "unspecified error" in Windows, commonly indicating 
            # - permission issues, 
            # - blocked network access, 
            # - failed updates. 
            # Solutions include running the Windows Update troubleshooter, enabling SMBv1, checking file permissions

            Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "
            Logging "$FctName ERROR - $($_.Exception.ErrorCode) = $ErrCodeHex"  
            Logging "$FctName ERROR - $($_.Exception.Message)"   
            Logging "$FctName ERROR - Please improve the script"
            $_.Exception | select * | FL | Out-String | Logging # Write-Host
            
            }
        catch
            {
            Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "
            Logging "$FctName ERROR - Exception Type        : $($_.Exception.GetType().FullName)" -ForegroundColor Red
            # Logging "$FctName ERROR - FullyQualifiedErrorId : $($_.FullyQualifiedErrorId) "Exception Type: $($_.Exception.GetType().Name)"     -ForegroundColor Red     # 'SessionStateException' or 'DriveNotFoundException'
            Logging "$FctName ERROR - Exception Message     : $($_.Exception.Message)"         -ForegroundColor Red
        
                <#
                If ( $iReturn -eq -2147023570 ) 
                    {
			        Logging "$FctName ERROR - verify values of script constant 'cCSVFileReadAccount'='$cCSVFileReadAccount'"
			        Logging "$FctName ERROR - verify values of script constant 'cCSVFileReadPWD'    =<not-displayed-in-log-file>"
		            }

		        If ( $iReturn -eq -2147023677 ) 
                    {
			        Logging "$FctName ERROR - multible connections to a share using different accounts is not possible"
			        Logging "$FctName ERROR - please disconnect all network drives to '$NetworkShare' before you restart the script"
		            }
                #>
            }
        }
	If ( $NewPsDrive -eq $Null)  { Logging "$FctName ERROR - Terminating script"  }
	Logging "NetworkShare_Connect() - END"
    Return $NewPsDrive 
    }


#=============================================================================
#=============================================================================
Function Query-SmbMapping_Existence
    {
    param ( [ValidateSet("DriveLetter","RemotePath")]$SearchProperty, [String]$SearchForString, [Switch]$MyDebug )
    
    # Query-SmbMapping_Existence -SearchProperty DriveLetter -SearchForString 'C'
    # Query-SmbMapping_Existence -SearchProperty Root        -SearchForString '\\Win11-c\share$'

    # return:  either $NULL or SmbMapping object

    
    $FctName = "Query-SmbMapping_Existence() -"
    If ( $MyDebug  ) { Logging "$FctName START" }
	Logging "$FctName SmbMapping where '$SearchProperty'='$SearchForString' search ..."
   
    # $SmbMappings    = Get-SmbMapping -LocalPath $LL -ErrorAction SilentlyContinue  # | select *
    $SmbMappings      = get-SmbMapping | select LocalPath,RemotePath,Status,GlobalMapping
    
    <#
            $SmbMappings | ft
            
            LocalPath RemotePath       Status GlobalMapping
            --------- ----------       ------ -------------
                      \\Win11-c\share$     OK         False
    #>
    if ($SearchProperty -eq 'DriveLetter' ) { $MyProperty  = 'LocalPath' } else {   $MyProperty  = 'RemotePath' }
    $Exists = $SmbMappings | Where { $_.$MyProperty -eq $SearchForString }
    If ( $Exists -eq $Null ) { Logging "$FctName SmbMapping where '$SearchProperty'='$SearchForString' does not exist"  } else {  Logging "$FctName SmbMapping where '$SearchProperty'='$SearchForString' exists"  }
    If ( $MyDebug  ) { Logging "$FctName END" }
    return $Exists
    }


#=============================================================================
#=============================================================================
Function Query-PSDrive_Existence
    {

    # Query-PSDrive_Existence -SearchProperty DriveLetter -SearchForString 'C'
    # Query-PSDrive_Existence -SearchProperty Root        -SearchForString '\\Win11-c\share$'

    # return:  either $NULL or PSDrive object

    param ( [ValidateSet("DriveLetter","Root")]$SearchProperty, [String]$SearchForString, [Switch]$MyDebug )
    $FctName = "Query-PSDrive_Existence() -"
    If ( $MyDebug  ) { Logging "$FctName START" }
	Logging "$FctName PsDrive where '$SearchProperty'='$SearchForString' searching ..."
    $PSDrives = Get-PSDrive -PSProvider FileSystem   # | select Name, Root, Description,DisplayRoot
    <#
            $PSDrives | ft
            
            Name Root             Description DisplayRoot
            ---- ----             ----------- -----------
            C    C:\              OS                     
            R    \\Win11-c\share$                        
    #>
    if ($SearchProperty -eq 'DriveLetter' ) { $MyProperty  = 'Name' } else {   $MyProperty  = 'Root' }
    $Exists = $PSDrives | Where { $_.$MyProperty -eq $SearchForString }
    If ( $Exists -eq $Null ) { Logging "$FctName PsDrive where '$SearchProperty'='$SearchForString' does not exist"  } else {  Logging "$FctName PsDrive where '$SearchProperty'='$SearchForString' exists"  }
    If ( $MyDebug  ) { Logging "$FctName END" }
    return $Exists
    }


# =============================================================================
# =============================================================================
Function Logging 
    {
    param ( [String]$Message )

   # $Message = $Message.Replace("$cAppName - ","")
    

    if     ( ($Message.IndexOf("ERROR -")) -ge 0 )         { Write-Host ("$Message")  -ForegroundColor RED }
    elseif ( ($Message.IndexOf("- ERROR -")) -ge 0 )       { Write-Host ("$Message")  -ForegroundColor RED }
    elseif ( ($Message.IndexOf("INFORMATION -")) -ge 0 )   { Write-Host ("$Message")  -ForegroundColor Yellow }
    elseif ( ($Message.IndexOf("- INFORMATION -")) -ge 0 ) { Write-Host ("$Message")  -ForegroundColor Yellow }
    else                                                   { Write-Host ("$Message") }
    $Message | Out-File -FilePath  $Logfile -Append
	}


return Main
