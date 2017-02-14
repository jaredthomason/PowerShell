#=============================================================================
# Script:   autg.ps1
#
# Purpose:  Reads a list of usernames from a text file and adds them to the
#           specified group in Active Directory.
#
# Usage:    autg <group_name> <user_file_path> <log_file_path> [ /di]
#
#           By default it validates the group name and user names, but does
#           not actually add them.  To perform the add you must use the
#           /di (DoIt) switch.

#           The use of a log file is mandatory - to document what was done
#           in case of the need to roll-back.
#
# Example:  autg sales_team newusers.txt logfile.txt /di
#
# Note:     This script will only work if the user running the script has the
#           necessary rights to add the users to the group.
#
# Author:   Jim Roberts
#
# Date:     27th September 2008
#
# Versions: 0.9  First version to be tested at work
#           0.91 Corrected spelling mistake in comment
#=============================================================================
$verNum    = "0.91"



#=============================================================================
# Help text
#=============================================================================
Function GiveHelp( )
  {
  $helpText = @"
  Script:   autg.ps1
  Purpose:  Reads a list of usernames from a text file and adds them to the
            specified group in Active Directory.
  Usage:    autg <group_name> <user_file_path> <log_file_path> [ /di]
            By default it validates the group name and user names, but does
            not actually add them.  To perform the add you must use the
            /di (DoIt) switch.
            The use of a log file is mandatory - to document what was done
            in case of the need to roll-back.
  Example: autg sales_team newusers.txt logfile.txt /di
  Version: $verNum
  Author:  Jim Roberts
  Date:    20th September 08
"@
  $HelpText
}



#=============================================================================
# Only used in debugging
#=============================================================================
Function Debug( )
  {
  Write-Host `n
  "------- DEBUG ------------------"
  "Script variables:"
  "  Help    = $Script:Help"
  "  getVer  = $Script:getVer"
  "  DoIt    = $Script:DoIt"
  "  Names:length = $($Script:Names.length)"
  "  Names = "
  for ( $i=0; $i -lt $Script:Names.Length; $i++ )
    {
    "          " + $Script:Names[ $i ]
    }
  "--------------------------------"
  Write-Host `n
}



#=============================================================================
# Powershell function to read and parse command line arguments.
# Sets the script Booleans and adds any other elements into the script
# $Names array
#=============================================================================
function readArguments( )
  {
  If ( $Script:args.length -gt 0 )
    {
    foreach( $Token in $script:Args )
      {
      $Token = $Token.ToUpper()
      If( ($Token -eq "/?") -or ($Token -eq "-?") -or
          ($Token -eq "/H") -or ($Token -eq "-H") -or
          ($Token -eq "/HELP") -or ($Token -eq "-HELP")
        )
        { $Script:Help = $True }
      ElseIf( $Token -eq "/VER" )
        { $Script:getVer = $True }
      ElseIf( ($Token -eq "/DI") -or ($Token -eq "-DI") )
        { $Script:DoIt = $True }
      Else
        {
        $Script:Names += $Token
        }
      } # ForEach
    } # If
  } # Function




#=============================================================================
# Powershell function to get a positive integer from the console.
#=============================================================================
Function GetPosInt
  {
  Param(
    [string] $prompt = $(throw "Param 'prompt' required in GetPosInt.")
    )
  do
    {
    $a = Read-Host $prompt
    } # do
  until( $a -match "^\d{0,5}$" ) # Zero to five digits
  If( $a.length -gt 0 )
    { $b = [int]$a }
  else
    { $b = 0 }
  $b
  }



#=============================================================================
# Powershell function to get a 'Y' or a 'N' from the console.
#=============================================================================
Function GetYN
  {
  Param(
    [string] $prompt = $(throw "Param 'prompt' required in GetYN.")
    )
  do
    {
    $c = (Read-Host $Prompt).ToUpper()
    }
  until( ($c -eq "Y") -or ($c -eq "N") )
  $c
  }



#=============================================================================
# Powershell function to test if a file is open.
# Version 1.1  Powershell's current directory is set at the time you launch
#              Powershell.  You can view it by typing
#              [environment]::currentdirectory   Using CD to switch to another
#              location does not change this.  Your prompt may say
#              PS D:\code\ps> but your current directory will probably still
#              be something like "C:\Documents and Settings\jim".  Try it.
#              This is may be a problem when you use system.io.file as its
#              functions prepend the current directory to the front of
#              relative paths - and this will probably not be what you think
#              it is.  (Powershell's native file functions do not do this.)
#              Hence the workaround in v 1.1
#=============================================================================
function isFileOpen
  {
  Param(
    [string] $Path  = $(throw "Param 'path' required in isFileOpen.")
    )

  trap
    {
    # If trap is called - the openRead failed - so the file is already open
    # and we set a function variable to indicate this.
    # Due to Powershells' scope rules, the only mechanism (other than using
    # a global variable) I could find to have the trap fucntion communicate
    # with its caller was to use set-variable with the -scope paramater.  This
    # enables it access a variable in the calling scope.
    Set-Variable -name alreadyOpen -value $True -scope 1
    continue
    }

  $alreadyOpen = $False

  # Verify file exists
  if ( !(Test-Path $path) )
    {
    Write-Host "'isFileOpen' called with non-existent file."
    exit
    }

  # If $path is relative, then prepend the
  # current location as shown by the Powershell prompt.
  if ( (Split-Path $path) -eq "" )
    {
    $path = Join-Path (Get-Location) $path
    }

  # Try to open the file - to see if an exception is thrown
  $f = [System.IO.File]::OpenRead( $path )

  # If we opened the file as part of our test, then close it again.
  if( !($alreadyOpen) )
    {
    $f.close()
    }
  $alreadyOpen
  }




#=============================================================================
# Powershell script to extract the OU part from a distinguished name.
# Pass it the distinguished name, the common name (which will be removed
# from the front) and the domain object (whose name will be removed from the
# the end).  Also removes attribute type identifiers: OU= etc
#=============================================================================
Function DNtoOU
  {
  param(
    [string] $DN = $(throw "Param 'DN' required in DNtoOU."),
    [string] $CN = $(throw "Param 'CN' required in DNtoOU."),
    [System.DirectoryServices.DirectoryEntry] $Domain = $(throw "Param 'Domain' required in DNtoOU.")
    )
  # Get the distinguished name of the domain
  $domainName = [string]$Domain.distinguishedname

  # Chop it off the end of the distinguished name of the user
  $noDom = $DN.substring( 0, $DN.length - $domainName.length - 1 )

  # (Every comma in the CN is preceded with an \ in the DN
  #  so we need to count them then add the total to the length
  #  to be removed.)
  $regex = New-Object -typename System.Text.RegularExpressions.Regex -argumentlist ","
  $commaCount = $regex.Matches( $CN ).count

  $CNlength = $CN.length + $commaCount + 3 # Plus three for prefix

  # Now chop the CN off the front of the DN
  $OU = $noDom.substring( $CNlength, $noDom.length - $CNlength )

  # By now, if the object is in the root of the domain, $OU will be "",
  # otherwise it will still have a leading comma that needs to be removed.
  if ( $OU -eq "" )
    { "<None>" }
  else
    {
    # If there is a leading comma, remove it
    if ( $OU.substring(0,1) -eq "," )
      {
      $OU = $OU.substring( 1, $OU.length - 1 )
      }

    # Now cut out the OU= and CN=
    $OU = [System.Text.RegularExpressions.Regex]::Replace( $OU, "OU=" ,"")
    $OU = [System.Text.RegularExpressions.Regex]::Replace( $OU, "CN=" ,"")
    $OU
    }
  }


#=============================================================================
# Powershell function to look up all of the groups that match the common name
# entered as a command line parameter.  If there are multiple groups with this
# common name it puts them in a local array of objects and writes them to the
# console.  It then gets the user to choose the desired one.
# It then returns the distinguished name of the chosen group.
# (Beware that this function returns a string so anything written to the console
# within the function must use 'write-host' otherwise they will not be
# 'consumed' so will form part of the return value.)
#=============================================================================
function selectGroup
  {
  param(
    [string] $groupName = $(throw "Param 'groupName' required in selectGroup.")
    )

  $gl = @()

  # Create a new .net DirectorySearcher based on our domain
  $searcher = New-Object System.DirectoryServices.DirectorySearcher( $script:domain )
  # Specify the attributes to be returned
  $searcher.PropertiesToLoad.AddRange( $groupAttributes )
  # Set the filter property of the DirectorySearcher object
  $typeClause = "(objectclass=group)"
  $CNClause = "(cn=$groupName)"
  # Put it all together
  $searcher.filter = "(&$typeClause$CNClause)"
  $searcher.PageSize = 1000

  $groups = $searcher.findall()
  $count = $groups.count
  switch( $count )
    {
    0 {
      Write-Host "Cannot find group with this name."
      ""
      }
    1 {
      $groups[0].properties.distinguishedname
      }
    default
      {
      # Multiple groups returned.  List them so that the user can choose.
      $l1 = $l2 = $l3 = 0
      foreach( $g in $groups )
        {
          $o = New-Object object
        $SAM = [string]$g.properties.samaccountname
        If ( $SAM.length -gt $l2 ) { $l2 = $SAM.length }
        Add-Member -in $o noteproperty SAM $SAM
        $CN = [string]$g.properties.cn
        If ( $CN.length -gt $l1 )  { $l1 = $CN.length }
        Add-Member -in $o noteproperty CN $CN
        $DN = $g.properties.distinguishedname
        $OU = DNtoOU $DN $CN $Script:Domain
        If ( $OU.length -gt $l3 )  { $l3 = $OU.length }
        Add-Member -in $o noteproperty DN $DN
        Add-Member -in $o noteproperty OU $OU
        $Desc = $g.properties.description
        Add-Member -in $o noteproperty desc $Desc
        # Add the object to the groups array
        $gl += $o
        } # foreach
      # Sort them by SAM
      $gl = $gl | Sort-Object -property SAM
      # Write column headers
      $Pad1 = " " * ($l1 + 1 - "CN".length )
      $Pad2 = " " * ($l2 + 1 - "Pre-W2K".length )
      $Pad3 = " " * ($l3 + 1 - "OU".length )
      $header = "`t"+"CN  "+$Pad1+"Pre-W2K  "+$pad2+"OU  "+$pad3+"Desc"
      Write-Host $header
      $header = "`t"+"==  "+$Pad1+"=======  "+$pad2+"==  "+$pad3+"===="
      Write-Host $header

     $i = 1
     foreach( $o in $gl )
        {
        $Pad1 = " " * ($l1 + 1 - $($($o.CN).length))
        $Pad2 = " " * ($l2 + 1 - $($($o.SAM).length))
        $Pad3 = " " * ($l3 + 1 - $($($o.OU).length))
        Write-Host "$i`t$($o.CN)  $Pad1$($o.SAM)  $pad2$($o.OU)  $pad3$($o.desc)"
        $i++
        }

      do
        {
        $Choice = GetPosInt( "Enter number to retrieve group details (Enter to exit)" )
        }
      Until( $Choice -le $Count )
      # write-host " "
      if( $choice -eq 0 )
        {
        ""
        }
      else
        {
        ($gl[$Choice-1]).DN
        }
      } # default
    } # Switch
  } # function



#=============================================================================
# Checks that the specified path to svae the log file is valid.
#=============================================================================
Function TestFileSavePath
  {
  Param(
    [string] $path = $(throw "Param 'path' required in TestFileSavePath.")
    )
    $OK = $false

    # Check path is not empty
    if ( $path -eq "" )
      {$path = Read-Host "Please provide a path for the log file to be saved. (e.g. c:\autgLog.txt)";
	if ( $path -eq "" ) { return $OK }}

    # Check path has a valid form
    if ( Test-Path $path -isValid )
      {
      # Check folder part points to a valid location
      $folder = Split-Path $path
      if ( $folder -eq "" ) { $folder = ".\" }
      if ( Test-Path $folder )
        { $OK = $True }
      else
        { Write-Host "Cannot find target folder for log file." }
      }
   else
      { Write-Host "Log file path is not valid." }

  # If you get here the path is valid, so check if there is already
  # a file with this path name.
  if ( Test-Path $path )
    {
    $yn = getYN "Log file exists.  Overwrite? (Y/N)"
    if ( $yn -eq "N" )
      {
      # Escape the function and return an empty string
      return $False
      }
    }
  # Return result
  $OK
  }



#=============================================================================
# Sets up the directory searcher that is used for the look ups.
#=============================================================================
Function InitialiseSearcher
  {
  # Sets of attributes to retrieve for different search types
  $Attributes  = @( "samaccountname", "distinguishedname")

  # Create a new .net DirectorySearcher based on our domain
  $script:searcher = New-Object System.DirectoryServices.DirectorySearcher( $script:domain )
  # Specify the attributes to be returned
  $script:searcher.PropertiesToLoad.AddRange( $Attributes )
  $script:Searcher.PageSize = 1000
  }



#=============================================================================
# Pass in a username and it will return a user object.
# Returns null if no matching user
#=============================================================================
function mapUsernameToObject
  {
  Param(
    [string] $SAMName = $(throw "Param 'SAMName' required in mapUsernameToObject.")
    )
  # Set the filter property of the DirectorySearcher object
  $typeClause = "(objectClass=user)(objectcategory=person)"
  $SAMClause = "(sAMAccountName=$SAMName)"
  $script:searcher.filter = "(&$typeClause$SAMclause)"
  # Call the findall() method of the DirectorySearcher object then return the result.
  # The return type will be of type SearchResultCollection which has a property
  # called count
  $users = $script:searcher.findall()
  $Count = $users.Count
  switch( $Count )
    {
    0 { # No records returned
      $null
      }
    1 { # Just one user returned
      $users[0]
      }
    default # Multiple users returned
      { # This should not happen!
      "Internal error: Multiple users with same SAM name"
      exit( 1 )
      } # default
    } # switch
  }



#=============================================================================
# Check that there are no blank lines in the input.
# If there are, then exit.
#=============================================================================
function CheckForBlanks
  {
  $i = 1
  $AllOK = $True
  Get-Content $inputFile |
    foreach {
      If ( $_.trim() -eq "" )
        {
        Write-Host "Line $i is blank."
        $AllOK = $False
        } # If
      $i++
      }    # Foreach
  $AllOk
  }


#=============================================================================
# Checks for duplicate lines and user names that do not exist.
# Assumes already checked for blank lines.
#=============================================================================
function validateInputFile
  {
  $i = 1
  $AllOK = $True
  Get-Content $inputFile |
    foreach {
      # Remove white space
      $userSAM = $_.trim()
      # check for duplicates
      If ( $script:unl -contains $userSAM  )
        {
        Write-Host "$userSAM is a duplicate on line $i"
        $AllOK = $False
        }
      else
        {
        $script:unl += $userSAM
        # Check user exists in directory
        $user = mapUsernameToObject $userSAM
        if ( $user -eq $null )
          {
          Write-Host "$_ does not exist on line $i"
          $AllOK = $False
          } # if
        } # else
      $i++
    } # for each
    $AllOK
  }



#=============================================================================
# If the /di switch is present it add users to the group.
# Without the switch it just simulates the addition.
#=============================================================================
function AddToGroup
  {
  $i = 1
  $AllOK = $True
  $group = [ADSI]"LDAP://$script:groupDN"

  trap
    {
    # If trap is called something failed - presumably the adding
    # of a user to the group.
    Write-Host "Unable to add user to group.  Do you have rights?"
    exit
    }

  if ( $script:doIt )
    {     if ( $userInfoTextFile -eq "y" ) {
    Get-Content $inputFile |
      foreach {
        $user = mapUsernameToObject $_
        $userDn = $user.properties.distinguishedname
        if ( $user -eq $null )
          {
          Write-Host "Internal error:  User $_ not found in function 'AddToGroup"
          exit
          }
        if ( $group.member -contains $userDN )
          {
          $msg = ">>> $_ is already a member of $($script:groupCN) <<<"
          Write-Host $msg
          $msg
          }
        else
          {
          $group.Add( "LDAP://$userDN" )
          $msg = "$_`tadded to group $($script:groupCN)"
          Write-Host $msg
          $msg
          }
       } | Set-Content $script:logfile
       $group.psbase.CommitChanges()}
    else {
    $user = mapUsernameToObject $userName
    $userDn = $user.properties.distinguishedname
        if ( $user -eq $null )
          {
          Write-Host "Internal error:  User $_ not found in function 'AddToGroup"
          exit
          }
        if ( $group.member -contains $userDN )
          {
          $msg = ">>> $_ is already a member of $($script:groupCN) <<<"
          Write-Host $msg
          $msg
          }
        else
          {
          $group.Add( "LDAP://$userDN" )
          $msg = "$_`tadded to group $($script:groupCN)"
          Write-Host $msg
          $msg
          } | Set-Content $script:logfile
       $group.psbase.CommitChanges()}
    }
  else
    {	if ( $userInfoTextFile -eq "y" ) {
    Get-Content $inputFile |
      foreach {
        $user = mapUsernameToObject $_
        $userDn = $user.properties.distinguishedname
        if ( $user -eq $null )
          {
          Write-Host "Internal error:  User $_ not found in function 'AddToGroup"
          exit
          }
        if ( $group.member -contains $userDN )
          {
          Write-Host "$_ is already a member of $($script:groupCN)"
          }
        else
          {
          Write-Host "$_`t will be added to group $($script:groupCN)"
          }
       }} # for each
    else {$user = mapUsernameToObject $userName
    $userDn = $user.properties.distinguishedname
        if ( $user -eq $null )
          {
          Write-Host "Internal error:  User $_ not found in function 'AddToGroup"
          exit
          }
        if ( $group.member -contains $userDN )
          {
          Write-Host "$_ is already a member of $($script:groupCN)"
          }
        else
          {
          Write-Host "$_`t will be added to group $($script:groupCN)"
          }}
      Write-Host "Completed in test mode only (no log file written)."
      Write-Host "Use /di switch to actually add members (Requires appropriate rights)."
    }
  }



#================================================================================================
#   M A I N   P R O G R A M
#================================================================================================


# Script scope variables
$Help      = $False   # Has the user asked for help?
$getVer    = $False   # Has the user asked for version info?
$DoIt      = $False   # Actually modify group membership?
$domain    = [adsi]"" # Bind to the root of the domain
$Names     = @()      # Array to hold names entered on command line
$unl       = @()      # List of user names
$userInfoTextFile = "y"

# Constants
# How many groups before we ask for confirmation?
Set-Variable Threshold 5 -option constant


# Sets of attributes to retrieve for different search types
$groupAttributes   = @( "samaccountname", "cn", "distinguishedname", "description" )


readArguments
If( $Help )
  { GiveHelp }
else
  {
  if ( $getVer )
    { "Version $verNum" ; exit }
  if ( $help )
    { GiveHelp ; exit }
  $argCount = $script:Names.length
  if ( ( $argCount -eq 3 ) -or ( ( $argCount -eq 4 ) -and ( $Names[3] -eq "/DI" ) ) )
    {
    $groupCN   = $Names[0]
    $inputFile = $Names[1]
    $logFile   = $Names[2]

    # Check paths are OK
    if ( ! ( Test-Path $inputFile ) )
      {$userInfoTextFile = Read-Host "Do you want to provide the location of a text file with a list of users to add? (y/n) Default is yes.";
	if ($userInfoTextFile -eq "y")
	  {$inputFile = Read-Host "Please provide the location of the text file with a list of users to add.";
	    if ( ! ( Test-Path $inputFile ) )
		{"Cannot find input file." ; exit }}
	else
	  {if ($userInfoTextFile -eq "n")
		{$userName = Read-Host "Please enter the name of the user you wish to add."}}}
    if ( TestFileSavePath $logFile )
      {
      # File already exists - so see if we can write to it.
      if ( Test-Path $logFile )
        {
        if ( isFileOpen( $logFile ) )
          { "Cannot open log file.  Is it already open?" ; exit }
        }
      }
    if ( $userInfoTextFile -eq "y" ) {
    # Check no blank lines in the input file
    if ( ! ( CheckForBlanks ) )
      { "Please remove blank lines then try again." ; exit }}

    # Check group name is OK
    $groupDN   = selectGroup $groupCN
    if ( $groupDN -eq "" )
      { $groupDN = "Please enter the name of the group you want users to be added to.";
	if ( $groupDN -eq "" ) {exit 0 }}

    # Set up a .net Directory Searcher
    InitialiseSearcher

    if ( $userInfoTextFile -eq "y" ) {
    # Check that there are no duplicates in the input file and that all
    # accounts exist
    if ( ! ( validateInputFile ) )
      { "Please correct errors in input file then try again." ; exit }}

    AddToGroup
    }
  else
    { GiveHelp }
  }
