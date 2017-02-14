# This script takes an optional parameter of a path.  If excluded, the path will
# be taken from the location where the PowerShell script is called.

param([String[]] $path=(get-location).Path)

# The following information determines the user/group that will be used and the
# permissions that will be set accordingly at every location that matches the
# criteria of having inheritance blocked.
# $Identity               – Identity of the user or group being granted permissions
# $Rights                   – The rights to be assigned
# $Inheritance     – Defines if subsequent objects and containers will receive the
#                                         the new permissions.
# $Propagation     – Defines how permissions are propagated to all child objects
# $Type                      – Defines whether access is Allowed or Denied

$Identity = "DOMAIN\USER"
$Rights = "Read"
$Inheritance = @("ObjectInherit"�, "ContainerInherit")
$Propagation = "None"
$Type = "Allow"
$Rule = New-Object System.Security.AccessControl.FileSystemAccessRule( `
$Identity, $Rights, $Inheritance, $Propagation, $Type)

# Get the security descriptor for the $path or location where the script was
# run from and add the previous rule to the location.

$TopACL = Get-ACL $path
$TopACL.AddAccessRule($Rule)
Set-ACL $path -AclObject $TopACL

# Recurse all folders and sub-folders contained in the original path or location
# where the script was called

get-childitem $path -recurse -force | where-object { $_.PsIsContainer } |
foreach-object {

# Get the security descriptor for the child object

$ACL = Get-ACL $_.FullName

# Get the access rules for the child object

$access = (get-acl $_.FullName).Access

# Select the IsInherited field and set $inherit to the number of times it
# occurs, should be 1 or 0, or TRUE/FALSE respectively.

$inherit = $access.Count -eq ($access | where-object { $_.IsInherited }).Count

# If inheritance is blocked, $inherit is NOT true, add the accessrule to the
# child object and apply it.

if( !$inherit )
 {
  $ACL.AddAccessRule($Rule)
  Set-ACL $_.FullName -AclObject $ACL
  }
}
