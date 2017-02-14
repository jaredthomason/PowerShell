$i = ("aos200s6")
ping ($i)


function Ping (  [string] $strComputer )
{$timeout=120;  trap { continue;}
$ping = new-object System.Net.NetworkInformation.Ping 
$reply = new-object System.Net.NetworkInformation.PingReply
$reply = $ping.Send($strComputer, $timeout);
if( $reply.Status -eq "Success"  ) 
{return $true;} 
return $false;
}
