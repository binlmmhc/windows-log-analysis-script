#define a analysis result class
Class AnaResult{
    [DateTime]$start_time
    [DateTime]$end_time
    [string]$source_ip = ""
    [Int32]$try_count = 0
    [System.Collections.ArrayList]$username = $Null
}

#read Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx
$rdpevent = Get-WinEvent -Path ".\\Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx"
$allconnections = New-Object System.Collections.ArrayList
if($rdpevent.count -gt 0){
    ForEach($event in $rdpevent){
        if($event.Id -eq 1149){
            $conn = [AnaResult]::new()
            if($event.Message -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"){
                $conn.source_ip = $matches.Values
                $conn.start_time = $event.TimeCreated
                $conn.end_time = $event.TimeCreated
                $conn.try_count = 1

                if($conn.username -eq $Null){
                    $conn.username = New-Object System.Collections.ArrayList
                    $splitinfo = $event.Message.Split(": ")
                    if($splitinfo[5] -match "[A-Za-z0-9]{1,40}"){
                        $username = $matches.Values
                        $conn.username = $username
                    }
                }
            }

            if($allconnections.Length -gt 0){
                $already_exist = $False
                ForEach($eachconn in $allconnections){
                    if($conn.source_ip -eq $eachconn.source_ip){
                        $eachconn.try_count += $conn.try_count

                        if($conn.end_time -gt $eachconn.end_time){
                            $eachconn.end_time = $conn.end_time
                        }

                        if($conn.start_time -lt $eachconn.end_time){
                            $eachconn.start_time = $conn.start_time
                        }

                        if($eachconn.username.count -gt 0){
                            $foundname = $True
                            ForEach($uname in $eachconn.username){
                                if($uname -ne $conn.username){
                                    $foundname = $False
                                }
                                else{ 
                                    $foundname = $True
                                }
                            }
                            if($foundname -eq $False){
                                $eachconn.username.Add($conn.username)
                            }
                        }
                        

                        $already_exist = $True
                        break
                    }
                   
                }
                if($already_exist -eq $False){
                    $allconnections.Add($conn) | Out-Null
                }
            }
            else{ 
                $allconnections.Add($conn) | Out-Null
            }
        }
    }
}
else{ 
    Write-Host "[*]No remote desktop connection event.\n" -Foreground Yellow -Background Black
}

if($allconnections.count -gt 0){
    $allconnections | Out-GridView
}