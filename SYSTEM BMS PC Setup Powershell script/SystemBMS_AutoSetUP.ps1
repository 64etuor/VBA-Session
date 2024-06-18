Set-ExecutionPolicy RemoteSigned
# English (United States)를 유일한 언어로 설정
$LangList = New-Object System.Collections.Generic.List[System.Globalization.CultureInfo]
$LangList.Add('en-US')
Set-WinUserLanguageList $LangList -Force
Rename-LocalUser -Name "abc" -NewName "kokam"
$password = ConvertTo-SecureString "!0601@KokamAdmin" -AsPlainText -Force
Set-LocalUser -Name "kokam" -Password $password
# UAC 설정을 "Never Notify"로 변경
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 0
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value 0
Rename-Computer -NewName "Bank1Main-PC" -Force #-Restart
# 원격 지원 허용
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Remote Assistance" -Name "fAllowToGetHelp" -Value 1
# 원격 데스크톱 허용
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
# NLA 요구
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Value 1
# 원격 데스크톱 방화벽 규칙 활성화
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
# 원격 지원 방화벽 규칙 활성화
Enable-NetFirewallRule -DisplayGroup "Remote Assistance"
# 변수 설정
$username = "kokam"
$password = "!0601@KokamAdmin"
# 레지스트리를 통한 자동 로그인 설정
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d $username /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d $password /f
# 암호 만료 비활성화
net accounts /maxpwage:unlimited
# 윈도우 업데이트 서비스 비활성화
Stop-Service -Name wuauserv -Force
Set-Service -Name wuauserv -StartupType Disabled
# 자동 업데이트 설정 변경
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 1
# "High Performance" 전원 계획의 GUID 찾기
$highPerformancePlan = Get-CimInstance -ClassName Win32_PowerPlan -Namespace root\cimv2\power -Filter "ElementName='High performance'"
if ($highPerformancePlan) {
    $guid = $highPerformancePlan.InstanceID -replace 'Microsoft:PowerPlan\\{', '' -replace '}', ''
    powercfg /setactive $guid
    
} else {
    Write-Output "High Performance Power plan does not exist."
}
# IP v4 프로토콜 변경 네트워크 이름 Ethernet, 수정 필요하면 아래 코드 수정
New-NetIPAddress -InterfaceAlias "Ethernet" -IPAddress 192.168.127.200 -PrefixLength 24 -AddressFamily IPv4
# UAC 관리자 모드 활성화
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 1
