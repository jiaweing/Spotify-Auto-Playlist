WScript.sleep 5000
Set WshShell = WScript.CreateObject("WScript.Shell")
Comandline = ":\Users\myst\AppData\Roaming\Spotify\Spotify.exe"
WScript.sleep 500

' Get Current Time
Dim dt
dt = now

' Get Current Hour
Dim currentHour
currentHour = hour(dt)

' Open Playlist
' Play a morning playlist if it is before 10 AM
If currentHour <= 10 Then
	CreateObject("WScript.Shell").Run("spotify:user:mysterioan:playlist:37i9dQZF1DX2MyUCsl25eb")
Else
	CreateObject("WScript.Shell").Run("spotify:user:mysterioan:playlist:37i9dQZEVXbK4gjvS1FjPY")
End If
WScript.sleep 1000

' Mute Volume
' It takes 16 x CTRL + DOWN to get from 100 to 0
for i=1 to 16
	WshShell.SendKeys("^{DOWN}")
next
WScript.sleep 500

' Turn Up the Volume
If currentHour <= 10 OR currentHour >= 19 Then
	' 25 in the morning and night
	for i=1 to 4
		WshShell.SendKeys("^{UP}")
	next
Else
	' 50 in the afternoon
	for i=1 to 8
		WshShell.SendKeys("^{UP}")
	next
End If
WScript.sleep 1000

' Find the Play button in the Playlist
' Some playlists have an additional "Created by Spotify" link which requires an additional TAB
WshShell.SendKeys "{TAB}"
WScript.sleep 100
WshShell.SendKeys "{ENTER}"
WScript.sleep 100
' In this case, my morning playlist has the additional "Created by Spotify" link so another TAB is required to get to the Play button
If currentHour <= 10 Then
	WshShell.SendKeys "{TAB}"
	WScript.sleep 100
End If
WshShell.SendKeys "{TAB}"
WScript.sleep 100

' Play Current Playlist
WshShell.SendKeys "{ENTER}"
' WScript.sleep 100

' Next Track
' WshShell.SendKeys "^{RIGHT}"

' WScript.sleep 100

' Play
' WshShell.SendKeys " "
