Set WshShell = WScript.CreateObject("WScript.Shell")
Comandline = ":\Users\myst\AppData\Roaming\Spotify\Spotify.exe"
WScript.sleep 500

' Open Top 50 - Singapore Playlist
CreateObject("WScript.Shell").Run("spotify:user:mysterioan:playlist:37i9dQZEVXbK4gjvS1FjPY")
WScript.sleep 1000

' Turn down the volume to 0
for i=1 to 20
	WshShell.SendKeys("^{DOWN}")
next

' Turn up the volume to 20
for i=1 to 5
	WshShell.SendKeys("^{UP}")
next
WScript.sleep 100

' Next Track
WshShell.SendKeys "^{RIGHT}"
' WScript.sleep 100

' Play
' WshShell.SendKeys " "
