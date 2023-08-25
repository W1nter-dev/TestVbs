Dim url, outputPath

' URL del File
url = "https://raw.githubusercontent.com/W1nter-dev/TestVbs/main/audio.mp3"

' Path del File
outputPath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\audio.mp3"

Set http = CreateObject("Microsoft.XMLHTTP")
http.Open "GET", url, False
http.Send

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1 ' Binary data
stream.Open
stream.Write http.ResponseBody
stream.SaveToFile outputPath, 2 ' Overwrite the file if it exists
stream.Close

Set http = Nothing

Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

Do While True

	' Randomize
	rand = Int(((900-300+1)*Rnd+300)*1000)
	' riprodure audio
	oPlayer.URL = outputPath
	oPlayer.controls.play 
	While oPlayer.playState <> 1 ' 1 = Stopped
	WScript.Sleep 100
	Wend

	' Rilascia file
	oPlayer.close

	WScript.Sleep rand
Loop
