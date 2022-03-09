' // - Include 'Functions' file - //
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile("Utils.vbs", 1)
    ExecuteGlobal objTextFile.ReadAll
    objTextFile.Close
    Set objFSO = Nothing
    Set objTextFile = Nothing
' // ----------------------------- //




 ' Get saved genre
selectedGenreNum = GetGenreFromFile()

 ' IF genre is 0, means 'All', so get random genre num
If selectedGenreNum = 0 Then
  selectedGenreNum = GetRandomGenreNum()
End If

 ' Get genre's folder path
genrePath = GetNthGenrePath(selectedGenreNum)

 ' Get random song
randomSongPath = GetRandomSongFromFolder(genrePath)

' Play file
audioPlayer.URL = randomSongPath
audioPlayer.controls.play 
While audioPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

' Close audio player
audioPlayer.close