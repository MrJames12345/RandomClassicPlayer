
' // - Include 'Functions' file - //
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    thisFolderPath = fsObj.GetParentFolderName(WScript.ScriptFullName)
    Set thisFolder = fsObj.GetFolder(thisFolderPath)
    Set utilsFile = fsObj.OpenTextFile(thisFolder & "\Utils.vbs", 1)
    ExecuteGlobal utilsFile.ReadAll
    utilsFile.Close
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