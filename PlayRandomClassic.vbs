' ' Variables
Set fsObj = CreateObject("Scripting.FileSystemObject")
Set audioPlayer = CreateObject("WMPlayer.OCX")
thisFolderPath = fsObj.GetParentFolderName(WScript.ScriptFullName)
Set thisFolder = fsObj.GetFolder(thisFolderPath)

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




' Functions


Function GetGenreFromFile()
  Set genreFile = fsObj.OpenTextFile(thisFolder & "\genre.txt", 1)
  genreNum = genreFile.ReadAll()
  genreFile.Close
  Set genreFile = Nothing
  GetGenreFromFile = genreNum

End Function



Function GetGenreFoldersList()

  Set GetGenreFoldersList = thisFolder.SubFolders

End Function



Function GetNthGenrePath(inNum)

  i = 1
  For Each genreFolder in GetGenreFoldersList()
    If Int(i) = Int(inNum) Then
      genre = genreFolder
    End If
    i = i + 1
  Next

  GetNthGenrePath = genre

End Function



Function GetRandomGenreNum()

  GetRandomGenreNum = GetRandomNum(1, GetGenreFoldersList().Count)

End Function



Function GetRandomSongFromFolder(inFolder)

  Set songsFolder = fsObj.GetFolder(inFolder)
  Set songFiles = songsFolder.Files
  totalSongs = songFiles.Count

  randSongNum = GetRandomNum(1, totalSongs)

  i = 1
  For Each song in songFiles
      If Int(i) = Int(randSongNum) Then
        songPath = song
      End If
      i = i + 1
  Next

  GetRandomSongFromFolder = songPath

End Function



Function GetRandomNum(inFrom, inTo)

  Randomize
  GetRandomNum = Int((inTo - inFrom + 1) * Rnd + inFrom)

End Function