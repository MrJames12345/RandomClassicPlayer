
 ' Variables

Set fsObj = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject ("WScript.Shell")
Set audioPlayer = CreateObject("WMPlayer.OCX")
thisFolderPath = fsObj.GetParentFolderName(WScript.ScriptFullName)
Set thisFolder = fsObj.GetFolder(thisFolderPath)



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



Function IsGenreFolder(inFolder)
 ' Only a genre folder if has only mp3 files in it

  check = true
  Set folderFiles = inFolder.Files

  For Each file in folderFiles
    If Not Right(file.Name, 4) = ".mp3" Then
      check = false
    End If
  Next

  IsGenreFolder = check

End Function



Function GetNumGenres()

  i = 0
  Set allFolders = GetGenreFoldersList()
  For Each folder in allFolders
    If IsGenreFolder(folder) Then
      i = i + 1
    End If
  Next
  GetNumGenres= i

End Function



Function GetNthGenrePath(inNum)

  i = 1
  For Each genreFolder in GetGenreFoldersList()
    If IsGenreFolder(genreFolder) Then
      If Int(i) = Int(inNum) Then
        genre = genreFolder
      End If
      i = i + 1
    End If
  Next

  GetNthGenrePath = genre

End Function



Function GetRandomGenreNum()

  GetRandomGenreNum = GetRandomNum(1, GetNumGenres())

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