
' // - Include 'Functions' file - //
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    thisFolderPath = fsObj.GetParentFolderName(WScript.ScriptFullName)
    Set thisFolder = fsObj.GetFolder(thisFolderPath)
    Set utilsFile = fsObj.OpenTextFile(thisFolder & "\Utils.vbs", 1)
    ExecuteGlobal utilsFile.ReadAll
    utilsFile.Close
' // ----------------------------- //




output = ""

 ' IF no argument, show help
If WScript.Arguments.Count = 0 Then
    
    output = HelpMessage()

 ' Else set genre
Else

     ' Get genre num from argument
    genreNum = WScript.Arguments.Item(0)

    ' If new genre num > num folders, reset to 0
    If Int(genreNum) > Int(GetNumGenres()) Then
        genreNum = 0
        output = output & "The genre number you entered is not an option." & vbCrLf
    End If

    ' Set new genre num
    Set genreFile = fsObj.OpenTextFile(thisFolder & "\genre.txt", 2)
    genreFile.Write(genreNum)
    genreFile.Close
    Set genreFile = Nothing

     ' Get genre name
    If genreNum = 0 Then
        genreName = "Random Genre"
    Else
        i = 1
        For Each genreFolder in GetGenreFoldersList()
            If IsGenreFolder(genreFolder) Then
                If Int(i) = Int(genreNum) Then
                    genreName = genreFolder.Name
                End If
                i = i + 1
            End If
        Next
    End If

    output = output & "Your genre is set to " & Int(genreNum) & " (" & genreName & ")."

End If


' Output
Wscript.Echo output








' Functions

Function HelpMessage()

    indent = "    "
    outString = vbCrLf & "Set the genre using one of these values:" & vbCrLf & indent & "0 - Random Genre"

    i = 1
    For Each folder in GetGenreFoldersList()
        If IsGenreFolder(folder) Then
            outString = outString & vbCrLf & indent & i & " - " & folder.Name
            i = i + 1
        End If
    Next

    HelpMessage = outString

End Function