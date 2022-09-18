Option Explicit
' https://github.com/oicu
' 2021/12/21, 2022/05/16
' Backup Windows Key, Windows Product Key Finder, Windows OEM Find Finder

Dim objshell, path, DigitalID, Result
Dim ProductName, ProductID, ProductKey, ProductData, OEMKey, OEMKeyDesc
OEMKey = "OEM Key: "
OEMKeyDesc = "Description: "
Set objshell = CreateObject("WScript.Shell")
Path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
DigitalID = objshell.RegRead(Path & "DigitalProductId")
' DigitalID = objshell.RegRead(Path & "DigitalProductId4")
ProductName = "Product Name: " & objshell.RegRead(Path & "ProductName")
ProductID = "Product ID: " & objshell.RegRead(Path & "ProductID")
ProductKey = "Installed Key: " & ConvertToKey(DigitalID)
GetOEMKey()
ProductData = ProductName  & vbNewLine & ProductID  & vbNewLine & ProductKey & vbNewLine & vbNewLine & OEMKey & vbNewLine & OEMKeyDesc
'If vbYes = MsgBox(ProductData  & vblf & vblf & "Save to a file?", vbYesNo + vbQuestion, "BackUp Windows Key Information") Then
If vbYes = MsgBox("查看 Win7、Win8、Win10 使用的序列号、主板 OEM 序列号。" & vblf & vblf & ProductData & vblf & vblf & "是否保存信息到桌面文件 WindowsKeyInfo.txt ？", vbYesNo + vbQuestion, "备份 Windows 序列号信息 - C老师") Then
   Save ProductData
End If

' Credits: nononsence @ MDL
' Convert binary to chars
Function ConvertToKey(Key)
    Const KeyOffset = 52
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert
    'Check if OS is Windows 8
    isWin8 = (Key(66) \ 6) And 1
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
    i = 24
    Maps = "BCDFGHJKMPQRTVWXY2346789"
    Do
        Current= 0
        j = 14
        Do
           Current = Current* 256
           Current = Key(j + KeyOffset) + Current
           Key(j + KeyOffset) = (Current \ 24)
           Current=Current Mod 24
            j = j -1
        Loop While j >= 0
        i = i -1
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput
        Last = Current
    Loop While i >= 0

    If (isWin8 = 1) Then
        keypart1 = Mid(KeyOutput, 2, Last)
        insert = "N"
        KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
        If Last = 0 Then KeyOutput = insert & KeyOutput
    End If

    ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)

End Function

'Save data to a file
Function Save(Data)
    Dim fso, fName, txt, objshell, UserProfile
    Set objshell = CreateObject("wscript.shell")
    'Get current user's home
    UserProfile = objshell.ExpandEnvironmentStrings("%USERPROFILE%")
    'Create a text file on desktop
    fName = UserProfile & "\Desktop\WindowsKeyInfo.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(fName)
    txt.Writeline Data
    txt.Close
End Function

' Credits: oicu @ github
Sub GetOEMKey()
    Dim strComputer
    Dim strWMINamespace
    Dim strWMIquery
    Dim objWMIService
    Dim colItems
    Dim objItem
    Dim boolPropertyExists
    boolPropertyExists = False
    strComputer = "."
    strWMINamespace = "\root\cimv2"
    strWMIquery = ":SoftwareLicensingService"
    Set objWMIService = GetObject ( "winmgmts:\\" & strComputer & strWMINamespace & strWMIQuery)
    For Each objItem In objWMIService.Properties_
        If objItem.Name = "OA3xOriginalProductKey" Then
            boolPropertyExists = True
        End If
    Next
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & strWMINamespace)
    Set colItems = objWMIService.ExecQuery("Select * from SoftwareLicensingService")
    For Each objItem In colItems
        If boolPropertyExists Then
            If objItem.OA3xOriginalProductKey <> "" Then
                OEMKey = OEMKey & objItem.OA3xOriginalProductKey
                OEMKeyDesc = OEMKeyDesc & objItem.OA3xOriginalProductKeyDescription
            End If
        End If
    Next
End Sub
