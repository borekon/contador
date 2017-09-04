Attribute VB_Name = "LeeINI"
Option Explicit
'*TO USE: Copy&Paste, the Key is the setting, and Value obviously is the value.
'*EXAMPLE: Var = ReadINI(filename, key)  -- & --  WriteINI filename, key, value
'*WHATSNEW: Can write complete ini, can write header tags, BUGS FIXED!
'*NOTE: To write header tags just leave the [value] out of WriteINI and the [Key] is the heading (wrote to file in order used)

Public Function ReadINI(File As String, Key As String) As String
 Dim fnum As Integer, Data As String
 If (Dir(File) = "") Then Exit Function
 fnum = FreeFile
 Open File For Input As fnum
  Do While Not EOF(fnum)
    Line Input #fnum, Data
    If (Left(LCase(Data), Len(Key)) = LCase(Key)) Then
      ReadINI = Trim(Right(Data, Len(Data) - Len(Key) - 1))
      Exit Do
    End If
  Loop
 Close fnum
End Function

Public Sub WriteINI(File As String, Key As String, Optional Value As String = "@~")
 Dim fn1 As Integer, fn2 As Integer, Data As String, File2 As String
 Dim Found As Boolean, NewFile As Boolean
 fn1 = FreeFile
 If (Dir(File) = "") Then Open File For Output As #fn1: Close fn1: NewFile = True
 fn2 = fn1 + 1
 File2 = Left(File, Len(File) - 4) & ".tmp"
 Open File For Input As fn1
 Open File2 For Output As fn2
  Do While Not EOF(fn1)
    Line Input #fn1, Data
    If (Value = "@~") Then
      If (Left(LCase(Data), Len(Key) + 2) = "[" & LCase(Key) & "]") Then
        Print #fn2, Data: Found = True
      Else: Print #fn2, Data
      End If
    Else
      If (Left(LCase(Data), Len(Key) + 1) = LCase(Key) & "=") Then
        Print #fn2, Key & "=" & Value: Found = True
      Else: Print #fn2, Data
      End If
    End If
  Loop
  If Not Found Then
    If (Value = "@~") Then
      If Not NewFile Then Print #fn2,
      Print #fn2, "[" & Key & "]"
    Else
      Print #fn2, Key & "=" & Value
    End If
  End If
 Close fn1, fn2
 FileCopy File2, File
 Kill File2
End Sub



