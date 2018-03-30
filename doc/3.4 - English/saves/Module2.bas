Attribute VB_Name = "Module2"
'Download by http://www.NewXing.com
Public Function FileExists(ByVal sFileName As String) As Integer
'判断所请求的文件是否存在
Dim i As Integer
On Error Resume Next
    i = Len(Dir$(sFileName))
    If Err Or i = 0 Then
        FileExists = False
        Else
            FileExists = True
    End If
End Function
Public Function ReplaceStr(ByVal strMain As String, strFind As String, strReplace As String) As String
    '替代函数
    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    ReplaceStr$ = strNew$
End Function
Public Function text_read(filename)
'读取文件函数
Dim f
Dim textda
Dim cha

On Error Resume Next
f = FreeFile
textda = ""
If FileExists(filename) Then
    If Len(filename) Then
        Open filename For Input As #f
        Do While Not EOF(f)
            cha = Input(1, #f)
             textda = "" & textda & cha
        Loop
        Close #f
    End If
text_read = textda
Else
text_read = ""
End If

End Function
Public Sub timeout(ByVal nSecond As Single)
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer

      dummy = DoEvents()
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub
Public Function html_404error()
'这个函数的作用是，当用户请求的页面不存在时，我们可以发送这个页面过去
Dim x As String
x = ""
x = x & "<html>" & vbCrLf
x = x & "" & vbCrLf
x = x & "<head>" & vbCrLf
x = x & "<style>" & vbCrLf
x = x & "a:link          {font:8pt/11pt verdana; color:red; text-decoration:none}" & vbCrLf
x = x & "a:visited       {font:8pt/11pt verdana; color:red}" & vbCrLf
x = x & "a:hover          {font:8pt/11pt verdana; color:red; text-decoration:underline}" & vbCrLf
x = x & "</style>" & vbCrLf
x = x & "<meta HTTP-EQUIV=""Content-Type"" Content=""text-html; charset=Windows-1252"">" & vbCrLf
x = x & "<title>HTTP 404 Not Found</title>" & vbCrLf
x = x & "</head>" & vbCrLf
x = x & "" & vbCrLf
x = x & "<body bgcolor=""#FFFFFF"">" & vbCrLf
x = x & "<p><font  size=""2""><b><font color=""#FF0000"">The" & vbCrLf
x = x & "  找不到请求页面 </font></b></font></p>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
x = x & "<p><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"">请求页面" & vbCrLf
x = x & "  可能不存在或已经被转移或改名。" & vbCrLf
x = x & "  unavailable. </font></p>" & vbCrLf
x = x & "<p align=""center"">&nbsp;</p>" & vbCrLf
x = x & "<p align=""center""><font size=""1""  color=""#0000FF""><i><font color=""#000000"">HTTP" & vbCrLf
x = x & "  404 - 文件没有找到</font></i></font></p>" & vbCrLf
x = x & "</body>" & vbCrLf
x = x & "</html>" & vbCrLf & vbCrLf & vbCrLf
html_404error = x
End Function


