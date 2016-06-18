Attribute VB_Name = "Module1"
Public connectinstring  As String

Public Sub main()
  connectinstring = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=KHACC_0015;Data Source=PIDATACENTER\SQL2014"
  MDIForm1.Show
End Sub
Public Function GetAllNum(ByVal TX As String, Optional xChar As String) As Double
    Dim X, Y
    
    On Error Resume Next
    
    X = Left("" & TX, 35)
    'If InStr(X, ".") > 0 Then MsgBox "SDFSDFWERWERWRWER"
    Y = ""
    If xChar = "" Then
        While Len(X) > 0
            
            If (Left(X, 1) >= "0" And Left(X, 1) <= "9") Or Left(X, 1) = DecimalPointChar Then ' "." Then
                Y = Y + Left(X, 1)
            End If
            X = Mid(X, 2)
         Wend
    Else
        Y = Replace(X, xChar, "")
    End If
    If Y = "" Then Y = "0"
    If Y = DecimalPointChar Then Y = "0"
    If Len(TX) = 8 And Mid(TX, 3, 1) = "/" And Mid(TX, 6, 1) = "/" Then
        Y = Val(Replace(Y, "/", ""))
    Else
        Y = Val(Replace(Y, DecimalPointChar, "."))
'        If Right(TX, 1) = "0" And InStr(TX, DecimalPointChar) > 0 Then
'            Y = Y & ".0"
'        End If
    End If
    GetAllNum = Y * IIf(InStr(TX, "-") > 0, -1, 1)
End Function

