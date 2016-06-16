VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Private rs As ADODB.Recordset

Public Function GetAccount() As ADODB.Recordset
 cnn.Open
 'Set rs = New ADODB.Recordset
 Set rs = cnn.Execute("select * from Account1 order by ac_num1")
 Set rs.ActiveConnection = Nothing
 'rs.CursorLocation = adUseClient
 Set GetAccount = rs
 
 cnn.Close
 
End Function



Private Sub Class_Initialize()
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.ConnectionString = connectinstring
    
    
End Sub


Public Function getDelete(strcode As String)
    cnn.Open
    cnn.Execute "delete from Account1 where ac_num1='" & strcode & "'"
    cnn.Close
    
End Function

Public Function getInsert(strcode As String, strname As String)
    cnn.Open
    cnn.Execute "insert into Account1 (ac_num1,ac_desc,ac_naghd,ac_bed,ac_bes,ac_moj,ac_tedad) values ('" & strcode & "' , '" & strname & "',0,0,0,0,0)"
    cnn.Close
End Function

Public Function getUpdate(strcode As String, strname As String, stroldcode As String)
    cnn.Open
    cnn.Execute "update account1 set ac_num1='" & strcode & "' , ac_desc='" & strname & "' where ac_num1='" & stroldcode & "'"
    cnn.Close
    
End Function

Public Function readtextbox(strcode As String, strname As String)
    cnn.Open
    
   Set rs = cnn.Execute("select ac_num1 ,ac_desc from account1 where ac_num1='" & strcode & "'")
   Set rs.ActiveConnection = Nothing
     If Not rs.EOF Then
      readtextbox = True
      Else
       readtextbox = False
     End If
     Set rs = Nothing
     cnn.Close
    
End Function




