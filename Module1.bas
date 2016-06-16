Attribute VB_Name = "Module1"
Public connectinstring  As String

Public Sub main()
  connectinstring = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=KHACC_0015;Data Source=PIDATACENTER\SQL2014"
  MDIForm1.Show
End Sub
