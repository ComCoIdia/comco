VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   7710
   ClientTop       =   3120
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnu_file 
      Caption         =   "file"
      Begin VB.Menu mnu_openfile 
         Caption         =   "open file"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "edit"
   End
   Begin VB.Menu mnu_calculater 
      Caption         =   "calculater"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_calculater_Click()
    Form4.Show
End Sub

Private Sub mnu_openfile_Click()
Form1.Show

End Sub
