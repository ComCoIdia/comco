VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3885
   ClientLeft      =   8355
   ClientTop       =   2370
   ClientWidth     =   5130
   LinkTopic       =   "Form4"
   ScaleHeight     =   3885
   ScaleWidth      =   5130
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
    Text3.Text = GetAllNum(Text1.Text) + GetAllNum(Text2.Text)
End Sub

Private Sub Text2_Change()
    Text3.Text = GetAllNum(Text1.Text) + GetAllNum(Text2.Text)
End Sub
