VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   7005
   ClientTop       =   2985
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command4 
      Caption         =   "previously"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "last"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "next"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "first"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton delelt 
      Caption         =   "delete"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton read 
      Caption         =   "read"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton insert 
      Caption         =   "insert"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Txtname 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Txtcode 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "äÇã"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "˜Ï"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acc As accounts
Private rs As ADODB.Recordset

Private Sub Command1_Click()
 rs.MoveFirst
 refreshform rs!ac_num1, rs!ac_desc
    
End Sub

Public Sub refreshform(strcode As String, strname As String)
 Txtcode.Text = "strcode"
 Txtname.Text = "strname"
End Sub
