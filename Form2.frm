VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exitform2 
      Caption         =   "����"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton sabtcommand 
      Caption         =   "���"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtcode 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "�����"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�����"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bln As Boolean
Private strcode As String
Private stroldcode As String
Public acc As accounts

Private Sub exitform2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 bln = True
 
End Sub

Private Sub sabtcommand_Click()
  Set acc = New accounts
  
  
    If Trim(txtcode.Text) = "" Then
     MsgBox "���� ����� �� ���"
    Exit Sub
    
    If Len(txtcode.Text) > 5 Then
     MsgBox "���ǘ����� ����� ���� ���� ����� ��� �� �� ���� �� ����"
     Exit Sub
     
     End If
     
    End If
    If Trim(txtname.Text) = "" Then
     MsgBox "���� ����� �� ���"
     Exit Sub
    If Len(txtname.Text) > 30 Then
     MsgBox "���ǘ����� ����� ���� ���� ����� ��� �� �� ���� �� ����"
     Exit Sub
    End If
    
    End If
    
    If bln Or stroldcode <> txtcode.Text Then
    
       If acc.readtextbox(txtcode.Text, "") Then
            txtcode.Text = stroldcode
            MsgBox " ������ ��� ����� �� ���� ����� ʘ���� �� ���� "
            Exit Sub
        End If
    End If
         
    If bln Then
    
        acc.getInsert txtcode.Text, txtname.Text
        
     Else
        acc.getUpdate txtcode.Text, txtname.Text, stroldcode
        
            
    End If
    
    Unload Me
    
        
End Sub


Public Sub loadForm(strcode As String, strname As String)
    Load Me
    txtcode.Text = strcode
    txtname.Text = strname
    bln = False
    stroldcode = strcode
    Me.Show vbModal
    
    
End Sub
