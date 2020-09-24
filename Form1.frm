VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   675
      Width           =   1140
   End
   Begin Project1.YsVSrcrollBar YsVSrcrollBar1 
      Height          =   2220
      Left            =   495
      TabIndex        =   0
      Top             =   360
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   3916
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Me.YsVSrcrollBar1.Max = 500
Me.YsVSrcrollBar1.LargeChange = 120
End Sub

Private Sub YsVSrcrollBar1_Change()
Me.Text1.Text = Me.YsVSrcrollBar1.value
End Sub

Private Sub YsVSrcrollBar1_Scroll(value As Long)
Me.Text1.Text = value
End Sub
