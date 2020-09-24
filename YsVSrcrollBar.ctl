VERSION 5.00
Begin VB.UserControl YsVSrcrollBar 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   ScaleHeight     =   3615
   ScaleWidth      =   315
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A3B39D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   45
      ScaleHeight     =   3075
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   270
      Width           =   195
      Begin VB.CommandButton Command1 
         BackColor       =   &H00008000&
         Height          =   195
         Left            =   -180
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Image Org2 
      Height          =   225
      Left            =   2205
      Picture         =   "YsVSrcrollBar.ctx":0000
      Top             =   2745
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Org1 
      Height          =   225
      Left            =   2205
      Picture         =   "YsVSrcrollBar.ctx":03C6
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   1935
      Picture         =   "YsVSrcrollBar.ctx":078C
      Top             =   2745
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   1935
      Picture         =   "YsVSrcrollBar.ctx":0B52
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   3195
      Left            =   0
      Picture         =   "YsVSrcrollBar.ctx":0F18
      Stretch         =   -1  'True
      Top             =   180
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   0
      Picture         =   "YsVSrcrollBar.ctx":0FD2
      Top             =   3375
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   0
      Picture         =   "YsVSrcrollBar.ctx":1398
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "YsVSrcrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private OldY As Integer
Private OldX As Integer
Private MoveControl As Boolean
'Default Property Values:
Const m_def_LargeChange = 10
Const m_def_Max = 100
Const m_def_Min = 1
Const m_def_Value = 0
Const m_def_Enabled = True
'Property Variables:
Dim m_LargeChange As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_Value As Long
Dim m_Enabled As Boolean


Public Event Change()
Public Event Scroll(value As Long)
Function GetCrnt() As Long

Dim z As Long
Dim crnt As Long
Dim Final As Long

z = Picture1.Height - Command1.Height
crnt = Command1.Top

Dim CrntRatio As Single
CrntRatio = crnt / z * 100
Final = CLng(Me.Max * CrntRatio / 100)

If Final < 1 Then Final = 1
If Final > Me.Max Then Final = Me.Max
GetCrnt = Final

End Function


Sub DrawCrnt()

Dim z As Long
Dim crnt As Long
Dim Final As Long

'z = Picture1.Height - Command1.Height
z = Me.Max

'crnt = Command1.Top
crnt = Me.value

Dim CrntRatio As Single
CrntRatio = crnt / z * 100
'Final = CLng(Me.Max * CrntRatio / 100)
Final = CLng((Picture1.Height - Command1.Height) * CrntRatio / 100)

If Final < 0 Then Final = 0
If Final > (Picture1.Height - Command1.Height) Then Final = (Picture1.Height - Command1.Height)
Command1.Top = Final

End Sub

Private Sub MoveDown(Optional Large As Boolean = False)
Dim mVale As Long
If Large Then
   mValue = Me.LargeChange
Else
   mValue = Me.Min
End If

Set Image2.Picture = Image5.Picture

If Me.value < Me.Max Then
   If Me.value + mValue <= Me.Max Then
      Me.value = Me.value + mValue
   Else
      Me.value = Me.value + 1
   End If
   Me.DrawCrnt
End If
End Sub

Private Sub MoveUp(Optional Large As Boolean = False)

Dim mVale As Long
If Large Then
   mValue = Me.LargeChange
Else
   mValue = Me.Min
End If

Set Image1.Picture = Image4.Picture

If Me.value > 1 Then
   If Me.value - mValue >= 1 Then
      Me.value = Me.value - mValue
   Else
      Me.value = Me.value - 1
   End If
   Me.DrawCrnt
End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 OldY = Y
 OldX = X
 MoveControl = True
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewY As Long
If MoveControl = True Then
    NewY = Command1.Top - OldY + Y
    
    If NewY < 0 Then NewY = 0
    If NewY > Picture1.Height - Command1.Height Then NewY = Picture1.Height - Command1.Height
    Command1.Top = NewY
    'Me.Value = Me.GetCrnt
    Dim mValue As Long
    mValue = Me.GetCrnt
    RaiseEvent Scroll(mValue)
End If
End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveControl = False
Me.value = GetCrnt
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Me.Enabled Then Exit Sub
Set Image1.Picture = Image4.Picture
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Me.Enabled Then Exit Sub
MoveUp
Set Image1.Picture = Org1.Picture
End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Me.Enabled Then Exit Sub
Set Image2.Picture = Image5.Picture
End Sub


Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Me.Enabled Then Exit Sub
MoveDown
Set Image2.Picture = Org2.Picture

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Me.Enabled Then Exit Sub

If Y < Command1.Top Then
   MoveUp True
Else
   MoveDown True
End If

End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Image1.Picture = Org1.Picture
Set Image2.Picture = Org2.Picture
End Sub


Private Sub UserControl_Resize()
Image1.Top = 0
Image1.Left = 0
Image2.Left = 0
Image2.Top = UserControl.Height - Image2.Height
Image3.Left = 0
Image3.Top = Image1.Height
Image3.Height = UserControl.Height - Image1.Height - Image2.Height
Picture1.Top = Image3.Top
Picture1.Height = Image3.Height
UserControl.Width = Image1.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
    Me.value = 1
    Me.DrawCrnt
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get value() As Long
    value = m_Value
End Property

Public Property Let value(ByVal New_Value As Long)
    If New_Value > Me.Max Or New_Value < 1 Then Exit Property
    m_Value = New_Value
    PropertyChanged "Value"
    RaiseEvent Change
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    Command1.Visible = Me.Enabled
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Value = m_def_Value
    m_Enabled = m_def_Enabled
    m_LargeChange = m_def_LargeChange
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
End Sub

Private Sub UserControl_Show()
Command1.Visible = Me.Enabled
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,10
Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

