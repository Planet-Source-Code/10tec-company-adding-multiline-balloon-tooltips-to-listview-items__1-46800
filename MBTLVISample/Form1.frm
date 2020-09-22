VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column 3"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Dim TT As CTooltip
Dim m_lCurItemIndex As Long

Private Sub Form_Load()
   With ListView1.ListItems
      .Add Text:="Test item #1"
      .Add Text:="Test item #2"
      .Add Text:="Long long long test item #3"
   End With

   Set TT = New CTooltip
   TT.Style = TTBalloon
   TT.Icon = TTIconInfo
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   
   lvhti.pt.x = x / Screen.TwipsPerPixelX
   lvhti.pt.y = y / Screen.TwipsPerPixelY
   lItemIndex = SendMessage(ListView1.hwnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
         TT.Title = "Multiline tooltip"
         TT.TipText = ListView1.ListItems(m_lCurItemIndex).Text
         TT.Create ListView1.hwnd
      End If
   End If
End Sub
