VERSION 5.00
Begin VB.UserControl BottomBar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   KeyPreview      =   -1  'True
   ScaleHeight     =   1095
   ScaleWidth      =   6165
   ToolboxBitmap   =   "BottomBar.ctx":0000
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5400
      Top             =   0
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdLastPage 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.Image imgPage 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "BottomBar.ctx":0312
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblPageCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Item"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   600
         Width           =   870
      End
   End
End
Attribute VB_Name = "BottomBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CYRIL'S BOTTOM BAR CONTROL.
'Finished, Friday November 17 2000, 10:45 AM
'Started, Tuesday November 13 2000, 1:50 PM
Option Explicit
Private m_SelectedItem As Long
Private DoRightScroll As Boolean
Private DoLeftScroll As Boolean
Private m_BackColor As OLE_COLOR
Private m_ItemCount As Integer
Event ItemClick(ItemIndex As Long)
Event ItemDblClick(ItemIndex As Long)

Public Sub AddItem(Capt As String)
Dim newItem As Long

newItem = imgPage.Count
Load imgPage(newItem)
Load lblPageCaption(newItem)

Deselect newItem

lblPageCaption(newItem).Caption = Capt
imgPage(newItem).Visible = True
lblPageCaption(newItem).Visible = True
AlignControls
End Sub

Public Sub RemoveItem(inde As Long)
Dim i As Long

If inde = 0 Then
    imgPage(0).Visible = False
    lblPageCaption(0).Visible = False
    Exit Sub
End If

Unload imgPage(inde)
Unload lblPageCaption(inde)
End Sub

Public Sub Clear()
Dim i As Long

For i = 1 To imgPage.Count - 1
    Unload imgPage(i)
    Unload lblPageCaption(i)
Next i

ItemCount = 0
m_SelectedItem = 0
End Sub

Private Sub cmdLastPage_Click()
If EnableLeftScroll Then picBack.Left = picBack.Left + 200
End Sub

Private Sub cmdLastPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
timScroll.Enabled = True
DoLeftScroll = True
End Sub

Private Sub cmdLastPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
timScroll.Enabled = False
DoLeftScroll = False
End Sub

Private Sub cmdNextPage_Click()
If EnableRightScroll Then picBack.Left = picBack.Left - 200
End Sub

Private Sub cmdNextPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
timScroll.Enabled = True
DoRightScroll = True
End Sub

Private Sub cmdNextPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
timScroll.Enabled = False
DoRightScroll = False
End Sub

Private Sub imgPage_Click(Index As Integer)
DoClick Index
End Sub

Private Sub imgPage_DblClick(Index As Integer)
RaiseEvent ItemDblClick(CLng(Index))
End Sub

Private Sub lblPageCaption_Click(Index As Integer)
DoClick Index
End Sub

Private Sub lblPageCaption_DblClick(Index As Integer)
RaiseEvent ItemDblClick(CLng(Index))
End Sub

Private Sub timScroll_Timer()

If DoRightScroll Then
  If EnableRightScroll Then picBack.Left = picBack.Left - 200
ElseIf DoLeftScroll Then
    If EnableLeftScroll Then picBack.Left = picBack.Left + 200
End If
End Sub

Private Sub UserControl_Initialize()
SelectedItem = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long
BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
ItemCount = PropBag.ReadProperty("ItemCount", 1)

For i = 0 To imgPage.Count - 1
    lblPageCaption(i) = PropBag.ReadProperty("ItemCapt" & i, "The Item")
Next i
    picBack.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    UserControl.BorderStyle = PropBag.ReadProperty("Style", 1)
End Sub

Private Sub UserControl_Resize()
AlignControls
End Sub

Private Sub AlignControls()
Dim i As Long
Dim lAdditive As Long

cmdNextPage.Left = UserControl.ScaleWidth - cmdNextPage.Width
cmdLastPage.Left = 0
cmdNextPage.Top = 0
cmdLastPage.Top = 0
cmdNextPage.Height = UserControl.ScaleHeight
cmdLastPage.Height = UserControl.ScaleHeight

For i = 0 To imgPage.Count - 1
    If Not imgPage(i) Is Nothing Then
        imgPage(i).Top = (UserControl.Height - imgPage(0).Height - lblPageCaption(0).Height - 40) / 2
        lblPageCaption(i).Top = imgPage(i).Top + imgPage(i).Height + 2
    
        If i = 0 Then
            imgPage(i).Left = (i * (lblPageCaption(0).Width)) + 500 + (i * 60)
            lblPageCaption(i).Left = imgPage(i).Left - 220
        Else
            imgPage(i).Left = (i * (lblPageCaption(i - 1).Width)) + 500 + (i * 60)
            lblPageCaption(i).Left = imgPage(i).Left - 220
        End If
    End If
Next

picBack.Width = (imgPage.Count * (lblPageCaption(0).Width)) + 700 + (i * 60)
End Sub

Private Sub SetSelectedItem(vItem As Long)
Deselect SelectedItem

lblPageCaption(vItem).BackStyle = 1
lblPageCaption(vItem).BackColor = vbHighlight
lblPageCaption(vItem).ForeColor = vbHighlightText
m_SelectedItem = vItem
End Sub

Private Sub Deselect(inde As Long)
lblPageCaption(inde).BackStyle = 0
lblPageCaption(inde).ForeColor = vbWindowText
'lblPageCaption(Inde).BorderStyle = 0
End Sub

Private Sub DoClick(ByVal Index As Integer)
SetSelectedItem CLng(Index)
DoEvents
RaiseEvent ItemClick(CLng(Index))
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long

PropBag.WriteProperty "BackColor", m_BackColor, vbApplicationWorkspace
PropBag.WriteProperty "ItemCount", m_ItemCount, 1

For i = 0 To imgPage.Count - 1
    PropBag.WriteProperty "ItemCapt" & i, CStr(lblPageCaption(i)), "The Item"
Next i
    Call PropBag.WriteProperty("ToolTipText", picBack.ToolTipText, "")
    Call PropBag.WriteProperty("Style", UserControl.BorderStyle, 1)
End Sub

Private Sub MakeVisible(vItem As Long)
Dim vStart As Long
Dim vEnd As Long
Dim imgVal As Long
Dim ItemWidth As Long
Dim ItemLeft As Long


vStart = Abs(picBack.Left)
vEnd = UserControl.Width

ItemWidth = lblPageCaption(vItem).Width
ItemLeft = lblPageCaption(vItem).Left


If ItemLeft < vStart Then
    picBack.Left = -(ItemLeft - 400)
ElseIf (ItemLeft + ItemWidth - vStart) > vEnd Then
    picBack.Left = vEnd - (ItemLeft + ItemWidth) - 400
End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = picBack.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    picBack.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(vNewColor As OLE_COLOR)
picBack.BackColor = vNewColor
UserControl.BackColor = vNewColor
m_BackColor = vNewColor
PropertyChanged "BackColor"
End Property

Public Property Get About()
MsgBox "BottomBar ver 1.00(freeware)" & vbCrLf & _
"Developed by Cyril M Gupta" & vbCrLf & _
"Email: psl@nde.vsnl.net.in" & vbCrLf & _
"Source is available for 10$."
End Property

Private Property Get EnableLeftScroll()
If picBack.Left >= 0 Then EnableLeftScroll = False Else EnableLeftScroll = True
End Property
Private Property Get EnableRightScroll()
If picBack.Left <= UserControl.Width - picBack.Width Then EnableRightScroll = False Else EnableRightScroll = True
End Property

Public Property Get ItemCount() As Long
ItemCount = m_ItemCount
End Property

Public Property Let ItemCount(vNewCount As Long)
Dim i As Long

m_ItemCount = vNewCount

If vNewCount = 1 Then
    imgPage(0).Visible = True
    lblPageCaption(0).Visible = True
End If

If vNewCount > imgPage.Count Then
    imgPage(0).Visible = True
    lblPageCaption(0).Visible = True

    For i = imgPage.Count To vNewCount - 1
        AddItem "The Item"
    Next i
ElseIf vNewCount < imgPage.Count Then
    For i = imgPage.Count To vNewCount + 1 Step -1
        RemoveItem i - 1
    Next i
End If

PropertyChanged "ItemCount"
End Property

Public Property Get ItemCaption() As String
ItemCaption = lblPageCaption(SelectedItem).Caption
End Property

Public Property Let ItemCaption(vNewCaption As String)
lblPageCaption(SelectedItem).Caption = vNewCaption
PropertyChanged "ItemCaption"
End Property

Public Property Get SelectedItem() As Long
SelectedItem = m_SelectedItem
End Property

Public Property Let SelectedItem(vItem As Long)
SetSelectedItem vItem
m_SelectedItem = vItem

MakeVisible vItem

PropertyChanged "SelecteItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets the border style for an object."
    Style = UserControl.BorderStyle
End Property

Public Property Let Style(ByVal New_Style As Integer)
    UserControl.BorderStyle() = New_Style
    PropertyChanged "Style"
End Property

