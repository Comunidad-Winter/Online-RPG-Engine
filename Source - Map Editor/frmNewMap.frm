VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Map..."
   ClientHeight    =   3945
   ClientLeft      =   3705
   ClientTop       =   2790
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Map size"
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      Begin VB.TextBox MapHeight 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "50"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox MapWidth 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "50"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "height"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "width"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.ListBox GrhList 
      Height          =   1620
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox AlphaBlendChk 
      Caption         =   "Alpha Blending"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox NewMapBaseLightColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2880
      ScaleHeight     =   300
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton NewMapSelectColorCmd 
      Caption         =   "Change..."
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "Base light"
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
      Begin VB.Label Label6 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   310
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Angle"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   3375
      Begin VB.PictureBox picRotate 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   825
         ScaleWidth      =   945
         TabIndex        =   17
         Top             =   240
         Width           =   975
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            Height          =   135
            Left            =   430
            Top             =   360
            Width           =   135
         End
         Begin VB.Line LineRotate 
            BorderColor     =   &H00FFFFFF&
            X1              =   500
            X2              =   500
            Y1              =   360
            Y2              =   0
         End
      End
      Begin VB.CommandButton GrhIncreaseAngleCmd 
         Caption         =   "+"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton GrhDecreaseAngleCmd 
         Caption         =   "-"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox GrhAngleTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Text            =   "0"
         Top             =   505
         Width           =   975
      End
      Begin MSComctlLib.Slider GrhAngleSlider 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Max             =   360
         TickStyle       =   3
         TickFrequency   =   5
      End
   End
End
Attribute VB_Name = "frmNewMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp_angle As Long
Dim temp_light As Long

Private Sub CancelButton_Click()

'Place GrhViewer back on top
If frmMain.GrhViewerMnuChk.Checked Then
    General_Form_On_Top_Set frmGrhViewer, True
End If

'Place Minimap back on top
If frmMain.MiniMapMnuChk.Checked Then
    General_Form_On_Top_Set frmMap, True
End If

Unload Me

End Sub

Private Sub Form_Load()

'Set all data acording to user´s default
MapWidth = map_width
MapHeight = map_height

temp_light = &HFFFFFF 'White

'Load Grh list
Engine.Grh_Add_GrhList_To_ListBox GrhList

End Sub

Private Sub GrhAngleSlider_Scroll()
'Code taken from Fredrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
temp_angle = GrhAngleSlider.value

With LineRotate
 .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(temp_angle * PI / 180)
 .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(temp_angle * PI / 180)
 .x1 = picRotate.width / 2
 .y1 = picRotate.height / 2
End With
GrhAngleTxt.text = Str(temp_angle)

End Sub

Private Sub GrhAngleTxt_Change()
temp_angle = Val(GrhAngleTxt.text)
GrhAngleSlider.value = temp_angle

With LineRotate
 .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(temp_angle * PI / 180)
 .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(temp_angle * PI / 180)
End With

End Sub

Private Sub GrhDecreaseAngleCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Code taked from Fredrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
If Button = vbLeftButton Then
    temp_angle = temp_angle - 1
End If
If Button = vbRightButton Then
    temp_angle = temp_angle - 5
End If

If temp_angle < 0 Then
    temp_angle = 360 + temp_angle
End If

With LineRotate
 .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(temp_angle * PI / 180)
 .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(temp_angle * PI / 180)
 .x1 = picRotate.width / 2
 .y1 = picRotate.height / 2
End With
GrhAngleTxt.text = Str(temp_angle)
GrhAngleSlider.value = temp_angle

End Sub

Private Sub GrhIncreaseAngleCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Code taked from Fredrik´s Map Editor and edited by Juan Martín Sotuyo Dodero
If Button = vbLeftButton Then
    temp_angle = temp_angle + 1
End If
If Button = vbRightButton Then
    temp_angle = temp_angle + 5
End If

While temp_angle > 360
   temp_angle = temp_angle - 360
Wend

With LineRotate
 .x2 = picRotate.width / 2 + (picRotate.width / 2.5) * Sin(temp_angle * PI / 180)
 .y2 = picRotate.height / 2 - (picRotate.height / 2.5) * Cos(temp_angle * PI / 180)
 .x1 = picRotate.width / 2
 .y1 = picRotate.height / 2
End With
GrhAngleTxt.text = Str(temp_angle)
frmMain.GrhAngleSlider.value = temp_angle
End Sub

Private Sub MapHeight_KeyDown(KeyCode As Integer, Shift As Integer)
If Val(MapHeight.text) < 1 Then
    MapHeight.text = 1
End If
'MUST be an int
MapHeight.text = Int(Val(MapHeight.text))

End Sub

Private Sub MapWidth_KeyDown(KeyCode As Integer, Shift As Integer)
If Val(MapWidth.text) < 1 Then
    MapWidth.text = map_width
End If
'MUST be an int
MapWidth.text = Int(Val(MapWidth.text))

End Sub

Private Sub NewMapSelectColorCmd_Click()

On Error GoTo ErrHandler:

'GrhViewer shouldn´t be topmost anymore, or it will cover the dialog
If frmMain.GrhViewerMnuChk.Checked Then
    General_Form_On_Top_Set frmGrhViewer
End If

frmMain.Dialog.CancelError = True

Call ArrangeDialog(frmMain.Dialog, 3)

Dim r As Integer
Dim g As Integer
Dim b As Integer

General_Long_Color_to_RGB frmMain.Dialog.color, r, g, b

temp_light = RGB(b, g, r)
NewMapBaseLightColor.BackColor = temp_light

ErrHandler:

'Set GrhViewer back to it´s original state
If frmMain.GrhViewerMnuChk.Checked Then
    General_Form_On_Top_Set frmGrhViewer, True
End If

End Sub

Private Sub OKButton_Click()

'Disable Walk Mode
If Walk_Mode Then
    Toggle_Walk_Mode
    frmMain.WalkModeChk.value = 0
End If

'Ask to save if necessary
If Modified = True Then
    If MsgBox("Changes have been made since this map was last saved. If you don´t save, changes will be lost. Do you want to save now?", vbYesNo) = vbYes Then
        frmMain.SaveMnu_Click
    End If
End If

If Engine.Map_Create(Val(MapWidth.text), Val(MapHeight.text)) Then
    Engine.Map_Fill Val(GrhList.text), 1, temp_light, AlphaBlendChk.value, Val(GrhAngleTxt.text)
End If

'Reset current map
Current_Map = ""

'Re-load map list
Load_Maps_To_ComboBox frmMain.ExitMapsList

'Erase action list (we can undo or redo anything, since it´s a new map)
Clear_Action_List

'Reset view pos
Engine.Engine_View_Pos_Set 5, 5

'Create mini map
If frmMap.Visible Then
    Dim X As Long
    Dim Y As Long
    Engine.Map_Bounds_Get X, Y
    frmMap.picmain.width = X
    frmMap.picmain.height = Y
    frmMap.height = (frmMap.picmain.height + 28) * Screen.TwipsPerPixelY
    frmMap.width = (frmMap.picmain.width + 10) * Screen.TwipsPerPixelX
    
    frmMap.picmain.Cls
    
    'Reset minimap rect
    frmMap.shparea.top = -2
    frmMap.shparea.left = -4
    
    Engine.Engine_Render_Mini_Map_To_hDC frmMap.picmain.hdc
End If

CancelButton_Click

End Sub
