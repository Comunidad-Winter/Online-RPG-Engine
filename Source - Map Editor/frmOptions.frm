VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4800
   ClientLeft      =   3690
   ClientTop       =   2385
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox UseResFilesChk 
      Caption         =   "Use Resource Files"
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox UseIniFilesChk 
      Caption         =   "Save and load .ini files"
      Height          =   195
      Left            =   1080
      TabIndex        =   25
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resource Path"
      Height          =   735
      Left            =   360
      TabIndex        =   23
      Top             =   2760
      Width           =   3615
      Begin VB.TextBox ResourcePathTxt 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   280
         Width           =   3375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Autosave delay"
      Height          =   735
      Left            =   2280
      TabIndex        =   20
      Top             =   1225
      Width           =   1695
      Begin VB.TextBox DelayInSecsTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   290
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Secs."
         Height          =   195
         Left            =   1080
         TabIndex        =   22
         Top             =   330
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Size"
      Height          =   1095
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Width           =   1695
      Begin VB.TextBox MapWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   280
         Width           =   855
      End
      Begin VB.TextBox MapHeight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Top             =   640
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   105
      End
   End
   Begin VB.CommandButton ApplyCmd 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tile Size"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2010
      Width           =   1695
      Begin VB.TextBox TileSizeTxt 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   280
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Pixels"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   330
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Border Size"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      Begin VB.TextBox XBorder 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   280
         Width           =   855
      End
      Begin VB.TextBox YBorder 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Top             =   640
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   105
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Char label"
      Height          =   735
      Left            =   360
      TabIndex        =   11
      Top             =   1225
      Width           =   1695
      Begin VB.TextBox CharLabel 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   290
         Width           =   1455
      End
   End
   Begin VB.TextBox EngineSpeed 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   2325
      Width           =   735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Engine´s speed"
      Height          =   735
      Left            =   2280
      TabIndex        =   14
      Top             =   2010
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Error As Boolean
Dim tile_size_changed As Boolean

Private Sub ApplyCmd_Click()

'Check data is correct
'Engine speed
If Val(EngineSpeed) <= 0 Then
    MsgBox "Engine speed must be higher than 0. The decimal separator is period (.)"
    Error = True
    Exit Sub
End If

'Check if tile size was changed
If tile_size_changed = True Then
    MsgBox "In order to change the tile size you must close down, and run the Map Editor again."
End If

'Set general map data
Engine.Engine_Base_Speed_Set Val(EngineSpeed.text)
Engine.Char_Label_Set User_Char_Index, CharLabel.text, 1
base_speed = Val(EngineSpeed.text)
x_border = Val(XBorder.text)
y_border = Val(YBorder.text)
tile_size = Val(TileSizeTxt.text)
map_width = MapWidth.text
map_height = MapHeight.text
autosave_delay = Val(DelayInSecsTxt.text) * 1000
resource_path = ResourcePathTxt.text
use_ini_files = CBool(UseIniFilesChk.value)
use_resource_files = CBool(UseResFilesChk.value)

Call Save_User_Defined_Data

ApplyCmd.Enabled = False

'Render map to see changes
prgRun = Engine.Engine_Render_Start
prgRun = Engine.Engine_Render_End

Error = False

End Sub

Private Sub CancelButton_Click()

'Place GrhViewer back on top if necessary
If frmMain.GrhViewerMnuChk.Checked Then
    General_Form_On_Top_Set frmGrhViewer, True
End If
'Same with minimap
If frmMain.MiniMapMnuChk.Checked Then
    General_Form_On_Top_Set frmMap, True
End If

Unload Me

End Sub

Private Sub CharLabel_Change()

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub DelayInSecsTxt_Change()
If Val(DelayInSecsTxt.text) < 0 Then
    DelayInSecsTxt.text = 10
End If

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub EngineSpeed_Change()
'Arrange the decimal saparator
Dim temp As String
temp = General_Field_Read(1, EngineSpeed.text, 44)
If Not temp = "" And Not Len(temp) = Len(EngineSpeed) Then
    EngineSpeed = temp & "." & Right(EngineSpeed.text, Len(EngineSpeed.text) - Len(temp) - 1)
    EngineSpeed.SelLength = 0
    EngineSpeed.SelStart = Len(EngineSpeed.text)
End If

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub Form_Load()

'Re-load all User Defined data
Call Load_User_Defined_Data

'Set all data
MapWidth.text = map_width
MapHeight.text = map_height
XBorder.text = x_border
YBorder.text = y_border
TileSizeTxt.text = tile_size
EngineSpeed = base_speed
CharLabel.text = char_label
DelayInSecsTxt.text = autosave_delay / 1000
ResourcePathTxt.text = resource_path
If use_ini_files Then
    UseIniFilesChk.value = vbChecked
Else
    UseIniFilesChk.value = vbUnchecked
End If
If use_resource_files Then
    UseResFilesChk.value = vbChecked
Else
    UseResFilesChk.value = vbUnchecked
End If

'Disable Aplly button
ApplyCmd.Enabled = False

tile_size_changed = False

End Sub

Private Sub MapHeight_Change()
If Val(MapHeight.text) < 1 Then
    MapHeight = 1
End If
MapHeight.text = Int(Val(MapHeight.text))

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub MapWidth_Change()
If Val(MapWidth.text) < 1 Then
    MapWidth = 1
End If
'MUST be an int
MapWidth.text = Int(Val(MapWidth.text))

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub OKButton_Click()

Call ApplyCmd_Click

If Error = False Then
    Unload Me
End If

'Place GrhViewer back on top
General_Form_On_Top_Set frmGrhViewer, True

End Sub

Private Sub ResourcePathTxt_Change()

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub TileSizeTxt_Change()

If Val(TileSizeTxt.text) < 1 Then
    TileSizeTxt.text = 1
End If
'MUST be an int
TileSizeTxt.text = Int(Val(TileSizeTxt.text))

tile_size_changed = True

End Sub

Private Sub UseIniFilesChk_Click()

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub XBorder_Change()
If Val(XBorder.text) < 0 Then
    XBorder.text = 0
End If
'MUST be an int
XBorder.text = Int(Val(XBorder.text))

'Enable Apply button
ApplyCmd.Enabled = True

End Sub

Private Sub YBorder_Change()
If Val(YBorder.text) < 0 Then
    YBorder.text = 0
End If
'MUST be an int
YBorder.text = Int(Val(YBorder.text))

'Enable Apply button
ApplyCmd.Enabled = True

End Sub
