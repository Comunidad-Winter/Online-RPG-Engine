VERSION 5.00
Begin VB.Form frmAdjacentMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Adjacent Map"
   ClientHeight    =   5160
   ClientLeft      =   2565
   ClientTop       =   1800
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame12 
      Caption         =   "Y"
      Height          =   975
      Left            =   3000
      TabIndex        =   38
      Top             =   3600
      Width           =   1695
      Begin VB.TextBox WestYFinish 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   40
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox WestYStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   39
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Finish at:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   645
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Start at:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   285
         Width           =   555
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Y"
      Height          =   975
      Left            =   3000
      TabIndex        =   33
      Top             =   1440
      Width           =   1695
      Begin VB.TextBox EastYStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   35
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox EastYFinish 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   34
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Start at:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Finish at:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "X"
      Height          =   975
      Left            =   3000
      TabIndex        =   28
      Top             =   2520
      Width           =   1695
      Begin VB.TextBox SouthXStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   30
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox SouthXFinish 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   29
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Start at:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Finish at:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "X"
      Height          =   975
      Left            =   3000
      TabIndex        =   23
      Top             =   360
      Width           =   1695
      Begin VB.TextBox NorthXFinish 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Text            =   "100"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox NorthXStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Finish at:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   645
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Start at:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   285
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Dimensions"
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
      Begin VB.Label WestMaxY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   150
      End
      Begin VB.Label WestMaxX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dimensions"
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
      Begin VB.Label SouthMaxY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   150
      End
      Begin VB.Label SouthMaxX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dimensions"
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
      Begin VB.Label EastMaxY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   150
      End
      Begin VB.Label EastMaxX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dimensions"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   840
      Width           =   1695
      Begin VB.Label NorthMaxY 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   150
      End
      Begin VB.Label NorthMaxX 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.ComboBox EastMap 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox SouthMap 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ComboBox WestMap 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox NorthMap 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "East:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   525
      TabIndex        =   6
      Top             =   1635
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "South:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   2700
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "West:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   525
      TabIndex        =   4
      Top             =   3795
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "North:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   555
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Set map to the..."
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "frmAdjacentMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim map_max_x As Long
Dim map_max_y As Long

Private Sub CancelButton_Click()

General_Form_On_Top_Set frmGrhViewer, True

Unload Me

End Sub

Private Sub EastMap_Click()
'Load map´s dimensions
Dim max_x As Long
Dim max_y As Long

If EastMap.text = "Current Map" Then
    Engine.Map_Bounds_Get max_x, max_y
ElseIf EastMap.text = "None" Then
    'Set caption
    EastMaxX.Caption = "X: -"
    EastMaxY.Caption = "Y: -"
    Exit Sub
Else
    Engine.Map_Bounds_Get_From_File EastMap.text, max_x, max_y
End If
'Set caption
EastMaxX.Caption = "X: " & max_x
EastMaxY.Caption = "Y: " & max_y

If EastYStart.text > max_y Then
    EastYStart.text = max_y
End If

If EastYFinish.text > max_y Then
    EastYFinish.text = max_y
End If

End Sub

Private Sub EastYFinish_Change()
Dim start As Long
Dim finish As Long

start = Val(EastYStart.text)
finish = Val(EastYFinish.text)

Check_Values_Y start, finish

EastYStart.text = start
EastYFinish.text = finish

End Sub

Private Sub EastYStart_Change()
Dim start As Long
Dim finish As Long

start = Val(EastYStart.text)
finish = Val(EastYFinish.text)

Check_Values_Y start, finish

EastYStart.text = start
EastYFinish.text = finish

End Sub

Private Sub Form_Load()

'Load maps
Load_Maps_To_ComboBox NorthMap
Load_Maps_To_ComboBox EastMap
Load_Maps_To_ComboBox SouthMap
Load_Maps_To_ComboBox WestMap

'Get current map´s dimensions
Engine.Map_Bounds_Get map_max_x, map_max_y

End Sub

Private Sub NorthMap_Click()
'Load map´s dimensions
Dim max_x As Long
Dim max_y As Long

If NorthMap.text = "Current Map" Then
    Engine.Map_Bounds_Get max_x, max_y
ElseIf NorthMap.text = "None" Then
    'Set caption
    NorthMaxX.Caption = "X: -"
    NorthMaxY.Caption = "Y: -"
    Exit Sub
Else
    Engine.Map_Bounds_Get_From_File NorthMap.text, max_x, max_y
End If
'Set caption
NorthMaxX.Caption = "X: " & max_x
NorthMaxY.Caption = "Y: " & max_y

If NorthXStart.text > max_x Then
    NorthXStart.text = max_x
End If

If NorthXFinish.text > max_x Then
    NorthXFinish.text = max_x
End If

End Sub

Private Sub NorthXFinish_Change()
Dim start As Long
Dim finish As Long

start = Val(NorthXStart.text)
finish = Val(NorthXFinish.text)

Check_Values_X start, finish

NorthXStart.text = start
NorthXFinish.text = finish

End Sub

Private Sub NorthXStart_Change()
Dim start As Long
Dim finish As Long

start = Val(NorthXStart.text)
finish = Val(NorthXFinish.text)

Check_Values_X start, finish

NorthXStart.text = start
NorthXFinish.text = finish

End Sub

Private Sub OKButton_Click()
Dim X As Long
Dim Y As Long
Dim LoopC As Long

'North
If NorthMap.text <> "None" Then
    'Place an exit on all unblocked tiles
    Y = y_border + 1
    For X = Val(NorthXStart.text) To Val(NorthXFinish.text)
        If Not Engine.Map_Blocked_Get(X, Y) Then
            Engine.Map_Exit_Add X, Y, NorthMap.text, X, map_max_y - Y
            LoopC = LoopC + 1
            store_action exits, fill, X, Y, LoopC, , , , , , , , , , , True
            Modified = True
        End If
    Next X
End If

'East
If EastMap.text <> "None" Then
    'Place an exit on all unblocked tiles
    X = map_max_x - x_border
    For Y = Val(EastYStart.text) To Val(EastYFinish.text)
        If Not Engine.Map_Blocked_Get(X, Y) Then
            Engine.Map_Exit_Add X, Y, EastMap.text, x_border + 1, Y
            LoopC = LoopC + 1
            store_action exits, fill, X, Y, LoopC, , , , , , , , , , , True
            Modified = True
        End If
    Next Y
End If

'South
If SouthMap.text <> "None" Then
    'Place an exit on all unblocked tiles
    Y = map_max_y - y_border
    For X = Val(SouthXStart.text) To Val(SouthXFinish.text)
        If Not Engine.Map_Blocked_Get(X, Y) Then
            Engine.Map_Exit_Add X, Y, SouthMap.text, X, y_border + 1
            LoopC = LoopC + 1
            store_action exits, fill, X, Y, LoopC, , , , , , , , , , , True
            Modified = True
        End If
    Next X
End If

'West
If WestMap.text <> "None" Then
    'Place an exit on all unblocked tiles
    X = x_border + 1
    For Y = Val(WestYStart.text) To Val(WestYFinish.text)
        If Not Engine.Map_Blocked_Get(X, Y) Then
            Engine.Map_Exit_Add X, Y, WestMap.text, map_max_x - X, Y
            LoopC = LoopC + 1
            store_action exits, fill, X, Y, LoopC, , , , , , , , , , , True
            Modified = True
        End If
    Next Y
End If

'Hide form
If frmMain.GrhViewerMnuChk.Checked Then
    General_Form_On_Top_Set frmGrhViewer, True
End If
If frmMain.MiniMapMnuChk.Checked Then
    General_Form_On_Top_Set frmMap, True
End If
Unload Me

End Sub

Private Sub SouthMap_Click()
'Load map´s dimensions
Dim max_x As Long
Dim max_y As Long

If SouthMap.text = "Current Map" Then
    Engine.Map_Bounds_Get max_x, max_y
ElseIf SouthMap.text = "None" Then
    'Set caption
    SouthMaxX.Caption = "X: -"
    SouthMaxY.Caption = "Y: -"
    Exit Sub
Else
    Engine.Map_Bounds_Get_From_File SouthMap.text, max_x, max_y
End If
'Set caption
SouthMaxX.Caption = "X: " & max_x
SouthMaxY.Caption = "Y: " & max_y

If SouthXStart.text > max_x Then
    SouthXStart.text = max_x
End If

If SouthXFinish.text > max_x Then
    SouthXFinish.text = max_x
End If

End Sub

Private Sub SouthXFinish_Change()
Dim start As Long
Dim finish As Long

start = Val(SouthXStart.text)
finish = Val(SouthXFinish.text)

Check_Values_X start, finish

SouthXStart.text = start
SouthXFinish.text = finish

End Sub

Private Sub SouthXStart_Change()
Dim start As Long
Dim finish As Long

start = Val(SouthXStart.text)
finish = Val(SouthXFinish.text)

Check_Values_X start, finish

SouthXStart.text = start
SouthXFinish.text = finish

End Sub

Private Sub Check_Values_X(ByRef start As Long, ByRef finish As Long)

If start < 0 Or start > finish Then
    start = 1
End If
'MUST be an int
start = Int(start)
If start > map_max_x Then
    start = map_max_x
End If

If finish < 0 Then
    finish = map_max_x
End If
If finish > map_max_x Then
    finish = map_max_x
End If
finish = Int(finish)
'start MUST be smaller than Finish
If start > finish Then
    MsgBox "Start point can´t be greater than finish point."
    start = 1
    finish = map_max_x
End If

End Sub

Private Sub Check_Values_Y(ByRef start As Long, ByRef finish As Long)

If start < 0 Or start > finish Then
    start = 1
End If
'MUST be an int
start = Int(start)
If start > map_max_y Then
    start = map_max_y
End If

If finish < 0 Then
    finish = map_max_y
End If
If finish > map_max_y Then
    finish = map_max_y
End If
finish = Int(finish)
'start MUST be smaller than Finish
If start > finish Then
    MsgBox "Start point can´t be greater than finish point."
    start = 1
    finish = map_max_y
End If

End Sub

Private Sub WestMap_Click()
'Load map´s dimensions
Dim max_x As Long
Dim max_y As Long

If WestMap.text = "Current Map" Then
    Engine.Map_Bounds_Get max_x, max_y
ElseIf WestMap.text = "None" Then
    'Set caption
    WestMaxX.Caption = "X: -"
    WestMaxY.Caption = "Y: -"
    Exit Sub
Else
    Engine.Map_Bounds_Get_From_File WestMap.text, max_x, max_y
End If
'Set caption
WestMaxX.Caption = "X: " & max_x
WestMaxY.Caption = "Y: " & max_y

If WestYStart.text > max_y Then
    WestYStart.text = max_y
End If

If WestYFinish.text > max_y Then
    WestYFinish.text = max_y
End If

End Sub

Private Sub WestYFinish_Change()
Dim start As Long
Dim finish As Long

start = Val(WestYStart.text)
finish = Val(WestYFinish.text)

Check_Values_Y start, finish

WestYStart.text = start
WestYFinish.text = finish

End Sub

Private Sub WestYStart_Change()
Dim start As Long
Dim finish As Long

start = Val(WestYStart.text)
finish = Val(WestYFinish.text)

Check_Values_Y start, finish

WestYStart.text = start
WestYFinish.text = finish

End Sub
