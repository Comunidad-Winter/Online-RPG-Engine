VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTileGroups 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile Groups"
   ClientHeight    =   2775
   ClientLeft      =   1875
   ClientTop       =   1950
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Fill Map"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox ShowGridChk 
      Caption         =   "Show Grid"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton SetGridColorCmd 
      Caption         =   "Set Grid Color..."
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton UpDown 
      Caption         =   "Down"
      Height          =   405
      Index           =   0
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   600
   End
   Begin VB.CommandButton UpDown 
      Caption         =   "Up"
      Height          =   405
      Index           =   1
      Left            =   7080
      TabIndex        =   1
      Top             =   480
      Width           =   600
   End
   Begin VB.PictureBox TileGroupViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   1950
      Left            =   3600
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   0
      Top             =   120
      Width           =   3390
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   7
         Visible         =   0   'False
         X1              =   224
         X2              =   0
         Y1              =   31
         Y2              =   31
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   9
         Visible         =   0   'False
         X1              =   224
         X2              =   0
         Y1              =   95
         Y2              =   95
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   8
         Visible         =   0   'False
         X1              =   224
         X2              =   0
         Y1              =   63
         Y2              =   63
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   6
         Visible         =   0   'False
         X1              =   191
         X2              =   191
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   5
         Visible         =   0   'False
         X1              =   159
         X2              =   159
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         Visible         =   0   'False
         X1              =   127
         X2              =   127
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         Visible         =   0   'False
         X1              =   95
         X2              =   95
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         Visible         =   0   'False
         X1              =   63
         X2              =   63
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Line Grid 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         Visible         =   0   'False
         X1              =   31
         X2              =   31
         Y1              =   0
         Y2              =   128
      End
   End
   Begin MSComctlLib.TreeView tree 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmTileGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GridColor As Long

Private Sub Command1_Click()
'Check if a Grh was chosen
If current_grh = -1 Then
    MsgBox "Must select the Grh to use first.", vbOKOnly
    Exit Sub
End If

'Make sure it´s not a decorative layer
If frmMain.GrhLayerList.ListIndex = 0 Or frmMain.GrhLayerList.ListIndex = 2 Or frmMain.GrhLayerList.ListIndex = 4 Then
    If MsgBox(frmMain.GrhLayerList.text & " layer will be filled with Grh " & current_grh & ". Are you sure?", vbOKCancel) = vbOK Then
        'store_action grh, fill, , , , frmMain.GrhLayerList.ListIndex + 1
        Engine.Map_Fill current_grh, frmMain.GrhLayerList.ListIndex + 1, , frmMain.GrhAlphaBlendingChk.value, Val(frmMain.GrhAngleTxt), frmMain.GrhHCenteredChk.value, frmMain.GrhVCenteredChk.value
        Modified = True
    End If
Else
    MsgBox "Can´t fill a decoration layer with a Grh. select a valid layer."
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    'Allow it to be hidden
    General_Form_On_Top_Set Me
    Me.Hide
    
    'Set the GrhViewer back on top if visible
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    If frmMain.MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
End Sub

Private Sub SetGridColorCmd_Click()
On Error GoTo ErrHandler:

    frmMain.Dialog.CancelError = True
    
    'Make sure nothing covers the color dialog before displaying it
    General_Form_On_Top_Set Me
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer
    End If
    
    ArrangeDialog frmMain.Dialog, 3
    
    GridColor = frmMain.Dialog.color
    Draw_Tile_Group ShowGridChk.value, GridColor
    
ErrHandler:
    General_Form_On_Top_Set Me, True
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
End Sub

Private Sub ShowGridChk_Click()

Draw_Tile_Group ShowGridChk.value, GridColor

End Sub

Private Sub TileGroupViewer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

current_grh = Tile_Group_Index_Get(X, Y)

If frmMain.GrhViewerMnuChk.Checked Then
    frmGrhViewer.Cls
    Engine.Grh_Render_To_Hdc current_grh, frmGrhViewer.hdc, 0, 0
End If

End Sub

Private Sub tree_NodeClick(ByVal Node As MSComctlLib.Node)
On Local Error GoTo ErrHandler
    Dim grh As Long
    Dim CurrentNode As Node
    Dim LastSibling As Node
    Dim LoopC As Long
    
    'A Grh was clicked, load it´s group
    If left$(Node.Key, 3) = "grh" Then
        'Load parent´s group
        Current_Group.Name = Node.Parent.text
        
        Erase Current_Group.GrhIndexes
        ReDim Current_Group.GrhIndexes(1 To Node.Parent.Children)
        Set CurrentNode = Node.FirstSibling
        Set LastSibling = Node.LastSibling
        
        'Set it as the selected Grh
        current_grh = Val(Right(Node.text, Len(Node.text) - 4))
        'Draw it in the GrhViewer
        If frmMain.GrhViewerMnuChk.Checked Then
            Engine.Grh_Render_To_Hdc current_grh, frmGrhViewer.hdc, 0, 0
        End If
        
    'A parent node was clicked, load all Grhs inmediately inside (not grandchildren)
    Else
        'Load parent´s group
        Current_Group.Name = Node.text
        
        ReDim Current_Group.GrhIndexes(1 To Node.Children)
        Set CurrentNode = Node.Child
        Set LastSibling = Node.Child.LastSibling
    End If
    
    Do Until CurrentNode.Key = LastSibling.Key
        If left$(CurrentNode.text, 4) = "Grh " Then
            LoopC = LoopC + 1
            Current_Group.GrhIndexes(LoopC) = Val(Right(CurrentNode.text, Len(CurrentNode.text) - 4))
        End If
        Set CurrentNode = CurrentNode.Next
    Loop
    
    'Set last one
    If left$(LastSibling.text, 4) = "Grh " Then
        Current_Group.GrhIndexes(LoopC + 1) = Val(Right(LastSibling.text, Len(CurrentNode.text) - 4))
    End If
    
    'Reset offset
    TileGroupOffset = 0
    
    'Draw it
    Draw_Tile_Group ShowGridChk.value, GridColor
    
    Set CurrentNode = Nothing
    Set LastSibling = Nothing
Exit Sub

ErrHandler:
    Set CurrentNode = Nothing
    Set LastSibling = Nothing
End Sub

Private Sub UpDown_Click(index As Integer)

Select Case index
    Case 0:
        TileGroupOffset = TileGroupOffset + 1
    Case 1:
        If TileGroupOffset > 0 Then TileGroupOffset = TileGroupOffset - 1
End Select

Draw_Tile_Group ShowGridChk.value, GridColor

End Sub
