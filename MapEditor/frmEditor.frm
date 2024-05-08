VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ore Map Editor"
   ClientHeight    =   7215
   ClientLeft      =   135
   ClientTop       =   705
   ClientWidth     =   9600
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.TextBox txtMapDesc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Untitled Map"
      Top             =   120
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonD 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   638
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin VB.Timer AutosaveTimer 
         Left            =   120
         Top             =   600
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Map"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Map"
      End
      Begin VB.Menu mnuStripe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Map"
      End
      Begin VB.Menu mnuStripe2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuGrhViewer 
         Caption         =   "Grh Viewer"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmEditor.frm - ORE Map Editor - v0.9.3
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'***************************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/17/2003
'   -Second official/unofficial release
'   -Change: Added NPC, Item and Exit support.
'            Also remade the settings dialouge.
'
'Aaron Perkins(aaron@baronsoft.com) - 5/12/2003
'   -First offical release
'   -Change: Moved the cRad and AlwaysOnTop functions to the
'               modGeneral.bas
'
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/8/2003
'   -Last unoffical release
'   -Orginal creator
'
'*****************************************************************
Option Explicit

Public engine As clsTileEngineX
Dim running As Boolean
Public ToolUsed As Integer
Public ParticelCCount As Long
Dim RESDIR As String
Public MapModified As Boolean
Public MapBaseLight As Long

Private Sub CheckKeyboard()

If engine.Input_Key_Get(vbKeyUp) Then
 engine.Engine_View_Move 1
End If
    
If engine.Input_Key_Get(vbKeyPageUp) Then
 engine.Engine_View_Move 2
End If
    
If engine.Input_Key_Get(vbKeyRight) Then
 engine.Engine_View_Move 3
End If
    
If engine.Input_Key_Get(vbKeyPageDown) Then
 engine.Engine_View_Move 4
End If
    
If engine.Input_Key_Get(vbKeyDown) Then
 engine.Engine_View_Move 5
End If
    
If engine.Input_Key_Get(vbKeyEnd) Then
 engine.Engine_View_Move 6
End If
    
If engine.Input_Key_Get(vbKeyLeft) Then
 engine.Engine_View_Move 7
End If
    
If engine.Input_Key_Get(vbKeyHome) Then
 engine.Engine_View_Move 8
End If

If engine.Input_Key_Get(vbKeyF1) Then
 Select_Tool (1)
End If
If engine.Input_Key_Get(vbKeyF2) Then
 Select_Tool (2)
End If
If engine.Input_Key_Get(vbKeyF3) Then
 Select_Tool (3)
End If
If engine.Input_Key_Get(vbKeyF4) Then
 Select_Tool (4)
End If
If engine.Input_Key_Get(vbKeyF5) Then
 Select_Tool (5)
End If
End Sub

Private Sub AutosaveTimer_Timer()
engine.Map_Save_Map_To_File App.Path & "\BACKUPMAP.map"
End Sub

Private Sub Form_Load()
Dim AutosaveTimerString As String
Dim ParticelCCound As Long
Dim numNPCs As Long
Dim NPCNum As Long
ParticelCCound = -1
Set engine = New clsTileEngineX

'Load Autosave stuff
AutosaveTimerString = General_Var_Get(App.Path & "\settings.ini", "OPTIONS", "ASINTERVAL")
AutosaveTimer.Interval = Val(AutosaveTimerString)
ToolUsed = 1

'Initialize Ore
running = engine.Engine_Initialize(frmEditor.hWnd, Picture1.hWnd, 1, App.Path & General_Var_Get(App.Path & "\settings.ini", "PATH", "RES"))
resource_path = General_Var_Get(App.Path & "\settings.ini", "PATH", "RES")
'Add grh to lists
engine.Grh_Add_GrhList_To_ListBox frmSettings.LstGrh
engine.Grh_Add_GrhList_To_ListBox frmSettings.lstPGrh
'Adds all the NPCs to a listbox
NPC_Add_List_To_Settings
Item_Add_List_To_Settings

'Show all the forms
frmSettings.Show
frmToolBox.Show
frmGrhViewer.Show
Me.Show

' Set some stuff
engine.Engine_Stats_Show_Toggle
engine.Engine_Base_Speed_Set 0.03
engine.Engine_View_Pos_Set 11, 11
engine.Engine_Blocked_Tiles_Show_Toggle
engine.Engine_Special_Tiles_Show_Toggle
engine.Map_Base_Light_Fill RGB(190, 190, 190)
MapBaseLight = RGB(190, 190, 190)
Select_Tool (1)
'Main loop

Do While running
 engine.Engine_Render_Start
 engine.GUI_Box_Filled_Render 5, 5, 174, 24, RGB(255, 140, 10), RGB(210, 140, 0), RGB(210, 140, 0), RGB(210, 140, 0), True
 engine.GUI_Box_Outline_Render 5, 5, 174, 24, 1, RGB(255, 255, 255)
 engine.Engine_Render_End
 CheckKeyboard
 DoEvents
Loop

engine.Engine_Deinitialize
Set engine = Nothing
Unload Me
End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim tempResponse As Integer
If MapModified = True Then
 tempResponse = MsgBox("Your recent updates ain't saved. So are you totally sure that you want to quit?", 4)
  Select Case tempResponse
   Case 6
   running = False
   Cancel = 1
   Case 7
   Cancel = 1
   Exit Sub
  End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
running = False
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Dim tempResponse As Integer
If MapModified = True Then
 tempResponse = MsgBox("Your recent updates ain't saved. So are you totally sure that you want to quit?", 4)
  Select Case tempResponse
   Case 6
   running = False
   Exit Sub
   Case 7
   Exit Sub
  End Select
End If
running = False
End Sub

Private Sub mnuGrhViewer_Click()
frmGrhViewer.Show
mnuGrhViewer.Checked = True
End Sub

Private Sub mnuLoad_Click()
On Error GoTo Cancel
CommonD.Filter = "OreMapFiles (*.map)|*.map|All files (*.*)|*.*"
CommonD.ShowOpen

If CommonD.FileName <> "" Then
 engine.Map_Load_Map_From_File CommonD.FileName, True
 txtMapDesc.text = engine.Map_Description_Get
End If
Cancel:
End Sub

Private Sub mnuNew_Click()
frmNew.Show
End Sub

Private Sub mnuSave_Click()
On Error GoTo Cancel
    CommonD.Filter = "OreMapFiles (*.map)|*.map|All files (*.*)|*.*"
    CommonD.ShowSave
    If CommonD.FileName <> "" Then
                        'I know it should be to
    engine.Map_Save_Map_To_File CommonD.FileName, True
    MapModified = False
    End If

Cancel:
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo Cancel
If engine.Input_Mouse_In_View Then
 Dim temp_x As Long
 Dim temp_y As Long
 engine.Input_Mouse_Map_Get temp_x, temp_y
 
 If engine.Input_Mouse_Button_Left_Get Then
 MapModified = True
'********** Grh Tool **********************
  If ToolUsed = 1 Then
   If frmSettings.LstGrh.ListIndex <> -1 Then
    If frmSettings.chkAlpha.value = 1 Then
     engine.Map_Grh_Set temp_x, temp_y, frmSettings.LstGrh.List(frmSettings.LstGrh.ListIndex), frmSettings.cmbLayer.text, True, General_Covert_Degrees_To_Radians(frmSettings.angle) 'If alphablending used
    Else
     engine.Map_Grh_Set temp_x, temp_y, frmSettings.LstGrh.List(frmSettings.LstGrh.ListIndex), frmSettings.cmbLayer.text, , General_Covert_Degrees_To_Radians(frmSettings.angle) 'If not alphablending used
    End If
    If frmSettings.chkBlocked.value = 1 Then
     engine.Map_Blocked_Set temp_x, temp_y, True
    End If
    If frmSettings.chkBlocked.value = 0 Then
     engine.Map_Blocked_Set temp_x, temp_y, False
    End If
   End If
  End If
'********** Light Tool**********************
  If ToolUsed = 2 Then
   If frmSettings.txtLightId = "" Then
    engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text
   End If
   If frmSettings.txtLightId <> "" Then
    engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text, frmSettings.txtLightId.text
   End If
  End If
'********** Particel Tool ******************
  If ToolUsed = 3 Then
   If ParticelCCount <> 0 Then
    Dim temp_list() As Long
    Dim i As Long
    ReDim temp_list(1 To ParticelCCount)
   
    For i = 1 To ParticelCCount
     temp_list(i) = Val(frmSettings.lstParticels.List(i - 1))
    Next i
    If frmSettings.cmbStyle.text = "Star Burst" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 2
    ElseIf frmSettings.cmbStyle.text = "Fountain" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 1
    ElseIf frmSettings.cmbStyle.text = "Insects" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 3
    ElseIf frmSettings.cmbStyle.text = "Water Fall" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 4
    ElseIf frmSettings.cmbStyle.text = "Smoke" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 5
    ElseIf frmSettings.cmbStyle.text = "Fire" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 6
    End If
   
   End If
  End If
'********** Block Tool *********************
  If ToolUsed = 4 Then
   engine.Map_Blocked_Set temp_x, temp_y, True
  End If
'**** Shadow Tool *************************
  If ToolUsed = 5 Then
   If frmSettings.ChkL1.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 0
   End If
   If frmSettings.ChkL2.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 1
   End If
   If frmSettings.ChkL3.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 2
   End If
   If frmSettings.ChkL4.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 3
   End If
  End If
'*** Add NPC *******************************
  If ToolUsed = 6 Then
   If frmSettings.lstNPC.ListIndex > -1 Then
    If Not frmSettings.lstNPC.ListCount < 1 Then
     engine.Map_NPC_Add temp_x, temp_y, NPC_ListBoxData(frmSettings.lstNPC.ListIndex + 1)
    End If
   End If
  End If
  
'*** Add Exit ******************************
 If ToolUsed = 7 Then
  engine.Map_Exit_Add temp_x, temp_y, frmSettings.txtMap, CLng(frmSettings.txtEX), CLng(frmSettings.txtEX)
 End If
'*** Add Item ******************************
 If ToolUsed = 8 Then
  If frmSettings.lstItem.ListIndex > -1 Then
   engine.Map_Item_Add temp_x, temp_y, Item_ListBoxData(frmSettings.lstItem.ListIndex + 1), CLng(frmSettings.txtIAmount)
  End If
 End If
End If
'******************************************
 
 
 If engine.Input_Mouse_Button_Right_Get Then
  MapModified = True
  If ToolUsed = 1 Then
  engine.Map_Grh_UnSet temp_x, temp_y, frmSettings.cmbLayer.text
  End If
  
  If ToolUsed = 2 Then
  engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text ' It works better if it creates a light and then delites it. I don't know why
  engine.Light_Remove engine.Map_Light_Get(temp_x, temp_y)
  End If

  If ToolUsed = 3 Then
  engine.Particle_Group_Remove engine.Map_Particle_Group_Get(temp_x, temp_y)
  End If
  
  If ToolUsed = 4 Then
  engine.Map_Blocked_Set temp_x, temp_y, False
  End If
  
  If ToolUsed = 5 Then
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 0
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 1
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 2
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 3
  End If
  
  If ToolUsed = 6 Then
   engine.Map_NPC_Remove temp_x, temp_y
  End If
  
  If ToolUsed = 7 Then
    engine.Map_Exit_Remove temp_x, temp_y
  End If
 End If
End If
Exit Sub
Cancel:
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo Cancel
If engine.Input_Mouse_In_View Then
 Dim temp_x As Long
 Dim temp_y As Long
 engine.Input_Mouse_Map_Get temp_x, temp_y
 
 If engine.Input_Mouse_Button_Left_Get Then
 MapModified = True
'********** Grh Tool **********************
  If ToolUsed = 1 Then
   If frmSettings.LstGrh.ListIndex <> -1 Then
    If frmSettings.chkAlpha.value = 1 Then
     engine.Map_Grh_Set temp_x, temp_y, frmSettings.LstGrh.List(frmSettings.LstGrh.ListIndex), frmSettings.cmbLayer.text, True, General_Covert_Degrees_To_Radians(frmSettings.angle) 'If alphablending used
    Else
     engine.Map_Grh_Set temp_x, temp_y, frmSettings.LstGrh.List(frmSettings.LstGrh.ListIndex), frmSettings.cmbLayer.text, , General_Covert_Degrees_To_Radians(frmSettings.angle) 'If not alphablending used
    End If
    If frmSettings.chkBlocked.value = 1 Then
     engine.Map_Blocked_Set temp_x, temp_y, True
    End If
    If frmSettings.chkBlocked.value = 0 Then
     engine.Map_Blocked_Set temp_x, temp_y, False
    End If
   End If
  End If
'********** Light Tool**********************
  If ToolUsed = 2 Then
   If frmSettings.txtLightId = "" Then
    engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text
   End If
   If frmSettings.txtLightId <> "" Then
    engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text, frmSettings.txtLightId.text
   End If
  End If
'********** Particel Tool ******************
  If ToolUsed = 3 Then
   If ParticelCCount <> 0 Then
    Dim temp_list() As Long
    Dim i As Long
    ReDim temp_list(1 To ParticelCCount)
   
    For i = 1 To ParticelCCount
     temp_list(i) = Val(frmSettings.lstParticels.List(i - 1))
    Next i
    If frmSettings.cmbStyle.text = "Star Burst" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 2
    ElseIf frmSettings.cmbStyle.text = "Fountain" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 1
    ElseIf frmSettings.cmbStyle.text = "Insects" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 3
    ElseIf frmSettings.cmbStyle.text = "Water Fall" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 4
    ElseIf frmSettings.cmbStyle.text = "Smoke" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 5
    ElseIf frmSettings.cmbStyle.text = "Fire" Then
     engine.Particle_Group_Create temp_x, temp_y, temp_list(), frmSettings.txtNumParticels.text, 6
    End If
   
   End If
  End If
'********** Block Tool *********************
  If ToolUsed = 4 Then
   engine.Map_Blocked_Set temp_x, temp_y, True
  End If
'**** Shadow Tool *************************
  If ToolUsed = 5 Then
   If frmSettings.ChkL1.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 0
   End If
   If frmSettings.ChkL2.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 1
   End If
   If frmSettings.ChkL3.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 2
   End If
   If frmSettings.ChkL4.value = 1 Then
    engine.Map_Base_Light_Set temp_x, temp_y, RGB(frmSettings.SldBCS.value, frmSettings.SldGCS.value, frmSettings.SldRCS.value), 3
   End If
  End If
'*** Add NPC *******************************
  If ToolUsed = 6 Then
   If frmSettings.lstNPC.ListIndex > -1 Then
    If Not frmSettings.lstNPC.ListCount < 1 Then
     engine.Map_NPC_Add temp_x, temp_y, NPC_ListBoxData(frmSettings.lstNPC.ListIndex + 1)
    End If
   End If
  End If
  
'*** Add Exit ******************************
 If ToolUsed = 7 Then
  engine.Map_Exit_Add temp_x, temp_y, frmSettings.txtMap, CLng(frmSettings.txtEX), CLng(frmSettings.txtEX)
 End If
'*** Add Item ******************************
 If ToolUsed = 8 Then
  If frmSettings.lstItem.ListIndex > -1 Then
   engine.Map_Item_Add temp_x, temp_y, Item_ListBoxData(frmSettings.lstItem.ListIndex + 1), CLng(frmSettings.txtIAmount)
  End If
 End If
End If
'******************************************
 
 
 If engine.Input_Mouse_Button_Right_Get Then
  MapModified = True
  If ToolUsed = 1 Then
  engine.Map_Grh_UnSet temp_x, temp_y, frmSettings.cmbLayer.text
  End If
  
  If ToolUsed = 2 Then
  engine.Light_Create temp_x, temp_y, RGB(frmSettings.SlidB.value, frmSettings.SlidG.value, frmSettings.SlidR.value), frmSettings.cmbRadius.text ' It works better if it creates a light and then delites it. I don't know why
  engine.Light_Remove engine.Map_Light_Get(temp_x, temp_y)
  End If

  If ToolUsed = 3 Then
  engine.Particle_Group_Remove engine.Map_Particle_Group_Get(temp_x, temp_y)
  End If
  
  If ToolUsed = 4 Then
  engine.Map_Blocked_Set temp_x, temp_y, False
  End If
  
  If ToolUsed = 5 Then
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 0
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 1
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 2
    engine.Map_Base_Light_Set temp_x, temp_y, MapBaseLight, 3
  End If
  
  If ToolUsed = 6 Then
   engine.Map_NPC_Remove temp_x, temp_y
  End If
  
  If ToolUsed = 7 Then
    engine.Map_Exit_Remove temp_x, temp_y
  End If
 End If
End If
Exit Sub
Cancel:

End Sub

Private Sub txtMapDesc_Change()
engine.Map_Description_Set txtMapDesc.text
End Sub
