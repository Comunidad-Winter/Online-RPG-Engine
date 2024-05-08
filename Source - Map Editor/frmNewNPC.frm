VERSION 5.00
Begin VB.Form frmNewNPC 
   Caption         =   "New NPC"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5295
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox npc_char_data_index 
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   945
      Width           =   2550
   End
   Begin VB.Frame Frame1 
      Caption         =   "Char"
      Height          =   2655
      Left            =   2910
      TabIndex        =   7
      Top             =   135
      Width           =   2250
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   2340
         Left            =   105
         ScaleHeight     =   2280
         ScaleWidth      =   1995
         TabIndex        =   8
         Top             =   225
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   2355
      Width           =   2535
   End
   Begin VB.TextBox npc_name 
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   315
      Width           =   2535
   End
   Begin VB.TextBox npc_AI_script 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   1515
      Width           =   2535
   End
   Begin VB.CommandButton npc_save_btn 
      Caption         =   "Save NPC"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   1875
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "NPC name:"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   75
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "NPC char data index:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   675
      Width           =   1515
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "NPC AI script:"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   1275
      Width           =   990
   End
End
Attribute VB_Name = "frmNewNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    If MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    Unload Me
End Sub

Private Sub npc_save_btn_Click()
    If npc_name.text <> "" And npc_char_data_index.text <> "" And npc_AI_script.text <> "" Then
        Dim npc_ini_path As String
        Dim npc_str_tmp As String
        Dim str_tmp As String
        Dim t_count As Long
        
        npc_ini_path = resource_path & "\scripts\npc.ini"
        npc_str_tmp = "NPC" & CStr(frmMain.NPCList.ListCount + 1)
        t_count = Val(General_Var_Get(npc_ini_path, "GENERAL", "npc_count"))
        
        General_Var_Write npc_ini_path, npc_str_tmp, "npc_name", npc_name.text
        General_Var_Write npc_ini_path, npc_str_tmp, "npc_char_data_index", npc_char_data_index.text
        General_Var_Write npc_ini_path, npc_str_tmp, "npc_ai_script", npc_AI_script.text
        General_Var_Write npc_ini_path, "GENERAL", "npc_count", CStr(t_count + 1)
        
        frmMain.NPCList.Clear
        Load_NPC_Data frmMain.NPCList
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Static Heading As Long
    Heading = Heading + 1
    If Heading > 8 Then Heading = 1
    If npc_char_data_index.ListIndex <> -1 Then
        Picture1.Cls
        Engine.Grh_Render_To_Hdc Engine.Char_Data_Grh_Index_Get(CLng(npc_char_data_index.text), 1, Heading), _
                                                Picture1.hdc, 50, 50, True
    End If
End Sub
