VERSION 5.00
Begin VB.Form frmNewItem 
   Caption         =   "New Item"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   2895
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox item_grh_index 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox grh_render_pb 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   1560
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox item_name 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "GRH Nro:"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Item Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmNewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If item_name.text <> "" And item_grh_index.text <> "" Then
        Dim item_ini_path As String
        Dim item_str_tmp As String
        Dim str_tmp As String
        Dim t_count As Long
        
        item_ini_path = resource_path & "\scripts\item.ini"
        t_count = Val(General_Var_Get(item_ini_path, "GENERAL", "item_count"))
        item_str_tmp = "ITEM" & CStr(t_count + 1)
        General_Var_Write item_ini_path, item_str_tmp, "item_name", item_name.text
        General_Var_Write item_ini_path, item_str_tmp, "item_grh_index", item_grh_index.text
        General_Var_Write item_ini_path, "GENERAL", "item_count", CStr(t_count + 1)
        frmMain.OBJList.Clear
    End If
End Sub

Private Sub Command2_Click()
    If frmMain.GrhViewerMnuChk.Checked Then
        General_Form_On_Top_Set frmGrhViewer, True
    End If
    If frmMain.MiniMapMnuChk.Checked Then
        General_Form_On_Top_Set frmMap, True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Engine.Grh_Add_GrhList_To_ListBox List1
End Sub

Private Sub List1_Click()
    grh_render_pb.Cls
    Engine.Grh_Render_To_Hdc List1.text, grh_render_pb.hdc, 1, 1, True
    item_grh_index.text = List1.text
End Sub
