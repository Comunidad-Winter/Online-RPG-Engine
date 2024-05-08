Attribute VB_Name = "modMapEditor"
Option Explicit
Private Const PATH_GRAPHICS = "\graphics"
Private Const PATH_MAPS = "\maps"
Private Const PATH_SOUNDS = "\sounds"
Private Const PATH_SCRIPTS = "\scripts"
Public resource_path As String
Public NPC_ListBoxData() As Long
Public Item_ListBoxData() As Long

Public Function NPC_Ini_Number_Get() As Long
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'
'**************************************************************
    NPC_Ini_Number_Get = CLng(General_Var_Get(App.Path & resource_path & PATH_SCRIPTS & "\npc.ini", "GENERAL", "npc_count"))
End Function
Public Function NPC_Ini_Name_Get(ByVal s_npc_data_index As Long) As String
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'
'**************************************************************
    NPC_Ini_Name_Get = General_Var_Get(App.Path & resource_path & PATH_SCRIPTS & "\npc.ini", "NPC" & s_npc_data_index, "npc_name")
End Function
Public Sub NPC_Add_List_To_Settings()
'*****************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/17/2003
'Add NPC list to list box
'*****************************************************************
Dim NPC As Long
Dim NPCNum As Long
On Local Error GoTo Cancel
    ReDim Preserve NPC_ListBoxData(0 To NPC_Ini_Number_Get())
    For NPC = 1 To NPC_Ini_Number_Get()
        If Not NPC_Ini_Name_Get(NPC) = "" Then
            NPCNum = NPCNum + 1
            frmSettings.lstNPC.AddItem NPC_Ini_Name_Get(NPC)
            NPC_ListBoxData(NPCNum) = NPC
            End If
    Next NPC
Exit Sub
Cancel:
MsgBox "Error loading NPC list.", vbCritical
End Sub
Public Sub Item_Add_List_To_Settings()
'*****************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/17/2003
'Add Iem list to list box
'*****************************************************************
Dim Item As Long
Dim ItemNum As Long
On Local Error GoTo Cancel
    ReDim Preserve Item_ListBoxData(0 To Item_Ini_Number_Get())
    For Item = 1 To Item_Ini_Number_Get()
        If Not Item_Ini_Name_Get(Item) = "" Then
            ItemNum = ItemNum + 1
            frmSettings.lstItem.AddItem Item_Ini_Name_Get(Item)
            Item_ListBoxData(ItemNum) = Item
            End If
    Next Item
Exit Sub
Cancel:
MsgBox "Error loading Item list.", vbCritical
End Sub
Public Function Item_Ini_Number_Get() As Long
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Item_Ini_Number_Get = CLng(General_Var_Get(App.Path & resource_path & PATH_SCRIPTS & "\item.ini", "GENERAL", "item_count"))
End Function
Public Function Item_Ini_Name_Get(ByVal s_item_data_index As Long) As String
'**************************************************************
'Author: Fredrik Alexandersson
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Item_Ini_Name_Get = General_Var_Get(App.Path & resource_path & PATH_SCRIPTS & "\item.ini", "ITEM" & s_item_data_index, "item_name")
End Function
Public Sub Select_Tool(ByVal tool As Integer)
If tool = 1 Then
    frmSettings.tabGrh.Visible = True
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 1
ElseIf tool = 2 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = True
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 2
ElseIf tool = 3 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = True
    frmSettings.tabBlock.Visible = False
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 3
ElseIf tool = 4 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = True
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 4
ElseIf tool = 5 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabShadow.Visible = True
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 5
ElseIf tool = 6 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = True
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 6
ElseIf tool = 7 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabExit.Visible = True
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabItem.Visible = False
    frmEditor.ToolUsed = 7
ElseIf tool = 8 Then
    frmSettings.tabGrh.Visible = False
    frmSettings.tabLight.Visible = False
    frmSettings.tabParticle.Visible = False
    frmSettings.tabBlock.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabShadow.Visible = False
    frmSettings.tabNPC.Visible = False
    frmSettings.tabExit.Visible = False
    frmSettings.tabItem.Visible = True
    frmEditor.ToolUsed = 8

End If
End Sub
