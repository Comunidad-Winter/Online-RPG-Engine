VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "ORE Server"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Frame fmeInfo 
      Caption         =   "Server Information"
      Height          =   4275
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2295
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label lblPlayerCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current Players: 0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame fmeControls 
      Caption         =   "Server Controls"
      Height          =   4275
      Left            =   2400
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      Begin VB.CommandButton cmdworldsave 
         Caption         =   "World save"
         Height          =   375
         Left            =   75
         TabIndex        =   6
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1680
         Top             =   3240
      End
      Begin VB.Timer timStatusUpdate 
         Interval        =   500
         Left            =   840
         Top             =   3660
      End
      Begin VB.CommandButton cmdResetScriptEngine 
         Caption         =   "Reset Script Engine"
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   840
         Width           =   2355
      End
      Begin VB.CommandButton cmdResetServer 
         Caption         =   "Reset Server"
         Height          =   435
         Left            =   75
         TabIndex        =   1
         Top             =   300
         Width           =   2355
      End
      Begin MSWinsockLib.Winsock Slave 
         Index           =   0
         Left            =   345
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   111
      End
      Begin MSWinsockLib.Winsock master 
         Left            =   915
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   111
      End
      Begin OREServer.ctlDirectPlayServer dps 
         Left            =   240
         Top             =   3600
         _ExtentX        =   873
         _ExtentY        =   873
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'OREServer - v0.5.0
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
'
' David Justus - 8/14/2004
'   - Add: cmdworldsave
'   - Add: web - server communication
'
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

Private server_guid As String
Private server_port As Long
Private server_max_players As Long
Private server_name As String
Private server_resource_path As String
Private server_Hide As Boolean
Private server_uptime As Long

Private Sub cmdworldsave_Click()
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    dps.WorldSave
    General_Write_To_TextBox frmcommand.txtlog, "World has been saved!"
End Sub

Private Sub Form_Load()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    'Load configuration file
    If Server_Ini_Load = False Then
        Unload Me
        Exit Sub
    End If
    frmcommand.Show
    frmcommand.txtlog.text = "Server Started"
    'Start server
    If dps.Initialize(server_guid, server_port, server_max_players, server_name, server_resource_path) = False Then
        Unload Me
    End If
    
    'Hide server if needed
    Load_Set_Data
    If server_Hide Then
        Me.Hide
    End If

    Dim i As Long
    For i = 1 To 200
        Load Slave(i)
    Next i
    DoEvents
    master.Listen
    
    'systray
    CreateSystemTrayIcon Me, "ORE 1.0 Server"
End Sub

Private Sub cmdResetScriptEngine_Click()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    If dps.Script_Engine_Reset = False Then
        Unload Me
    End If
End Sub

Private Sub cmdResetServer_Click()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
   'Stop server
    dps.Deinitialize
    
    'Load configuration file
    If Server_Ini_Load = False Then
        Unload Me
    End If
    
    'Start server
    If dps.Initialize(server_guid, server_port, server_max_players, server_name, server_resource_path) = False Then
        Unload Me
    End If
End Sub

Public Sub Reset_Server()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    'Stop server
    dps.Deinitialize
    
    'Load configuration file
    If Server_Ini_Load = False Then
        Unload Me
    End If
    
    'Start server
    If dps.Initialize(server_guid, server_port, server_max_players, server_name, server_resource_path) = False Then
        Unload Me
    End If
End Sub

Public Sub Shutdown_Server()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    dps.Deinitialize
    
    'Load configuration file
    If Server_Ini_Load = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    DeleteSystemTrayIcon Me
    End
End Sub

Private Sub Master_ConnectionRequest(ByVal requestID As Long)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    Dim i As Long
    
    For i = 0 To 200
        If Slave(i).State = sckClosed Then
            Slave(i).Close
            Slave(i).Accept requestID
            Exit Sub
        End If
    Next i
End Sub

Private Sub Slave_DataArrival(Index As Integer, ByVal BytesTotal As Long)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    Dim strData As String
    Dim strGet As String
    Dim spc2 As Long
    Dim page As String
    Dim beginpost As Integer
    Dim LenPost As Long 'Length of the Post data
    Dim PostStuff As String
    Dim Account As clsAccount
    Dim made As Boolean
    
    Set Account = New clsAccount
    
    Slave(Index).GetData strData
    
    strData = ConvertUTF8toASCII(strData)
    
    'For the Post server command
    If Left(strData, 4) = "POST" Then
        beginpost = InStr(1, strData, vbCrLf + vbCrLf) + 4
        LenPost = Len(strData) - beginpost + 1
        PostStuff = Mid(strData, beginpost, LenPost)
        
        Account.Account_Initialize server_resource_path & "\accounts"
        'Password does not match
        If Not Web_Field_Read_Values(Web_Field_Read(2, PostStuff)) = Web_Field_Read_Values(Web_Field_Read(3, PostStuff)) Then
            Slave(Index).SendData Error2Data(dps.Player_Count, 1, 1)
        End If
        made = Account.Account_Create(Web_Field_Read_Values(Web_Field_Read(1, PostStuff)), Web_Field_Read_Values(Web_Field_Read(2, PostStuff)), Web_Field_Read_Values(Web_Field_Read(5, PostStuff)), Web_Field_Read_Values(Web_Field_Read(6, PostStuff)), Web_Field_Read_Values(Web_Field_Read(4, PostStuff)))
        'error
        If made = False Then
            Slave(Index).SendData Error1Data(dps.Player_Count, 1, 1)
        Else
            frmcommand.txtlog.text = frmcommand.txtlog.text & vbNewLine & "New Account"
            Slave(Index).SendData Welcome(dps.Player_Count, 1, 1)
        End If
    End If
    
    'For the Get server command
    If Mid(strData, 1, 3) = "GET" Then
        strGet = InStr(strData, "GET ")
        spc2 = InStr(strGet + 5, strData, " ")
        page = Trim(Mid(strData, strGet + 5, spc2 - (strGet + 4)))
        If Right(page, 1) = "/" Then page = Left(page, Len(page) - 1)
        If page = "/" Then page = "index.html"
        If page = "" Then page = "index.html"
        If page = "index.html" Then
            Slave(Index).SendData IndexData(dps.Player_Count, 1, 1)
        End If
        'send page
        If LCase(page) = "account.html" Then
            Slave(Index).SendData AccountData(dps.Player_Count, 1, 1)
        End If
    End If
End Sub

Private Sub Slave_SendComplete(Index As Integer)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    Slave(Index).Close
End Sub

Private Sub Timer1_Timer()
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    server_uptime = server_uptime + 1
End Sub

Private Sub timStatusUpdate_Timer()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    lblStatus.Caption = "Status: "
    
    If dps.ServerStatus = s_s_listening Then
        lblStatus.Caption = lblStatus.Caption & "Listening"
    End If

    If dps.ServerStatus = s_s_shutting_down Then
        lblStatus.Caption = lblStatus.Caption & "Shutting down ..."
    End If

    If dps.ServerStatus = s_s_closed Then
        lblStatus.Caption = lblStatus.Caption & "Closed ..."
    End If
End Sub

Private Sub dps_ServerConnectionAdded(player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    lblPlayerCount.Caption = "Current Players: " & dps.Player_Count
End Sub

Private Sub dps_ServerConnectionRemoved(player_id As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    lblPlayerCount.Caption = "Current Players: " & dps.Player_Count
End Sub

Private Function Server_Ini_Load() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    Dim ini_path As String
    ini_path = App.Path & "\" & "server.ini"
    
    If General_File_Exists(ini_path, vbNormal) Then
        server_guid = General_Var_Get(ini_path, "GENERAL", "guid")
        server_port = Val(General_Var_Get(ini_path, "GENERAL", "port"))
        server_max_players = Val(General_Var_Get(ini_path, "GENERAL", "max_players"))
        server_name = General_Var_Get(ini_path, "GENERAL", "name")
        server_resource_path = App.Path & General_Var_Get(ini_path, "GENERAL", "resource_path")
        server_Hide = CBool(General_Var_Get(ini_path, "GENERAL", "Hide"))
        Server_Ini_Load = True
    Else
        dps.Log_Event "frmMain", "Server_Ini_Load", "Error - server.ini not found: " & ini_path
    End If
    
End Function

Public Function Get_Feild(text As String, feild As String)
'**************************************************************
'Author: David Justus
'Last Modify Date: 8/14/2004
'
'**************************************************************
    Dim strTemp As String
    Dim j As Long
    Dim i As Long
    
    For i = 1 To Len(text)
        If Mid(text, i, Len(feild)) = feild Then
            strTemp = Right(text, Len(text) - Len(feild) - 1)
            For j = 1 To Len(strTemp)
                If Mid(strTemp, j, 1) = "&" Then Exit Function
                Get_Feild = Get_Feild + Mid(strTemp, j, 1)
            Next j
            Exit Function
        End If
    Next i
End Function
