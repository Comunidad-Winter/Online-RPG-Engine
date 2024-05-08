VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ORE Server"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
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
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   2355
      End
      Begin OREServer.ctlDirectPlayServer dps 
         Left            =   180
         Top             =   3600
         _ExtentX        =   873
         _ExtentY        =   979
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
'Aaron Perkins(aaron@baronsoft.com) - 8/04/2003
'   - First Release
'*****************************************************************
Option Explicit

Private server_guid As String
Private server_port As Long
Private server_max_players As Long
Private server_name As String
Private server_resource_path As String
Private server_hide As Boolean

Private Sub Form_Load()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/8/2003
'
'**************************************************************
    'Load configuration file
    If Server_Ini_Load = False Then
        Unload Me
    End If
    
    'Start server
    If dps.Initialize(server_guid, server_port, server_max_players, server_name, server_resource_path) = False Then
        Unload Me
    End If
    
    'Hide server if needed
    If server_hide Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/29/2003
'
'**************************************************************
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
        server_hide = CBool(General_Var_Get(ini_path, "GENERAL", "hide"))
        Server_Ini_Load = True
    Else
        dps.Log_Event "frmMain", "Server_Ini_Load", "Error - server.ini not found: " & ini_path
    End If
End Function
