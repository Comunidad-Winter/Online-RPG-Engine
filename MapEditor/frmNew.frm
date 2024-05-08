VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Map"
   ClientHeight    =   2070
   ClientLeft      =   3945
   ClientTop       =   2835
   ClientWidth     =   1965
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "50"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "50"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Width"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmNew.frm - ORE Map Editor
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
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/17/2003
'   -Second official/unofficial release
'
'Aaron Perkins(aaron@baronsoft.com) - 5/12/2003
'   -First offical release
'
'Fredrik Alexandersson (fredrik@oraklet.zzn.com) - 5/12/2003
'   -Last unoffical release
'
'*****************************************************************
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
General_Form_On_Top_Set frmNew, True
End Sub

Private Sub OKButton_Click()
On Local Error GoTo Cancel
frmEditor.engine.Map_Create txtHeight, txtWidth
frmEditor.engine.Map_Base_Light_Fill RGB(190, 190, 190)
frmEditor.txtMapDesc = "Untitled Map"
Unload Me
Exit Sub
Cancel:
MsgBox "Error:" & Error
End Sub

Private Sub txtHeight_Change()
txtHeight.text = Val(txtHeight.text)
End Sub

Private Sub txtWidth_Change()
txtWidth.text = Val(txtWidth.text)
End Sub
