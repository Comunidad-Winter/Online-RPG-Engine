VERSION 5.00
Begin VB.Form frmToolBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   3870
   ClientLeft      =   9930
   ClientTop       =   420
   ClientWidth     =   495
   ControlBox      =   0   'False
   Icon            =   "frmToolBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   33
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image17 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":0CCE
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":1910
      Top             =   4200
      Width           =   480
   End
   Begin VB.Image imgItem 
      Height          =   495
      Left            =   0
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":2552
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image13 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":3194
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image imgExitTool 
      Height          =   495
      Left            =   0
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":3DD6
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":4A18
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgNPC 
      Height          =   495
      Left            =   0
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":565A
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":629C
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image imgColorStuff 
      Height          =   495
      Left            =   0
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image imgBlock 
      Height          =   495
      Left            =   0
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgParticel 
      Height          =   495
      Left            =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgGrh 
      Height          =   495
      Left            =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Image imgLight 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":6EDE
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":7B20
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":8762
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":93A4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":9FE6
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":AC28
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1080
      Picture         =   "frmToolBox.frx":B86A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "frmToolBox.frx":C4AC
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'frmToolBox.frm - ORE Map Editor
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

Private Sub Form_Load()
General_Form_On_Top_Set Me, True
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image4.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
End Sub


Private Sub imgBlock_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image8.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
Select_Tool (4)
imgItem.Picture = Image16.Picture
End Sub

Private Sub imgColorStuff_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image11.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
Select_Tool (5)
End Sub

Private Sub imgExitTool_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image14.Picture
imgItem.Picture = Image16.Picture
Select_Tool (7)
End Sub

Private Sub imgGrh_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image4.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
Select_Tool (1)
End Sub

Private Sub imgItem_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image17.Picture
Select_Tool (8)
End Sub

Private Sub imgLight_Click()
imgLight.Picture = Image2.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
Select_Tool (2)
End Sub

Private Sub imgNPC_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image5.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image12.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
Select_Tool (6)
End Sub

Private Sub imgParticel_Click()
imgLight.Picture = Image1.Picture
imgGrh.Picture = Image3.Picture
imgParticel.Picture = Image6.Picture
imgBlock.Picture = Image7.Picture
imgColorStuff.Picture = Image10.Picture
imgNPC.Picture = Image9.Picture
imgExitTool.Picture = Image13.Picture
imgItem.Picture = Image16.Picture
Select_Tool (3)
End Sub
