VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsGump.cls - ORE Gump Handling Object - v0.5.0
'
'Handles the gump server side
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/200X - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 8/14/2004
'   -Change: Removed some stuff never used
'
'David Justus (big.david@txun.net) - 3/23/2004
'   - First Release, converted from some C# code from Merlin
'*****************************************************************
Option Explicit

Public Enum button_type
    page = 0
    reply = 2
End Enum

Private Type Gump_Button
    X As Integer
    Y As Integer
    id1 As Integer
    id2 As Integer
    button_id As Long
    button_type As button_type
End Type

Private Type Gump_Check
    X As Integer
    Y As Integer
    id1 As Integer
    id2 As Integer
    initialstate As Boolean
    Switch_id As Boolean
End Type

Private Type Gump_Image
    X As Integer
    Y As Integer
    gump_id As Long
End Type

Private Type Gump_Image_Tiled
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    gump_id As Long
End Type

Private Type Gump_Item
    X As Integer
    Y As Integer
    item_id As Integer
End Type

Private Type Gump_label
    X As Integer
    Y As Integer
    hue As Integer
    text As String
End Type

Private Type Gump_label_Cropped
    X As Integer
    Y As Integer
    hue As Integer
    width As Integer
    height As Integer
    text As String
End Type

Private Type Gump_page
    id As Integer
End Type

Private Type Gump_radio
    X As Integer
    Y As Integer
    id1 As Integer
    id2 As Integer
    initialstate As Boolean
    switchid As Integer
End Type

Private Type Gump_Text
    entryID As Integer
    height As Integer
    width As Integer
    hue As Integer
    initial_text As String
    X As Integer
    Y As Integer
End Type

Private Type Gump_Alpharegion
    height As Integer
    width As Integer
    X As Integer
    Y As Integer
End Type

Private Type Gump_Background
    GumpID As Long
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
End Type

Private Type gump
    closable As Boolean
    Disposable As Boolean
    Dragable As Boolean
    resizable As Boolean
    strings() As String
    serial As Integer
    typeid As Integer
    X As Integer
    Y As Integer
End Type
    
Private dps As ctlDirectPlayServer
Private script_engine As clsScriptEngine
Private map_list As clsList
Private player_list As clsList
Private npc_list As clsList
Private speech_list As clsList

Private button(0 To 100) As Gump_Button
Private button_index As Integer
Private check(0 To 100) As Gump_Check
Private check_index As Integer
Private image(0 To 100) As Gump_Image
Private image_index As Integer
Private imageT(0 To 100) As Gump_Image_Tiled
Private image_indexT As Integer
Private Item(0 To 100) As Gump_Item
Private item_index As Integer
Private label(0 To 100) As Gump_label
Private label_index As String
Private labelC(0 To 100) As Gump_label_Cropped
Private labelC_index As String
Private pages(0 To 100) As Gump_page
Private page_index As Integer
Private radio(0 To 100) As Gump_radio
Private radio_index As Integer
Private text(0 To 100) As Gump_Text
Private text_index As Integer
Private Alpharegion(0 To 100) As Gump_Alpharegion
Private Alpharegion_index As Integer
Private background(0 To 100) As Gump_Background
Private background_index As Integer

Private gump_list As clsList
Dim gump As gump

Public Function Initialize(ByRef s_dps As ctlDirectPlayServer, ByVal s_gump_list As clsList) As Boolean
'**************************************************************
'Author: David justus
'Last Modify Date: 3/23/2004
'Sets stuff up
'**************************************************************
    Set dps = s_dps
    Set gump_list = s_gump_list
    
    button_index = 0
    check_index = 0
    image_index = 0
    image_indexT = 0
    item_index = 0
    label_index = 0
    labelC_index = 0
    page_index = 0
    radio_index = 0
    text_index = 0
    Alpharegion_index = 0
    background_index = 0
    
    Initialize = True
End Function

Public Function Clear_Gump()
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Sets indexes to 0 clearing gump thE lazy way!
'**************************************************************
    button_index = 0
    check_index = 0
    image_index = 0
    image_indexT = 0
    item_index = 0
    label_index = 0
    labelC_index = 0
    page_index = 0
    radio_index = 0
    text_index = 0
    Alpharegion_index = 0
    background_index = 0
    Clear_Gump = True
End Function

Public Function Deinitialize()
'**************************************************************
'Author: David justus
'Last Modify Date: 3/23/2004
'Sets stuff up
'**************************************************************
    Set dps = Nothing
    Set gump_list = Nothing
End Function

Public Function Create_Gump(ByVal X As Long, ByVal Y As Long, ByVal closable As Boolean, ByVal Disposable As Boolean, ByVal Dragable As Boolean, ByVal resizable As Boolean, ByVal serial As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Creates the a GUMP page
'**************************************************************
    gump.X = X
    gump.Y = Y
    gump.closable = closable
    gump.Disposable = Disposable
    gump.Dragable = Dragable
    gump.resizable = resizable
    gump.serial = serial
    Create_Gump = True
End Function

Public Function Compile_Gump() As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Sets stuff up
'**************************************************************
    Dim line As String
    Dim i As Integer
    
    On Error Resume Next
    
    If gump.Dragable = True Then line = line & "{ nomove }"
    If gump.closable = True Then line = line & "{ noclose }"
    If gump.Disposable = True Then line = line & "{ nodispose }"
    If gump.resizable = True Then line = line & "{ noresize }"
    line = line & "{ " & gump.serial & " }"
    
    For i = 1 To button_index
        line = line & Compile_Gump_Button(i)
    Next i
    
    For i = 1 To check_index
        line = line & Compile_Gump_Check(i)
    Next i
    
    For i = 1 To image_index
        line = line & Compile_Gump_Image(i)
    Next i
    
    For i = 1 To image_indexT
        line = line & Compile_Gump_ImageT(i)
    Next i
    
    For i = 1 To item_index
        line = line & Compile_Gump_Item(i)
    Next i
    
    For i = 1 To label_index
        line = line & Compile_Gump_Label(i)
    Next i
    
    For i = 1 To labelC_index
        line = line & Compile_Gump_LabelC(i)
    Next i
    
    For i = 1 To radio_index
        line = line & Compile_Gump_Radio(i)
    Next i
    
    For i = 1 To text_index
        line = line & Compile_Gump_Text(i)
    Next i
    
    For i = 1 To Alpharegion_index
        line = line & Compile_Gump_Alpharegion(i)
    Next i
    
    For i = 1 To background_index
        line = line & Compile_Gump_Background(i)
    Next i
    
    Compile_Gump = line
End Function

Public Function Add_Gump_Button(ByVal X As Long, ByVal Y As Long, ByVal id1 As Long, ByVal id2 As Long, ByVal button_types As button_type) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds a button to the current gump
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Dim Index As Integer
    
    Index = button_index + 1
    button(Index).button_id = Index
    button(Index).button_type = button_types
    button(Index).id1 = id1
    button(Index).id2 = id2
    button(Index).X = X
    button(Index).Y = Y
    button_index = Index
End Function

Public Function Compile_Gump_Button(ByVal button_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for a button
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Compile_Gump_Button = "{button/" & button(button_id).X & "/" & button(button_id).Y & "/" & button(button_id).id1 & "/" & button(button_id).id2 & "/" & button(button_id).button_id & "/" & button(button_id).button_type & "}"
End Function

Public Function Add_Gump_Check(ByVal X As Long, ByVal Y As Long, ByVal id1 As Long, ByVal id2 As Long, ByVal initialstate As Boolean, ByVal Switch_id As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds a Checkbox to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = check_index + 1
    check(Index).X = X
    check(Index).Y = Y
    check(Index).id1 = id1
    check(Index).id2 = id2
    check(Index).initialstate = initialstate
    check(Index).Switch_id = Switch_id
    check_index = Index
End Function

Public Function Compile_Gump_Check(ByVal check_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'compiles a string for a check
'**************************************************************
    Compile_Gump_Check = "{checkbox/" & check(check_id).X & "/" & check(check_id).Y & "/" & check(check_id).id1 & "/" & check(check_id).id2 & "/" & check(check_id).initialstate & "/" & check(check_id).Switch_id & "/}"
End Function

Public Function Add_Gump_Image(ByVal X As Long, ByVal Y As Long, ByVal gump_id As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an Image to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = image_index + 1
    image(Index).X = X
    image(Index).Y = Y
    image(Index).gump_id = gump_id
    
    image_index = Index
End Function

Public Function Compile_Gump_Image(ByVal image_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an Image
'**************************************************************
    Compile_Gump_Image = "{gumppic/" & image(image_id).X & "/" & image(image_id).Y & "/" & image(image_id).gump_id & "/}"
End Function

Public Function Add_Gump_ImageT(ByVal X As Long, ByVal Y As Long, ByVal gump_id As Long, ByVal height As Long, ByVal width As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an Image to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = image_indexT + 1
    imageT(Index).X = X
    imageT(Index).Y = Y
    imageT(Index).height = height
    imageT(Index).width = width
    imageT(Index).gump_id = gump_id
    
    image_indexT = Index
End Function

Public Function Compile_Gump_ImageT(ByVal image_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an Image
'**************************************************************
    Compile_Gump_ImageT = "{gumppictiled/" & imageT(image_id).X & "/" & imageT(image_id).Y & "/" & imageT(image_id).height & "/" & imageT(image_id).width & "/" & imageT(image_id).gump_id & "/}"
End Function

Public Function Add_Gump_Item(ByVal X As Long, ByVal Y As Long, ByVal itemid As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an Item to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = item_index + 1
    Item(Index).X = X
    Item(Index).Y = Y
    Item(Index).item_id = itemid
    item_index = Index
End Function

Public Function Compile_Gump_Item(ByVal item_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an Item
'**************************************************************
    Compile_Gump_Item = "{tilepic/" & Item(item_id).X & "/" & Item(item_id).Y & "/" & Item(item_id).item_id & "/}"
End Function

Public Function Add_Gump_Label(ByVal X As Long, ByVal Y As Long, ByVal hue As Long, ByVal text As String) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an label to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = label_index + 1
    label(Index).X = X
    label(Index).Y = Y
    label(Index).hue = hue
    label(Index).text = text
    label_index = Index
End Function

Public Function Compile_Gump_Label(ByVal label_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an label
'**************************************************************
    Compile_Gump_Label = "{text/" & label(label_id).X & "/" & label(label_id).Y & "/" & label(label_id).hue & "/" & label(label_id).text & "/}"
End Function

Public Function Add_Gump_LabelC(ByVal X As Long, ByVal Y As Long, ByVal hue As Long, ByVal height As Long, ByVal width As Long, ByVal text As String) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an labelC to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = labelC_index + 1
    labelC(Index).X = X
    labelC(Index).Y = Y
    labelC(Index).hue = hue
    labelC(Index).height = height
    labelC(Index).width = width
    labelC(Index).text = text
    labelC_index = Index
End Function

Public Function Compile_Gump_LabelC(ByVal labelc_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an labelC
'**************************************************************
    Compile_Gump_LabelC = "{croppedtext/" & labelC(labelc_id).X & "/" & labelC(labelc_id).Y & "/" & labelC(labelc_id).hue & "/" & labelC(labelc_id).height & "/" & labelC(labelc_id).width & "/" & labelC(labelc_id).text & "/}"
End Function

Public Function Add_Gump_Radio(ByVal X As Long, ByVal Y As Long, ByVal id1 As Long, ByVal id2 As Long, ByVal initialstate As Boolean, ByVal switchid As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an radio to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = radio_index + 1
    radio(Index).X = X
    radio(Index).Y = Y
    radio(Index).id1 = id1
    radio(Index).id2 = id2
    radio(Index).initialstate = initialstate
    radio(Index).switchid = switchid
    radio_index = Index
End Function

Public Function Compile_Gump_Radio(ByVal radio_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a radio for an labelC
'**************************************************************
    Compile_Gump_Radio = "{radio/" & radio(radio_id).X & "/" & radio(radio_id).Y & "/" & radio(radio_id).id1 & "/" & radio(radio_id).id2 & "/" & radio(radio_id).initialstate & "/" & radio(radio_id).switchid & "/}"
End Function

Public Function Add_Gump_Text(ByVal X As Long, ByVal Y As Long, ByVal hue As Long, ByVal height As Long, ByVal width As Long, ByVal itext As String) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an text to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = text_index + 1
    text(Index).X = X
    text(Index).Y = Y
    text(Index).hue = hue
    text(Index).height = height
    text(Index).width = width
    text(Index).initial_text = itext
    text_index = Index
End Function

Public Function Compile_Gump_Text(ByVal text_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an labelC
'**************************************************************
    Compile_Gump_Text = "{textentry/" & text(text_id).X & "/" & text(text_id).Y & "/" & text(text_id).hue & "/" & text(text_id).height & "/" & text(text_id).width & "/" & text(text_id).initial_text & "/}"
End Function

Public Function Add_Gump_Alpharegion(ByVal X As Long, ByVal Y As Long, ByVal height As Long, ByVal width As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an Alpharegion to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = Alpharegion_index + 1
    Alpharegion(Index).X = X
    Alpharegion(Index).Y = Y
    Alpharegion(Index).height = height
    Alpharegion(Index).width = width
    Alpharegion_index = Index
End Function

Public Function Compile_Gump_Alpharegion(ByVal Alpharegion_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an Alpharegion
'**************************************************************
    Compile_Gump_Alpharegion = "{checkertrans/" & Alpharegion(Alpharegion_id).X & "/" & Alpharegion(Alpharegion_id).Y & "/" & Alpharegion(Alpharegion_id).height & "/" & Alpharegion(Alpharegion_id).width & "/}"
End Function

Public Function Add_Gump_Background(ByVal X As Long, ByVal Y As Long, ByVal height As Long, ByVal width As Long, ByVal GumpID As Long) As Boolean
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Adds an Background to the current gump
'**************************************************************
    Dim Index As Integer
    
    Index = background_index + 1
    background(Index).X = X
    background(Index).Y = Y
    background(Index).height = height
    background(Index).width = width
    background(Index).GumpID = GumpID
    background_index = Index
End Function

Public Function Compile_Gump_Background(ByVal Background_id As Long) As String
'**************************************************************
'Author: David Justus
'Last Modify Date: 3/23/2004
'Compiles a string for an Background
'**************************************************************
    Compile_Gump_Background = "{resizepic/" & background(Background_id).X & "/" & background(Background_id).Y & "/" & background(Background_id).height & "/" & background(Background_id).width & "/" & background(Background_id).GumpID & "/}"
End Function
