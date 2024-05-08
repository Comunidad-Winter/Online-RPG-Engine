Attribute VB_Name = "modDeclares"
Option Explicit

'The engine object (used to load grh list, and load and display grhs)
Public engine As New clsTileEngineX

Public fso As New FileSystemObject, f As Scripting.TextStream
Public Tree_Cur_Grh As Long

Public graphics_path As String
Public scripts_path As String

Public Modified As Boolean

Public NodeCount As Long

'We use different separators from the ones used in the script files because ">"
'is used when setting the layer and may cause errors or confusion
Public Const Separator = "\"
Public Const Child = "\\"
Public Const ChildrenEnd = "\EOF"
