Attribute VB_Name = "modCompression"
'*****************************************************************
'modCompression.bas - v1.0.0
'
'All methods to handle resource files
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
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 10/13/2004
'   - First Release
'*****************************************************************
Option Explicit

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    intNumFiles As Integer              'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    grh
    MIDI
    MP3
    WAV
    Scripts
    Patch
End Enum

Const GRAPHIC_PATH = "\Graphics\"
Const MIDI_PATH = "\Midis\"
Const MP3_PATH = "\Mp3\"
Const WAV_PATH = "\Wavs\"
Const SCRIPT_PATH = "\Scripts\"
Const PATCH_PATH = "\Patches\"
Const OUTPUT_PATH = "\Output\"

Public Type file_handle
    handle As Integer
    file_name As String
End Type

'Arrays used to keep all file handles so they keep locked and safe
Public GRH_Handles() As file_handle
Public MP3_Handles() As file_handle
Public MIDI_Handles() As file_handle
Public WAV_Handles() As file_handle
Public Scripts_Handles() As file_handle

'True if a new file of the stated format was downloaded in a patch
Public GraphicsDSO As Boolean
Public MP3DSO As Boolean
Public MIDIDSO As Boolean
Public WAVDSO As Boolean
Public ScriptsDSO As Boolean

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long

Public Sub Compress_Data(ByRef data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Static compression_rate As Single
    Dim LoopC As Long
    
    If compression_rate = 0 Then compression_rate = 0.05
    
    Dimensions = UBound(data)
    
    DimBuffer = Dimensions * compression_rate
    
    ReDim BufTemp(DimBuffer)
    
    compress BufTemp(0), DimBuffer, data(0), Dimensions
    
    'Check if there was data loss
    ReDim BufTemp2(Dimensions)
    
    uncompress BufTemp2(0), Dimensions, BufTemp(0), UBound(BufTemp) + 1
    
    For LoopC = 0 To UBound(data)
        If data(LoopC) <> BufTemp2(LoopC) Then
            'Clear memory
            Erase BufTemp
            Erase BufTemp2
            
            'If we have reached 1, then just copy the data
            If compression_rate < 1 Then
                'Increase compression rate
                compression_rate = compression_rate + 0.05
                'Try again
                Compress_Data data
            End If
            
            'Reset compression rate and exit
            compression_rate = 0.05
            Exit Sub
        End If
    Next LoopC
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    Erase BufTemp2
    
    'Encrypt the first byte of the compressed data for extra security
    data(0) = data(0) Xor 12
End Sub

Public Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor 12
    
    uncompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Sub Encrypt_File_Header(ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 12345
        .lngFileSize = .lngFileSize Xor 1234567890
    End With
End Sub

Public Sub Encrypt_Info_Header(ByRef InfoHead As INFOHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    Dim EncryptedFileName As String
    Dim LoopC As Long
    
    For LoopC = 1 To Len(InfoHead.strFileName)
        If LoopC Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr(Asc(Mid(InfoHead.strFileName, LoopC, 1)) Xor 123)
        Else
            EncryptedFileName = EncryptedFileName & Chr(Asc(Mid(InfoHead.strFileName, LoopC, 1)) Xor 12)
        End If
    Next LoopC
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 1234567890
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 1234567890
        .lngFileStart = .lngFileStart Xor 123456789
        .strFileName = EncryptedFileName
    End With
End Sub

Public Function Extract_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByRef ResourcePrgbar As ProgressBar, ByRef GeneralPrgBar As ProgressBar, ByRef GeneralLbl As Label, Optional ByVal UseOutputFolder As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************
    Dim LoopC As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    Dim file_handler_list() As file_handle
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case grh
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Graphics.ORE"
            Else
                SourceFilePath = resource_path & "\Graphics.ORE"
            End If
            OutputFilePath = resource_path & GRAPHIC_PATH
            
        Case MIDI
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Low-Def Music.ORE"
            Else
                SourceFilePath = resource_path & "\Low-Def Music.ORE"
            End If
            OutputFilePath = resource_path & MIDI_PATH
        
        Case MP3
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Hi-Def Music.ORE"
            Else
                SourceFilePath = resource_path & "\Hi-Def Music.ORE"
            End If
            OutputFilePath = resource_path & MP3_PATH
        
        Case WAV
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Sounds.ORE"
            Else
                SourceFilePath = resource_path & "\Sounds.ORE"
            End If
            OutputFilePath = resource_path & WAV_PATH
        
        Case Scripts
            If UseOutputFolder Then
                SourceFilePath = resource_path & OUTPUT_PATH & "Scripts.ORE"
            Else
                SourceFilePath = resource_path & "\Scripts.ORE"
            End If
            OutputFilePath = resource_path & SCRIPT_PATH
        
        Case Else
            Exit Function
    End Select
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For LoopC = 0 To UBound(InfoHead)
        'Desencrypt each Info Header before accessing the data
        Encrypt_Info_Header InfoHead(LoopC)
        RequiredSpace = RequiredSpace + InfoHead(LoopC).lngFileSizeUncompressed
    Next LoopC
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
        Erase InfoHead
        Close SourceFile
        MsgBox "There is not enough drive space to extract the compressed files.", , "Error"
        Exit Function
    End If
    
    'Size the file handler array
    ReDim file_handler_list(UBound(InfoHead))
    
    If Not ResourcePrgbar Is Nothing Then ResourcePrgbar.Max = FileHead.intNumFiles
    
    'Extract all of the files from the binary file
    For LoopC = 0 To UBound(InfoHead)
        'Resize the byte data array
        ReDim SourceData(InfoHead(LoopC).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(LoopC).lngFileStart, SourceData
        
        'Decompress all data
        If InfoHead(LoopC).lngFileSize < InfoHead(LoopC).lngFileSizeUncompressed Then
            Decompress_Data SourceData, InfoHead(LoopC).lngFileSizeUncompressed
        End If
        
        'Get a free handler
        file_handler_list(LoopC).handle = FreeFile
        
        Open OutputFilePath & InfoHead(LoopC).strFileName For Binary As file_handler_list(LoopC).handle
        
        file_handler_list(LoopC).file_name = InfoHead(LoopC).strFileName
        
        Put file_handler_list(LoopC).handle, , SourceData
        
        'We leave the files open so they are locked. To unlock them call Close_All
        
        Erase SourceData
        
        'Update progress bars
        If Not ResourcePrgbar Is Nothing Then ResourcePrgbar.value = ResourcePrgbar.value + 1
        If Not GeneralPrgBar Is Nothing Then GeneralPrgBar.value = GeneralPrgBar.value + 1
        If Not GeneralLbl Is Nothing Then GeneralLbl.Caption = "Loading File " & LoopC & " of " & FileHead.intNumFiles & "..."
        DoEvents
    Next LoopC
    
    'Close the binary file
    Close SourceFile
    
    'Copy handler list
    Select Case file_type
        Case grh
            ReDim GRH_Handles(UBound(file_handler_list()))
            GRH_Handles = file_handler_list
            
        Case MIDI
            ReDim MIDI_Handles(UBound(file_handler_list()))
            MIDI_Handles = file_handler_list
        
        Case MP3
            ReDim MP3_Handles(UBound(file_handler_list()))
            MP3_Handles = file_handler_list
        
        Case WAV
            ReDim WAV_Handles(UBound(file_handler_list()))
            WAV_Handles = file_handler_list
        
        Case Scripts
            ReDim Scripts_Handles(UBound(file_handler_list()))
            Scripts_Handles = file_handler_list
    End Select
    
    Erase InfoHead
    
    Extract_Files = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_Patch(ByVal resource_path As String, ByVal file_name As String, ByRef ResourcePrgbar As ProgressBar, ByRef GeneralPrgBar As ProgressBar, ByRef GeneralLbl As Label) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a patch file
'*****************************************************************
    Dim LoopC As Long
    Dim OutputFilePath As String
    Dim OutputFile As Integer
    Dim SourceFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim RequiredSpace As Currency
    
    '************************************************************************************************
    'This is similar to Extract, but has some small differences to make sure what is being updated
    '************************************************************************************************
'Set up the error handler
On Local Error GoTo ErrHandler
    
    'Open the binary file
    SourceFile = FreeFile
    SourceFilePath = resource_path & "\" & file_name
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt File Header
    Encrypt_File_Header FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
    
    'Check if there is enough hard drive space to extract all files
    For LoopC = 0 To UBound(InfoHead)
        'Desencrypt each Info Header before accessing the data
        Encrypt_Info_Header InfoHead(LoopC)
        RequiredSpace = RequiredSpace + InfoHead(LoopC).lngFileSizeUncompressed
    Next LoopC
    
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
        Erase InfoHead
        MsgBox "There is not enough drive space to extract the compressed files.", , "Error"
        Exit Function
    End If
    
    'Extract all of the files from the binary file
    For LoopC = 0 To UBound(InfoHead)
        'Resize the byte data array
        ReDim SourceData(InfoHead(LoopC).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(LoopC).lngFileStart, SourceData
        
        'Decompress all data
        If InfoHead(LoopC).lngFileSize < InfoHead(LoopC).lngFileSizeUncompressed Then
            Decompress_Data SourceData, InfoHead(LoopC).lngFileSizeUncompressed
        End If
        
        'Get a free handle
        OutputFile = FreeFile
        
        'Check what kind of file it is
        Select Case Right(InfoHead(LoopC).strFileName, 3)
            Case Is = "bmp"
                OutputFilePath = resource_path & GRAPHIC_PATH
                ReDim Preserve GRH_Handles(0 To UBound(GRH_Handles()) + 1)
                GRH_Handles(UBound(GRH_Handles)).handle = OutputFile
                GraphicsDSO = True
            
            Case Is = "mp3"
                OutputFilePath = resource_path & MP3_PATH
                ReDim Preserve MP3_Handles(0 To UBound(MP3_Handles()) + 1)
                MP3_Handles(UBound(MP3_Handles)).handle = OutputFile
                MP3DSO = True
            
            Case Is = "mid"
                OutputFilePath = resource_path & MIDI_PATH
                ReDim Preserve MIDI_Handles(0 To UBound(MIDI_Handles()) + 1)
                MIDI_Handles(UBound(MIDI_Handles)).handle = OutputFile
                MIDIDSO = True
            
            Case Is = "wav"
                OutputFilePath = resource_path & WAV_PATH
                ReDim Preserve WAV_Handles(0 To UBound(WAV_Handles()) + 1)
                WAV_Handles(UBound(WAV_Handles)).handle = OutputFile
                WAVDSO = True
            
            Case Else
                OutputFilePath = resource_path & SCRIPT_PATH
                ReDim Preserve Scripts_Handles(0 To UBound(Scripts_Handles()) + 1)
                Scripts_Handles(UBound(Scripts_Handles)).handle = OutputFile
                ScriptsDSO = True
        End Select
        
        Open OutputFilePath & InfoHead(LoopC).strFileName For Binary Access Write Lock Read As OutputFile
        
        Put OutputFile, 1, SourceData
        
        Erase SourceData
        
        'Update progress bars
        If Not ResourcePrgbar Is Nothing Then ResourcePrgbar.value = ResourcePrgbar.value + 1
        If Not GeneralPrgBar Is Nothing Then GeneralPrgBar.value = GeneralPrgBar.value + 1
        If Not GeneralLbl Is Nothing Then GeneralLbl.Caption = "Loading File " & GeneralPrgBar.value & " of " & GeneralPrgBar.Max & "..."
        DoEvents
    Next LoopC
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    Extract_Patch = True
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Compress_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String, ByRef GeneralPrgBar As ProgressBar) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************
    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim lngFileStart As Long
    Dim LoopC As Long
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case grh
            SourceFilePath = resource_path & GRAPHIC_PATH
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Graphics.ORE"
        
        Case MIDI
            SourceFilePath = resource_path & MIDI_PATH
            SourceFileExtension = ".mid"
            OutputFilePath = dest_path & "Low-Def Music.ORE"
        
        Case MP3
            SourceFilePath = resource_path & MP3_PATH
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "Hi-Def Music.ORE"
        
        Case WAV
            SourceFilePath = resource_path & WAV_PATH
            SourceFileExtension = ".map"
            OutputFilePath = dest_path & "Sounds.ORE"
        
        Case Scripts
            SourceFilePath = resource_path & SCRIPT_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Scripts.ORE"
        
        Case Patch
            SourceFilePath = resource_path & PATCH_PATH
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Patch.ORE"
    End Select
    
    SourceFileName = Dir(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve InfoHead(FileHead.intNumFiles - 1)
        InfoHead(FileHead.intNumFiles - 1).strFileName = SourceFileName
        
        Open SourceFilePath & SourceFileName For Binary As SourceFile
        
        Close SourceFile
        
        'Search new file
        SourceFileName = Dir
    Wend
    
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    GeneralPrgBar.Max = FileHead.intNumFiles
    GeneralPrgBar.value = 0
    
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For LoopC = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & InfoHead(LoopC).strFileName For Binary Access Read Lock Write As SourceFile
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(LoopC).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(LoopC).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        'Update progress bar
        If Not GeneralPrgBar Is Nothing Then GeneralPrgBar.value = GeneralPrgBar.value + 1
        DoEvents
    Next LoopC
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + FileHead.intNumFiles * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + FileHead.intNumFiles * Len(InfoHead(0)) + 1
    For LoopC = 0 To FileHead.intNumFiles - 1
        InfoHead(LoopC).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(LoopC).lngFileSize
        'Once an InfoHead index is ready, we encrypt it
        Encrypt_Info_Header InfoHead(LoopC)
    Next LoopC
    
    'Encrypt the FileHeader
    Encrypt_File_Header FileHead
    
    '************ Write Data
    
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Sub Close_Handler_List(ByRef file_handler_list() As file_handle, Optional ByVal EmptyFile As Boolean = False)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Closes all files in a handler list. clears the files if asked to.
'*****************************************************************
On Local Error GoTo ErrHandler
    Dim LoopC As Long
    Dim data() As Byte
    
    'Fill all files with 0s and close them
    For LoopC = 0 To UBound(file_handler_list())
        If EmptyFile Then
            ReDim data(LOF(file_handler_list(LoopC).handle) - 1)
            Get file_handler_list(LoopC).handle, 1, data
            ZeroMemory data(0), LOF(file_handler_list(LoopC).handle)
            Put file_handler_list(LoopC).handle, 1, data
        End If
        Close file_handler_list(LoopC).handle
    Next LoopC
ErrHandler:
End Sub

Public Sub Close_All(Optional ByVal EmptyFile As Boolean = False)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Closes all files in all handler lists
'*****************************************************************
On Local Error Resume Next
    Close_Handler_List GRH_Handles
    Close_Handler_List MP3_Handles
    Close_Handler_List MIDI_Handles
    Close_Handler_List WAV_Handles
    Close_Handler_List Scripts_Handles
End Sub

Public Sub Delete_Resources(ByVal resource_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Deletes all resource files
'*****************************************************************
On Local Error Resume Next
    Close_All True
    
    Kill resource_path & GRAPHIC_PATH & "*.bmp"
    Kill resource_path & MP3_PATH & "*.mp3"
    Kill resource_path & MIDI_PATH & "*.mid"
    Kill resource_path & SCRIPT_PATH & "*.*"
End Sub
