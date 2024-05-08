VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacterCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsCharacterCreation.cls - v1.0.0
'
'Has all methods to create a char using D&D second edition rules
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
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 2/24/2004
'   - First Release
'*****************************************************************
Option Explicit

'***********
'Race
'***********
Public Enum races
    Human
    HalfElf
    HalfGiant
    Mul
    ThriKreen
    Elf
    Halfling
    Dwarf
End Enum

'***********
'Class
'***********
Public Enum classes
    Gladiator
    Fighter
    Thief
    Ranger
    Psionicist
    Preserver
    Defiler
    Cleric
    Druid
End Enum

'***********
'Alignment
'***********
Public Enum Alignment
    TrueNeutral
    ChaoticGood
    ChaoticNeutral
    LawfulGood
    LawfulNeutral
    NeutralGood
End Enum

'***********
'Sphere
'***********
Public Enum spheres
    None
    Earth
    Air
    Fire
    Water
End Enum

'***********
'Psionic powers
'***********
Public Enum psionic_powers
    Kinetic
    Telepathic
    Metabolic
    All
End Enum

'Used to check stats´s limits of a char according to it´s race
Private Type Stats
    Min_Strength As Byte
    Max_Strength As Byte
    Min_Dexterity As Byte
    Max_Dexterity As Byte
    Min_Constitution As Byte
    Max_Constitution As Byte
    Min_Intelligence As Byte
    Max_Intelligence As Byte
    Min_Wisdom As Byte
    Max_Wisdom As Byte
    Min_Charisma As Byte
    Max_Charisma As Byte
End Type

Dim CharStats(0 To 7) As Stats

'**********************
'Character´s race stuff
'**********************
Dim CharRace As races
Dim RaceNames(0 To 7) As String

'**********************
'Character´s alignment stuff
'**********************
Dim CharAlign As Alignment
Dim AlignmentNames(0 To 5) As String

'**********************
'Character´s class
'**********************
Dim CharClass As classes
Dim ClassNames(0 To 8) As String

'**********************
'Character´s sphere
'**********************
Dim CharSphere As spheres

'**********************
'Character´s psionic power
'**********************
Dim CharPsionicPower As psionic_powers

'**********************
'Character´s stats
'**********************
Dim char_STR As Byte
Dim char_DEX As Byte
Dim char_CON As Byte
Dim char_INT As Byte
Dim char_WIS As Byte
Dim char_CHR As Byte

'Frre points
Dim FreePoints As Long

Public Function Race_Change(ByVal Decrease As Boolean) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If CharRace > 0 Then
            CharRace = CharRace - 1
        Else
            CharRace = 7
        End If
    Else
        If CharRace < 7 Then
            CharRace = CharRace + 1
        Else
            CharRace = 0
        End If
    End If
    
    'Reset all values since the race was changed
    CharClass = Gladiator
    CharSphere = None
    CharPsionicPower = Kinetic
    
    Race_Change = RaceNames(CharRace)
End Function

Public Function Char_STR_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_STR > CharStats(CharRace).Min_Strength Then
            char_STR = char_STR - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_STR < CharStats(CharRace).Max_Strength And FreePoints > 0 Then
            char_STR = char_STR + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_STR_Change = char_STR
End Function

Public Sub Char_STR_Set(ByVal newSTR As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_STR = newSTR
End Sub

Public Function Char_DEX_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_DEX > CharStats(CharRace).Min_Dexterity Then
            char_DEX = char_DEX - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_DEX < CharStats(CharRace).Max_Dexterity And FreePoints > 0 Then
            char_DEX = char_DEX + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_DEX_Change = char_DEX
End Function

Public Sub Char_DEX_Set(ByVal newDEX As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_DEX = newDEX
End Sub

Public Function Char_CON_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_CON > CharStats(CharRace).Min_Constitution Then
            char_CON = char_CON - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_CON < CharStats(CharRace).Max_Constitution And FreePoints > 0 Then
            char_CON = char_CON + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_CON_Change = char_CON
End Function

Public Sub Char_CON_Set(ByVal newCON As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_CON = newCON
End Sub

Public Function Char_INT_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_INT > CharStats(CharRace).Min_Intelligence Then
            char_INT = char_INT - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_INT < CharStats(CharRace).Max_Intelligence And FreePoints > 0 Then
            char_INT = char_INT + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_INT_Change = char_INT
End Function

Public Sub Char_INT_Set(ByVal newINT As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_INT = newINT
End Sub

Public Function Char_WIS_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_WIS > CharStats(CharRace).Min_Wisdom Then
            char_WIS = char_WIS - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_WIS < CharStats(CharRace).Max_Wisdom And FreePoints > 0 Then
            char_WIS = char_WIS + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_WIS_Change = char_WIS
End Function

Public Sub Char_WIS_Set(ByVal newWIS As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_WIS = newWIS
End Sub

Public Function Char_CHR_Change(ByVal Decrease As Boolean) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If char_CHR > CharStats(CharRace).Min_Charisma Then
            char_CHR = char_CHR - 1
            FreePoints = FreePoints + 1
        End If
    Else
        If char_CHR < CharStats(CharRace).Max_Charisma And FreePoints > 0 Then
            char_CHR = char_CHR + 1
            FreePoints = FreePoints - 1
        End If
    End If
    
    Char_CHR_Change = char_CHR
End Function

Public Sub Char_CHR_Set(ByVal newCHR As Long)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    char_CHR = newCHR
End Sub

Public Function Alignment_Change(ByVal Decrease As Boolean) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Decrease Then
        If CharAlign = 0 Then
            CharAlign = 5
        Else
            CharAlign = CharAlign - 1
        End If
    Else
        If CharAlign = 5 Then
            CharAlign = 0
        Else
            CharAlign = CharAlign + 1
        End If
    End If
    
    Alignment_Change = AlignmentNames(CharAlign)
End Function

Public Function Alignment_Set(ByVal align As Alignment) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If align < TrueNeutral Or align > NeutralGood Then
        Alignment_Set = False
        Exit Function
    End If
    
    CharAlign = align
End Function

Public Function Psionic_Power_Set(ByVal psionic_power As psionic_powers) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    'Make sure it´s a valid psionic power
    If psionic_power < Kinetic Or psionic_power > All Then
        Psionic_Power_Set = False
        Exit Function
    End If
    
    CharPsionicPower = psionic_power
End Function

Public Function Sphere_Set(ByVal sphere As spheres) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    'Make sure it´s a valid sphere
    If sphere < None Or sphere > Water Then
        Sphere_Set = False
        Exit Function
    End If
    
    CharSphere = sphere
    
    Sphere_Set = True
End Function

Public Function Class_Set(ByVal Class As Long) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    'Make sure it´s a valid class for the race
    If Class < Gladiator Or Class > Druid Then
        Class_Set = False
        Exit Function
    End If
    
    Select Case CharRace
        Case HalfGiant
            If Class = Defiler Or Class = Preserver Or Class = Druid Or Class = Thief Then
                Class_Set = False
                Exit Function
            End If
        Case Mul
            If Class = Defiler Or Class = Preserver Or Class = Ranger Then
                Class_Set = False
                Exit Function
            End If
        Case ThriKreen
            If Class = Defiler Or Class = Preserver Or Class = Thief Then
                Class_Set = False
                Exit Function
            End If
        Case Elf
            If Class = Druid Then
                Class_Set = False
                Exit Function
            End If
        Case Halfling
            If Class = Defiler Or Class = Preserver Then
                Class_Set = False
                Exit Function
            End If
        Case Dwarf
            If Class = Defiler Or Class = Preserver Or Class = Ranger Or Class = Druid Then
                Class_Set = False
                Exit Function
            End If
    End Select
    
    CharClass = Class
    
    Class_Set = True
End Function

Public Function Race_Set(ByVal race As races) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If race < Human Or race > Dwarf Then
        Race_Set = False
        Exit Function
    End If
    
    CharRace = race
End Function

Public Function Character_Check() As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
'Checks if all values are correct
    'Check if race and class match
    Select Case CharRace
        Case HalfGiant
            If CharClass = Defiler Or CharClass = Preserver Or CharClass = Druid Or CharClass = Thief Then
                Character_Check = False
                Exit Function
            End If
        Case Mul
            If CharClass = Defiler Or CharClass = Preserver Or CharClass = Ranger Then
                Character_Check = False
                Exit Function
            End If
        Case ThriKreen
            If CharClass = Defiler Or CharClass = Preserver Or CharClass = Thief Then
                Character_Check = False
                Exit Function
            End If
        Case Elf
            If CharClass = Druid Then
                Character_Check = False
                Exit Function
            End If
        Case Halfling
            If CharClass = Defiler Or CharClass = Preserver Then
                Character_Check = False
                Exit Function
            End If
        Case Dwarf
            If CharClass = Defiler Or CharClass = Preserver Or CharClass = Ranger Or CharClass = Druid Then
                Character_Check = False
                Exit Function
            End If
    End Select
    
    'Make sure spheres and psionic powers are right with class
    If CharClass = Cleric Or CharClass = Druid Or CharClass = Ranger Then
        'Any sphere is OK
    ElseIf CharSphere <> None Then
        Character_Check = False
        Exit Function
    End If
    
    If CharClass <> Psionicist And CharPsionicPower = All Then
        Character_Check = False
        Exit Function
    End If
    
    'Check stats are between limits according to race
    If char_STR < CharStats(CharRace).Min_Strength Or char_STR > CharStats(CharRace).Max_Strength Then
        Character_Check = False
        Exit Function
    End If
    
    If char_DEX < CharStats(CharRace).Min_Dexterity Or char_DEX > CharStats(CharRace).Max_Dexterity Then
        Character_Check = False
        Exit Function
    End If
    
    If char_CON < CharStats(CharRace).Min_Constitution Or char_CON > CharStats(CharRace).Max_Constitution Then
        Character_Check = False
        Exit Function
    End If
    
    If char_INT < CharStats(CharRace).Min_Intelligence Or char_INT > CharStats(CharRace).Max_Intelligence Then
        Character_Check = False
        Exit Function
    End If
    
    If char_WIS < CharStats(CharRace).Min_Wisdom Or char_WIS > CharStats(CharRace).Max_Wisdom Then
        Character_Check = False
        Exit Function
    End If
    
    If char_CHR < CharStats(CharRace).Min_Charisma Or char_CHR > CharStats(CharRace).Max_Charisma Then
        Character_Check = False
        Exit Function
    End If
    
    Character_Check = True
End Function

Public Function Class_Get() As classes
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Class_Get = CharClass
End Function

Public Function Psionic_Power_Get() As psionic_powers
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Psionic_Power_Get = CharPsionicPower
End Function

Public Function Race_Get() As races
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Race_Get = CharRace
End Function

Public Function Race_Get_Name(ByVal race As races) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Race_Get_Name = RaceNames(race)
End Function

Public Function Class_Get_Name(ByVal Class As classes) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Class_Get_Name = ClassNames(Class)
End Function

Public Function Sphere_Get() As spheres
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Sphere_Get = CharSphere
End Function

Public Function Alignment_Get() As Alignment
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Alignment_Get = CharAlign
End Function

Public Function Alignment_Get_Name(ByVal align As Alignment) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Alignment_Get_Name = AlignmentNames(align)
End Function

Public Function Stats_Roll(ByVal race As races)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim result As Long
    Dim points As Long
    
    result = Round(General_Random_Number(1, 100), 0)
    
    Select Case result
        'See how many point correspond to the obtained number
        Case 1 To 40
            points = 66
        Case 41 To 70
            points = 72
        Case 71 To 86
            points = 78
        Case 87 To 96
            points = 86
        Case 97 To 99
            points = 92
        Case 100
            points = 98
    End Select
    
    Stats_Roll = points
End Function

Public Function FreePoints_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    FreePoints_Get = FreePoints
End Function

Public Function Char_STR_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_STR_Get = char_STR
End Function

Public Function Char_DEX_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_DEX_Get = char_DEX
End Function

Public Function Char_CON_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CON_Get = char_CON
End Function

Public Function Char_INT_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_INT_Get = char_INT
End Function

Public Function Char_WIS_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_WIS_Get = char_WIS
End Function

Public Function Char_CHR_Get() As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CHR_Get = char_CHR
End Function

Public Function Char_CHR_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CHR_Get_Min = CharStats(race).Min_Charisma
End Function

Public Function Char_STR_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_STR_Get_Min = CharStats(race).Min_Strength
End Function

Public Function Char_DEX_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_DEX_Get_Min = CharStats(race).Min_Dexterity
End Function

Public Function Char_CON_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CON_Get_Min = CharStats(race).Min_Constitution
End Function

Public Function Char_INT_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_INT_Get_Min = CharStats(race).Min_Intelligence
End Function

Public Function Char_WIS_Get_Min(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_WIS_Get_Min = CharStats(race).Min_Wisdom
End Function

Public Function Char_CHR_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CHR_Get_Max = CharStats(race).Max_Charisma
End Function

Public Function Char_STR_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_STR_Get_Max = CharStats(race).Max_Strength
End Function

Public Function Char_DEX_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_DEX_Get_Max = CharStats(race).Max_Dexterity
End Function

Public Function Char_CON_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_CON_Get_Max = CharStats(race).Max_Constitution
End Function

Public Function Char_INT_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_INT_Get_Max = CharStats(race).Max_Intelligence
End Function

Public Function Char_WIS_Get_Max(ByVal race As races) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Char_WIS_Get_Max = CharStats(race).Max_Wisdom
End Function

Public Sub Initialize()
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    'Race´s names
    RaceNames(0) = "Human"
    RaceNames(1) = "Half-Elf"
    RaceNames(2) = "Half-Giant"
    RaceNames(3) = "Mul"
    RaceNames(4) = "Thri-Kreen"
    RaceNames(5) = "Elf"
    RaceNames(6) = "Halfling"
    RaceNames(7) = "Dwarf"
    
    'Alignments
    AlignmentNames(0) = "True Neutral"
    AlignmentNames(1) = "Chaotic Good"
    AlignmentNames(2) = "Chaotic Neutral"
    AlignmentNames(3) = "Lawful Good"
    AlignmentNames(4) = "Lawful Neutral"
    AlignmentNames(5) = "Neutral Good"
    
    'Classes
    ClassNames(0) = "Gladiator"
    ClassNames(1) = "Fighter"
    ClassNames(2) = "Thief"
    ClassNames(3) = "Ranger"
    ClassNames(4) = "Psionicist"
    ClassNames(5) = "Preserver"
    ClassNames(6) = "Defiler"
    ClassNames(7) = "Cleric"
    ClassNames(8) = "Druid"
    
    'Stats limits
    'Human
    CharStats(0).Min_Strength = 5
    CharStats(0).Max_Strength = 20
    CharStats(0).Min_Dexterity = 5
    CharStats(0).Max_Dexterity = 20
    CharStats(0).Min_Constitution = 5
    CharStats(0).Max_Constitution = 20
    CharStats(0).Min_Intelligence = 5
    CharStats(0).Max_Intelligence = 20
    CharStats(0).Min_Wisdom = 5
    CharStats(0).Max_Wisdom = 20
    CharStats(0).Min_Charisma = 5
    CharStats(0).Max_Charisma = 20
    
    'Half-Elf
    CharStats(1).Min_Strength = 5
    CharStats(1).Max_Strength = 20
    CharStats(1).Min_Dexterity = 8
    CharStats(1).Max_Dexterity = 21
    CharStats(1).Min_Constitution = 5
    CharStats(1).Max_Constitution = 19
    CharStats(1).Min_Intelligence = 5
    CharStats(1).Max_Intelligence = 20
    CharStats(1).Min_Wisdom = 5
    CharStats(1).Max_Wisdom = 20
    CharStats(1).Min_Charisma = 5
    CharStats(1).Max_Charisma = 20
    
    'Half-Giant
    CharStats(2).Min_Strength = 17
    CharStats(2).Max_Strength = 24
    CharStats(2).Min_Dexterity = 5
    CharStats(2).Max_Dexterity = 15
    CharStats(2).Min_Constitution = 5
    CharStats(2).Max_Constitution = 22
    CharStats(2).Min_Intelligence = 5
    CharStats(2).Max_Intelligence = 13
    CharStats(2).Min_Wisdom = 5
    CharStats(2).Max_Wisdom = 15
    CharStats(2).Min_Charisma = 5
    CharStats(2).Max_Charisma = 15
    
    'Mul
    CharStats(3).Min_Strength = 10
    CharStats(3).Max_Strength = 22
    CharStats(3).Min_Dexterity = 5
    CharStats(3).Max_Dexterity = 20
    CharStats(3).Min_Constitution = 10
    CharStats(3).Max_Constitution = 21
    CharStats(3).Min_Intelligence = 5
    CharStats(3).Max_Intelligence = 19
    CharStats(3).Min_Wisdom = 5
    CharStats(3).Max_Wisdom = 20
    CharStats(3).Min_Charisma = 5
    CharStats(3).Max_Charisma = 18
    
    'Thri-Kreen
    CharStats(4).Min_Strength = 8
    CharStats(4).Max_Strength = 20
    CharStats(4).Min_Dexterity = 15
    CharStats(4).Max_Dexterity = 22
    CharStats(4).Min_Constitution = 5
    CharStats(4).Max_Constitution = 20
    CharStats(4).Min_Intelligence = 5
    CharStats(4).Max_Intelligence = 19
    CharStats(4).Min_Wisdom = 5
    CharStats(4).Max_Wisdom = 21
    CharStats(4).Min_Charisma = 5
    CharStats(4).Max_Charisma = 15
    
    'Elf
    CharStats(5).Min_Strength = 5
    CharStats(5).Max_Strength = 20
    CharStats(5).Min_Dexterity = 5
    CharStats(5).Max_Dexterity = 22
    CharStats(5).Min_Constitution = 8
    CharStats(5).Max_Constitution = 18
    CharStats(5).Min_Intelligence = 8
    CharStats(5).Max_Intelligence = 21
    CharStats(5).Min_Wisdom = 5
    CharStats(5).Max_Wisdom = 19
    CharStats(5).Min_Charisma = 5
    CharStats(5).Max_Charisma = 20
    
    'Halfling
    CharStats(6).Min_Strength = 5
    CharStats(6).Max_Strength = 18
    CharStats(6).Min_Dexterity = 12
    CharStats(6).Max_Dexterity = 22
    CharStats(6).Min_Constitution = 5
    CharStats(6).Max_Constitution = 19
    CharStats(6).Min_Intelligence = 5
    CharStats(6).Max_Intelligence = 20
    CharStats(6).Min_Wisdom = 7
    CharStats(6).Max_Wisdom = 22
    CharStats(6).Min_Charisma = 5
    CharStats(6).Max_Charisma = 19
    
    'Dwarf
    CharStats(7).Min_Strength = 10
    CharStats(7).Max_Strength = 21
    CharStats(7).Min_Dexterity = 5
    CharStats(7).Max_Dexterity = 19
    CharStats(7).Min_Constitution = 14
    CharStats(7).Max_Constitution = 22
    CharStats(7).Min_Intelligence = 5
    CharStats(7).Max_Intelligence = 20
    CharStats(7).Min_Wisdom = 5
    CharStats(7).Max_Wisdom = 20
    CharStats(7).Min_Charisma = 5
    CharStats(7).Max_Charisma = 18
End Sub
