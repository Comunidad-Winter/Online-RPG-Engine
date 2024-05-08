Attribute VB_Name = "modItemDB"
Public Enum ItemType
    itWeapon = 1
    itShield
    itArmor
    itPotion
End Enum

Public Type Item
    item_name As String
    item_grh As Long
    max_dam As Long
    min_dam As Long
    item_type As ItemType
End Type

Public item_list() As Item
