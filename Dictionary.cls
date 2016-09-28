VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '
Private dict_pDictionary As Object

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
    Attribute CompareMode.VB_Description = "Set or get the string comparison method."
    CompareMode = dict_pDictionary.CompareMode
End Property

Public Property Let CompareMode(Value As CompareMethod)
    dict_pDictionary.CompareMode = Value
End Property

Public Property Get Count() As Long
    Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
    Count = dict_pDictionary.Count
End Property

Public Property Get Item(Key As Variant) As Variant
    Attribute Item.VB_Description = "Set or get the item for a given key."
    Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(dict_pDictionary.Item(Key)) Then
        Set Item = dict_pDictionary.Item(Key)
    Else
        Item = dict_pDictionary.Item(Key)
    End If
End Property

Public Property Let Item(Key As Variant, Value As Variant)
    dict_pDictionary.Item(Key) = Value
End Property

Public Property Set Item(Key As Variant, Value As Variant)
    Set dict_pDictionary.Item(Key) = Value
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    Attribute Key.VB_Description = "Change a key to a different key."
    dict_pDictionary.Key(Previous) = Updated
End Property


' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(Key As Variant, Item As Variant)
    Attribute Add.VB_Description = "Add a new key and item to the dictionary."
    dict_pDictionary.Add Key, Item
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
    Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
    Exists = dict_pDictionary.Exists(Key)
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
    Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
    Items = dict_pDictionary.Items
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
    Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
    Keys = dict_pDictionary.Keys
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
    Attribute Remove.VB_Description = "Remove a given key from the dictionary."
    dict_pDictionary.Remove Key
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
    Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
    dict_pDictionary.RemoveAll
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    Set dict_pDictionary = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set dict_pDictionary = Nothing
End Sub
