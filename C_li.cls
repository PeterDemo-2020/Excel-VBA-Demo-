VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_li"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private val_ As New C_a
Private key_ As New C_a
Private ix_ As New C_a

Private Sub Class_Initialize()
    Me.Reset = 21
End Sub

Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Call checkKey(key)
    Call Assign(Item, val_(IIf(TypeName(key) = "String", Ix(key), key)))
End Property

Public Property Let Item(ByVal key As Variant, ByRef newVal As Variant)
Attribute Item.VB_UserMemId = 0
    Call checkKey(key)
    val_(IIf(TypeName(key) = "String", Ix(key), key)) = newVal
End Property

Public Property Let Add(Optional key As String = "", ByRef newVal As Variant)
    val_.Add = newVal
    If key <> "" Then
        key_.Add = key
        ix_.Add = val_.Last
    End If
End Property

Public Property Let Reset(ByVal newLength As Long)
    val_.Reset = newLength
    key_.Reset = newLength
    ix_.Reset = newLength
End Property

Private Sub checkKey(ByVal key As String)
    Dim keyType As String:    keyType = TypeName(key)
    If keyType <> "String" And keyType <> "Long" And keyType <> "Integer" Then Call MsgBox(" ? C_li: (Get) Item -> key 's type is error", vbExclamation, "Error"):    End
End Sub

Private Function Ix(ByVal key As String) As Long
    Ix = -1
    Dim i As Long
    For i = 0 To key_.Last
        If key_(i) = key Then Ix = ix_(i):    Exit Function
    Next
End Function

Private Sub Assign(ByRef thisRef As Variant, ByRef thisVal As Variant)
    If IsObject(thisVal) Then Set thisRef = thisVal:    Exit Sub
    thisRef = thisVal
End Sub
