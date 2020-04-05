VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private val_() As Variant
Private last_ As Long
Private Sub Class_Initialize()
    Me.Reset = 21
End Sub
Public Property Get Last() As Long
    Last = last_
End Property
Public Property Get Item(ByVal ix As Long) As Variant
    Call checkIx(ix, "(Get) Item")
    Call Assign(Item, val_(ix))
End Property
Public Property Let Item(ByVal ix As Long, ByRef newVal As Variant)
    Call checkIx(ix, "(Let) Item")
    Call Assign(val_(ix), newVal)
End Property
Public Property Let Add(ByRef newVal As Variant)
    last_ = last_ + 1
    If last_ > UBound(val_) Then ReDim Preserve val_(UBound(val_) * 1.5 \ 1)
    Call Assign(val_(last_), newVal)
End Property
Public Property Let Reset(ByVal newLength As Long)
    ReDim val_(newLength - 1)
    last_ = -1
End Property
Private Sub checkIx(ByVal ix As Long, ByVal Class As String)
    If ix < 0 Or ix > last_ Then Call MsgBox(" ? C_a: " & Class & " -> ix is overflow (checkIx)"):    End
End Sub
Private Sub Assign(ByRef thisRef As Variant, ByRef thisVal As Variant)
    If IsObject(thisVal) Then Set thisRef = thisVal:    Exit Sub
    thisRef = thisVal
End Sub
