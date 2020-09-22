Attribute VB_Name = "modListObjectFunction"
Option Explicit

Public Function ItemPosInListObj(ListObj As Object, ItemValue As String) As Long
ItemPosInListObj = 0
If ListObj.ListCount = 0 Then Exit Function

Dim Found As Boolean, a As Long
With ListObj
    For a = 0 To .ListCount - 1
        If .List(a) = ItemValue Then Found = True
        If Found Then Exit For
    Next
End With
If Found Then ItemPosInListObj = a + 1
End Function
