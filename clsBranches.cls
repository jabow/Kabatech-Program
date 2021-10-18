VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBranches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colBranches As New Collection

Public Function Add(objBranch As clsbranch)
   Call colBranches.Add(Item:=objBranch, Key:=objBranch.branchNr)
End Function

Public Sub Remove(vItem As Variant)
   colBranches.Remove vItem
End Sub

Public Property Get Item(vItem As Variant) As clsbranch
   Set Item = colBranches(vItem)
End Property

Public Property Get Count() As Long
   Count = colBranches.Count
End Property
