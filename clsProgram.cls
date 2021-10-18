VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_customer As String
Private m_range As String
Private m_model As String
Private m_location As String
Private m_program_no As String
Private m_part_no As String
Private m_drawer As String
Private m_revision As Long
Private m_word_created_date As Date
Private m_final_branch_tape_length As Long
Private m_word_saved As Date
Private m_word_overwritten As Boolean
Private m_prog_saved As Date
Private m_prog_overwritten As Boolean
Private m_input_sheet_saved As Date
Private m_input_sheet_overwritten As Boolean



'Private m_branch As New Branch

'Private m_branch As New Branch
Private colBranches As New Collection

Public Sub Remove(vItem As Variant)
   m_branch.Remove vItem
End Sub

Public Property Get Item(vItem As Variant) As clsbranch
   Set Item = m_branch(vItem)
End Property

Public Property Get Count() As Long
   Count = m_branch.Count
End Property



Property Let Customer(C As String)
    m_customer = C
End Property

Property Get Customer() As String
    Customer = m_customer
End Property

Property Let Range(r As String)
    m_range = r
End Property

Property Get Range() As String
    Range = m_range
End Property

Property Let Model(m As String)
    m_model = m
End Property

Property Get Model() As String
    Model = m_model
End Property

Property Let Location(l As String)
    m_location = l
End Property

Property Get Location() As String
    Location = m_location
End Property

Property Let Program_no(p As String)
    m_program_no = p
End Property

Property Get Program_no() As String
    Program_no = m_program_no
End Property

Property Let Part_no(p As String)
    m_part_no = p
End Property

Property Get Part_no() As String
    Part_no = m_part_no
End Property

Property Let Drawer(d As String)
    m_drawer = d
End Property

Property Get Drawer() As String
    Drawer = m_drawer
End Property

Property Let Revision(r As String)
    m_revision = r
End Property

Property Get Revision() As String
    Revision = m_revision
End Property

Property Let Final_branch_tape_length(fbl As String)
    m_final_branch_tape_length = fbl
End Property

Property Get Final_branch_tape_length() As String
    Final_branch_tape_length = m_final_branch_tape_length
End Property

Property Let Word_Created_Date(wcd As String)
    m_word_created_date = wcd
End Property

Property Get Word_Created_Date() As String
    Word_Created_Date = m_word_created_date
End Property

'Property Get Branches() As Collection
'    Set Branches = m_branches
'End Property
'
'Property Set Branch(newBranch As clsbranch)
'    m_branches.Add newBranches
'End Property

Property Let WordSaved(ws As Date)
    m_word_saved = ws
End Property

Property Get WordSaved() As Date
    WordSaved = m_word_saved
End Property

Property Let WordOverwritten(wo As Boolean)
    m_word_overwritten = wo
End Property

Property Get WordOverwritten() As Boolean
    WordOverwritten = m_word_overwritten
End Property

Property Let ProgSaved(ps As Date)
    m_prog_saved = ps
End Property

Property Get ProgSaved() As Date
    ProgSaved = m_prog_saved
End Property

Property Let ProgOverwritten(po As Boolean)
    m_prog_overwritten = po
End Property

Property Get ProgOverwritten() As Boolean
    ProgOverwritten = m_prog_overwritten
End Property

Property Let InputSheetSaved(iss As Date)
    m_input_sheet_saved = iss
End Property

Property Get InputSheetSaved() As Date
    InputSheetSaved = m_input_sheet_saved
End Property

Property Let InputSheetOverwritten(iso As Boolean)
    m_input_sheet_overwritten = iso
End Property

Property Get InputSheetOverwritten() As Boolean
    InputSheetOverwritten = m_input_sheet_overwritten
End Property




'Property Set branches(b() As branch, arraySize As Long)
'    ReDim m_branches(arraySize)
'    m_branches = b
'End Property
'
'Property Get branches() As branch
'    branches = m_branches
'End Property
