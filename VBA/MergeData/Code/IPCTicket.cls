VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPCTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'All columns for each row in the new sheet
Public ChangeID As String
Public ChangeType As String
Public StartTime As String
Public EndTime As String
Public Summary As String
Public Impact As String
Public RequesterName As String
Const NewLine As String = "n"

Public Sub Load(ByRef Row As Range)
    Dim Item As String
    Item = Rows.Cells(1, 2)
    
    Me.ChangeID = Row.Cells(1, 1) & vbNewLine & Item
    Me.StartTime = Row.Cells(1, 3)
    Me.EndTime = Row.Cells(1, 4)
    Me.Summary = Row.Cells(1, 5)
    Me.ChangeType = Row.Cells(1, 20)
    Me.RequesterName = Row.Cells(1, 18)
    Me.Impact = Item & " - " & Row.Cells(1, 8) & " - " & Me.StartTime & " to " & Me.EndTime
    Me.Impact = Me.Impact + vbNewLine + Row.Cells(1, 11)
    
End Sub

'This should do the concatenation bit
Public Sub Merge(ByRef Ticket As IPCTicket)
    Me.Impact = Me.Impact & vbNewLine & Ticket.Impact
End Sub

Public Sub PrintTicket()
    Debug.Print "Ticket ["
    Debug.Print "'" & Me.ChangeID & "',"
    Debug.Print "'" & Me.StartTime & "',"
    Debug.Print "'" & Me.EndTime & "',"
    Debug.Print "'" & Me.ChangeType & "',"
    Debug.Print "'" & Me.Summary & "',"
    Debug.Print "'" & Me.RequesterName & "',"
    Debug.Print "'" & Me.Impact & "'"
    Debug.Print "]"
End Sub








