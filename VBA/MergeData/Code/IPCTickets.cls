VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPCTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Tickets As New Collection

Sub Add(ByVal Ticket As IPCTicket)
    Tickets.Add Item:=Ticket, Key:=Ticket.ChangeID
    Debug.Print "Adding Ticket: " & Ticket.ChangeID
End Sub


'Find if exists
Function Find(ChangeID As String, Optional Item As String = Empty) As IPCTicket
    
    If Item = Empty Then
    Else
        ChangeID = ChangeID & vbNewLine & Item
    End If
    
    Dim Ticket As IPCTicket
    
    For Each Ticket In Tickets
        If Ticket.ChangeID = ChangeID Then
            'Debug.Print "Found Ticket Matching: " & ChangeID
            Set Find = Ticket
            Exit Function
        End If
    Next
    
    'Debug.Print "Not Found Ticket Matching: " & ChangeID
    Set Find = Nothing
End Function

Function Size() As Integer
    Size = Tickets.Count
End Function

Function All() As Collection
 Set All = Tickets
End Function


