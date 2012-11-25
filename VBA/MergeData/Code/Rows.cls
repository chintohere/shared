VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Tickets As New Collection

Public Sub AddTicket(Ticket As IPCTicket)
    Tickets.Add Ticket
End Sub

'Find if exists
Public Function Find(ChangeID As String) As IPCTicket
    Dim Ticket As IPCTicket
    For Each Ticket In Tickets
        If Ticket.ChangeID = ChangeID Then
            Debug.Print "Found Ticket"
            Set Find = Ticket
            Exit Function
        End If
    Next
    
    Debug.Print "Not Found Ticket"
    Set Find = Nothing
End Function

