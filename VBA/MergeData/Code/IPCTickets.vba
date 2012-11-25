Option Explicit
Private Tickets As Collection

Public Sub Add(Ticket As IPCTicket)
    Tickets.Add Ticket
    Debug.Print "Adding Ticket"
End Sub



'Find if exists
Public Function Find(ChangeID As String) As IPCTicket
    Dim Ticket As IPCTicket
    For Each Ticket In Tickets
        If Ticket.ChangeID = ChangeID Then
            Debug.Print "Found Ticket"
            Set Find = Ticket
        End If
    Next
    
    Debug.Print "Not Found Ticket"
    Set Find = Empty
End Function


