Attribute VB_Name = "UnitTests"

Public Sub TestRun()

    Dim Tickets As New IPCTickets
    
    Dim Ticket1 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    Tickets.Add Ticket1
    
    Dim Ticket2 As New IPCTicket
    Ticket1.ChangeID = "Change2"
    Tickets.Add Ticket2
    
    Dim Ticket3 As New IPCTicket
    Ticket1.ChangeID = "Change3"
    Tickets.Add Ticket3
    

    Debug.Print Tickets.Find("Change2")
    
    Debug.Print Tickets.Find("Change4")
    
    
    
    
    
End Sub
