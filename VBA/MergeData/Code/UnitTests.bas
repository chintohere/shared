Attribute VB_Name = "UnitTests"
Option Explicit

Public Sub TestRun()
       
    
    Debug.Print "Creating Ticket1"
    Dim Ticket1 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    Tickets.Add Ticket1
    
    Debug.Print "Finding Ticket"
    Dim LookingFor As IPCTicket
    
    Debug.Print Tickets.Find("Change1").ChangeID
    
    Set LookingFor = Tickets.Find("Change1")
    
    Debug.Print LookingFor.ChangeID
    
    Tickets.Find ("Change2")
    
    
    
End Sub

