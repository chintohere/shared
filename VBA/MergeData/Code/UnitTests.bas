Attribute VB_Name = "UnitTests"
Option Explicit

Public Sub RunAll()
    Call TestAdd
    Call TestMerge
    Call TestFind
    Call TestFindWithItemName
End Sub

Public Sub TestAdd()
       
    Debug.Print "Creating Ticket1"
    Dim Ticket1 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    
    Debug.Print Tickets.Size()
    Debug.Assert Tickets.Size() = 0
    
    Tickets.Add Ticket1
    
    Debug.Print Tickets.Size()
    Debug.Assert Tickets.Size() = 1
    
End Sub

Public Sub TestAddMultiple()
       
    Dim Ticket1 As New IPCTicket
    Dim Ticket2 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    Ticket2.ChangeID = "Change2"
    Ticket1.Impact = "Impact 1"
    Ticket2.Impact = "Impact 2"
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    
    Debug.Print Tickets.Size()
    Debug.Assert Tickets.Size() = 0
    
    Tickets.Add Ticket1
    
    
    Debug.Print Tickets.Size()
    Debug.Assert Tickets.Size() = 1
    
    Tickets.Add Ticket2
    
    Debug.Print Tickets.Size()
    Debug.Assert Tickets.Size() = 2
    
End Sub

Public Sub TestMerge()
       
    Debug.Print "Creating Ticket1"
    Dim Ticket1 As New IPCTicket
    Dim Ticket2 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    Ticket2.ChangeID = "Change1"
    Ticket1.Impact = "Impact 1"
    Ticket2.Impact = "Impact 2"
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    
    Call Ticket1.Merge(Ticket2)
    
    Debug.Assert Ticket1.Impact = "Impact 1" & vbNewLine & "Impact 2"
    
    
End Sub
Public Sub TestFindWithItemName()
       
    Debug.Print "Creating Ticket1"
    Dim Ticket1 As New IPCTicket
    Dim Id As String
    Id = "Change" & vbNewLine & "1"
    
    Ticket1.ChangeID = Id
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    Tickets.Add Ticket1
    
    Debug.Print "Finding Ticket"
    Dim LookingFor As IPCTicket
    
    Set LookingFor = Tickets.Find("Change", "1")
    
    Debug.Assert Not IsEmpty(LookingFor)
    Debug.Assert LookingFor.ChangeID = Id
    
    Set LookingFor = Tickets.Find("Change", "2")
    
    Debug.Assert LookingFor Is Nothing
End Sub
Public Sub TestFind()
       
    
    Debug.Print "Creating Ticket1"
    Dim Ticket1 As New IPCTicket
    Ticket1.ChangeID = "Change1"
    
    Debug.Print "Adding Ticket1"
    Dim Tickets As New IPCTickets
    Tickets.Add Ticket1
    
    Debug.Print "Finding Ticket"
    Dim LookingFor As IPCTicket
    
    Set LookingFor = Tickets.Find("Change1")
    
    Debug.Assert LookingFor.ChangeID = "Change1"
    
    Set LookingFor = Tickets.Find("Change2")
    
    Debug.Assert LookingFor Is Nothing
End Sub
