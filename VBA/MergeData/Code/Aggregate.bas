Attribute VB_Name = "Aggregate"
Sub Aggregate()
    Dim Change As Worksheet
    Dim Row As Range
    Dim Index As Integer
    Dim Tickets As New IPCTickets
    
    Set Change = Worksheets("Change") 'Get worksheet
    
    Index = 2 'Start from
    
    'go through the sheet till change id is empty
    Do Until Change.Range("A" & Index) = ""
    
        'Get row
        Set Row = Change.Range("A" & Index & ":" & "T" & Index)
        
        Debug.Print Row.Address ' e.g. print $A$2:$T$2
        
        'Build ticket
        Dim Ticket As New IPCTicket
        Dim ExistingTicket As IPCTicket
        
        Call Ticket.Load(Row)
        
        'Check for ticket
        Set ExistingTicket = Tickets.Find(Ticket.ChangeID)
               
        If (ExistingTicket Is Nothing) Then
            Call Tickets.Add(Ticket)
            'Print Ticket
            Call Ticket.PrintTicket
        Else
            Call ExistingTicket.Merge(Ticket)
            Call ExistingTicket.PrintTicket
        End If
    

                
        'Check
        Debug.Print "Tickets Size: " & Tickets.Size()
        
        Index = Index + 1
        
        'Temporary debug exit
        If Index > 10 Then
            Exit Sub
        End If
        
    Loop
    
End Sub
