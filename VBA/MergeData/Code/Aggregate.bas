Attribute VB_Name = "Aggregate"
Sub Aggregate()
    Dim Change As Worksheet
    Dim Output As Worksheet
    Dim Row As Range
    Dim Index As Integer
    Dim Tickets As New IPCTickets
    
    Debug.Print "Starting"
    
    'Get worksheets
    Set Change = Worksheets("Change")
    Set Output = Worksheets("Output")
    
    Index = 2 'Start from
    
    Debug.Print "Change Worksheet Size:" & Change.UsedRange.Rows.Count
    
    
    'go through the sheet till change id is empty
    Do Until Change.Range("A" & Index) = ""
    
        'Get row
        Set Row = Change.Range("A" & Index & ":" & "T" & Index)
        
        Debug.Print Row.Address ' e.g. print $A$2:$T$2
        
        'Build ticket
        Dim Ticket As New IPCTicket
        Dim ExistingTicket As IPCTicket
        
        Call Ticket.ReadFromRow(Row)
        
        'Check for ticket
        Set ExistingTicket = Tickets.Find(Ticket.ChangeID)
               
        If (ExistingTicket Is Nothing) Then
            Tickets.Add Ticket
            'Print Ticket
            Call Ticket.PrintTicket
        Else
            Call ExistingTicket.Merge(Ticket)
            Call ExistingTicket.PrintTicket
        End If
    
        'Check
        Debug.Print "Tickets Size: " & Tickets.Size()
        
        'Clear values
        Set Ticket = Nothing
        Set ExistingTicket = Nothing
        
        Index = Index + 1
        
        'Safety exit
        If Index > 10 Then
            'Exit Sub
        End If
                
    Loop
    
    Call WriteToSheet(Output, Tickets, 2)
    
End Sub

Sub WriteToSheet(Output As Worksheet, Tickets As IPCTickets, Optional Index As Integer = 1)

    Call Output.Cells.ClearContents
    Output.Cells(1, 1) = "Cache Id"
    Output.Cells(1, 2) = "Type"
    Output.Cells(1, 3) = "Start Time"
    Output.Cells(1, 4) = "End Time"
    Output.Cells(1, 5) = "Summary"
    Output.Cells(1, 6) = "Impact"
    Output.Cells(1, 6) = "Requestor Name"
    
    
    Dim Ticket As IPCTicket

    For Each Ticket In Tickets.All
        Dim Row As Range
        Set Row = Output.Rows(Index)
        Call Ticket.WriteToRow(Row)
        
        Index = Index + 1
        
        Set Row = Nothing
    Next

End Sub
