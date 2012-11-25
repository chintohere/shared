Option Explicit

'All columns for each row in the new sheet
Public ChangeID As String
Public ChangeType As String
Public StratTime As String
Public EndTime As String
Public Summary As String
Public Impact As String
Public RequesterName As String

'This should do the concatenation bit
Public Sub Merge(ByRef Ticket As IPCTicket)
    Me.Impact = Me.Impact & Ticket.Impact
End Sub






