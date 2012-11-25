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
Public StratTime As String
Public EndTime As String
Public Summary As String
Public Impact As String
Public RequesterName As String

'This should do the concatenation bit
Public Sub Merge(ByRef Ticket As IPCTicket)
    Me.Impact = Me.Impact & Ticket.Impact
End Sub






