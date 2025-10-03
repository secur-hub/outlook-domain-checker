' === DICHIARAZIONI API DNS ===
Private Declare PtrSafe Function DnsQuery Lib "dnsapi.dll" Alias "DnsQuery_W" ( _
    ByVal lpstrName As LongPtr, _
    ByVal wType As Integer, _
    ByVal fOptions As Long, _
    ByVal aipServers As LongPtr, _
    ByRef ppQueryResultsSet As LongPtr, _
    ByVal pReserved As LongPtr) As Long

Private Declare PtrSafe Sub DnsRecordListFree Lib "dnsapi.dll" ( _
    ByVal pRecordList As LongPtr, _
    ByVal FreeType As Integer)

Private Const DNS_TYPE_MX As Integer = &HF    ' 15 decimale
Private Const DNS_TYPE_A As Integer = 1
Private Const DNS_TYPE_AAAA As Integer = 28
Private Const DNS_QUERY_STANDARD As Long = 0&
Private Const DNS_FREE_RECORDLIST As Integer = 1

' === FUNZIONE PER VERIFICARE SE IL DOMINIO ESISTE ===
Private Function DomainExists(domain As String) As Boolean
    Dim pResults As LongPtr
    Dim ret As Long
    
    On Error GoTo ErrHandler
    
    ' 1) Query MX
    ret = DnsQuery(StrPtr(domain), DNS_TYPE_MX, DNS_QUERY_STANDARD, 0, pResults, 0)
    If ret = 0 And pResults <> 0 Then
        DnsRecordListFree pResults, DNS_FREE_RECORDLIST
        DomainExists = True
        Exit Function
    End If
    
    ' 2) Query A
    ret = DnsQuery(StrPtr(domain), DNS_TYPE_A, DNS_QUERY_STANDARD, 0, pResults, 0)
    If ret = 0 And pResults <> 0 Then
        DnsRecordListFree pResults, DNS_FREE_RECORDLIST
        DomainExists = True
        Exit Function
    End If
    
    ' 3) Query AAAA
    ret = DnsQuery(StrPtr(domain), DNS_TYPE_AAAA, DNS_QUERY_STANDARD, 0, pResults, 0)
    If ret = 0 And pResults <> 0 Then
        DnsRecordListFree pResults, DNS_FREE_RECORDLIST
        DomainExists = True
        Exit Function
    End If
    
    ' Nessun record trovato
    DomainExists = False
    Exit Function
    
ErrHandler:
    DomainExists = False
End Function

' === EVENTO IN OUTLOOK: PRIMA DELL'INVIO ===
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim rec As Outlook.Recipient
    Dim addr As String, domain As String
    Dim i As Integer
    
    ' Controlla tutti i destinatari
    For Each rec In Item.Recipients
        addr = rec.Address
        ' Se è nel formato SMTP classico (con @)
        If InStr(1, addr, "@") > 0 Then
            domain = Mid$(addr, InStrRev(addr, "@") + 1)
            If Not DomainExists(domain) Then
                MsgBox "Invio annullato: il dominio """ & domain & """ non esiste o non ha record DNS validi.", vbCritical
                Cancel = True
                Exit For
            End If
        End If
    Next
End Sub

