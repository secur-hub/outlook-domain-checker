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

Private Const DNS_TYPE_MX As Integer = &HF   ' 15 decimale
Private Const DNS_TYPE_A As Integer = 1
Private Const DNS_TYPE_AAAA As Integer = 28
Private Const DNS_QUERY_STANDARD As Long = 0&
Private Const DNS_FREE_RECORDLIST As Integer = 1


' === FUNZIONE PER VERIFICARE SE IL DOMINIO ESISTE ===
Private Function DomainExists(domain As String) As Boolean
    Dim pResults As LongPtr
    Dim ret As Long
    
    On Error GoTo ErrHandler
    
    ' Query MX
    ret = DnsQuery(StrPtr(domain), DNS_TYPE_MX, DNS_QUERY_STANDARD, 0, pResults, 0)
    If ret = 0 And pResults <> 0 Then
        DnsRecordListFree pResults, DNS_FREE_RECORDLIST
        DomainExists = True
        Exit Function
    End If
    
    ' Query A
    ret = DnsQuery(StrPtr(domain), DNS_TYPE_A, DNS_QUERY_STANDARD, 0, pResults, 0)
    If ret = 0 And pResults <> 0 Then
        DnsRecordListFree pResults, DNS_FREE_RECORDLIST
        DomainExists = True
        Exit Function
    End If
    
    ' Query AAAA
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


' === FUNZIONE DI PULIZIA INDIRIZZI EMAIL ===
Private Function CleanEmailAddress(rawAddr As String) As String
    Dim addr As String
    
    ' Rimuove spazi e apici singoli
    addr = Trim(Replace(rawAddr, "'", ""))
    
    ' Rimuove eventuali caratteri invisibili (Unicode LRM, ecc.)
    addr = Replace(addr, ChrW(&H200E), "")
    addr = Replace(addr, ChrW(&H200F), "")
    
    ' Se ha nome visualizzato tipo "Mario Rossi" <mario@azienda.it>
    If InStr(addr, "<") > 0 And InStr(addr, ">") > 0 Then
        addr = Mid$(addr, InStr(addr, "<") + 1, InStr(addr, ">") - InStr(addr, "<") - 1)
    End If
    
    ' Rimuove spazi residui
    addr = Trim(addr)
    
    CleanEmailAddress = addr
End Function


' === EVENTO IN OUTLOOK: PRIMA DELL'INVIO ===
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim rec As Outlook.Recipient
    Dim addr As String, domain As String
    
    For Each rec In Item.Recipients
        addr = CleanEmailAddress(rec.Address)
        
        ' Controlla che contenga una @ valida
        If InStr(1, addr, "@") > 0 Then
            domain = Mid$(addr, InStrRev(addr, "@") + 1)
            
            If Not DomainExists(domain) Then
                MsgBox "Invio annullato: il dominio """ & domain & """ non esiste o non ha record DNS validi.", vbCritical
                Cancel = True
                Exit For
            End If
        Else
            MsgBox "Invio annullato: indirizzo email non valido (" & addr & ").", vbExclamation
            Cancel = True
            Exit For
        End If
    Next
End Sub



