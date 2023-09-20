Sub fullReminder()
    'declare variables
    Dim nextPMList As Range
    Dim nextPM As Range
    Dim PM As Range
    Dim diff As Integer
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim strbody As String
    Dim strbody6 As String
    Dim strbody3 As String
    Dim strbody1 As String
    Dim strbodyD As String
    Dim i As Integer
    
    'assign variables
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    curDate = Date
    Set nextPMList = Range("I2:I150")
    i = 1
    Count = 0
    curRow = 1
    For Each nextPM In nextPMList
        curRow = curRow + 1
        nextPMDate = nextPM.Value
        'check if "PM DUE"
        If nextPMDate = "PM DUE" Then
            Count = Count + 1
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than a month"
            strbodyD = nextPM.EntireRow.Cells(2) & nextPMDate
            strbody = strbody & vbNewLine & strbodyD
        End If
        
        If Not nextPMDate = "" And Not nextPMDate = "N/A" And IsDate(nextPMDate) Then
        diff = DateDiff("d", curDate, nextPMDate)
            'check if next PM is less than 6 months away
            If diff < 180 And diff > 130 Then
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than 6 months"
            strbody6 = nextPM.EntireRow.Cells(2) & " PM due on " & nextPMDate & " in less than 6 months"
            Count = Count + 1
            strbody = strbody & vbNewLine & strbody6
            End If
            
            'check if next pm is less than 3 months away
            If diff < 90 And diff > 58 Then
            Count = Count + 1
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than 3 months"
            strbody3 = nextPM.EntireRow.Cells(2) & " PM due on " & nextPMDate & " in less than 3 months"
            strbody = strbody & vbNewLine & strbody3
            End If
            
            'check if next PM is less than 1 month away
            If diff < 32 And diff > 0 Then
            Count = Count + 1
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than a month"
            strbody1 = nextPM.EntireRow.Cells(2) & " PM due on " & nextPMDate & " in less than a month"
            strbody = strbody & vbNewLine & strbody1
            End If
        End If
        
    Next
    Debug.Print strbody
    'Debug.Print Count
    
    If Not IsEmpty(strbody) Then
        On Error Resume Next
        With OutMail
            .To = "test@outlook.com"
            .CC = ""
            .BCC = ""
            .Subject = "Things are due"
            .Body = strbody
            .send
            .Display
        End With
        On Error GoTo 0
 
        Set OutMail = Nothing
        Set OutApp = Nothing
    End If
End Sub

      Sub oneMonth()
    'declare variables
    Dim nextPMList As Range
    Dim nextPM As Range
    Dim PM As Range
    Dim diff As Integer
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim strbody As String
    Dim strbody6 As String
    Dim strbody3 As String
    Dim strbody1 As String
    Dim strbodyD As String
    Dim i As Integer
    
    'assign variables
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    curDate = Date
    Set nextPMList = Range("I2:I150")
    i = 1
    Count = 0
    curRow = 1
    For Each nextPM In nextPMList
        curRow = curRow + 1
        nextPMDate = nextPM.Value
        'check if "PM DUE"
        If nextPMDate = "PM DUE" Then
            Count = Count + 1
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than a month"
            strbodyD = nextPM.EntireRow.Cells(2) & " " & nextPMDate
            strbody = strbody & vbNewLine & strbodyD
        End If
        
        If Not nextPMDate = "" And Not nextPMDate = "N/A" And IsDate(nextPMDate) Then
        diff = DateDiff("d", curDate, nextPMDate)
            'check if next PM is less than 1 month away
            If diff < 32 And diff > 0 Then
            Count = Count + 1
            'Debug.Print curRow
            'Debug.Print nextPM.EntireRow.Cells(2); "PM due on "; nextPMDate; " in less than a month"
            strbody1 = nextPM.EntireRow.Cells(2) & " PM due on " & nextPMDate & " in less than a month"
            strbody = strbody & vbNewLine & strbody1
            End If
        End If
        
    Next
    Debug.Print strbody
    'Debug.Print Count
    
    If Not IsEmpty(strbody) Then
        On Error Resume Next
        With OutMail
            .To = "test@outlook.com"
            .CC = ""
            .BCC = ""
            .Subject = "Upcoming Projects Due"
            .Body = strbody
            .send
            .Display
        End With
        On Error GoTo 0
 
        Set OutMail = Nothing
        Set OutApp = Nothing
    End If
End Sub

Private Sub CommandButton1_Click()
    Call fullReminder
End Sub

Private Sub CommandButton2_Click()
    Call oneMonth
End Sub
