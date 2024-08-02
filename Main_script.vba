
Private Sub Document_Open()
    On Error GoTo ErrorHandler
    InsertHiddenInfo
    Exit Sub

ErrorHandler:
    MsgBox "Error occurred: " & Err.Description
End Sub

Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
    ' Trigger macro when entering a specific content control
    If ContentControl.Title = "TriggerControl" Then
        MsgBox "TriggerControl Entered!" ' Confirmation message
        InsertHiddenInfo
    End If
End Sub

Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    ' Trigger macro when exiting a specific content control
    If ContentControl.Title = "TriggerControl" Then
        MsgBox "TriggerControl Exited!" ' Confirmation message
        InsertHiddenInfo
    End If
End Sub

Sub InsertHiddenInfo()
    MsgBox "InsertHiddenInfo is running!" ' Confirmation message
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Get information
    Dim macAddress As String
    macAddress = GetMyMACAddress
    
    Dim ipAddress As String
    ipAddress = GetMyLocalIP
    
    Dim username As String
    username = GetUsername
    
    Dim timestamps As String
    timestamps = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Access the primary header
    Dim header As HeaderFooter
    Set header = doc.Sections(1).Headers(wdHeaderFooterPrimary)
    
    ' Remove existing text boxes in the header
    Dim shp As Shape
    For Each shp In header.Shapes
        If shp.Type = msoTextBox Then shp.Delete
    Next shp
    
    ' Add a new text box to the header
    Dim textBox As Shape
    Set textBox = header.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                           Left:=10, Top:=10, Width:=200, Height:=80)
                                           
    ' Set text box properties
    With textBox
        .TextFrame.TextRange.Text = "MAC Address: " & macAddress & vbCrLf & _
                                    "IP Address: " & ipAddress & vbCrLf & _
                                    "Username: " & username & vbCrLf & _
                                    "Timestamp: " & timestamps
        .TextFrame.TextRange.Font.Size = 8
        .TextFrame.TextRange.Font.Color = RGB(255, 255, 255) ' White font
        .Fill.Visible = msoFalse ' No fill
        .Line.Visible = msoFalse ' No border
        .TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphLeft ' Align text
        .Top = 10 ' Distance from the top of the header
        
        ' Position the text box in the right corner of the header
        Dim pageWidth As Single
        Dim textBoxWidth As Single
        pageWidth = doc.PageSetup.PageWidth
        textBoxWidth = .Width
        .Left = pageWidth - textBoxWidth - 10 ' 10 points from the right edge
    End With
End Sub

Function GetUsername() As String
    GetUsername = Environ$("USERNAME")
End Function

Function GetMyMACAddress() As String
    Dim oWMI As Object
    Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Dim oCols As Object
    Set oCols = oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
    Dim oCol As Object
    For Each oCol In oCols
        If oCol.IPEnabled Then
            GetMyMACAddress = oCol.macAddress
            Exit Function
        End If
    Next oCol
End Function

Function GetMyLocalIP() As String
    Dim oWMI As Object
    Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Dim oCols As Object
    Set oCols = oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
    Dim oCol As Object
    For Each oCol In oCols
        If oCol.IPEnabled Then
            GetMyLocalIP = oCol.ipAddress(0)
            Exit Function
        End If
    Next oCol
End Function
