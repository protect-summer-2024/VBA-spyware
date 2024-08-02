# VBA-spyware
 Functionality of a VBA macro designed to automate the insertion of hidden information into a Word document. This macro fetches the MAC address, IP address, username, and timestamp, and displays this information in a header text box.

# VBA Document Automation

This project involves a VBA macro designed to automate the insertion of hidden information (MAC address, IP address, username, and timestamp) into a Word document header.

## Table of Contents
1. [Introduction](#introduction)
2. [System Architecture](#system-architecture)
3. [Implementation](#implementation)
4. [Testing and Results](#testing-and-results)
5. [Ethical Considerations](#ethical-considerations)
6. [Future Enhancements](#future-enhancements)
7. [Conclusion](#conclusion)
8. [Uses of VBA Automation](#uses-of-vba-automation)
9. [Disclaimer](#disclaimer)
10. [Limitations](#limitations)
11. [Appendix](#appendix)
12. [Installation](#installation)
13. [References](#references)
14. [Security](#security)


## Introduction

### Background
This VBA macro automates the insertion of hidden information into a Word document. It fetches the MAC address, IP address, username, and timestamp, and displays this information in a header text box.

## System Architecture

### High-Level Architecture
The architecture consists of the following components:
- `Document_Open`: Triggers the insertion of hidden information when the document is opened.
- `ContentControlOnEnter`: Activates when a specific content control is entered.
- `ContentControlOnExit`: Activates when a specific content control is exited.
- `InsertHiddenInfo`: Inserts the hidden information into the document header.
- **Helper Functions**: Functions to get the username, MAC address, and IP address.

### Component Descriptions
- `Document_Open`: Ensures hidden information is inserted upon opening the document.
- `ContentControlOnEnter/Exit`: Provides real-time updates when interacting with specific content controls.
- `InsertHiddenInfo`: Gathers system information and inserts it into a text box in the document header.

## Implementation

### Environment Setup
1. **Microsoft Word**: Ensure Microsoft Word is installed on your computer.
2. **VBA Editor**: Access the VBA editor through Microsoft Word. No additional libraries are required.
3. **Developer Mode**: Enable Developer Mode in Microsoft Word to access the VBA editor and other developer tools.

#### Steps to Enable Developer Mode
1. Open Microsoft Word.
2. Go to the File Menu: Click on File in the top left corner.
3. Access Options: Select Options at the bottom of the left-hand menu.
4. Customize Ribbon: In the Word Options dialog, go to Customize Ribbon.
5. Enable Developer Tab: Check the box next to Developer in the right column under Main Tabs.
6. Close and Apply: Click OK to close the dialog and apply the changes.

The Developer tab should now be visible in the ribbon, providing access to the VBA editor and other developer tools.




### Key Code Snippets

```vba
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
```
## Testing and Results

### Test Scenarios
- **Scenario 1**: Open the document and verify that the hidden information is inserted correctly.
- **Scenario 2**: Enter and exit the specific content control to ensure real-time updates.
- **Scenario 3**: Test on different systems to confirm consistent functionality.

### Results Analysis
- **Scenario 1**: Successful insertion of hidden information upon document opening.
- **Scenario 2**: Real-time updates confirmed with accurate information.
- **Scenario 3**: Consistent results across multiple systems.

## Ethical Considerations

### Legal Implications
Unauthorized access or use of system information can violate privacy laws and regulations.

### Responsible Usage
Users must ensure compliance with local laws and obtain necessary permissions before using this macro.

## Future Enhancements

### Planned Features
- **Enhanced User Interface**: Develop a more sophisticated GUI for easier control and monitoring.
- **Dynamic Content Controls**: Allow dynamic interaction with various content controls.

### Potential Improvements
- **Integration with Other Office Applications**: Enable interoperability with other Microsoft Office applications for comprehensive functionality.

## Conclusion
This project demonstrates the potential of VBA macros to automate tasks within Microsoft Word documents. By embedding hidden information such as MAC address, IP address, username, and timestamp, this macro showcases how VBA can enhance document automation and security features.

## Uses of VBA Automation
This VBA macro can be used in various scenarios including:
- Automated document generation with embedded metadata.
- Dynamic content insertion based on user interaction.
- Enhanced document security and tracking by embedding system information.

## Disclaimer
This VBA macro is intended for educational and informational purposes only. Users must ensure compliance with local laws and obtain necessary permissions before using it. The author is not responsible for any misuse of this macro.

## Limitations
- The macro is designed to work with Microsoft Word documents only.
- It may not function as intended in other document formats or software.
- The macro requires the user to enable macros in Microsoft Word, which may pose a security risk.

## Appendix
For additional resources and further reading on VBA macros and automation, refer to the official Microsoft documentation and other reputable sources.

## Installation
1. Download and open the Word document containing the VBA macro.
2. Enable macros when prompted.
3. The macro will automatically run and insert the hidden information into the document header.

## References
- Microsoft Documentation on VBA: [https://docs.microsoft.com/en-us/office/vba/api/overview/](https://docs.microsoft.com/en-us/office/vba/api/overview/)
- VBA Macro Security: [https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/macro-security](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/macro-security)

## Security
Ensure that you understand the security implications of enabling and running macros in Microsoft Word. Only run macros from trusted sources and be aware of the potential risks involved.
