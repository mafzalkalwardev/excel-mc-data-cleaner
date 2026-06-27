Sub MC_formatting()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found!", vbExclamation
        GoTo Cleanup
    End If

    ' Create new worksheet for final result
    Dim wsFinal As Worksheet
    Set wsFinal = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wsFinal.Name = "MC_Cleaned"
    
    ' Copy headers
    wsFinal.Cells(1, 1).Value = "MC Number"
    wsFinal.Cells(1, 2).Value = "Legal Name"
    wsFinal.Cells(1, 3).Value = "States"
    wsFinal.Cells(1, 4).Value = "Address"
    wsFinal.Cells(1, 5).Value = "Phone"
    wsFinal.Cells(1, 6).Value = "Email"

    ' Copy data from original columns and validate emails
    Dim outRow As Long: outRow = 2
    Dim validEmailCount As Long: validEmailCount = 0
    
    For i = 2 To lastRow
        ' CORRECTED: Column X (24) = Email
        Dim emailVal As String
        emailVal = Trim(ws.Cells(i, 24).Value) ' Column X = Email
        
        ' CORRECTED: Column W (23) = Phone
        Dim phoneVal As String
        phoneVal = Trim(ws.Cells(i, 23).Value) ' Column W = Phone
        
        ' Skip rows with invalid or empty emails
        If emailVal = "" Or InStr(emailVal, "@") = 0 Or InStr(emailVal, ".") = 0 Then
            GoTo NextRow
        End If
        
        ' Copy data - WITH CORRECTED COLUMNS
        wsFinal.Cells(outRow, 1).Value = ws.Cells(i, 11).Value  ' Column K = DBA Name -> MC Number
        wsFinal.Cells(outRow, 2).Value = ws.Cells(i, 21).Value  ' Column U = Legal Name
        wsFinal.Cells(outRow, 4).Value = ws.Cells(i, 22).Value  ' Column V = Address
        wsFinal.Cells(outRow, 5).Value = phoneVal               ' Column W = Phone
        wsFinal.Cells(outRow, 6).Value = emailVal               ' Column X = Email
        
        ' Extract state from address
        Dim addressText As String
        addressText = Trim(ws.Cells(i, 22).Value) ' Column V = Address
        If addressText <> "" Then
            Dim commaPos As Long
            commaPos = InStrRev(addressText, ",")
            If commaPos > 0 Then
                Dim afterComma As String
                afterComma = Trim(Mid(addressText, commaPos + 1))
                If Len(afterComma) >= 2 Then
                    Dim stateCode As String
                    stateCode = UCase(Mid(afterComma, 1, 2))
                    stateCode = Replace(stateCode, " ", "")
                    stateCode = Replace(stateCode, vbCr, "")
                    stateCode = Replace(stateCode, vbLf, "")
                    If Len(stateCode) = 2 And stateCode Like "[A-Z][A-Z]" Then
                        wsFinal.Cells(outRow, 3).Value = stateCode
                    End If
                End If
            End If
        End If
        
        outRow = outRow + 1
        validEmailCount = validEmailCount + 1
        
NextRow:
    Next i

    ' Format as table
    lastRow = wsFinal.Cells(wsFinal.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then
        Dim tbl As ListObject
        Set tbl = wsFinal.ListObjects.Add(xlSrcRange, wsFinal.Range("A1:F" & lastRow), , xlYes)
        tbl.Name = "Table1"
        tbl.TableStyle = "TableStyleLight15"
        
        ' Adjust column widths
        wsFinal.Columns("C").ColumnWidth = 8
        wsFinal.Columns("D").ColumnWidth = 30
        wsFinal.Columns.AutoFit
        
        MsgBox "Success! " & validEmailCount & " records with valid emails processed.", vbInformation
    Else
        MsgBox "No valid data found after email validation.", vbExclamation
    End If

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
