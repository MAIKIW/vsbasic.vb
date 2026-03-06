Sub SplitDataExactLayout()
    Dim wsMain As Worksheet, wsDest As Worksheet
    Dim lastRowMain As Long, i As Long
    Dim destRow As Long
    Dim studioName As String, safeSheetName As String
    Dim dictSheets As Object
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "E").End(xlUp).Row
    
    If lastRowMain < 3 Then
        MsgBox "ไม่พบข้อมูลหนังตั้งแต่แถวที่ 3 เป็นต้นไปครับ", vbExclamation
        Exit Sub
    End If
    
    Set dictSheets = CreateObject("Scripting.Dictionary")
    
    ' *** อัปเกรดความฉลาด: สั่งให้หน่วยความจำไม่สนใจตัวพิมพ์เล็ก/ใหญ่ (A = a) ***
    dictSheets.CompareMode = 1 
    
    For i = 3 To lastRowMain
        ' *** อัปเกรดการทำความสะอาด: ลบช่องว่างส่วนเกินทั้งหน้า หลัง และตรงกลางคำทิ้ง ***
        studioName = Application.WorksheetFunction.Trim(wsMain.Cells(i, 5).Value)
        
        If studioName <> "" Then
            
            safeSheetName = Left(studioName, 31)
            safeSheetName = Replace(safeSheetName, "/", "_")
            safeSheetName = Replace(safeSheetName, "\", "_")
            safeSheetName = Replace(safeSheetName, "?", "")
            safeSheetName = Replace(safeSheetName, "*", "")
            safeSheetName = Replace(safeSheetName, "[", "")
            safeSheetName = Replace(safeSheetName, "]", "")
            safeSheetName = Replace(safeSheetName, ":", "")
            
            If Not dictSheets.Exists(safeSheetName) Then
                On Error Resume Next
                Set wsDest = ThisWorkbook.Sheets(safeSheetName)
                On Error GoTo 0
                
                If wsDest Is Nothing Then
                    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                    wsDest.Name = safeSheetName
                Else
                    wsDest.Cells.Clear
                End If
                
                wsMain.Rows("1:2").Copy Destination:=wsDest.Rows("1:2")
                Set dictSheets(safeSheetName) = wsDest
            End If
            
            Set wsDest = dictSheets(safeSheetName)
            destRow = wsDest.Cells(wsDest.Rows.Count, "E").End(xlUp).Row + 1
            
            wsMain.Range("A" & i & ":E" & i).Copy Destination:=wsDest.Range("A" & destRow)
            
        End If
    Next i
    
    Dim key As Variant
    For Each key In dictSheets.keys
        Set wsDest = dictSheets(key)
        wsDest.Columns("C:D").NumberFormat = "dd/mm/yyyy"
        wsDest.Columns.AutoFit
    Next key
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "แยกข้อมูลสำเร็จ! ขจัดปัญหาตัวพิมพ์เล็ก-ใหญ่และช่องว่างเรียบร้อยครับ", vbInformation
End Sub
