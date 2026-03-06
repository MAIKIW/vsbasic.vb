Sub SplitDataSuperCleanerFixed()
    Dim wsMain As Worksheet, wsDest As Worksheet
    Dim lastRowMain As Long, i As Long
    Dim destRow As Long
    Dim rawName As String, safeSheetName As String, dictKey As String
    Dim dictSheets As Object
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsMain = ThisWorkbook.Sheets("Main")
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "E").End(xlUp).Row
    
    If lastRowMain < 3 Then
        MsgBox "No data", vbExclamation
        Exit Sub
    End If
    
    Set dictSheets = CreateObject("Scripting.Dictionary")
    
    For i = 3 To lastRowMain
        rawName = wsMain.Cells(i, 5).Value
        
        If Trim(rawName) <> "" Then
            
            ' 1. ทำความสะอาดเบื้องต้นสำหรับตั้งชื่อชีท
            safeSheetName = Application.WorksheetFunction.Trim(rawName)
            safeSheetName = Left(safeSheetName, 31)
            safeSheetName = Replace(safeSheetName, "/", "_")
            safeSheetName = Replace(safeSheetName, "\", "_")
            safeSheetName = Replace(safeSheetName, "?", "")
            safeSheetName = Replace(safeSheetName, "*", "")
            safeSheetName = Replace(safeSheetName, "[", "")
            safeSheetName = Replace(safeSheetName, "]", "")
            safeSheetName = Replace(safeSheetName, ":", "")
            
            ' 2. เครื่องซักล้างคำขั้นเด็ดขาด
            dictKey = UCase(safeSheetName)
            dictKey = Replace(dictKey, " ", "")
            dictKey = Replace(dictKey, "-", "")
            dictKey = Replace(dictKey, "S", "") 
            
            ' จัดกลุ่มพิเศษสำหรับกิจกรรมร่วมค้า
            If InStr(dictKey, "กิจก") > 0 Or InStr(dictKey, "กิจก") > 0 Then
                dictKey = "JOINTVENTURE"
                safeSheetName = "กิจกรรมร่วมการค้า"
            End If
            
            ' เช็กว่าในหน่วยความจำมีกุญแจนี้หรือยัง
            If Not dictSheets.Exists(dictKey) Then
                
                ' ****************************************************
                ' จุดแก้บั๊กสำคัญ: สั่งให้ล้างความจำ (ทิ้งชีทเก่าในมือ) ทุกครั้ง
                Set wsDest = Nothing 
                ' ****************************************************
                
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
                Set dictSheets(dictKey) = wsDest
            End If
            
            Set wsDest = dictSheets(dictKey)
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
    
    MsgBox "finish", vbInformation
End Sub
