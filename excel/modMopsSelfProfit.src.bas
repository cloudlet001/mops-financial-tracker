Attribute VB_Name = "modMopsSelfProfit"
Option Explicit
'==========================================================================
' MOPS financial-stock monthly self-reported profit workbook (pure VBA)
'  1. Fetch latest month self-profit list (ajax_t138sb01), extract financial
'     stocks -> sheet [list]
'  2. For each target stock fetch per-company table (ajax_t138sb02) for
'     previous + current year, month aligned -> one sheet per stock
'  3. Back-fill latest month values into [list]
' Entry point: RunFinancialSelfProfit
'==========================================================================

Private Const API_URL As String = "https://mops.twse.com.tw/mops/api/redirectToOld"
Private Const UA As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
Private Const DELAY_SEC As Double = 1.5
Private Const MAX_RETRY As Long = 3

Private logRow As Long
Private gSilent As Boolean

' Used for automated testing (no message boxes)
Public Sub RunSilent()
    gSilent = True
    RunFinancialSelfProfit
    gSilent = False
End Sub

'---------------------------------------------------------------- names
Private Function CFG_NAME() As String: CFG_NAME = "設定": End Function
Private Function LIST_NAME() As String: LIST_NAME = "自結清單": End Function
Private Function FIN_MARK() As String: FIN_MARK = "不適用": End Function

'================================================================ main
Public Sub RunFinancialSelfProfit()
    Dim cfg As Worksheet
    Set cfg = ThisWorkbook.Worksheets(CFG_NAME())

    Dim market As String, ymIn As String
    market = Trim$(CStr(cfg.Range("B3").Value))
    If market = "" Then market = "1"
    ymIn = Trim$(CStr(cfg.Range("B4").Value))

    Dim targets As New Collection, r As Long
    r = 7
    Do While Trim$(CStr(cfg.Cells(r, 2).Value)) <> ""
        targets.Add Trim$(CStr(cfg.Cells(r, 2).Value))
        r = r + 1
    Loop

    cfg.Range("D6:D300").ClearContents
    logRow = 6
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    On Error GoTo EH

    SetStatus "執行中..."
    Dim ym As String, lst As Collection
    If ymIn <> "" Then
        ym = ymIn
        LogMsg "查詢指定年月 " & ym
        Set lst = FetchList(ym, market)
        If lst.Count = 0 Then Err.Raise vbObjectError + 1, , "指定年月查無資料: " & ym
    Else
        ym = FindLatestYM(market, lst)
    End If

    Dim curY As Long, curM As Long, prevY As Long
    curY = CLng(Left$(ym, 3)): curM = CLng(Mid$(ym, 4, 2)): prevY = curY - 1
    LogMsg "最新月份: " & curY & "年" & curM & "月, 清單共 " & lst.Count & " 筆"

    '------------------------------------------------ financial filter
    Dim fin As New Collection, rec As Variant
    For Each rec In lst
        If InStr(rec(3), FIN_MARK()) > 0 Then fin.Add rec
    Next
    LogMsg "金融股共 " & fin.Count & " 檔"

    DeleteOutputSheets

    '------------------------------------------------ list sheet
    Dim ymLbl As String
    ymLbl = curY & "/" & Format$(curM, "00")
    Dim wsL As Worksheet
    Set wsL = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsL.Name = LIST_NAME()
    Dim hdr As Variant
    hdr = Array("金融股", "報表類型", ymLbl & " 本月稅前損益(清單)", ymLbl & " 本年累計稅前損益(清單)", _
                prevY & "年度 本月稅前損益", prevY & "年度 本年累計稅前損益", _
                curY & "年度 本月稅前損益", curY & "年度 本年累計稅前損益", "累計稅前 YoY")
    wsL.Range("A1").Resize(1, 9).Value = hdr
    Dim i As Long
    If fin.Count > 0 Then
        Dim arr() As Variant
        ReDim arr(1 To fin.Count, 1 To 4)
        i = 0
        For Each rec In fin
            i = i + 1
            arr(i, 1) = rec(0) & " " & rec(1)
            arr(i, 2) = rec(2)
            arr(i, 3) = ToNum(rec(5))
            arr(i, 4) = ToNum(rec(6))
        Next
        wsL.Range("A2").Resize(fin.Count, 4).Value = arr
        wsL.Range("A1").Resize(fin.Count + 1, 9).Sort Key1:=wsL.Range("A2"), Order1:=xlAscending, Header:=xlYes
    End If

    '------------------------------------------------ per-stock sheets
    Dim co As Variant, nm As String, nm2 As String
    Dim pMon(1 To 12) As Variant, pCum(1 To 12) As Variant
    Dim cMon(1 To 12) As Variant, cCum(1 To 12) As Variant
    For Each co In targets
        SetStatus "抓取 " & co & " ..."
        LogMsg "抓取 " & co & " 的 " & prevY & "/" & curY & " 個股自結損益"
        Erase pMon: Erase pCum: Erase cMon: Erase cCum
        FetchCompanyYear CStr(co), prevY, nm, pMon, pCum
        FetchCompanyYear CStr(co), curY, nm2, cMon, cCum
        If nm = "" Then nm = nm2

        Dim wsC As Worksheet
        Set wsC = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsC.Name = SafeSheetName(co & " " & nm)
        wsC.Range("A1").Resize(1, 7).Value = Array("月份", _
            prevY & "年度 本月稅前損益", prevY & "年度 本年累計稅前損益", _
            curY & "年度 本月稅前損益", curY & "年度 本年累計稅前損益", "單月稅前 YoY", "累計稅前 YoY")
        Dim body(1 To 12, 1 To 7) As Variant, m As Long
        For m = 1 To 12
            body(m, 1) = m & "月"
            body(m, 2) = pMon(m): body(m, 3) = pCum(m)
            body(m, 4) = cMon(m): body(m, 5) = cCum(m)
            body(m, 6) = YoY(cMon(m), pMon(m))
            body(m, 7) = YoY(cCum(m), pCum(m))
        Next
        wsC.Range("A2").Resize(12, 7).Value = body
        FormatSheet wsC

        '-------------- back-fill
        Dim rowL As Long
        rowL = FindStockRow(wsL, CStr(co))
        If rowL = 0 Then
            rowL = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row + 1
            wsL.Cells(rowL, 1).Value = co & " " & nm
            LogMsg "  [!] " & co & " 不在金融股清單中, 已新增一列"
        End If
        wsL.Cells(rowL, 5).Value = pMon(curM)
        wsL.Cells(rowL, 6).Value = pCum(curM)
        wsL.Cells(rowL, 7).Value = cMon(curM)
        wsL.Cells(rowL, 8).Value = cCum(curM)
        wsL.Cells(rowL, 9).Value = YoY(cCum(curM), pCum(curM))

        Dim ok As Boolean
        ok = (CStr(wsL.Cells(rowL, 3).Value) = CStr(cMon(curM))) And (CStr(wsL.Cells(rowL, 4).Value) = CStr(cCum(curM)))
        LogMsg "  " & co & " " & nm & ": 交叉驗證 " & IIf(ok, "OK", "不一致")
    Next

    FormatSheet wsL
    wsL.Activate
    wsL.Range("A1").Select
    SetStatus "完成 (" & ymLbl & ", 金融股 " & fin.Count & " 檔, 個股 " & targets.Count & " 檔)  " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
    LogMsg "完成"
    GoTo Done

EH:
    SetStatus "錯誤: " & Err.Description
    LogMsg "錯誤: " & Err.Description
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    If Not gSilent Then MsgBox "執行失敗: " & Err.Description, vbExclamation
    Exit Sub
Done:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

'================================================================ fetchers
Private Function FindLatestYM(ByVal market As String, ByRef lst As Collection) As String
    Dim y As Long, m As Long, k As Long, ym As String
    y = Year(Date) - 1911: m = Month(Date)
    For k = 1 To 6
        m = m - 1
        If m = 0 Then y = y - 1: m = 12
        ym = y & Format$(m, "00")
        LogMsg "嘗試 " & ym & " ..."
        Set lst = FetchList(ym, market)
        If lst.Count > 0 Then
            FindLatestYM = ym
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 2, , "找不到最近 6 個月的月自結資料"
End Function

' returns Collection of Array(code, name, type, op, cumOp, pretax, cumPretax)
Private Function FetchList(ByVal ym As String, ByVal market As String) As Collection
    Dim html As String, segs() As String, seg As String
    Dim rtype As String, i As Long, res As New Collection
    html = MopsGet("ajax_t138sb01", "{""CK2"":""1"",""YM1"":""" & ym & """,""CK1"":""" & market & _
                   """,""encodeURIComponent"":1,""firstin"":1,""step"":1,""first"":1}")
    segs = Split(html, "<table")
    rtype = "個別"
    For i = 1 To UBound(segs)
        seg = segs(i)
        If InStr(seg, "（合併）") > 0 Then rtype = "合併"
        If InStr(seg, "（個別）") > 0 Then rtype = "個別"
        If InStr(seg, "公司代號") > 0 Then
            Dim rows As Object, rw As Object, tds As Object
            Set rows = RegexMatches(seg, "<tr[^>]*>([\s\S]*?)</tr>")
            For Each rw In rows
                Set tds = RegexMatches(rw.SubMatches(0), "<td[^>]*>([\s\S]*?)</td>")
                If tds.Count = 7 Then
                    res.Add Array(CleanText(tds(0).SubMatches(0)), CleanText(tds(1).SubMatches(0)), rtype, _
                                  CleanText(tds(2).SubMatches(0)), CleanText(tds(3).SubMatches(0)), _
                                  CleanText(tds(4).SubMatches(0)), CleanText(tds(5).SubMatches(0)))
                End If
            Next
        End If
    Next
    Set FetchList = res
End Function

Private Sub FetchCompanyYear(ByVal co As String, ByVal yr As Long, ByRef nm As String, _
                             ByRef mon() As Variant, ByRef cum() As Variant)
    Dim q1 As String, p As String, k As Variant, v As String
    q1 = MopsGet("ajax_t138sb02_q1", "{""co_id"":""" & co & """,""year"":""" & yr & _
                 """,""encodeURIComponent"":1,""step"":1,""firstin"":1,""off"":1}")
    p = "{""encodeURIComponent"":1,""step"":1,""firstin"":1,""off"":1"
    For Each k In Array("CK2", "BK1", "TYPEK", "YEAR", "COMP")
        v = RegexFirst(q1, "name='" & k & "' value='([^']*)'")
        If v <> "" Then p = p & ",""" & k & """:""" & v & """"
    Next
    p = p & "}"

    Dim html As String, segs() As String, seg As String, i As Long, curM As Long, mm As String
    html = MopsGet("ajax_t138sb02", p)
    nm = RegexFirst(html, "公司名稱：\s*([^\s<]+)")
    segs = Split(html, "<table")
    curM = 0
    For i = 1 To UBound(segs)
        seg = segs(i)
        If InStr(seg, "hasBorder") > 0 And curM > 0 Then
            Dim prs As Object, pr As Object, lbl As String, sval As String
            Set prs = RegexMatches(seg, "<th[^>]*>([\s\S]*?)</th>\s*<td[^>]*>([\s\S]*?)</td>")
            For Each pr In prs
                lbl = CleanText(pr.SubMatches(0))
                sval = CleanText(pr.SubMatches(1))
                If InStr(lbl, "稅前") > 0 Then
                    If InStr(lbl, "累計") > 0 Then
                        cum(curM) = ToNum(sval)
                    Else
                        mon(curM) = ToNum(sval)
                    End If
                End If
            Next
            curM = 0
        Else
            mm = RegexFirst(seg, yr & "年\s*(\d+)月")
            If mm <> "" Then curM = CLng(mm)
        End If
    Next
End Sub

'================================================================ HTTP
Private Function MopsGet(ByVal api As String, ByVal paramsJson As String) As String
    Dim attempt As Long, lastErr As String, html As String
    For attempt = 1 To MAX_RETRY
        lastErr = TryMopsGet(api, paramsJson, html)
        If lastErr = "" Then
            Pause DELAY_SEC
            MopsGet = html
            Exit Function
        End If
        Pause DELAY_SEC * (attempt + 1)
    Next
    Err.Raise vbObjectError + 3, , api & " 失敗: " & lastErr
End Function

Private Function TryMopsGet(ByVal api As String, ByVal paramsJson As String, ByRef html As String) As String
    On Error GoTo EH
    Dim x As Object, txt As String, url As String
    Set x = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    x.setTimeouts 15000, 15000, 30000, 30000
    x.Open "POST", API_URL, False
    x.setRequestHeader "Content-Type", "application/json"
    x.setRequestHeader "User-Agent", UA
    x.setRequestHeader "Referer", "https://mops.twse.com.tw/mops/"
    x.setRequestHeader "Origin", "https://mops.twse.com.tw"
    x.send "{""apiName"":""" & api & """,""parameters"":" & paramsJson & "}"
    If x.Status <> 200 Then TryMopsGet = "HTTP " & x.Status: Exit Function
    txt = Utf8(x.responseBody)
    url = RegexFirst(txt, """url""\s*:\s*""([^""]+)""")
    If url = "" Then TryMopsGet = "no url: " & Left$(txt, 200): Exit Function

    Set x = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    x.setTimeouts 15000, 15000, 30000, 30000
    x.Open "GET", url, False
    x.setRequestHeader "User-Agent", UA
    x.setRequestHeader "Referer", "https://mops.twse.com.tw/mops/"
    x.send
    If x.Status <> 200 Then TryMopsGet = "HTTP " & x.Status: Exit Function
    html = Utf8(x.responseBody)
    If InStr(html, "FOR SECURITY REASONS") > 0 Then TryMopsGet = "blocked by MOPS security": Exit Function
    TryMopsGet = ""
    Exit Function
EH:
    TryMopsGet = Err.Description
End Function

Private Function Utf8(ByVal bytes As Variant) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8 = st.ReadText
    st.Close
End Function

'================================================================ helpers
Private Function RegexMatches(ByVal s As String, ByVal pat As String) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True: re.Pattern = pat
    Set RegexMatches = re.Execute(s)
End Function

Private Function RegexFirst(ByVal s As String, ByVal pat As String) As String
    Dim mc As Object
    Set mc = RegexMatches(s, pat)
    If mc.Count > 0 Then RegexFirst = mc(0).SubMatches(0)
End Function

Private Function CleanText(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Pattern = "<[^>]+>"
    s = re.Replace(s, "")
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, ""): s = Replace(s, vbLf, "")
    CleanText = Trim$(s)
End Function

Private Function ToNum(ByVal s As Variant) As Variant
    Dim t As String
    t = Replace(Trim$(CStr(s)), ",", "")
    If t = "" Or InStr(t, FIN_MARK()) > 0 Or Not IsNumeric(t) Then
        ToNum = Empty
    Else
        ToNum = CDbl(t)
    End If
End Function

Private Function YoY(ByVal cur As Variant, ByVal prev As Variant) As Variant
    If IsEmpty(cur) Or IsEmpty(prev) Then YoY = Empty: Exit Function
    If prev = 0 Then YoY = Empty: Exit Function
    YoY = (cur - prev) / Abs(prev)
End Function

Private Function FindStockRow(ByVal ws As Worksheet, ByVal co As String) As Long
    Dim r As Long, last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        If Left$(CStr(ws.Cells(r, 1).Value), Len(co) + 1) = co & " " Then FindStockRow = r: Exit Function
    Next
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, c, "")
    Next
    SafeSheetName = Left$(s, 31)
End Function

Private Sub DeleteOutputSheets()
    Dim i As Long
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(i).Name <> CFG_NAME() Then ThisWorkbook.Worksheets(i).Delete
    Next
    Application.DisplayAlerts = True
End Sub

Private Sub FormatSheet(ByVal ws As Worksheet)
    Dim lastR As Long, lastC As Long, c As Long, h As String
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC))
        .Font.Name = "微軟正黑體"
        .Font.Size = 10
        .Borders.Color = RGB(217, 217, 217)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastC))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(1).RowHeight = 36
    For c = 2 To lastC
        h = CStr(ws.Cells(1, c).Value)
        If lastR >= 2 Then
            If InStr(h, "YoY") > 0 Then
                ws.Range(ws.Cells(2, c), ws.Cells(lastR, c)).NumberFormat = "0.00%;[Red]-0.00%"
            ElseIf InStr(h, "損益") > 0 Then
                ws.Range(ws.Cells(2, c), ws.Cells(lastR, c)).NumberFormat = "#,##0;[Red]-#,##0"
            End If
        End If
        ws.Columns(c).ColumnWidth = 17
    Next
    ws.Columns(1).ColumnWidth = 14
    ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 1)).HorizontalAlignment = xlCenter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub SetStatus(ByVal s As String)
    ThisWorkbook.Worksheets(CFG_NAME()).Range("E3").Value = s
    Application.StatusBar = s
    DoEvents
End Sub

Private Sub LogMsg(ByVal s As String)
    With ThisWorkbook.Worksheets(CFG_NAME())
        .Cells(logRow, 4).Value = Format$(Now, "hh:nn:ss") & "  " & s
    End With
    logRow = logRow + 1
    Application.StatusBar = s
    DoEvents
End Sub

Private Sub Pause(ByVal sec As Double)
    Dim t As Double
    t = Timer
    Do While Timer < t + sec And Timer >= t
        DoEvents
    Loop
End Sub

' Decode "\uXXXX" escapes (keeps the module locale independent)
Private Function U(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" Then
            out = out & ChrW(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function
