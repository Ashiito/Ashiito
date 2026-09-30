Attribute VB_Name = "modBuild"
'==============================================================================
' Контрольные задания по Microsoft Access (задания 01-10).
' Модуль создает в текущей базе данных все объекты:
'   таблицы, связи (с обеспечением целостности), тестовые данные,
'   запросы, формы, отчеты, управляющие формы заданий
'   и общую управляющую форму «Главная».
' Запуск: поставить курсор внутрь процедуры BuildAll и нажать F5.
'==============================================================================
Option Compare Database
Option Explicit

Private db As DAO.Database
Private gStep As String
Private g01Start As Date

' состояние построителя управляющих форм
Private cfTmp As String
Private cfY As Long
Private cfN As Integer
Private cfColY(2) As Long
Private Const CF_W As Long = 15600

'------------------------------------------------------------------------------
' ГЛАВНАЯ ПРОЦЕДУРА
'------------------------------------------------------------------------------
Public Sub BuildAll()
    On Error GoTo EH
    Set db = CurrentDb
    DoCmd.Hourglass True
    Application.Echo False
    DoCmd.SetWarnings False

    gStep = "Удаление старых объектов": Cleanup
    BuildT01
    BuildT02
    BuildT03
    BuildT04
    BuildT05
    BuildT06
    BuildT07
    BuildT08
    BuildPub "З09_", False
    BuildPub "З10_", True
    BuildMain

    gStep = "Параметры запуска"
    SetProp db, "StartupForm", dbText, "Главная"
    SetProp db, "AppTitle", dbText, "Контрольные задания MS Access"
    Application.RefreshTitleBar

    DoCmd.SetWarnings True
    Application.Echo True
    DoCmd.Hourglass False
    MsgBox "База данных построена." & vbCrLf & _
           "Открывается управляющая форма «Главная».", vbInformation
    DoCmd.OpenForm "Главная"
    Exit Sub
EH:
    DoCmd.SetWarnings True
    Application.Echo True
    DoCmd.Hourglass False
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Шаг: " & gStep, vbCritical, "Построение базы данных"
End Sub

'------------------------------------------------------------------------------
' Обработчик кнопок всех управляющих форм.
' spec = "Вид:Имя|Параметр1|Параметр2..."
'   F - форма, Q - запрос, R - отчет, B - в главное меню, X - выход
'   ПараметрN - имя поля управляющей формы, которое должно быть заполнено
'------------------------------------------------------------------------------
Public Function OpenObj(ByVal spec As String) As Variant
    Dim k As String, a As Variant, nm As String, i As Integer
    Dim frm As Form, c As Control
    On Error GoTo EH
    k = Left(spec, 1)
    a = Split(Mid(spec, 3), "|")
    If UBound(a) >= 0 Then nm = a(0)
    If UBound(a) >= 1 Then
        Set frm = Screen.ActiveForm
        For i = 1 To UBound(a)
            Set c = frm.Controls(a(i))
            If IsNull(c.Value) Then
                MsgBox "Сначала выберите значение в поле «" & c.Tag & "».", vbExclamation
                c.SetFocus
                Exit Function
            End If
        Next i
    End If
    Select Case k
        Case "F"
            DoCmd.OpenForm nm
        Case "Q"
            If CurrentData.AllQueries(nm).IsLoaded Then DoCmd.Close acQuery, nm
            DoCmd.OpenQuery nm
        Case "R"
            If CurrentProject.AllReports(nm).IsLoaded Then DoCmd.Close acReport, nm
            DoCmd.OpenReport nm, acViewPreview
        Case "B"
            nm = Screen.ActiveForm.Name
            DoCmd.Close acForm, nm
            DoCmd.OpenForm "Главная"
        Case "X"
            If MsgBox("Выйти из приложения?", vbQuestion + vbYesNo) = vbYes Then DoCmd.Quit
    End Select
    Exit Function
EH:
    MsgBox "Не удалось открыть «" & nm & "»: " & Err.Description, vbExclamation
End Function

'------------------------------------------------------------------------------
' Список авторов издания (для заданий 09 и 10), используется в запросах
'------------------------------------------------------------------------------
Public Function AuthorsList(ByVal pfx As String, ByVal id As Variant) As String
    Dim rs As DAO.Recordset, s As String
    If IsNull(id) Then Exit Function
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT p.Фамилия, p.Имя, p.Отчество FROM [" & pfx & "Авторство] AS a " & _
        "INNER JOIN [" & pfx & "Преподаватели] AS p ON a.КодПреподавателя = p.КодПреподавателя " & _
        "WHERE a.КодИздания = " & id & " ORDER BY p.Фамилия", dbOpenSnapshot)
    Do While Not rs.EOF
        If s <> "" Then s = s & ", "
        s = s & rs!Фамилия & " " & Left(Nz(rs!Имя, ""), 1) & "." & Left(Nz(rs!Отчество, ""), 1) & "."
        rs.MoveNext
    Loop
    rs.Close
    AuthorsList = s
End Function

'==============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ПРОЦЕДУРЫ: ТАБЛИЦЫ, СВЯЗИ, ДАННЫЕ, ЗАПРОСЫ
'==============================================================================
Private Sub ExecSQL(ByVal s As String)
    gStep = Left(s, 250)
    db.Execute s, dbFailOnError
End Sub

Private Sub MakeTable(ByVal t As String, ByVal cols As String)
    ExecSQL "CREATE TABLE [" & t & "] (" & cols & ")"
    db.TableDefs.Refresh
    FormatDates t
End Sub

' связь "один-ко-многим" с обеспечением целостности и каскадным обновлением
Private Sub MakeRel(ByVal pt As String, ByVal pf As String, ByVal ct As String, ByVal cf As String, _
                    Optional ByVal casc As Boolean = False)
    Dim rel As DAO.Relation, f As DAO.Field, attr As Long
    gStep = "Связь " & pt & " -> " & ct
    attr = dbRelationUpdateCascade
    If casc Then attr = attr Or dbRelationDeleteCascade
    Set rel = db.CreateRelation("rel_" & ct & "_" & cf, pt, ct, attr)
    Set f = rel.CreateField(pf)
    f.ForeignName = cf
    rel.Fields.Append f
    db.Relations.Append rel
End Sub

' поле со списком (подстановка) в таблице
Private Sub Lookup(ByVal t As String, ByVal f As String, ByVal src As String, _
                   Optional ByVal ncol As Integer = 2, Optional ByVal widths As String = "0;3402")
    Dim fd As DAO.Field
    gStep = "Подстановка " & t & "." & f
    db.TableDefs.Refresh
    Set fd = db.TableDefs(t).Fields(f)
    SetProp fd, "DisplayControl", dbInteger, 111
    SetProp fd, "RowSourceType", dbText, "Table/Query"
    SetProp fd, "RowSource", dbText, src
    SetProp fd, "BoundColumn", dbInteger, 1
    SetProp fd, "ColumnCount", dbInteger, ncol
    SetProp fd, "ColumnWidths", dbText, widths
    SetProp fd, "LimitToList", dbBoolean, True
End Sub

Private Sub FormatDates(ByVal t As String)
    Dim fd As DAO.Field
    For Each fd In db.TableDefs(t).Fields
        If fd.Type = dbDate Then SetProp fd, "Format", dbText, "Short Date"
    Next fd
End Sub

Private Sub SetProp(ByVal obj As Object, ByVal pname As String, ByVal ptype As Integer, ByVal pval As Variant)
    On Error Resume Next
    obj.Properties(pname) = pval
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        obj.Properties.Append obj.CreateProperty(pname, ptype, pval)
    End If
End Sub

Private Function GetProp(ByVal obj As Object, ByVal pname As String) As Variant
    GetProp = Null
    On Error Resume Next
    GetProp = obj.Properties(pname).Value
End Function

' вставка нескольких строк: строки разделяются "|"
' в строках допускаются подстановки:
'   {d-10}   - дата: сегодня минус 10 дней
'   {y-2}    - год: текущий минус 2
'   {m1990}  - 15-е число текущего месяца 1990 года (для именинников)
Private Sub InsRows(ByVal t As String, ByVal rows As String, Optional ByVal cols As String = "")
    Dim a As Variant, i As Integer, s As String
    a = Split(rows, "|")
    For i = 0 To UBound(a)
        s = "INSERT INTO [" & t & "] "
        If cols <> "" Then s = s & "(" & cols & ") "
        ExecSQL s & "VALUES (" & Tok(Trim(a(i))) & ")"
    Next i
End Sub

Private Function Tok(ByVal s As String) As String
    Dim p As Long, q As Long, tk As String, v As String, n As Long
    Do
        p = InStr(s, "{")
        If p = 0 Then Exit Do
        q = InStr(p, s, "}")
        tk = Mid(s, p + 1, q - p - 1)
        n = Val(Mid(tk, 2))
        Select Case Left(tk, 1)
            Case "d": v = SD(Date + n)
            Case "y": v = CStr(Year(Date) + n)
            Case "m": v = SD(DateSerial(n, Month(Date), 15))
        End Select
        s = Left(s, p - 1) & v & Mid(s, q + 1)
    Loop
    Tok = s
End Function

' дата в формате SQL: #мм/дд/гггг#
Private Function SD(ByVal d As Date) As String
    SD = "#" & Format(d, "mm\/dd\/yyyy") & "#"
End Function

' апострофы -> кавычки (для выражений в свойствах элементов)
Private Function Q(ByVal s As String) As String
    Q = Replace(s, "'", Chr(34))
End Function

Private Sub MakeQuery(ByVal n As String, ByVal sql As String)
    Dim qd As DAO.QueryDef
    gStep = "Запрос " & n & ": " & Left(sql, 200)
    Set qd = db.CreateQueryDef(n, sql)
End Sub

' удаление объектов, созданных ранее этим модулем
Private Sub Cleanup()
    Dim names As New Collection, i As Integer, n As Variant, ao As AccessObject

    For i = db.Relations.Count - 1 To 0 Step -1
        If db.Relations(i).Name Like "rel_*" Then db.Relations.Delete db.Relations(i).Name
    Next i

    For Each ao In CurrentProject.AllForms
        If IsOurs(ao.Name) Then names.Add ao.Name
    Next ao
    For Each n In names
        If CurrentProject.AllForms(n).IsLoaded Then DoCmd.Close acForm, n, acSaveNo
        DoCmd.DeleteObject acForm, n
    Next n

    Set names = New Collection
    For Each ao In CurrentProject.AllReports
        If IsOurs(ao.Name) Then names.Add ao.Name
    Next ao
    For Each n In names
        If CurrentProject.AllReports(n).IsLoaded Then DoCmd.Close acReport, n, acSaveNo
        DoCmd.DeleteObject acReport, n
    Next n

    Set names = New Collection
    For i = 0 To db.QueryDefs.Count - 1
        If IsOurs(db.QueryDefs(i).Name) Then names.Add db.QueryDefs(i).Name
    Next i
    For Each n In names
        db.QueryDefs.Delete n
    Next n

    Set names = New Collection
    For i = 0 To db.TableDefs.Count - 1
        If IsOurs(db.TableDefs(i).Name) Then names.Add db.TableDefs(i).Name
    Next i
    For Each n In names
        db.TableDefs.Delete n
    Next n
    db.TableDefs.Refresh
End Sub

Private Function IsOurs(ByVal n As String) As Boolean
    IsOurs = (n Like "З##_*") Or (n = "Главная")
End Function

'==============================================================================
' ФОРМЫ ДЛЯ ВВОДА И ПРОСМОТРА ДАННЫХ
'==============================================================================
' view: 0 - одиночная форма, 2 - таблица (для подчиненных форм)
' excl: поля через ";", которые не выводятся (например, поле связи с главной формой)
Private Sub MakeTableForm(ByVal fname As String, ByVal src As String, ByVal title As String, _
                          Optional ByVal view As Integer = 0, Optional ByVal excl As String = "")
    Dim frm As Form, tmp As String, c As Control, lb As Control
    Dim tdf As DAO.TableDef, fd As DAO.Field
    Dim y As Long, w As Long, h As Long, ct As Integer

    gStep = "Форма " & fname
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(src)
    Set frm = CreateForm()
    tmp = frm.Name
    frm.RecordSource = src
    frm.Caption = title
    frm.DefaultView = view
    frm.AutoCenter = True
    frm.Width = 8400
    frm.Section(acDetail).BackColor = RGB(242, 242, 242)

    y = 200
    If view <> 2 Then
        Set lb = CreateControl(tmp, acLabel, acDetail, , , 200, 150, 7800, 500)
        lb.Caption = title
        lb.FontSize = 16
        lb.FontBold = True
        lb.ForeColor = RGB(192, 80, 0)
        y = 850
    End If

    For Each fd In tdf.Fields
        If InStr(";" & excl & ";", ";" & fd.Name & ";") = 0 Then
            h = 330
            If Nz(GetProp(fd, "DisplayControl"), 0) = 111 Then
                ct = acComboBox: w = 4200
            ElseIf fd.Type = dbBoolean Then
                ct = acCheckBox: w = 300
            ElseIf fd.Type = dbMemo Then
                ct = acTextBox: w = 5200: h = 900
            ElseIf fd.Type = dbText Then
                ct = acTextBox: w = fd.Size * 90
                If w < 1400 Then w = 1400
                If w > 5200 Then w = 5200
            Else
                ct = acTextBox: w = 1700
            End If
            Set c = CreateControl(tmp, ct, acDetail, , , 2900, y, w, h)
            c.ControlSource = fd.Name
            c.Name = fd.Name
            If ct = acComboBox Then
                c.RowSourceType = "Table/Query"
                c.RowSource = GetProp(fd, "RowSource")
                c.ColumnCount = GetProp(fd, "ColumnCount")
                c.ColumnWidths = GetProp(fd, "ColumnWidths")
                c.BoundColumn = 1
                c.LimitToList = True
            End If
            If (fd.Attributes And dbAutoIncrField) <> 0 Then
                c.Locked = True
                c.BackColor = RGB(230, 230, 230)
            End If
            Set lb = CreateControl(tmp, acLabel, acDetail, fd.Name, , 200, y, 2600, h)
            lb.Caption = fd.Name
            y = y + h + 90
        End If
    Next fd

    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.Rename fname, acForm, tmp
End Sub

' добавить подчиненную форму (в режиме таблицы) в главную форму
Private Sub AddSubform(ByVal mainName As String, ByVal childName As String, _
                       ByVal master As String, ByVal child As String, ByVal cap As String)
    Dim frm As Form, c As Control, sf As Control, lb As Control, y As Long
    gStep = "Подчиненная форма " & childName & " в " & mainName
    DoCmd.OpenForm mainName, acDesign, , , , acHidden
    Set frm = Forms(mainName)
    For Each c In frm.Controls
        If c.Top + c.Height > y Then y = c.Top + c.Height
    Next c
    Set lb = CreateControl(mainName, acLabel, acDetail, , , 200, y + 200, 6000, 330)
    lb.Caption = cap
    lb.FontBold = True
    lb.ForeColor = RGB(192, 80, 0)
    Set sf = CreateControl(mainName, acSubform, acDetail, , , 200, y + 580, 9800, 3300)
    sf.SourceObject = childName
    sf.LinkMasterFields = master
    sf.LinkChildFields = child
    sf.Name = "Подчиненная"
    frm.Width = 10200
    DoCmd.Close acForm, mainName, acSaveYes
End Sub

'==============================================================================
' ОТЧЕТЫ
'==============================================================================
' title   - текст заголовка или выражение (начинается с "="; строки в апострофах)
' groups  - поля группировки через ";" (0..2 уровня); в примечании группы - количество
' flds    - поля области данных через ";"; widths - их ширина в твипах (567 = 1 см)
' sorts   - поля сортировки через ";" ("-Поле" - по убыванию)
' sums    - поля, по которым считаются суммы в примечаниях групп и отчета
' numbered- нумерация записей (внутри группы - сумма с накоплением "для группы")
Private Sub MakeReport(ByVal rname As String, ByVal src As String, ByVal title As String, _
                       ByVal groups As String, ByVal flds As String, ByVal widths As String, _
                       Optional ByVal sorts As String = "", Optional ByVal sums As String = "", _
                       Optional ByVal numbered As Boolean = True)
    Dim rpt As Report, tmp As String, c As Control
    Dim fa As Variant, wa As Variant, ga As Variant, sa As Variant, sm As Variant
    Dim xs() As Long, i As Integer, j As Integer, ng As Integer
    Dim x As Long, totalW As Long, numW As Long, yL As Long, gi As Long
    Dim hasRH As Boolean, secT As Integer, sName As String

    gStep = "Отчет " & rname
    fa = Split(flds, ";")
    wa = Split(widths, ";")
    ng = 0
    If groups <> "" Then
        ga = Split(groups, ";")
        ng = UBound(ga) + 1
    End If
    numW = 0
    If numbered Then numW = 550
    ReDim xs(UBound(fa))
    x = numW
    For i = 0 To UBound(fa)
        xs(i) = x
        x = x + CLng(wa(i))
    Next i
    totalW = x

    Set rpt = CreateReport()
    tmp = rpt.Name
    rpt.RecordSource = src
    rpt.Caption = Mid(rname, 5)
    rpt.Width = totalW
    If totalW > 10400 Then
        On Error Resume Next
        rpt.Printer.Orientation = acPRORLandscape
        On Error GoTo 0
    End If

    ' заголовок и примечание отчета
    On Error Resume Next
    DoCmd.RunCommand acCmdReportHdrFtr
    On Error GoTo 0
    hasRH = SectionOk(rpt, acHeader)

    If hasRH Then
        secT = acHeader
        rpt.Section(acHeader).Height = 900
        yL = 60
    Else
        secT = acPageHeader
        yL = 750
    End If
    rpt.Section(acPageHeader).Height = yL + 420

    ' заголовок
    If Left(title, 1) = "=" Then
        Set c = CreateReportControl(tmp, acTextBox, secT, , , 0, 100, totalW, 520)
        c.ControlSource = Q(title)
    Else
        Set c = CreateReportControl(tmp, acLabel, secT, , , 0, 100, totalW, 520)
        c.Caption = title
    End If
    c.FontSize = 16
    c.FontBold = True
    c.ForeColor = RGB(192, 80, 0)

    ' заголовки столбцов
    If numbered Then
        Set c = CreateReportControl(tmp, acLabel, acPageHeader, , , 0, yL, numW - 50, 330)
        c.Caption = "№"
        c.FontBold = True
    End If
    For i = 0 To UBound(fa)
        Set c = CreateReportControl(tmp, acLabel, acPageHeader, , , xs(i), yL, CLng(wa(i)) - 50, 330)
        c.Caption = fa(i)
        c.FontBold = True
    Next i
    Set c = CreateReportControl(tmp, acLine, acPageHeader, , , 0, yL + 360, totalW, 0)

    ' группировка
    For i = 0 To ng - 1
        gi = CreateGroupLevel(tmp, ga(i), True, True)
        rpt.Section(5 + 2 * i).Height = 430
        Set c = CreateReportControl(tmp, acTextBox, 5 + 2 * i, , , i * 300, 60, totalW - i * 300, 340)
        c.ControlSource = Q("='" & ga(i) & ": ' & [" & ga(i) & "]")
        c.FontBold = True
        c.FontSize = 12 - i
        c.ForeColor = RGB(31, 73, 125)
        rpt.Section(6 + 2 * i).Height = 450
        Set c = CreateReportControl(tmp, acTextBox, 6 + 2 * i, , , i * 300, 40, 3200, 330)
        c.ControlSource = Q("='Количество: ' & Count(*)")
        c.FontItalic = True
        AddSums tmp, 6 + 2 * i, sums, fa, xs, wa, "g" & i
    Next i

    ' сортировка
    If sorts <> "" Then
        sa = Split(sorts, ";")
        For i = 0 To UBound(sa)
            sName = sa(i)
            If Left(sName, 1) = "-" Then sName = Mid(sName, 2)
            gi = CreateGroupLevel(tmp, sName, False, False)
            If Left(sa(i), 1) = "-" Then rpt.GroupLevel(gi).SortOrder = True
        Next i
    End If

    ' область данных
    rpt.Section(acDetail).Height = 330
    If numbered Then
        Set c = CreateReportControl(tmp, acTextBox, acDetail, , , 0, 0, numW - 50, 300)
        c.ControlSource = "=1"
        If ng > 0 Then c.RunningSum = 1 Else c.RunningSum = 2
        c.Name = "НомерПП"
    End If
    For i = 0 To UBound(fa)
        Set c = CreateReportControl(tmp, acTextBox, acDetail, , , xs(i), 0, CLng(wa(i)) - 50, 300)
        c.ControlSource = fa(i)
        c.Name = fa(i)
        c.CanGrow = True
    Next i
    rpt.Section(acDetail).CanGrow = True

    ' примечание отчета - итог
    If hasRH Then
        rpt.Section(acFooter).Height = 500
        Set c = CreateReportControl(tmp, acTextBox, acFooter, , , 0, 100, 3200, 330)
        c.ControlSource = Q("='Всего записей: ' & Count(*)")
        c.FontBold = True
        AddSums tmp, acFooter, sums, fa, xs, wa, "r"
    End If

    ' нижний колонтитул - номера страниц
    rpt.Section(acPageFooter).Height = 400
    Set c = CreateReportControl(tmp, acTextBox, acPageFooter, , , totalW - 3000, 40, 3000, 300)
    c.ControlSource = Q("='Страница ' & [Page] & ' из ' & [Pages]")
    c.TextAlign = 3

    DoCmd.Close acReport, tmp, acSaveYes
    DoCmd.Rename rname, acReport, tmp
End Sub

Private Sub AddSums(ByVal tmp As String, ByVal sec As Integer, ByVal sums As String, _
                    ByVal fa As Variant, ByRef xs() As Long, ByVal wa As Variant, ByVal tag As String)
    Dim sm As Variant, i As Integer, j As Integer, c As Control
    If sums = "" Then Exit Sub
    sm = Split(sums, ";")
    For j = 0 To UBound(sm)
        For i = 0 To UBound(fa)
            If fa(i) = sm(j) Then
                Set c = CreateReportControl(tmp, acTextBox, sec, , , xs(i), 40, CLng(wa(i)) - 50, 330)
                c.ControlSource = "=Sum([" & sm(j) & "])"
                c.Name = "Сумма_" & tag & "_" & sm(j)
                c.FontBold = True
            End If
        Next i
    Next j
End Sub

Private Function SectionOk(ByVal rpt As Report, ByVal sec As Integer) As Boolean
    Dim h As Long
    On Error Resume Next
    h = rpt.Section(sec).Height
    SectionOk = (Err.Number = 0)
    Err.Clear
End Function

' отчет-поздравление с днем рождения (одна страница на запись)
' источник должен содержать поля [Обращение] и [ДатаРождения]
Private Sub MakeGreeting(ByVal rname As String, ByVal src As String, ByVal salute As String, _
                         ByVal msg As String, ByVal sign As String)
    Dim rpt As Report, tmp As String, c As Control
    gStep = "Отчет " & rname
    Set rpt = CreateReport()
    tmp = rpt.Name
    rpt.RecordSource = src
    rpt.Caption = Mid(rname, 5)
    rpt.Width = 10000
    rpt.Section(acPageHeader).Height = 0
    rpt.Section(acPageFooter).Height = 0
    rpt.Section(acDetail).Height = 6300
    rpt.Section(acDetail).ForceNewPage = 2

    Set c = CreateReportControl(tmp, acRectangle, acDetail, , , 200, 200, 9600, 5800)
    c.BorderWidth = 3
    c.BorderColor = RGB(192, 0, 0)
    c.BackStyle = 0

    Set c = CreateReportControl(tmp, acLabel, acDetail, , , 400, 600, 9200, 900)
    c.Caption = "С днём рождения!"
    c.FontSize = 32
    c.FontBold = True
    c.ForeColor = RGB(192, 0, 0)
    c.TextAlign = 2

    Set c = CreateReportControl(tmp, acTextBox, acDetail, , , 400, 1750, 9200, 600)
    c.ControlSource = Q("='" & salute & " ' & [Обращение] & '!'")
    c.FontSize = 20
    c.FontItalic = True
    c.TextAlign = 2

    Set c = CreateReportControl(tmp, acLabel, acDetail, , , 700, 2600, 8600, 1500)
    c.Caption = msg
    c.FontSize = 14
    c.TextAlign = 2

    Set c = CreateReportControl(tmp, acTextBox, acDetail, , , 400, 4300, 9200, 400)
    c.ControlSource = Q("='Дата рождения: ' & [ДатаРождения]")
    c.FontSize = 11
    c.TextAlign = 2

    Set c = CreateReportControl(tmp, acLabel, acDetail, , , 400, 5100, 9000, 400)
    c.Caption = sign
    c.FontSize = 12
    c.FontItalic = True
    c.TextAlign = 3

    DoCmd.Close acReport, tmp, acSaveYes
    DoCmd.Rename rname, acReport, tmp
End Sub

'==============================================================================
' УПРАВЛЯЮЩИЕ ФОРМЫ
' верхняя часть - параметры (поля со списком), ниже три колонки кнопок:
'   0 - Ввод и просмотр данных, 1 - Запросы, 2 - Отчеты
'==============================================================================
Private Sub CF_Begin(ByVal title As String)
    Dim frm As Form, c As Control
    gStep = "Управляющая форма: " & title
    Set frm = CreateForm()
    cfTmp = frm.Name
    frm.Caption = title
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Width = CF_W
    frm.Section(acDetail).BackColor = RGB(238, 236, 225)

    Set c = CreateControl(cfTmp, acLabel, acDetail, , , 0, 0, CF_W, 850)
    c.Caption = " " & title
    c.BackStyle = 1
    c.BackColor = RGB(228, 108, 10)
    c.ForeColor = RGB(255, 255, 255)
    c.FontSize = 20
    c.FontBold = True

    Set c = CreateControl(cfTmp, acLabel, acDetail, , , 300, 1000, 9000, 380)
    c.Caption = "Параметры отбора для запросов и отчетов"
    c.FontSize = 12
    c.FontBold = True
    c.ForeColor = RGB(192, 80, 0)
    cfY = 1450
    cfN = 0
End Sub

' kind: "C" - поле со списком из таблицы/запроса, "V" - список значений, "T" - текстовое поле
Private Sub CF_Param(ByVal nm As String, ByVal cap As String, ByVal src As String, _
                     Optional ByVal ncol As Integer = 2, Optional ByVal cw As String = "0;4500", _
                     Optional ByVal dflt As String = "", Optional ByVal kind As String = "C")
    Dim c As Control, lb As Control, x As Long, y As Long
    x = 300 + (cfN Mod 2) * 7700
    y = cfY + (cfN \ 2) * 450
    If kind = "T" Then
        Set c = CreateControl(cfTmp, acTextBox, acDetail, , , x + 2400, y, 4800, 330)
    Else
        Set c = CreateControl(cfTmp, acComboBox, acDetail, , , x + 2400, y, 4800, 330)
        If kind = "V" Then c.RowSourceType = "Value List" Else c.RowSourceType = "Table/Query"
        c.RowSource = src
        c.ColumnCount = ncol
        c.ColumnWidths = cw
        c.BoundColumn = 1
        c.LimitToList = True
        c.ListRows = 16
    End If
    c.Name = nm
    c.Tag = cap
    If dflt <> "" Then c.DefaultValue = dflt
    Set lb = CreateControl(cfTmp, acLabel, acDetail, nm, , x, y, 2300, 330)
    lb.Caption = cap & ":"
    cfN = cfN + 1
End Sub

Private Sub CF_Columns()
    Dim c As Control, i As Integer, y As Long, caps As Variant
    caps = Array("Ввод и просмотр данных", "Получение запросов", "Получение отчетов")
    y = cfY + ((cfN + 1) \ 2) * 450 + 250
    For i = 0 To 2
        Set c = CreateControl(cfTmp, acLabel, acDetail, , , 300 + i * 5100, y, 4800, 400)
        c.Caption = caps(i)
        c.FontSize = 13
        c.FontBold = True
        c.ForeColor = RGB(192, 80, 0)
        c.TextAlign = 2
        cfColY(i) = y + 480
    Next i
End Sub

Private Sub CF_Btn(ByVal col As Integer, ByVal cap As String, ByVal spec As String)
    Dim c As Control
    Set c = CreateControl(cfTmp, acCommandButton, acDetail, , , 300 + col * 5100, cfColY(col), 4800, 400)
    c.Caption = cap
    c.OnClick = "=OpenObj(""" & spec & """)"
    c.ForeColor = RGB(192, 80, 0)
    cfColY(col) = cfColY(col) + 450
End Sub

Private Sub CF_End(ByVal fname As String)
    Dim c As Control, y As Long, i As Integer
    For i = 0 To 2
        If cfColY(i) > y Then y = cfColY(i)
    Next i
    y = y + 250
    Set c = CreateControl(cfTmp, acCommandButton, acDetail, , , 300, y, 4800, 450)
    c.Caption = "<< Главное меню"
    c.OnClick = "=OpenObj(""B:"")"
    c.FontBold = True
    Set c = CreateControl(cfTmp, acLabel, acDetail, , , 5400, y + 40, 9900, 400)
    c.Caption = "Выберите параметры вверху, затем нажмите кнопку нужного запроса или отчета."
    c.FontItalic = True
    DoCmd.Close acForm, cfTmp, acSaveYes
    DoCmd.Rename fname, acForm, cfTmp
End Sub

'==============================================================================
' ГЛАВНАЯ УПРАВЛЯЮЩАЯ ФОРМА - объединяет все задания
'==============================================================================
Private Sub BuildMain()
    Dim frm As Form, tmp As String, c As Control, i As Integer, caps As Variant
    gStep = "Форма Главная"
    caps = Array("Расписание занятий", "Абитуриенты", "Телефонная книга", "Учебные планы", _
                 "Сотрудники", "Выпускники", "Преподаватели", "Библиотека", _
                 "Издательский отдел", "Информационный отдел")
    Set frm = CreateForm()
    tmp = frm.Name
    frm.Caption = "Главная управляющая форма"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Width = CF_W
    frm.Section(acDetail).BackColor = RGB(238, 236, 225)

    Set c = CreateControl(tmp, acLabel, acDetail, , , 0, 0, CF_W, 950)
    c.Caption = " Управляющая форма: контрольные задания по Microsoft Access"
    c.BackStyle = 1
    c.BackColor = RGB(228, 108, 10)
    c.ForeColor = RGB(255, 255, 255)
    c.FontSize = 20
    c.FontBold = True

    Set c = CreateControl(tmp, acLabel, acDetail, , , 300, 1150, 15000, 400)
    c.Caption = "Выберите задание. Откроется его управляющая форма с данными, запросами и отчетами."
    c.FontSize = 12
    c.ForeColor = RGB(192, 80, 0)

    For i = 0 To 9
        Set c = CreateControl(tmp, acCommandButton, acDetail, , , 300 + (i \ 5) * 7650, 1800 + (i Mod 5) * 700, 7350, 580)
        c.Caption = "Задание " & Format(i + 1, "00") & ".  " & caps(i)
        c.OnClick = "=OpenObj(""F:З" & Format(i + 1, "00") & "_Управляющая"")"
        c.FontSize = 12
        c.FontBold = True
        c.ForeColor = RGB(192, 80, 0)
    Next i

    Set c = CreateControl(tmp, acCommandButton, acDetail, , , 300 + 7650, 5500, 7350, 500)
    c.Caption = "Выйти из приложения"
    c.OnClick = "=OpenObj(""X:"")"

    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.Rename "Главная", acForm, tmp
End Sub

'==============================================================================
' ЗАДАНИЕ 01. База данных «Расписание занятий»
' Календарь на семестр. Расписание пар. Аудитории. Группы студентов.
' Учебные дисциплины и количество аудиторных часов. Преподаватели.
' Часы для групп и преподавателей по дисциплинам, для аудиторий.
' Расписание для группы, для преподавателя, для аудитории.
'==============================================================================
Private Sub BuildT01()
    Dim s As String, tpl As Variant, a As Variant, dn As Variant
    Dim d0 As Date, m0 As Date, d As Date, k As Long, j As Integer
    Dim wd As Integer, wk As Integer, y As Integer, study As Boolean
    Const FR As String = "[Forms]![З01_Управляющая]!"

    ' --- таблицы ---
    MakeTable "З01_Группы", "КодГруппы COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Группа TEXT(20) NOT NULL, Курс INTEGER, КолСтудентов INTEGER"
    MakeTable "З01_Аудитории", "КодАудитории COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Аудитория TEXT(10) NOT NULL, Корпус TEXT(40), Мест INTEGER, ТипАудитории TEXT(40)"
    MakeTable "З01_Преподаватели", "КодПреподавателя COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), Должность TEXT(40)"
    MakeTable "З01_Дисциплины", "КодДисциплины COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Дисциплина TEXT(80) NOT NULL, ЧасовПоПлану INTEGER"
    MakeTable "З01_Пары", "НомерПары INTEGER CONSTRAINT PrimaryKey PRIMARY KEY, Начало TEXT(5), Окончание TEXT(5)"
    MakeTable "З01_Календарь", "Дата DATETIME CONSTRAINT PrimaryKey PRIMARY KEY, ДеньНедели TEXT(15), НомерНедели INTEGER, Учебный BIT"
    MakeTable "З01_Занятия", "КодЗанятия COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Дата DATETIME NOT NULL, НомерПары INTEGER NOT NULL, " & _
        "КодГруппы LONG NOT NULL, КодДисциплины LONG NOT NULL, КодПреподавателя LONG NOT NULL, КодАудитории LONG NOT NULL, ВидЗанятия TEXT(20)"

    ' --- связи ---
    MakeRel "З01_Календарь", "Дата", "З01_Занятия", "Дата"
    MakeRel "З01_Пары", "НомерПары", "З01_Занятия", "НомерПары"
    MakeRel "З01_Группы", "КодГруппы", "З01_Занятия", "КодГруппы"
    MakeRel "З01_Дисциплины", "КодДисциплины", "З01_Занятия", "КодДисциплины"
    MakeRel "З01_Преподаватели", "КодПреподавателя", "З01_Занятия", "КодПреподавателя"
    MakeRel "З01_Аудитории", "КодАудитории", "З01_Занятия", "КодАудитории"

    ' --- поля со списком ---
    Lookup "З01_Занятия", "НомерПары", "SELECT НомерПары, Начало & '-' & Окончание FROM [З01_Пары] ORDER BY НомерПары", 2, "567;1700"
    Lookup "З01_Занятия", "КодГруппы", "SELECT КодГруппы, Группа FROM [З01_Группы] ORDER BY Группа"
    Lookup "З01_Занятия", "КодДисциплины", "SELECT КодДисциплины, Дисциплина FROM [З01_Дисциплины] ORDER BY Дисциплина"
    Lookup "З01_Занятия", "КодПреподавателя", "SELECT КодПреподавателя, Фамилия & ' ' & Имя & ' ' & Отчество FROM [З01_Преподаватели] ORDER BY Фамилия"
    Lookup "З01_Занятия", "КодАудитории", "SELECT КодАудитории, Аудитория & ' (' & Корпус & ')' FROM [З01_Аудитории] ORDER BY Аудитория"

    ' --- данные ---
    InsRows "З01_Группы", "1,'ИС-21',2,24|2,'ЭК-22',1,22|3,'ЮР-21',2,26|4,'ПС-23',1,20"
    InsRows "З01_Аудитории", "1,'101','Главный корпус',60,'Лекционная'|2,'205','Главный корпус',25,'Компьютерный класс'|" & _
        "3,'310','Корпус 2',30,'Практическая'|4,'12','Корпус 2',120,'Поточная'"
    InsRows "З01_Преподаватели", "1,'Иванов','Сергей','Петрович','Доцент'|2,'Петрова','Анна','Викторовна','Профессор'|" & _
        "3,'Сидоров','Олег','Иванович','Старший преподаватель'|4,'Кузнецова','Елена','Сергеевна','Доцент'|" & _
        "5,'Смирнов','Алексей','Николаевич','Ассистент'"
    InsRows "З01_Дисциплины", "1,'Базы данных',64|2,'Математика',72|3,'Экономическая теория',54|4,'Правоведение',36|" & _
        "5,'Психология',36|6,'Информатика',54|7,'Иностранный язык',36"
    InsRows "З01_Пары", "1,'08:30','10:00'|2,'10:10','11:40'|3,'12:20','13:50'|4,'14:00','15:30'|5,'15:40','17:10'"

    ' календарь на семестр (с 1 сентября, 17 недель) и занятия по недельному шаблону
    ' шаблон: день недели; пара; группа; дисциплина; преподаватель; аудитория; вид занятия
    tpl = Array("1;1;1;1;1;2;Лабораторная", "1;1;2;2;2;1;Лекция", "1;2;3;4;4;3;Лекция", "1;2;4;5;5;1;Лекция", _
                "2;1;1;2;2;1;Практика", "2;1;3;3;3;3;Лекция", "2;2;2;3;3;3;Практика", "2;2;4;6;1;2;Лабораторная", _
                "3;1;1;6;1;2;Лабораторная", "3;2;2;6;1;2;Лабораторная", "3;3;1;1;1;4;Лекция", _
                "4;1;3;2;2;1;Лекция", "4;2;4;2;2;1;Практика", "4;1;4;4;4;3;Практика", _
                "5;1;3;5;5;1;Практика", "5;2;2;4;4;3;Лекция", "5;1;1;3;3;3;Лекция")
    dn = Array("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье")
    y = Year(Date)
    If Month(Date) < 8 Then y = y - 1
    d0 = DateSerial(y, 9, 1)
    g01Start = d0
    m0 = d0 - (Weekday(d0, vbMonday) - 1)
    For k = 0 To 118
        d = d0 + k
        wd = Weekday(d, vbMonday)
        wk = CInt((d - m0) \ 7) + 1
        study = (wd <= 6) And Not (Month(d) = 11 And Day(d) = 4)
        ExecSQL "INSERT INTO [З01_Календарь] VALUES (" & SD(d) & ", '" & dn(wd - 1) & "', " & wk & ", " & IIf(study, "True", "False") & ")"
        If study Then
            For j = 0 To UBound(tpl)
                a = Split(tpl(j), ";")
                If CInt(a(0)) = wd Then
                    ExecSQL "INSERT INTO [З01_Занятия] (Дата, НомерПары, КодГруппы, КодДисциплины, КодПреподавателя, КодАудитории, ВидЗанятия) " & _
                        "VALUES (" & SD(d) & ", " & a(1) & ", " & a(2) & ", " & a(3) & ", " & a(4) & ", " & a(5) & ", '" & a(6) & "')"
                End If
            Next j
        End If
    Next k

    ' --- запросы ---
    s = "SELECT z.КодЗанятия, z.Дата, c.ДеньНедели, c.НомерНедели, z.НомерПары, pr.Начало, pr.Окончание, "
    s = s & "z.КодГруппы, g.Группа, z.КодДисциплины, d.Дисциплина, z.ВидЗанятия, z.КодПреподавателя, "
    s = s & "t.Фамилия & ' ' & Left(t.Имя,1) & '.' & Left(t.Отчество,1) & '.' AS Преподаватель, "
    s = s & "z.КодАудитории, a.Аудитория, a.Корпус, 2 AS Часы "
    s = s & "FROM ((((([З01_Занятия] AS z INNER JOIN [З01_Календарь] AS c ON z.Дата = c.Дата) "
    s = s & "INNER JOIN [З01_Пары] AS pr ON z.НомерПары = pr.НомерПары) "
    s = s & "INNER JOIN [З01_Группы] AS g ON z.КодГруппы = g.КодГруппы) "
    s = s & "INNER JOIN [З01_Дисциплины] AS d ON z.КодДисциплины = d.КодДисциплины) "
    s = s & "INNER JOIN [З01_Преподаватели] AS t ON z.КодПреподавателя = t.КодПреподавателя) "
    s = s & "INNER JOIN [З01_Аудитории] AS a ON z.КодАудитории = a.КодАудитории"
    MakeQuery "З01_Занятия подробно", s

    MakeQuery "З01_Часы групп по дисциплинам", _
        "SELECT q.Группа, q.Дисциплина, d.ЧасовПоПлану, Count(*) AS Пар, Sum(q.Часы) AS ЧасовПоРасписанию " & _
        "FROM [З01_Занятия подробно] AS q INNER JOIN [З01_Дисциплины] AS d ON q.КодДисциплины = d.КодДисциплины " & _
        "GROUP BY q.Группа, q.Дисциплина, d.ЧасовПоПлану ORDER BY q.Группа, q.Дисциплина"
    MakeQuery "З01_Часы преподавателей по дисциплинам", _
        "SELECT Преподаватель, Дисциплина, Группа, Count(*) AS Пар, Sum(Часы) AS Часов " & _
        "FROM [З01_Занятия подробно] GROUP BY Преподаватель, Дисциплина, Группа ORDER BY Преподаватель, Дисциплина, Группа"
    MakeQuery "З01_Нагрузка преподавателей", _
        "SELECT Преподаватель, Count(*) AS Пар, Sum(Часы) AS ВсегоЧасов " & _
        "FROM [З01_Занятия подробно] GROUP BY Преподаватель ORDER BY Sum(Часы) DESC"
    MakeQuery "З01_Часы аудиторий", _
        "SELECT Аудитория, Корпус, Count(*) AS Пар, Sum(Часы) AS Часов " & _
        "FROM [З01_Занятия подробно] GROUP BY Аудитория, Корпус ORDER BY Аудитория"
    MakeQuery "З01_Сводная часы групп по дисциплинам", _
        "TRANSFORM Sum(Часы) SELECT Дисциплина, Sum(Часы) AS Итого " & _
        "FROM [З01_Занятия подробно] GROUP BY Дисциплина PIVOT Группа"

    s = "PARAMETERS " & FR & "[cboС] DateTime, " & FR & "[cboПо] DateTime; "
    MakeQuery "З01_Расписание группы", s & _
        "SELECT Группа, Дата, ДеньНедели, НомерПары, Начало, Дисциплина, ВидЗанятия, Преподаватель, Аудитория " & _
        "FROM [З01_Занятия подробно] WHERE КодГруппы = " & FR & "[cboГруппа] " & _
        "AND Дата Between " & FR & "[cboС] And " & FR & "[cboПо] ORDER BY Дата, НомерПары"
    MakeQuery "З01_Расписание преподавателя", s & _
        "SELECT Преподаватель, Дата, ДеньНедели, НомерПары, Начало, Дисциплина, ВидЗанятия, Группа, Аудитория " & _
        "FROM [З01_Занятия подробно] WHERE КодПреподавателя = " & FR & "[cboПреподаватель] " & _
        "AND Дата Between " & FR & "[cboС] And " & FR & "[cboПо] ORDER BY Дата, НомерПары"
    MakeQuery "З01_Расписание аудитории", s & _
        "SELECT Аудитория, Корпус, Дата, ДеньНедели, НомерПары, Начало, Группа, Дисциплина, ВидЗанятия, Преподаватель " & _
        "FROM [З01_Занятия подробно] WHERE КодАудитории = " & FR & "[cboАудитория] " & _
        "AND Дата Between " & FR & "[cboС] And " & FR & "[cboПо] ORDER BY Дата, НомерПары"

    ' --- формы ---
    MakeTableForm "З01_Группы", "З01_Группы", "Группы студентов"
    MakeTableForm "З01_Аудитории", "З01_Аудитории", "Аудитории"
    MakeTableForm "З01_Преподаватели", "З01_Преподаватели", "Преподаватели"
    MakeTableForm "З01_Дисциплины", "З01_Дисциплины", "Учебные дисциплины"
    MakeTableForm "З01_Пары", "З01_Пары", "Расписание пар (звонков)"
    MakeTableForm "З01_Календарь", "З01_Календарь", "Календарь на семестр", 2
    MakeTableForm "З01_Занятия", "З01_Занятия", "Занятия (расписание)"

    ' --- отчеты ---
    MakeReport "З01_Расписание группы", "З01_Расписание группы", "='Расписание занятий группы ' & [Группа]", "Дата", _
        "НомерПары;Начало;Дисциплина;ВидЗанятия;Преподаватель;Аудитория", "650;750;3000;1650;2200;1000", "НомерПары", "", False
    MakeReport "З01_Расписание преподавателя", "З01_Расписание преподавателя", "='Расписание преподавателя ' & [Преподаватель]", "Дата", _
        "НомерПары;Начало;Дисциплина;ВидЗанятия;Группа;Аудитория", "650;750;3200;1650;1200;1000", "НомерПары", "", False
    MakeReport "З01_Расписание аудитории", "З01_Расписание аудитории", "='Расписание аудитории ' & [Аудитория] & ' (' & [Корпус] & ')'", "Дата", _
        "НомерПары;Начало;Группа;Дисциплина;ВидЗанятия;Преподаватель", "650;750;1100;3000;1650;2200", "НомерПары", "", False
    MakeReport "З01_Часы групп по дисциплинам", "З01_Часы групп по дисциплинам", "Количество часов групп по дисциплинам", "Группа", _
        "Дисциплина;ЧасовПоПлану;Пар;ЧасовПоРасписанию", "3800;1600;900;2100", "Дисциплина", "Пар;ЧасовПоРасписанию"
    MakeReport "З01_Часы преподавателей по дисциплинам", "З01_Часы преподавателей по дисциплинам", "Количество часов преподавателей по дисциплинам", "Преподаватель", _
        "Дисциплина;Группа;Пар;Часов", "3800;1500;900;1100", "Дисциплина;Группа", "Пар;Часов"
    MakeReport "З01_Часы аудиторий", "З01_Часы аудиторий", "Занятость аудиторий (часы)", "", _
        "Аудитория;Корпус;Пар;Часов", "1500;3000;1200;1200", "Аудитория", "Пар;Часов"

    ' --- управляющая форма ---
    CF_Begin "Задание 01. Расписание занятий"
    CF_Param "cboГруппа", "Группа", "SELECT КодГруппы, Группа FROM [З01_Группы] ORDER BY Группа", 2, "0;4500", "1"
    CF_Param "cboПреподаватель", "Преподаватель", "SELECT КодПреподавателя, Фамилия & ' ' & Имя & ' ' & Отчество FROM [З01_Преподаватели] ORDER BY Фамилия", 2, "0;4500", "1"
    CF_Param "cboАудитория", "Аудитория", "SELECT КодАудитории, Аудитория & ' (' & Корпус & ')' FROM [З01_Аудитории] ORDER BY Аудитория", 2, "0;4500", "1"
    CF_Param "cboС", "Период с", "SELECT Дата FROM [З01_Календарь] ORDER BY Дата", 1, "", SD(g01Start)
    CF_Param "cboПо", "Период по", "SELECT Дата FROM [З01_Календарь] ORDER BY Дата", 1, "", SD(g01Start + 13)
    CF_Columns
    CF_Btn 0, "Группы студентов", "F:З01_Группы"
    CF_Btn 0, "Аудитории", "F:З01_Аудитории"
    CF_Btn 0, "Преподаватели", "F:З01_Преподаватели"
    CF_Btn 0, "Учебные дисциплины", "F:З01_Дисциплины"
    CF_Btn 0, "Расписание пар (звонков)", "F:З01_Пары"
    CF_Btn 0, "Календарь на семестр", "F:З01_Календарь"
    CF_Btn 0, "Занятия (ввод расписания)", "F:З01_Занятия"
    CF_Btn 1, "Часы групп по дисциплинам", "Q:З01_Часы групп по дисциплинам"
    CF_Btn 1, "Часы преподавателей по дисциплинам", "Q:З01_Часы преподавателей по дисциплинам"
    CF_Btn 1, "Нагрузка преподавателей", "Q:З01_Нагрузка преподавателей"
    CF_Btn 1, "Часы аудиторий", "Q:З01_Часы аудиторий"
    CF_Btn 1, "Сводная: часы групп (перекрестный)", "Q:З01_Сводная часы групп по дисциплинам"
    CF_Btn 1, "Расписание группы", "Q:З01_Расписание группы|cboГруппа|cboС|cboПо"
    CF_Btn 1, "Расписание преподавателя", "Q:З01_Расписание преподавателя|cboПреподаватель|cboС|cboПо"
    CF_Btn 1, "Расписание аудитории", "Q:З01_Расписание аудитории|cboАудитория|cboС|cboПо"
    CF_Btn 2, "Расписание группы", "R:З01_Расписание группы|cboГруппа|cboС|cboПо"
    CF_Btn 2, "Расписание преподавателя", "R:З01_Расписание преподавателя|cboПреподаватель|cboС|cboПо"
    CF_Btn 2, "Расписание аудитории", "R:З01_Расписание аудитории|cboАудитория|cboС|cboПо"
    CF_Btn 2, "Часы групп по дисциплинам", "R:З01_Часы групп по дисциплинам"
    CF_Btn 2, "Часы преподавателей по дисциплинам", "R:З01_Часы преподавателей по дисциплинам"
    CF_Btn 2, "Занятость аудиторий", "R:З01_Часы аудиторий"
    CF_End "З01_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 02. База данных «Абитуриенты»
'==============================================================================
Private Sub BuildT02()
    Dim s As String
    Const FR As String = "[Forms]![З02_Управляющая]!"

    MakeTable "З02_НаселенныеПункты", "КодПункта COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, НаселенныйПункт TEXT(50) NOT NULL, Регион TEXT(50)"
    MakeTable "З02_Специальности", "КодСпециальности COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Шифр TEXT(10), Специальность TEXT(80) NOT NULL, ПланПриема INTEGER"
    MakeTable "З02_Абитуриенты", "КодАбитуриента COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), " & _
        "ДатаРождения DATETIME, КодПункта LONG, Адрес TEXT(100), Телефон TEXT(20), СреднийБалл DOUBLE, КодСпециальности LONG, ДатаПодачи DATETIME"
    MakeRel "З02_НаселенныеПункты", "КодПункта", "З02_Абитуриенты", "КодПункта"
    MakeRel "З02_Специальности", "КодСпециальности", "З02_Абитуриенты", "КодСпециальности"
    Lookup "З02_Абитуриенты", "КодПункта", "SELECT КодПункта, НаселенныйПункт FROM [З02_НаселенныеПункты] ORDER BY НаселенныйПункт"
    Lookup "З02_Абитуриенты", "КодСпециальности", "SELECT КодСпециальности, Специальность FROM [З02_Специальности] ORDER BY Специальность"

    InsRows "З02_НаселенныеПункты", "1,'Москва','Москва'|2,'Подольск','Московская обл.'|3,'Тула','Тульская обл.'|" & _
        "4,'Калуга','Калужская обл.'|5,'Серпухов','Московская обл.'|6,'Рязань','Рязанская обл.'|7,'Коломна','Московская обл.'"
    InsRows "З02_Специальности", "1,'09.03.03','Прикладная информатика',25|2,'38.03.01','Экономика',30|" & _
        "3,'40.03.01','Юриспруденция',20|4,'37.03.01','Психология',15|5,'44.03.01','Педагогическое образование',15"
    s = "1,'Александров','Дмитрий','Игоревич',#03/15/2008#,1,'ул. Ленина, 5-12','+7 (916) 111-22-33',4.8,1,{d-20}"
    s = s & "|2,'Белова','Мария','Сергеевна',{m2007},2,'ул. Советская, 10-3','+7 (903) 222-33-44',4.6,2,{d-18}"
    s = s & "|3,'Васильев','Игорь','Павлович',#11/02/2007#,3,'пр. Ленина, 45-8','+7 (910) 333-44-55',3.9,1,{d-17}"
    s = s & "|4,'Григорьева','Анна','Олеговна',#07/21/2008#,1,'ул. Тверская, 7-21','+7 (925) 444-55-66',4.2,3,{d-16}"
    s = s & "|5,'Дмитриев','Павел','Андреевич',#01/30/2008#,4,'ул. Кирова, 3-4','+7 (920) 555-66-77',3.4,2,{d-15}"
    s = s & "|6,'Егорова','Ольга','Викторовна',#05/12/2008#,5,'ул. Чехова, 12-1','+7 (915) 666-77-88',4.9,4,{d-14}"
    s = s & "|7,'Жуков','Никита','Сергеевич',#09/09/2007#,2,'ул. Мира, 8-15','+7 (916) 777-88-99',3.2,5,{d-12}"
    s = s & "|8,'Зайцева','Екатерина','Дмитриевна',{m2008},6,'ул. Есенина, 22-7','+7 (910) 888-99-00',4.4,1,{d-11}"
    s = s & "|9,'Иванов','Артём','Алексеевич',#12/05/2007#,1,'ул. Садовая, 1-9','+7 (903) 123-45-67',3.7,3,{d-10}"
    s = s & "|10,'Козлова','Дарья','Ивановна',#04/18/2008#,3,'ул. Пушкина, 14-2','+7 (953) 234-56-78',4.7,2,{d-9}"
    s = s & "|11,'Лебедев','Максим','Олегович',#02/27/2008#,4,'ул. Гагарина, 6-30','+7 (920) 345-67-89',3.1,1,{d-8}"
    s = s & "|12,'Морозова','Полина','Андреевна',#06/06/2008#,5,'ул. Победы, 9-5','+7 (915) 456-78-90',4,4,{d-7}"
    s = s & "|13,'Николаев','Кирилл','Романович',#10/14/2007#,6,'ул. Садовая, 18-11','+7 (910) 567-89-01',3.6,5,{d-5}"
    s = s & "|14,'Орлова','Виктория','Павловна',#08/25/2008#,1,'ул. Арбат, 20-4','+7 (916) 678-90-12',5,1,{d-4}"
    s = s & "|15,'Павлов','Егор','Дмитриевич',#03/03/2008#,2,'ул. Школьная, 4-17','+7 (903) 789-01-23',3.3,3,{d-3}"
    s = s & "|16,'Романова','София','Игоревна',#12/19/2007#,3,'ул. Октябрьская, 2-6','+7 (910) 890-12-34',4.5,2,{d-2}"
    InsRows "З02_Абитуриенты", s

    s = "SELECT a.КодАбитуриента, a.Фамилия & ' ' & a.Имя & ' ' & a.Отчество AS ФИО, a.Имя AS Обращение, a.ДатаРождения, "
    s = s & "a.КодПункта, n.НаселенныйПункт, a.Адрес, a.Телефон, a.СреднийБалл, "
    s = s & "IIf(a.СреднийБалл>=4.5,'Отличник',IIf(a.СреднийБалл>=3.5,'Хорошист','Троечник')) AS Успеваемость, "
    s = s & "a.КодСпециальности, s.Шифр, s.Специальность, a.ДатаПодачи "
    s = s & "FROM ([З02_Абитуриенты] AS a LEFT JOIN [З02_НаселенныеПункты] AS n ON a.КодПункта = n.КодПункта) "
    s = s & "LEFT JOIN [З02_Специальности] AS s ON a.КодСпециальности = s.КодСпециальности"
    MakeQuery "З02_Абитуриенты подробно", s
    MakeQuery "З02_Количество по населенным пунктам", _
        "SELECT n.НаселенныйПункт, n.Регион, Count(a.КодАбитуриента) AS Абитуриентов " & _
        "FROM [З02_НаселенныеПункты] AS n LEFT JOIN [З02_Абитуриенты] AS a ON n.КодПункта = a.КодПункта " & _
        "GROUP BY n.НаселенныйПункт, n.Регион ORDER BY Count(a.КодАбитуриента) DESC, n.НаселенныйПункт"
    MakeQuery "З02_Количество по специальностям", _
        "SELECT s.Шифр, s.Специальность, s.ПланПриема, Count(a.КодАбитуриента) AS Заявлений, " & _
        "IIf(Nz(s.ПланПриема,0)=0,Null,Round(Count(a.КодАбитуриента)/s.ПланПриема,2)) AS КонкурсНаМесто " & _
        "FROM [З02_Специальности] AS s LEFT JOIN [З02_Абитуриенты] AS a ON s.КодСпециальности = a.КодСпециальности " & _
        "GROUP BY s.Шифр, s.Специальность, s.ПланПриема ORDER BY s.Шифр"
    MakeQuery "З02_Количество по успеваемости", _
        "SELECT Успеваемость, Count(*) AS Абитуриентов, Round(Avg(СреднийБалл),2) AS СреднийБаллКатегории " & _
        "FROM [З02_Абитуриенты подробно] GROUP BY Успеваемость"
    MakeQuery "З02_Сводная пункты и специальности", _
        "TRANSFORM Count(КодАбитуриента) SELECT НаселенныйПункт, Count(КодАбитуриента) AS Всего " & _
        "FROM [З02_Абитуриенты подробно] GROUP BY НаселенныйПункт PIVOT Специальность"
    MakeQuery "З02_Абитуриенты специальности", _
        "SELECT Специальность, ФИО, НаселенныйПункт, Телефон, СреднийБалл, Успеваемость FROM [З02_Абитуриенты подробно] " & _
        "WHERE КодСпециальности = " & FR & "[cboСпециальность] ORDER BY СреднийБалл DESC, ФИО"
    MakeQuery "З02_Абитуриенты категории", _
        "SELECT Успеваемость, ФИО, Специальность, НаселенныйПункт, СреднийБалл FROM [З02_Абитуриенты подробно] " & _
        "WHERE Успеваемость = " & FR & "[cboКатегория] ORDER BY СреднийБалл DESC, ФИО"
    MakeQuery "З02_Именинники месяца", _
        "SELECT ФИО, ДатаРождения, Телефон, Специальность FROM [З02_Абитуриенты подробно] " & _
        "WHERE Month(ДатаРождения) = Month(Date()) ORDER BY Day(ДатаРождения)"
    MakeQuery "З02_Поздравление", _
        "SELECT ФИО, Обращение, ДатаРождения FROM [З02_Абитуриенты подробно] WHERE КодАбитуриента = " & FR & "[cboАбитуриент]"

    MakeTableForm "З02_НаселенныеПункты", "З02_НаселенныеПункты", "Населенные пункты"
    MakeTableForm "З02_Специальности", "З02_Специальности", "Специальности"
    MakeTableForm "З02_Абитуриенты", "З02_Абитуриенты", "Абитуриенты"
    MakeTableForm "З02_Абитуриенты подчиненная", "З02_Абитуриенты", "Абитуриенты", 2, "КодСпециальности"
    AddSubform "З02_Специальности", "З02_Абитуриенты подчиненная", "КодСпециальности", "КодСпециальности", "Абитуриенты, выбравшие специальность"

    MakeReport "З02_Количественный состав абитуриентов", "З02_Абитуриенты подробно", "Количественный состав абитуриентов по специальностям", "Специальность", _
        "ФИО;НаселенныйПункт;СреднийБалл;Успеваемость", "3700;2200;1400;1600", "ФИО"
    MakeReport "З02_Абитуриенты по успеваемости", "З02_Абитуриенты подробно", "Отличники, хорошисты, троечники", "Успеваемость", _
        "ФИО;Специальность;СреднийБалл", "3700;3400;1400", "-СреднийБалл;ФИО"
    MakeReport "З02_Абитуриенты по населенным пунктам", "З02_Абитуриенты подробно", "Абитуриенты по населенным пунктам", "НаселенныйПункт", _
        "ФИО;Специальность;Телефон", "3700;3400;2000", "ФИО"
    MakeReport "З02_Список абитуриентов специальности", "З02_Абитуриенты специальности", "='Список абитуриентов: ' & [Специальность]", "", _
        "ФИО;НаселенныйПункт;Телефон;СреднийБалл;Успеваемость", "3200;1800;2000;1300;1500", "-СреднийБалл;ФИО"
    MakeGreeting "З02_Поздравление с днем рождения", "З02_Поздравление", "Дорогой(ая)", _
        "Приёмная комиссия университета от всей души поздравляет Вас с днём рождения! Желаем крепкого здоровья, успешной сдачи вступительных испытаний и поступления на выбранную специальность!", _
        "Приёмная комиссия университета"

    CF_Begin "Задание 02. Абитуриенты"
    CF_Param "cboСпециальность", "Специальность", "SELECT КодСпециальности, Специальность FROM [З02_Специальности] ORDER BY Специальность", 2, "0;4500", "1"
    CF_Param "cboКатегория", "Успеваемость", "Отличник;Хорошист;Троечник", 1, "", """Отличник""", "V"
    CF_Param "cboАбитуриент", "Абитуриент", "SELECT КодАбитуриента, Фамилия & ' ' & Имя & ' ' & Отчество FROM [З02_Абитуриенты] ORDER BY Фамилия, Имя", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Абитуриенты", "F:З02_Абитуриенты"
    CF_Btn 0, "Специальности", "F:З02_Специальности"
    CF_Btn 0, "Населенные пункты", "F:З02_НаселенныеПункты"
    CF_Btn 1, "Абитуриенты (все сведения)", "Q:З02_Абитуриенты подробно"
    CF_Btn 1, "Количество по населенным пунктам", "Q:З02_Количество по населенным пунктам"
    CF_Btn 1, "Количество по специальностям", "Q:З02_Количество по специальностям"
    CF_Btn 1, "Отличники, хорошисты, троечники", "Q:З02_Количество по успеваемости"
    CF_Btn 1, "Сводная: пункты и специальности", "Q:З02_Сводная пункты и специальности"
    CF_Btn 1, "Абитуриенты специальности", "Q:З02_Абитуриенты специальности|cboСпециальность"
    CF_Btn 1, "Абитуриенты выбранной категории", "Q:З02_Абитуриенты категории|cboКатегория"
    CF_Btn 1, "Именинники текущего месяца", "Q:З02_Именинники месяца"
    CF_Btn 2, "Количественный состав абитуриентов", "R:З02_Количественный состав абитуриентов"
    CF_Btn 2, "Абитуриенты по успеваемости", "R:З02_Абитуриенты по успеваемости"
    CF_Btn 2, "Абитуриенты по населенным пунктам", "R:З02_Абитуриенты по населенным пунктам"
    CF_Btn 2, "Список абитуриентов специальности", "R:З02_Список абитуриентов специальности|cboСпециальность"
    CF_Btn 2, "Поздравление с днем рождения", "R:З02_Поздравление с днем рождения|cboАбитуриент"
    CF_End "З02_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 03. База данных «Телефонная книга»
'==============================================================================
Private Sub BuildT03()
    Dim s As String
    Const FR As String = "[Forms]![З03_Управляющая]!"

    MakeTable "З03_Факультеты", "КодФакультета COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Факультет TEXT(80) NOT NULL, Декан TEXT(60)"
    MakeTable "З03_Группы", "КодГруппы COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Группа TEXT(20) NOT NULL, КодФакультета LONG"
    MakeTable "З03_НаселенныеПункты", "КодПункта COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, НаселенныйПункт TEXT(50) NOT NULL"
    MakeTable "З03_Студенты", "КодСтудента COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), " & _
        "КодГруппы LONG, КодПункта LONG, Адрес TEXT(100), Телефон TEXT(20)"
    MakeRel "З03_Факультеты", "КодФакультета", "З03_Группы", "КодФакультета"
    MakeRel "З03_Группы", "КодГруппы", "З03_Студенты", "КодГруппы"
    MakeRel "З03_НаселенныеПункты", "КодПункта", "З03_Студенты", "КодПункта"
    Lookup "З03_Группы", "КодФакультета", "SELECT КодФакультета, Факультет FROM [З03_Факультеты] ORDER BY Факультет"
    Lookup "З03_Студенты", "КодГруппы", "SELECT КодГруппы, Группа FROM [З03_Группы] ORDER BY Группа"
    Lookup "З03_Студенты", "КодПункта", "SELECT КодПункта, НаселенныйПункт FROM [З03_НаселенныеПункты] ORDER BY НаселенныйПункт"

    InsRows "З03_Факультеты", "1,'Информационных технологий','Соколов В.А.'|2,'Экономический','Миронова Т.Н.'|3,'Юридический','Фролов Д.С.'"
    InsRows "З03_Группы", "1,'ИТ-11',1|2,'ИТ-21',1|3,'ЭК-11',2|4,'ЭК-21',2|5,'ЮР-11',3|6,'ЮР-21',3"
    InsRows "З03_НаселенныеПункты", "1,'Москва'|2,'Химки'|3,'Мытищи'|4,'Балашиха'|5,'Люберцы'"
    s = "1,'Андреев','Илья','Сергеевич',1,1,'ул. Вавилова, 12-45','+7 (916) 101-11-11'"
    s = s & "|2,'Борисова','Алина','Олеговна',1,2,'ул. Молодежная, 3-8','+7 (926) 202-22-22'"
    s = s & "|3,'Волков','Денис','Андреевич',2,3,'ул. Мира, 17-2','+7 (903) 303-33-33'"
    s = s & "|4,'Гусева','Ксения','Павловна',2,1,'Ленинский пр., 60-14','+7 (915) 404-44-44'"
    s = s & "|5,'Данилов','Роман','Игоревич',3,4,'ул. Советская, 25-9','+7 (925) 505-55-55'"
    s = s & "|6,'Ершова','Валерия','Денисовна',3,5,'Октябрьский пр., 101-33','+7 (916) 606-66-66'"
    s = s & "|7,'Захаров','Тимур','Русланович',4,1,'ул. Профсоюзная, 88-7','+7 (903) 707-77-77'"
    s = s & "|8,'Ильина','Софья','Максимовна',4,2,'ул. Победы, 5-19','+7 (926) 808-88-88'"
    s = s & "|9,'Комаров','Глеб','Витальевич',5,3,'ул. Колпакова, 11-40','+7 (915) 909-99-99'"
    s = s & "|10,'Лаптева','Анастасия','Юрьевна',5,4,'ш. Энтузиастов, 7-3','+7 (925) 111-00-22'"
    s = s & "|11,'Медведев','Арсений','Олегович',6,5,'ул. Кирова, 9-12','+7 (916) 222-00-33'"
    s = s & "|12,'Никитина','Ульяна','Артёмовна',6,1,'ул. Тверская, 15-6','+7 (903) 333-00-44'"
    s = s & "|13,'Осипов','Даниил','Евгеньевич',1,3,'ул. Юбилейная, 30-21','+7 (926) 444-00-55'"
    s = s & "|14,'Полякова','Вероника','Ильинична',2,4,'ул. Лесная, 4-1','+7 (915) 555-00-66'"
    s = s & "|15,'Родионов','Марк','Александрович',3,1,'ул. Гарибальди, 23-54','+7 (925) 666-00-77'"
    s = s & "|16,'Соболева','Алиса','Романовна',4,5,'ул. Южная, 8-16','+7 (916) 777-00-88'"
    s = s & "|17,'Тихонов','Лев','Николаевич',5,2,'ул. Библиотечная, 2-10','+7 (903) 888-00-99'"
    s = s & "|18,'Андреева','Ева','Сергеевна',6,1,'ул. Вавилова, 12-45','+7 (926) 999-00-11'"
    InsRows "З03_Студенты", s

    s = "SELECT s.КодСтудента, s.Фамилия, s.Имя, s.Отчество, s.Фамилия & ' ' & s.Имя & ' ' & s.Отчество AS ФИО, "
    s = s & "Left(s.Фамилия,1) AS Буква, s.Телефон, s.КодПункта, n.НаселенныйПункт, s.Адрес, s.КодГруппы, g.Группа, f.Факультет "
    s = s & "FROM (([З03_Студенты] AS s LEFT JOIN [З03_Группы] AS g ON s.КодГруппы = g.КодГруппы) "
    s = s & "LEFT JOIN [З03_Факультеты] AS f ON g.КодФакультета = f.КодФакультета) "
    s = s & "LEFT JOIN [З03_НаселенныеПункты] AS n ON s.КодПункта = n.КодПункта"
    MakeQuery "З03_Телефонная книга", s
    MakeQuery "З03_Поиск по фамилии", _
        "PARAMETERS " & FR & "[txtФамилия] Text ( 255 ); " & _
        "SELECT ФИО, Телефон, НаселенныйПункт, Адрес, Группа FROM [З03_Телефонная книга] " & _
        "WHERE Фамилия Like " & FR & "[txtФамилия] & '*' ORDER BY ФИО"
    MakeQuery "З03_Поиск по телефону", _
        "PARAMETERS " & FR & "[txtТелефон] Text ( 255 ); " & _
        "SELECT Телефон, ФИО, НаселенныйПункт, Адрес, Группа FROM [З03_Телефонная книга] " & _
        "WHERE Телефон Like '*' & " & FR & "[txtТелефон] & '*' ORDER BY ФИО"
    MakeQuery "З03_Студенты группы", _
        "SELECT Группа, ФИО, Телефон, НаселенныйПункт, Адрес FROM [З03_Телефонная книга] " & _
        "WHERE КодГруппы = " & FR & "[cboГруппа] ORDER BY ФИО"
    MakeQuery "З03_Студенты населенного пункта", _
        "SELECT НаселенныйПункт, ФИО, Группа, Телефон, Адрес FROM [З03_Телефонная книга] " & _
        "WHERE КодПункта = " & FR & "[cboПункт] ORDER BY ФИО"
    MakeQuery "З03_Количество студентов по группам", _
        "SELECT f.Факультет, g.Группа, Count(s.КодСтудента) AS Студентов " & _
        "FROM ([З03_Факультеты] AS f INNER JOIN [З03_Группы] AS g ON f.КодФакультета = g.КодФакультета) " & _
        "LEFT JOIN [З03_Студенты] AS s ON g.КодГруппы = s.КодГруппы GROUP BY f.Факультет, g.Группа ORDER BY f.Факультет, g.Группа"

    MakeTableForm "З03_Факультеты", "З03_Факультеты", "Факультеты"
    MakeTableForm "З03_Группы", "З03_Группы", "Группы"
    MakeTableForm "З03_НаселенныеПункты", "З03_НаселенныеПункты", "Населенные пункты"
    MakeTableForm "З03_Студенты", "З03_Студенты", "Студенты: адреса и телефоны"
    MakeTableForm "З03_Группы подчиненная", "З03_Группы", "Группы", 2, "КодФакультета"
    MakeTableForm "З03_Студенты подчиненная", "З03_Студенты", "Студенты", 2, "КодГруппы"
    AddSubform "З03_Факультеты", "З03_Группы подчиненная", "КодФакультета", "КодФакультета", "Группы факультета"
    AddSubform "З03_Группы", "З03_Студенты подчиненная", "КодГруппы", "КодГруппы", "Студенты группы"

    MakeReport "З03_Телефонная книга студентов", "З03_Телефонная книга", "Телефонная книга студентов", "Буква", _
        "ФИО;Телефон;Группа;НаселенныйПункт;Адрес", "3000;1900;900;1500;2500", "ФИО"
    MakeReport "З03_Студенты по факультетам", "З03_Телефонная книга", "Студенты по факультетам и группам", "Факультет;Группа", _
        "ФИО;Телефон;НаселенныйПункт;Адрес", "3200;1900;1600;2800", "ФИО"
    MakeReport "З03_Студенты группы", "З03_Студенты группы", "='Список студентов группы ' & [Группа]", "", _
        "ФИО;Телефон;НаселенныйПункт;Адрес", "3200;1900;1600;3000", "ФИО"
    MakeReport "З03_Студенты населенного пункта", "З03_Студенты населенного пункта", "='Студенты, проживающие: ' & [НаселенныйПункт]", "", _
        "ФИО;Группа;Телефон;Адрес", "3200;1000;1900;3500", "ФИО"

    CF_Begin "Задание 03. Телефонная книга"
    CF_Param "txtФамилия", "Фамилия (начало)", "", , , , "T"
    CF_Param "txtТелефон", "Телефон (часть номера)", "", , , , "T"
    CF_Param "cboГруппа", "Группа", "SELECT КодГруппы, Группа FROM [З03_Группы] ORDER BY Группа", 2, "0;4500", "1"
    CF_Param "cboПункт", "Населенный пункт", "SELECT КодПункта, НаселенныйПункт FROM [З03_НаселенныеПункты] ORDER BY НаселенныйПункт", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Студенты", "F:З03_Студенты"
    CF_Btn 0, "Группы", "F:З03_Группы"
    CF_Btn 0, "Факультеты", "F:З03_Факультеты"
    CF_Btn 0, "Населенные пункты", "F:З03_НаселенныеПункты"
    CF_Btn 1, "Телефонная книга", "Q:З03_Телефонная книга"
    CF_Btn 1, "Телефон и адрес по фамилии", "Q:З03_Поиск по фамилии|txtФамилия"
    CF_Btn 1, "Фамилия и адрес по телефону", "Q:З03_Поиск по телефону|txtТелефон"
    CF_Btn 1, "Студенты выбранной группы", "Q:З03_Студенты группы|cboГруппа"
    CF_Btn 1, "Студенты из населенного пункта", "Q:З03_Студенты населенного пункта|cboПункт"
    CF_Btn 1, "Количество студентов по группам", "Q:З03_Количество студентов по группам"
    CF_Btn 2, "Телефонная книга студентов", "R:З03_Телефонная книга студентов"
    CF_Btn 2, "Студенты по факультетам", "R:З03_Студенты по факультетам"
    CF_Btn 2, "Список студентов группы", "R:З03_Студенты группы|cboГруппа"
    CF_Btn 2, "Студенты населенного пункта", "R:З03_Студенты населенного пункта|cboПункт"
    CF_End "З03_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 04. База данных «Учебные планы»
'==============================================================================
Private Sub BuildT04()
    Dim s As String
    Const FR As String = "[Forms]![З04_Управляющая]!"

    MakeTable "З04_Специальности", "КодСпециальности COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Шифр TEXT(10), Специальность TEXT(80) NOT NULL"
    MakeTable "З04_Дисциплины", "КодДисциплины COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Дисциплина TEXT(80) NOT NULL"
    MakeTable "З04_Семестры", "НомерСеместра INTEGER CONSTRAINT PrimaryKey PRIMARY KEY, Курс INTEGER, Период TEXT(20)"
    MakeTable "З04_ФормыКонтроля", "КодФормы COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, ФормаКонтроля TEXT(20) NOT NULL"
    MakeTable "З04_УчебныйПлан", "КодЗаписи COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, КодСпециальности LONG NOT NULL, НомерСеместра INTEGER NOT NULL, " & _
        "КодДисциплины LONG NOT NULL, Лекции INTEGER, Практикум INTEGER, СамРабота INTEGER, КодФормы LONG, " & _
        "CONSTRAINT UqПлан UNIQUE (КодСпециальности, НомерСеместра, КодДисциплины)"
    MakeRel "З04_Специальности", "КодСпециальности", "З04_УчебныйПлан", "КодСпециальности", True
    MakeRel "З04_Семестры", "НомерСеместра", "З04_УчебныйПлан", "НомерСеместра"
    MakeRel "З04_Дисциплины", "КодДисциплины", "З04_УчебныйПлан", "КодДисциплины"
    MakeRel "З04_ФормыКонтроля", "КодФормы", "З04_УчебныйПлан", "КодФормы"
    Lookup "З04_УчебныйПлан", "КодСпециальности", "SELECT КодСпециальности, Шифр & ' ' & Специальность FROM [З04_Специальности] ORDER BY Шифр"
    Lookup "З04_УчебныйПлан", "НомерСеместра", "SELECT НомерСеместра, Курс & ' курс, ' & Период FROM [З04_Семестры] ORDER BY НомерСеместра", 2, "567;2000"
    Lookup "З04_УчебныйПлан", "КодДисциплины", "SELECT КодДисциплины, Дисциплина FROM [З04_Дисциплины] ORDER BY Дисциплина"
    Lookup "З04_УчебныйПлан", "КодФормы", "SELECT КодФормы, ФормаКонтроля FROM [З04_ФормыКонтроля] ORDER BY КодФормы"

    InsRows "З04_Специальности", "1,'09.03.03','Прикладная информатика'|2,'38.03.01','Экономика'|3,'40.03.01','Юриспруденция'"
    InsRows "З04_Дисциплины", "1,'Математика'|2,'Информатика'|3,'Иностранный язык'|4,'История'|5,'Философия'|6,'Базы данных'|" & _
        "7,'Программирование'|8,'Экономическая теория'|9,'Бухгалтерский учет'|10,'Статистика'|11,'Теория государства и права'|" & _
        "12,'Гражданское право'|13,'Физическая культура'"
    InsRows "З04_Семестры", "1,1,'Осенний'|2,1,'Весенний'|3,2,'Осенний'|4,2,'Весенний'|5,3,'Осенний'|6,3,'Весенний'|7,4,'Осенний'|8,4,'Весенний'"
    InsRows "З04_ФормыКонтроля", "1,'Зачет'|2,'Дифзачет'|3,'Экзамен'"
    ' специальность, семестр, дисциплина, лекции, практикум, самработа, форма контроля
    s = "1,1,1,36,36,72,3|1,1,2,18,36,54,3|1,1,3,0,36,36,1|1,1,4,18,18,36,2|1,2,1,36,36,72,3|1,2,7,18,36,54,2|1,2,3,0,36,36,1"
    s = s & "|1,3,6,36,36,72,3|1,3,7,18,36,54,3|1,3,5,18,18,36,1|1,4,6,18,36,54,2|1,4,10,18,18,36,1"
    s = s & "|2,1,1,36,36,72,3|2,1,8,36,18,54,3|2,1,3,0,36,36,1|2,1,4,18,18,36,2|2,2,8,18,36,54,3|2,2,2,18,18,36,1"
    s = s & "|2,3,9,36,36,72,3|2,3,10,18,36,54,2|2,4,9,18,36,54,3|2,4,5,18,18,36,1"
    s = s & "|3,1,11,36,36,72,3|3,1,4,18,18,36,2|3,1,3,0,36,36,1|3,2,12,36,36,72,3|3,2,5,18,18,36,1|3,2,2,18,18,36,1"
    InsRows "З04_УчебныйПлан", s, "КодСпециальности, НомерСеместра, КодДисциплины, Лекции, Практикум, СамРабота, КодФормы"

    s = "SELECT u.КодЗаписи, u.КодСпециальности, s.Шифр, s.Специальность, u.НомерСеместра AS Семестр, se.Курс, d.Дисциплина, "
    s = s & "u.Лекции, u.Практикум, u.СамРабота, Nz(u.Лекции,0)+Nz(u.Практикум,0) AS Аудиторных, "
    s = s & "Nz(u.Лекции,0)+Nz(u.Практикум,0)+Nz(u.СамРабота,0) AS Всего, f.ФормаКонтроля "
    s = s & "FROM ((([З04_УчебныйПлан] AS u INNER JOIN [З04_Специальности] AS s ON u.КодСпециальности = s.КодСпециальности) "
    s = s & "INNER JOIN [З04_Семестры] AS se ON u.НомерСеместра = se.НомерСеместра) "
    s = s & "INNER JOIN [З04_Дисциплины] AS d ON u.КодДисциплины = d.КодДисциплины) "
    s = s & "LEFT JOIN [З04_ФормыКонтроля] AS f ON u.КодФормы = f.КодФормы"
    MakeQuery "З04_Учебный план подробно", s
    MakeQuery "З04_План специальности", _
        "SELECT Шифр, Специальность, Семестр, Курс, Дисциплина, Лекции, Практикум, СамРабота, Всего, ФормаКонтроля " & _
        "FROM [З04_Учебный план подробно] WHERE КодСпециальности = " & FR & "[cboСпециальность] ORDER BY Семестр, Дисциплина"
    MakeQuery "З04_План семестра", _
        "SELECT Специальность, Семестр, Дисциплина, Лекции, Практикум, СамРабота, Всего, ФормаКонтроля " & _
        "FROM [З04_Учебный план подробно] WHERE КодСпециальности = " & FR & "[cboСпециальность] " & _
        "AND Семестр = " & FR & "[cboСеместр] ORDER BY Дисциплина"
    MakeQuery "З04_Итоги по семестрам", _
        "SELECT Специальность, Семестр, Count(*) AS Дисциплин, Sum(Лекции) AS Лекций, Sum(Практикум) AS Практикумов, " & _
        "Sum(СамРабота) AS СамРаботы, Sum(Всего) AS ВсегоЧасов, Sum(IIf(ФормаКонтроля='Экзамен',1,0)) AS Экзаменов, " & _
        "Sum(IIf(ФормаКонтроля='Дифзачет',1,0)) AS Дифзачетов, Sum(IIf(ФормаКонтроля='Зачет',1,0)) AS Зачетов " & _
        "FROM [З04_Учебный план подробно] GROUP BY Специальность, Семестр ORDER BY Специальность, Семестр"
    MakeQuery "З04_Сводный учебный план", _
        "TRANSFORM Sum(Всего) SELECT Специальность, Дисциплина, Sum(Всего) AS ИтогоЧасов " & _
        "FROM [З04_Учебный план подробно] GROUP BY Специальность, Дисциплина PIVOT 'Семестр ' & Семестр"
    MakeQuery "З04_Формы контроля по специальностям", _
        "TRANSFORM Count(КодЗаписи) SELECT Специальность, Count(КодЗаписи) AS Всего " & _
        "FROM [З04_Учебный план подробно] GROUP BY Специальность PIVOT ФормаКонтроля"
    MakeQuery "З04_Дисциплины вне учебных планов", _
        "SELECT d.Дисциплина FROM [З04_Дисциплины] AS d LEFT JOIN [З04_УчебныйПлан] AS u ON d.КодДисциплины = u.КодДисциплины " & _
        "WHERE u.КодДисциплины Is Null ORDER BY d.Дисциплина"

    MakeTableForm "З04_Специальности", "З04_Специальности", "Специальности"
    MakeTableForm "З04_Дисциплины", "З04_Дисциплины", "Учебные дисциплины"
    MakeTableForm "З04_Семестры", "З04_Семестры", "Семестры"
    MakeTableForm "З04_ФормыКонтроля", "З04_ФормыКонтроля", "Формы контроля"
    MakeTableForm "З04_УчебныйПлан", "З04_УчебныйПлан", "Учебный план (часы и формы контроля)"
    MakeTableForm "З04_УчебныйПлан подчиненная", "З04_УчебныйПлан", "Учебный план", 2, "КодСпециальности;КодЗаписи"
    AddSubform "З04_Специальности", "З04_УчебныйПлан подчиненная", "КодСпециальности", "КодСпециальности", "Учебный план специальности"

    MakeReport "З04_Учебные планы специальностей", "З04_Учебный план подробно", "Учебные планы по специальностям на все семестры", "Специальность;Семестр", _
        "Дисциплина;Лекции;Практикум;СамРабота;Всего;ФормаКонтроля", "3300;1000;1150;1250;900;1500", "Дисциплина", "Лекции;Практикум;СамРабота;Всего"
    MakeReport "З04_Учебный план специальности", "З04_План специальности", "='Учебный план: ' & [Шифр] & ' ' & [Специальность]", "Семестр", _
        "Дисциплина;Лекции;Практикум;СамРабота;Всего;ФормаКонтроля", "3300;1000;1150;1250;900;1500", "Дисциплина", "Лекции;Практикум;СамРабота;Всего"
    MakeReport "З04_Итоги по семестрам", "З04_Итоги по семестрам", "Итоги учебных планов по семестрам", "Специальность", _
        "Семестр;Дисциплин;Лекций;Практикумов;СамРаботы;ВсегоЧасов;Экзаменов;Зачетов", "950;1100;900;1350;1200;1300;1200;900", "Семестр", _
        "Дисциплин;Лекций;Практикумов;СамРаботы;ВсегоЧасов;Экзаменов;Зачетов", False

    CF_Begin "Задание 04. Учебные планы"
    CF_Param "cboСпециальность", "Специальность", "SELECT КодСпециальности, Шифр & ' ' & Специальность FROM [З04_Специальности] ORDER BY Шифр", 2, "0;4500", "1"
    CF_Param "cboСеместр", "Семестр", "SELECT НомерСеместра, НомерСеместра & ' семестр (' & Курс & ' курс)' FROM [З04_Семестры] ORDER BY НомерСеместра", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Учебный план (ввод)", "F:З04_УчебныйПлан"
    CF_Btn 0, "Специальности и их планы", "F:З04_Специальности"
    CF_Btn 0, "Дисциплины", "F:З04_Дисциплины"
    CF_Btn 0, "Семестры", "F:З04_Семестры"
    CF_Btn 0, "Формы контроля", "F:З04_ФормыКонтроля"
    CF_Btn 1, "Учебный план (все сведения)", "Q:З04_Учебный план подробно"
    CF_Btn 1, "План специальности", "Q:З04_План специальности|cboСпециальность"
    CF_Btn 1, "План специальности на семестр", "Q:З04_План семестра|cboСпециальность|cboСеместр"
    CF_Btn 1, "Итоги по семестрам", "Q:З04_Итоги по семестрам"
    CF_Btn 1, "Сводный план (перекрестный)", "Q:З04_Сводный учебный план"
    CF_Btn 1, "Формы контроля по специальностям", "Q:З04_Формы контроля по специальностям"
    CF_Btn 1, "Дисциплины вне учебных планов", "Q:З04_Дисциплины вне учебных планов"
    CF_Btn 2, "Учебные планы специальностей", "R:З04_Учебные планы специальностей"
    CF_Btn 2, "Учебный план специальности", "R:З04_Учебный план специальности|cboСпециальность"
    CF_Btn 2, "Итоги по семестрам", "R:З04_Итоги по семестрам"
    CF_End "З04_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 05. База данных «Сотрудники»
'==============================================================================
Private Sub BuildT05()
    Dim s As String
    Const FR As String = "[Forms]![З05_Управляющая]!"

    MakeTable "З05_Отделы", "КодОтдела COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Отдел TEXT(80) NOT NULL, Телефон TEXT(20)"
    MakeTable "З05_Образование", "КодОбразования COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Образование TEXT(60) NOT NULL, Уровень TEXT(20)"
    MakeTable "З05_НаселенныеПункты", "КодПункта COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, НаселенныйПункт TEXT(50) NOT NULL"
    MakeTable "З05_Сотрудники", "КодСотрудника COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), " & _
        "ДатаРождения DATETIME, КодОтдела LONG, Должность TEXT(60), КодОбразования LONG, КодПункта LONG, Адрес TEXT(100), Телефон TEXT(20)"
    MakeRel "З05_Отделы", "КодОтдела", "З05_Сотрудники", "КодОтдела"
    MakeRel "З05_Образование", "КодОбразования", "З05_Сотрудники", "КодОбразования"
    MakeRel "З05_НаселенныеПункты", "КодПункта", "З05_Сотрудники", "КодПункта"
    Lookup "З05_Сотрудники", "КодОтдела", "SELECT КодОтдела, Отдел FROM [З05_Отделы] ORDER BY Отдел"
    Lookup "З05_Сотрудники", "КодОбразования", "SELECT КодОбразования, Образование FROM [З05_Образование] ORDER BY КодОбразования"
    Lookup "З05_Сотрудники", "КодПункта", "SELECT КодПункта, НаселенныйПункт FROM [З05_НаселенныеПункты] ORDER BY НаселенныйПункт"

    InsRows "З05_Отделы", "1,'Ректорат','+7 (495) 100-01-01'|2,'Бухгалтерия','+7 (495) 100-02-02'|3,'Отдел кадров','+7 (495) 100-03-03'|" & _
        "4,'Учебный отдел','+7 (495) 100-04-04'|5,'Информационно-технический отдел','+7 (495) 100-05-05'|6,'Юридический отдел','+7 (495) 100-06-06'"
    InsRows "З05_Образование", "1,'Высшее (специалитет, магистратура)','Высшее'|2,'Высшее (бакалавриат)','Высшее'|" & _
        "3,'Среднее профессиональное','Среднее'|4,'Среднее общее','Среднее'"
    InsRows "З05_НаселенныеПункты", "1,'Москва'|2,'Королёв'|3,'Подольск'|4,'Красногорск'|5,'Одинцово'"
    s = "1,'Беляев','Андрей','Викторович',#04/12/1965#,1,'Ректор',1,1,'ул. Остоженка, 10-5','+7 (916) 500-01-01'"
    s = s & "|2,'Воронина','Татьяна','Ивановна',{m1978},2,'Главный бухгалтер',1,2,'пр. Космонавтов, 3-44','+7 (916) 500-02-02'"
    s = s & "|3,'Гришина','Наталья','Петровна',#09/23/1985#,2,'Бухгалтер',3,3,'ул. Кирова, 17-3','+7 (916) 500-03-03'"
    s = s & "|4,'Денисова','Ирина','Александровна',#01/15/1990#,3,'Начальник отдела кадров',2,1,'ул. Лесная, 8-12','+7 (916) 500-04-04'"
    s = s & "|5,'Ефимов','Сергей','Николаевич',{m1983},5,'Системный администратор',3,4,'ул. Ленина, 22-7','+7 (916) 500-05-05'"
    s = s & "|6,'Жаров','Константин','Юрьевич',#07/07/1992#,5,'Программист',2,1,'ул. Новая, 1-90','+7 (916) 500-06-06'"
    s = s & "|7,'Зуева','Людмила','Михайловна',#11/30/1970#,4,'Начальник учебного отдела',1,5,'ул. Маршала Жукова, 5-6','+7 (916) 500-07-07'"
    s = s & "|8,'Исаева','Марина','Олеговна',#03/05/1995#,4,'Методист',2,2,'ул. Циолковского, 9-15','+7 (916) 500-08-08'"
    s = s & "|9,'Карпов','Владимир','Анатольевич',#06/18/1958#,1,'Проректор',1,1,'Кутузовский пр., 30-2','+7 (916) 500-09-09'"
    s = s & "|10,'Лукина','Светлана','Геннадьевна',#10/10/1988#,3,'Специалист по кадрам',4,3,'ул. Садовая, 4-4','+7 (916) 500-10-10'"
    s = s & "|11,'Макаров','Павел','Дмитриевич',#02/22/1998#,5,'Техник',3,4,'ул. Школьная, 6-21','+7 (916) 500-11-11'"
    s = s & "|12,'Никифорова','Ольга','Васильевна',#12/01/1980#,4,'Диспетчер',4,5,'ул. Молодежная, 12-8','+7 (916) 500-12-12'"
    s = s & "|13,'Овчинникова','Юлия','Сергеевна',#08/14/1993#,2,'Кассир',3,1,'ул. Бутлерова, 2-17','+7 (916) 500-13-13'"
    s = s & "|14,'Панов','Игорь','Евгеньевич',#05/29/1975#,5,'Начальник ИТ-отдела',1,2,'ул. Гагарина, 14-30','+7 (916) 500-14-14'"
    InsRows "З05_Сотрудники", s

    s = "SELECT e.КодСотрудника, e.Фамилия & ' ' & e.Имя & ' ' & e.Отчество AS ФИО, e.Имя & ' ' & e.Отчество AS Обращение, "
    s = s & "e.ДатаРождения, Year(Date())-Year(e.ДатаРождения) AS ИсполняетсяЛет, e.КодОтдела, o.Отдел, e.Должность, "
    s = s & "ed.Образование, ed.Уровень, e.КодПункта, n.НаселенныйПункт, e.Адрес, e.Телефон "
    s = s & "FROM (([З05_Сотрудники] AS e LEFT JOIN [З05_Отделы] AS o ON e.КодОтдела = o.КодОтдела) "
    s = s & "LEFT JOIN [З05_Образование] AS ed ON e.КодОбразования = ed.КодОбразования) "
    s = s & "LEFT JOIN [З05_НаселенныеПункты] AS n ON e.КодПункта = n.КодПункта"
    MakeQuery "З05_Сотрудники подробно", s
    MakeQuery "З05_Общее количество сотрудников", _
        "SELECT Count(*) AS ВсегоСотрудников, Sum(IIf(Уровень='Высшее',1,0)) AS СВысшимОбразованием, " & _
        "Sum(IIf(Уровень='Среднее',1,0)) AS СоСреднимОбразованием FROM [З05_Сотрудники подробно]"
    MakeQuery "З05_Количество по отделам", _
        "SELECT o.Отдел, Count(e.КодСотрудника) AS Сотрудников FROM [З05_Отделы] AS o " & _
        "LEFT JOIN [З05_Сотрудники] AS e ON o.КодОтдела = e.КодОтдела GROUP BY o.Отдел ORDER BY o.Отдел"
    MakeQuery "З05_Количество по уровню образования", _
        "SELECT Уровень AS УровеньОбразования, Count(*) AS Сотрудников FROM [З05_Сотрудники подробно] GROUP BY Уровень"
    MakeQuery "З05_Сводная отделы и образование", _
        "TRANSFORM Count(КодСотрудника) SELECT Отдел, Count(КодСотрудника) AS Всего " & _
        "FROM [З05_Сотрудники подробно] GROUP BY Отдел PIVOT Уровень"
    MakeQuery "З05_Сотрудники отдела", _
        "SELECT Отдел, ФИО, Должность, Образование, Телефон, НаселенныйПункт, Адрес FROM [З05_Сотрудники подробно] " & _
        "WHERE КодОтдела = " & FR & "[cboОтдел] ORDER BY ФИО"
    MakeQuery "З05_Сотрудники населенного пункта", _
        "SELECT НаселенныйПункт, ФИО, Отдел, Телефон, Адрес FROM [З05_Сотрудники подробно] " & _
        "WHERE КодПункта = " & FR & "[cboПункт] ORDER BY ФИО"
    MakeQuery "З05_Именинники месяца", _
        "SELECT ФИО, ДатаРождения, ИсполняетсяЛет, Отдел, Телефон FROM [З05_Сотрудники подробно] " & _
        "WHERE Month(ДатаРождения) = Month(Date()) ORDER BY Day(ДатаРождения)"
    MakeQuery "З05_Поздравление", _
        "SELECT ФИО, Обращение, ДатаРождения FROM [З05_Сотрудники подробно] WHERE КодСотрудника = " & FR & "[cboСотрудник]"

    MakeTableForm "З05_Отделы", "З05_Отделы", "Отделы университета"
    MakeTableForm "З05_Образование", "З05_Образование", "Уровни образования"
    MakeTableForm "З05_НаселенныеПункты", "З05_НаселенныеПункты", "Населенные пункты"
    MakeTableForm "З05_Сотрудники", "З05_Сотрудники", "Сотрудники"
    MakeTableForm "З05_Сотрудники подчиненная", "З05_Сотрудники", "Сотрудники", 2, "КодОтдела"
    AddSubform "З05_Отделы", "З05_Сотрудники подчиненная", "КодОтдела", "КодОтдела", "Сотрудники отдела"

    MakeReport "З05_Сотрудники по отделам", "З05_Сотрудники подробно", "Сотрудники университета по отделам", "Отдел", _
        "ФИО;Должность;Образование;Телефон", "3000;2400;2400;1900", "ФИО"
    MakeReport "З05_Сотрудники по образованию", "З05_Сотрудники подробно", "Сотрудники по уровню образования", "Уровень", _
        "ФИО;Отдел;Образование", "3200;3300;3300", "ФИО"
    MakeReport "З05_Список сотрудников отдела", "З05_Сотрудники отдела", "='Сотрудники: ' & [Отдел]", "", _
        "ФИО;Должность;Телефон;Адрес", "3000;2500;1900;2400", "ФИО"
    MakeReport "З05_Сотрудники населенного пункта", "З05_Сотрудники населенного пункта", "='Сотрудники, проживающие: ' & [НаселенныйПункт]", "", _
        "ФИО;Отдел;Телефон;Адрес", "3000;2500;1900;2400", "ФИО"
    MakeGreeting "З05_Поздравление с днем рождения", "З05_Поздравление", "Уважаемый(ая)", _
        "Коллектив университета сердечно поздравляет Вас с днём рождения! Желаем Вам крепкого здоровья, благополучия, профессиональных успехов и исполнения всех желаний!", _
        "Ректорат и профсоюзный комитет"

    CF_Begin "Задание 05. Сотрудники"
    CF_Param "cboОтдел", "Отдел", "SELECT КодОтдела, Отдел FROM [З05_Отделы] ORDER BY Отдел", 2, "0;4500", "1"
    CF_Param "cboПункт", "Населенный пункт", "SELECT КодПункта, НаселенныйПункт FROM [З05_НаселенныеПункты] ORDER BY НаселенныйПункт", 2, "0;4500", "1"
    CF_Param "cboСотрудник", "Сотрудник", "SELECT КодСотрудника, Фамилия & ' ' & Имя & ' ' & Отчество FROM [З05_Сотрудники] ORDER BY Фамилия, Имя", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Сотрудники", "F:З05_Сотрудники"
    CF_Btn 0, "Отделы и их сотрудники", "F:З05_Отделы"
    CF_Btn 0, "Уровни образования", "F:З05_Образование"
    CF_Btn 0, "Населенные пункты", "F:З05_НаселенныеПункты"
    CF_Btn 1, "Сотрудники (все сведения)", "Q:З05_Сотрудники подробно"
    CF_Btn 1, "Общее количество сотрудников", "Q:З05_Общее количество сотрудников"
    CF_Btn 1, "Количество по отделам", "Q:З05_Количество по отделам"
    CF_Btn 1, "Количество: высшее / среднее", "Q:З05_Количество по уровню образования"
    CF_Btn 1, "Сводная: отделы и образование", "Q:З05_Сводная отделы и образование"
    CF_Btn 1, "Сотрудники выбранного отдела", "Q:З05_Сотрудники отдела|cboОтдел"
    CF_Btn 1, "Сотрудники из населенного пункта", "Q:З05_Сотрудники населенного пункта|cboПункт"
    CF_Btn 1, "Именинники текущего месяца", "Q:З05_Именинники месяца"
    CF_Btn 2, "Сотрудники по отделам", "R:З05_Сотрудники по отделам"
    CF_Btn 2, "Сотрудники по образованию", "R:З05_Сотрудники по образованию"
    CF_Btn 2, "Список сотрудников отдела", "R:З05_Список сотрудников отдела|cboОтдел"
    CF_Btn 2, "Сотрудники населенного пункта", "R:З05_Сотрудники населенного пункта|cboПункт"
    CF_Btn 2, "Поздравление с днем рождения", "R:З05_Поздравление с днем рождения|cboСотрудник"
    CF_End "З05_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 06. База данных «Выпускники»
'==============================================================================
Private Sub BuildT06()
    Dim s As String
    Const FR As String = "[Forms]![З06_Управляющая]!"

    MakeTable "З06_Специальности", "КодСпециальности COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Шифр TEXT(10), Специальность TEXT(80) NOT NULL"
    MakeTable "З06_Организации", "КодОрганизации COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Организация TEXT(100) NOT NULL, Отрасль TEXT(60), Город TEXT(40)"
    MakeTable "З06_Выпускники", "КодВыпускника COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), " & _
        "КодСпециальности LONG, ГодВыпуска INTEGER, Телефон TEXT(20), Email TEXT(60), КодОрганизации LONG, Должность TEXT(60), " & _
        "ПоСпециальности BIT, ДатаСведений DATETIME"
    MakeRel "З06_Специальности", "КодСпециальности", "З06_Выпускники", "КодСпециальности"
    MakeRel "З06_Организации", "КодОрганизации", "З06_Выпускники", "КодОрганизации"
    Lookup "З06_Выпускники", "КодСпециальности", "SELECT КодСпециальности, Специальность FROM [З06_Специальности] ORDER BY Специальность"
    Lookup "З06_Выпускники", "КодОрганизации", "SELECT КодОрганизации, Организация FROM [З06_Организации] ORDER BY Организация"

    InsRows "З06_Специальности", "1,'09.03.03','Прикладная информатика'|2,'38.03.01','Экономика'|3,'40.03.01','Юриспруденция'|" & _
        "4,'37.03.01','Психология'|5,'42.03.01','Реклама и связи с общественностью'"
    InsRows "З06_Организации", "1,'ООО «ТехноСофт»','Информационные технологии','Москва'|2,'АО «Финансовый дом»','Финансы','Москва'|" & _
        "3,'Коллегия адвокатов «Право»','Юриспруденция','Москва'|4,'ГБУ «Центр психологической помощи»','Социальная сфера','Москва'|" & _
        "5,'ПАО «Регионбанк»','Финансы','Тула'|6,'ООО «Медиагрупп»','Медиа и реклама','Москва'"
    s = "1,'Абрамов','Николай','Игоревич',1,{y-1},'+7 (916) 120-01-01','abramov@example.ru',1,'Программист',True,{d-30}"
    s = s & "|2,'Баранова','Елена','Сергеевна',1,{y-1},'+7 (916) 120-02-02','baranova@example.ru',1,'Тестировщик',True,{d-45}"
    s = s & "|3,'Виноградов','Антон','Павлович',1,{y-2},'+7 (916) 120-03-03','vinogradov@example.ru',Null,Null,False,{d-60}"
    s = s & "|4,'Голубева','Ирина','Андреевна',2,{y-1},'+7 (916) 120-04-04','golubeva@example.ru',2,'Экономист',True,{d-20}"
    s = s & "|5,'Давыдов','Кирилл','Олегович',2,{y-2},'+7 (916) 120-05-05','davydov@example.ru',5,'Кредитный специалист',True,{d-90}"
    s = s & "|6,'Ефремова','Алёна','Викторовна',2,{y-3},'+7 (916) 120-06-06','efremova@example.ru',6,'Менеджер по продажам',False,{d-120}"
    s = s & "|7,'Жданов','Олег','Николаевич',2,{y-1},'+7 (916) 120-07-07','zhdanov@example.ru',Null,Null,False,{d-15}"
    s = s & "|8,'Зимина','Марина','Юрьевна',3,{y-1},'+7 (916) 120-08-08','zimina@example.ru',3,'Юрист',True,{d-25}"
    s = s & "|9,'Калинин','Артур','Ильич',3,{y-2},'+7 (916) 120-09-09','kalinin@example.ru',Null,Null,False,{d-70}"
    s = s & "|10,'Ларионова','Дарья','Петровна',3,{y-3},'+7 (916) 120-10-10','larionova@example.ru',5,'Юрисконсульт',True,{d-100}"
    s = s & "|11,'Мартынов','Семён','Алексеевич',3,{y-2},'+7 (916) 120-11-11','martynov@example.ru',Null,Null,False,{d-80}"
    s = s & "|12,'Назарова','Кристина','Олеговна',4,{y-1},'+7 (916) 120-12-12','nazarova@example.ru',4,'Психолог',True,{d-10}"
    s = s & "|13,'Олейников','Вадим','Сергеевич',4,{y-2},'+7 (916) 120-13-13','oleynikov@example.ru',Null,Null,False,{d-50}"
    s = s & "|14,'Пестова','Юлия','Игоревна',4,{y-3},'+7 (916) 120-14-14','pestova@example.ru',6,'Специалист по персоналу',False,{d-130}"
    s = s & "|15,'Рябов','Григорий','Андреевич',1,{y-3},'+7 (916) 120-15-15','ryabov@example.ru',5,'Аналитик данных',True,{d-140}"
    s = s & "|16,'Сафонова','Анна','Дмитриевна',5,{y-1},'+7 (916) 120-16-16','safonova@example.ru',6,'PR-менеджер',True,{d-35}"
    s = s & "|17,'Титов','Евгений','Павлович',5,{y-2},'+7 (916) 120-17-17','titov@example.ru',Null,Null,False,{d-65}"
    s = s & "|18,'Устинова','Мария','Андреевна',5,{y-1},'+7 (916) 120-18-18','ustinova@example.ru',Null,Null,False,{d-40}"
    s = s & "|19,'Фомин','Роман','Сергеевич',1,{y-2},'+7 (916) 120-19-19','fomin@example.ru',2,'Разработчик',True,{d-75}"
    s = s & "|20,'Хохлова','Вера','Николаевна',2,{y-3},'+7 (916) 120-20-20','hohlova@example.ru',2,'Бухгалтер',True,{d-110}"
    InsRows "З06_Выпускники", s

    s = "SELECT v.КодВыпускника, v.Фамилия & ' ' & v.Имя & ' ' & v.Отчество AS ФИО, v.КодСпециальности, s.Специальность, v.ГодВыпуска, "
    s = s & "v.Телефон, v.Email, IIf(v.КодОрганизации Is Null,'Безработный','Трудоустроен') AS Статус, "
    s = s & "o.Организация, o.Отрасль, v.Должность, IIf(v.ПоСпециальности,'Да','Нет') AS РаботаПоСпециальности, v.ДатаСведений "
    s = s & "FROM ([З06_Выпускники] AS v LEFT JOIN [З06_Специальности] AS s ON v.КодСпециальности = s.КодСпециальности) "
    s = s & "LEFT JOIN [З06_Организации] AS o ON v.КодОрганизации = o.КодОрганизации"
    MakeQuery "З06_Выпускники подробно", s
    MakeQuery "З06_Количество выпускников", _
        "SELECT Count(*) AS ВсегоВыпускников, Sum(IIf(КодОрганизации Is Null,0,1)) AS Трудоустроено, " & _
        "Sum(IIf(КодОрганизации Is Null,1,0)) AS Безработных, " & _
        "Round(100*Sum(IIf(КодОрганизации Is Null,0,1))/Count(*),1) AS ПроцентТрудоустройства FROM [З06_Выпускники]"
    MakeQuery "З06_Количество по статусу", _
        "SELECT Статус, Count(*) AS Выпускников FROM [З06_Выпускники подробно] GROUP BY Статус"
    s = "SELECT s.Специальность, Count(v.КодВыпускника) AS Выпускников, Sum(IIf(v.КодОрганизации Is Null,0,1)) AS Трудоустроено, "
    s = s & "Sum(IIf(v.ПоСпециальности,1,0)) AS РаботаютПоСпециальности, "
    s = s & "Round(100*Sum(IIf(v.КодОрганизации Is Null,0,1))/Count(v.КодВыпускника),1) AS ПроцентТрудоустройства "
    s = s & "FROM [З06_Специальности] AS s INNER JOIN [З06_Выпускники] AS v ON s.КодСпециальности = v.КодСпециальности "
    s = s & "GROUP BY s.Специальность "
    s = s & "ORDER BY Round(100*Sum(IIf(v.КодОрганизации Is Null,0,1))/Count(v.КодВыпускника),1) DESC, Sum(IIf(v.ПоСпециальности,1,0)) DESC"
    MakeQuery "З06_Востребованность специальностей", s
    MakeQuery "З06_Безработные выпускники", _
        "SELECT ФИО, Специальность, ГодВыпуска, Телефон, Email FROM [З06_Выпускники подробно] WHERE Статус = 'Безработный' ORDER BY Специальность, ФИО"
    MakeQuery "З06_Трудоустроенные выпускники", _
        "SELECT ФИО, Специальность, ГодВыпуска, Организация, Должность, РаботаПоСпециальности FROM [З06_Выпускники подробно] " & _
        "WHERE Статус = 'Трудоустроен' ORDER BY Специальность, ФИО"
    MakeQuery "З06_Трудоустройство по отраслям", _
        "SELECT Отрасль, Count(*) AS Выпускников FROM [З06_Выпускники подробно] WHERE Статус = 'Трудоустроен' " & _
        "GROUP BY Отрасль ORDER BY Count(*) DESC"
    MakeQuery "З06_Сводная специальности и годы", _
        "TRANSFORM Count(КодВыпускника) SELECT Специальность, Count(КодВыпускника) AS Всего " & _
        "FROM [З06_Выпускники подробно] GROUP BY Специальность PIVOT ГодВыпуска"
    MakeQuery "З06_Выпускники специальности", _
        "SELECT Специальность, ФИО, ГодВыпуска, Статус, Организация, Должность, Телефон FROM [З06_Выпускники подробно] " & _
        "WHERE КодСпециальности = " & FR & "[cboСпециальность] ORDER BY ГодВыпуска DESC, ФИО"
    MakeQuery "З06_Выпускники года", _
        "SELECT ГодВыпуска, ФИО, Специальность, Статус, Организация FROM [З06_Выпускники подробно] " & _
        "WHERE ГодВыпуска = " & FR & "[cboГод] ORDER BY Специальность, ФИО"

    MakeTableForm "З06_Специальности", "З06_Специальности", "Специальности"
    MakeTableForm "З06_Организации", "З06_Организации", "Организации (места работы)"
    MakeTableForm "З06_Выпускники", "З06_Выпускники", "Выпускники и их трудоустройство"
    MakeTableForm "З06_Выпускники подчиненная", "З06_Выпускники", "Выпускники", 2, "КодСпециальности"
    AddSubform "З06_Специальности", "З06_Выпускники подчиненная", "КодСпециальности", "КодСпециальности", "Выпускники специальности"

    MakeReport "З06_Выпускники по специальностям", "З06_Выпускники подробно", "Выпускники по специальностям", "Специальность", _
        "ФИО;ГодВыпуска;Статус;Организация;Должность", "2900;900;1500;2800;2200", "-ГодВыпуска;ФИО"
    MakeReport "З06_Трудоустройство выпускников", "З06_Выпускники подробно", "Трудоустроенные и безработные выпускники", "Статус", _
        "ФИО;Специальность;ГодВыпуска;Организация", "2900;3000;900;3000", "Специальность;ФИО"
    MakeReport "З06_Востребованность специальностей", "З06_Востребованность специальностей", "Востребованность специальностей на рынке труда", "", _
        "Специальность;Выпускников;Трудоустроено;РаботаютПоСпециальности;ПроцентТрудоустройства", "3200;1300;1500;2300;2100", "", _
        "Выпускников;Трудоустроено;РаботаютПоСпециальности"
    MakeReport "З06_Выпускники специальности", "З06_Выпускники специальности", "='Выпускники специальности: ' & [Специальность]", "", _
        "ФИО;ГодВыпуска;Статус;Организация;Телефон", "2900;900;1500;2800;1900", "-ГодВыпуска;ФИО"

    CF_Begin "Задание 06. Выпускники"
    CF_Param "cboСпециальность", "Специальность", "SELECT КодСпециальности, Специальность FROM [З06_Специальности] ORDER BY Специальность", 2, "0;4500", "1"
    CF_Param "cboГод", "Год выпуска", "SELECT DISTINCT ГодВыпуска FROM [З06_Выпускники] ORDER BY ГодВыпуска DESC", 1, "", CStr(Year(Date) - 1)
    CF_Columns
    CF_Btn 0, "Выпускники", "F:З06_Выпускники"
    CF_Btn 0, "Специальности и выпускники", "F:З06_Специальности"
    CF_Btn 0, "Организации (места работы)", "F:З06_Организации"
    CF_Btn 1, "Выпускники (все сведения)", "Q:З06_Выпускники подробно"
    CF_Btn 1, "Количество выпускников по институту", "Q:З06_Количество выпускников"
    CF_Btn 1, "Трудоустроенные и безработные", "Q:З06_Количество по статусу"
    CF_Btn 1, "Востребованность специальностей", "Q:З06_Востребованность специальностей"
    CF_Btn 1, "Список безработных", "Q:З06_Безработные выпускники"
    CF_Btn 1, "Список трудоустроенных", "Q:З06_Трудоустроенные выпускники"
    CF_Btn 1, "Трудоустройство по отраслям", "Q:З06_Трудоустройство по отраслям"
    CF_Btn 1, "Сводная: специальности и годы", "Q:З06_Сводная специальности и годы"
    CF_Btn 1, "Выпускники специальности", "Q:З06_Выпускники специальности|cboСпециальность"
    CF_Btn 1, "Выпускники года", "Q:З06_Выпускники года|cboГод"
    CF_Btn 2, "Выпускники по специальностям", "R:З06_Выпускники по специальностям"
    CF_Btn 2, "Трудоустройство выпускников", "R:З06_Трудоустройство выпускников"
    CF_Btn 2, "Востребованность специальностей", "R:З06_Востребованность специальностей"
    CF_Btn 2, "Выпускники специальности", "R:З06_Выпускники специальности|cboСпециальность"
    CF_End "З06_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 07. База данных «Преподаватели»
'==============================================================================
Private Sub BuildT07()
    Dim s As String
    Const FR As String = "[Forms]![З07_Управляющая]!"

    MakeTable "З07_Кафедры", "КодКафедры COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Кафедра TEXT(80) NOT NULL, Заведующий TEXT(60), Телефон TEXT(20)"
    MakeTable "З07_УченыеСтепени", "КодСтепени COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, УченаяСтепень TEXT(40) NOT NULL"
    MakeTable "З07_Преподаватели", "КодПреподавателя COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), Отчество TEXT(30), " & _
        "КодКафедры LONG, Должность TEXT(40), КодСтепени LONG, Телефон TEXT(20)"
    MakeTable "З07_Дисциплины", "КодДисциплины COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Дисциплина TEXT(80) NOT NULL"
    MakeTable "З07_Преподавание", "КодПреподавателя LONG NOT NULL, КодДисциплины LONG NOT NULL, " & _
        "CONSTRAINT PrimaryKey PRIMARY KEY (КодПреподавателя, КодДисциплины)"
    MakeRel "З07_Кафедры", "КодКафедры", "З07_Преподаватели", "КодКафедры"
    MakeRel "З07_УченыеСтепени", "КодСтепени", "З07_Преподаватели", "КодСтепени"
    MakeRel "З07_Преподаватели", "КодПреподавателя", "З07_Преподавание", "КодПреподавателя", True
    MakeRel "З07_Дисциплины", "КодДисциплины", "З07_Преподавание", "КодДисциплины", True
    Lookup "З07_Преподаватели", "КодКафедры", "SELECT КодКафедры, Кафедра FROM [З07_Кафедры] ORDER BY Кафедра"
    Lookup "З07_Преподаватели", "КодСтепени", "SELECT КодСтепени, УченаяСтепень FROM [З07_УченыеСтепени] ORDER BY КодСтепени"
    Lookup "З07_Преподавание", "КодПреподавателя", "SELECT КодПреподавателя, Фамилия & ' ' & Имя & ' ' & Отчество FROM [З07_Преподаватели] ORDER BY Фамилия"
    Lookup "З07_Преподавание", "КодДисциплины", "SELECT КодДисциплины, Дисциплина FROM [З07_Дисциплины] ORDER BY Дисциплина"

    InsRows "З07_Кафедры", "1,'Информатики и информационных технологий','Орлов А.В.','+7 (495) 200-01-01'|" & _
        "2,'Высшей математики','Лебедева Н.С.','+7 (495) 200-02-02'|3,'Экономики и управления','Громов П.И.','+7 (495) 200-03-03'|" & _
        "4,'Иностранных языков','Кравец Е.Л.','+7 (495) 200-04-04'"
    InsRows "З07_УченыеСтепени", "1,'Кандидат наук'|2,'Доктор наук'"
    s = "1,'Орлов','Алексей','Владимирович',1,'Заведующий кафедрой',2,'+7 (916) 300-00-01'"
    s = s & "|2,'Белкин','Сергей','Юрьевич',1,'Доцент',1,'+7 (916) 300-00-02'"
    s = s & "|3,'Ковалёва','Инна','Петровна',1,'Старший преподаватель',Null,'+7 (916) 300-00-03'"
    s = s & "|4,'Лебедева','Наталья','Сергеевна',2,'Заведующий кафедрой',2,'+7 (916) 300-00-04'"
    s = s & "|5,'Мухин','Андрей','Олегович',2,'Доцент',1,'+7 (916) 300-00-05'"
    s = s & "|6,'Нестерова','Ольга','Ивановна',2,'Ассистент',Null,'+7 (916) 300-00-06'"
    s = s & "|7,'Громов','Пётр','Ильич',3,'Заведующий кафедрой',1,'+7 (916) 300-00-07'"
    s = s & "|8,'Сазонова','Елена','Викторовна',3,'Доцент',1,'+7 (916) 300-00-08'"
    s = s & "|9,'Тарасов','Игорь','Николаевич',3,'Старший преподаватель',Null,'+7 (916) 300-00-09'"
    s = s & "|10,'Кравец','Екатерина','Леонидовна',4,'Заведующий кафедрой',1,'+7 (916) 300-00-10'"
    s = s & "|11,'Уварова','Анна','Сергеевна',4,'Старший преподаватель',Null,'+7 (916) 300-00-11'"
    s = s & "|12,'Фёдоров','Максим','Андреевич',4,'Ассистент',Null,'+7 (916) 300-00-12'"
    InsRows "З07_Преподаватели", s
    InsRows "З07_Дисциплины", "1,'Информатика'|2,'Базы данных'|3,'Программирование'|4,'Математический анализ'|5,'Линейная алгебра'|" & _
        "6,'Теория вероятностей'|7,'Экономическая теория'|8,'Менеджмент'|9,'Английский язык'|10,'Немецкий язык'|" & _
        "11,'Криптография'|12,'Эконометрика'"
    InsRows "З07_Преподавание", "1,2|1,3|2,1|2,2|3,1|3,3|4,4|4,6|5,4|5,5|5,6|6,5|7,7|7,8|8,7|9,8|10,9|11,9|11,10|12,10"

    s = "SELECT p.КодПреподавателя, p.Фамилия & ' ' & p.Имя & ' ' & p.Отчество AS ФИО, p.КодКафедры, k.Кафедра, p.Должность, "
    s = s & "p.КодСтепени, Nz(st.УченаяСтепень,'Нет степени') AS Степень, p.Телефон "
    s = s & "FROM ([З07_Преподаватели] AS p LEFT JOIN [З07_Кафедры] AS k ON p.КодКафедры = k.КодКафедры) "
    s = s & "LEFT JOIN [З07_УченыеСтепени] AS st ON p.КодСтепени = st.КодСтепени"
    MakeQuery "З07_Преподаватели подробно", s
    MakeQuery "З07_Общее количество преподавателей", _
        "SELECT Count(*) AS ВсегоПреподавателей, Sum(IIf(КодСтепени Is Null,0,1)) AS СУченойСтепенью, " & _
        "Sum(IIf(КодСтепени Is Null,1,0)) AS БезСтепени FROM [З07_Преподаватели]"
    MakeQuery "З07_Количество по дисциплинам", _
        "SELECT d.Дисциплина, Count(pr.КодПреподавателя) AS Преподавателей FROM [З07_Дисциплины] AS d " & _
        "LEFT JOIN [З07_Преподавание] AS pr ON d.КодДисциплины = pr.КодДисциплины GROUP BY d.Дисциплина ORDER BY d.Дисциплина"
    MakeQuery "З07_Количество по ученым степеням", _
        "SELECT Степень, Count(*) AS Преподавателей FROM [З07_Преподаватели подробно] GROUP BY Степень"
    MakeQuery "З07_Дисциплины без преподавателей", _
        "SELECT d.Дисциплина FROM [З07_Дисциплины] AS d LEFT JOIN [З07_Преподавание] AS pr ON d.КодДисциплины = pr.КодДисциплины " & _
        "WHERE pr.КодДисциплины Is Null ORDER BY d.Дисциплина"
    s = "SELECT d.КодДисциплины, d.Дисциплина, q.ФИО, q.Кафедра, q.Должность, q.Степень "
    s = s & "FROM ([З07_Преподавание] AS pr INNER JOIN [З07_Преподаватели подробно] AS q ON pr.КодПреподавателя = q.КодПреподавателя) "
    s = s & "INNER JOIN [З07_Дисциплины] AS d ON pr.КодДисциплины = d.КодДисциплины"
    MakeQuery "З07_Преподаватели и дисциплины", s
    MakeQuery "З07_Преподаватели дисциплины", _
        "SELECT Дисциплина, ФИО, Кафедра, Должность, Степень FROM [З07_Преподаватели и дисциплины] " & _
        "WHERE КодДисциплины = " & FR & "[cboДисциплина] ORDER BY ФИО"
    MakeQuery "З07_Преподаватели с учеными степенями", _
        "SELECT ФИО, Степень, Кафедра, Должность FROM [З07_Преподаватели подробно] WHERE КодСтепени Is Not Null ORDER BY Степень DESC, ФИО"
    MakeQuery "З07_Кафедры и количество преподавателей", _
        "SELECT k.Кафедра, k.Заведующий, Count(p.КодПреподавателя) AS Преподавателей FROM [З07_Кафедры] AS k " & _
        "LEFT JOIN [З07_Преподаватели] AS p ON k.КодКафедры = p.КодКафедры GROUP BY k.Кафедра, k.Заведующий ORDER BY k.Кафедра"
    MakeQuery "З07_Преподаватели кафедры", _
        "SELECT Кафедра, ФИО, Должность, Степень, Телефон FROM [З07_Преподаватели подробно] " & _
        "WHERE КодКафедры = " & FR & "[cboКафедра] ORDER BY ФИО"
    MakeQuery "З07_Сводная кафедры и степени", _
        "TRANSFORM Count(КодПреподавателя) SELECT Кафедра, Count(КодПреподавателя) AS Всего " & _
        "FROM [З07_Преподаватели подробно] GROUP BY Кафедра PIVOT Степень"

    MakeTableForm "З07_Кафедры", "З07_Кафедры", "Кафедры"
    MakeTableForm "З07_УченыеСтепени", "З07_УченыеСтепени", "Ученые степени"
    MakeTableForm "З07_Преподаватели", "З07_Преподаватели", "Преподаватели"
    MakeTableForm "З07_Дисциплины", "З07_Дисциплины", "Учебные дисциплины"
    MakeTableForm "З07_Преподаватели подчиненная", "З07_Преподаватели", "Преподаватели", 2, "КодКафедры"
    MakeTableForm "З07_Дисциплины преподавателя", "З07_Преподавание", "Дисциплины", 2, "КодПреподавателя"
    MakeTableForm "З07_Преподаватели дисциплины подчиненная", "З07_Преподавание", "Преподаватели", 2, "КодДисциплины"
    AddSubform "З07_Кафедры", "З07_Преподаватели подчиненная", "КодКафедры", "КодКафедры", "Преподаватели кафедры"
    AddSubform "З07_Преподаватели", "З07_Дисциплины преподавателя", "КодПреподавателя", "КодПреподавателя", "Преподаваемые дисциплины"
    AddSubform "З07_Дисциплины", "З07_Преподаватели дисциплины подчиненная", "КодДисциплины", "КодДисциплины", "Преподаватели дисциплины"

    MakeReport "З07_Преподаватели кафедр", "З07_Преподаватели подробно", "Кафедры: список и количество преподавателей", "Кафедра", _
        "ФИО;Должность;Степень;Телефон", "3300;2600;1900;1900", "ФИО"
    MakeReport "З07_Преподаватели по дисциплинам", "З07_Преподаватели и дисциплины", "Преподаватели по учебным дисциплинам", "Дисциплина", _
        "ФИО;Кафедра;Степень", "3300;4000;1900", "ФИО"
    MakeReport "З07_Преподаватели с учеными степенями", "З07_Преподаватели с учеными степенями", "Преподаватели с учеными степенями", "Степень", _
        "ФИО;Кафедра;Должность", "3300;4000;2300", "ФИО"
    MakeReport "З07_Список преподавателей кафедры", "З07_Преподаватели кафедры", "='Кафедра ' & [Кафедра]", "", _
        "ФИО;Должность;Степень;Телефон", "3300;2600;1900;1900", "ФИО"

    CF_Begin "Задание 07. Преподаватели"
    CF_Param "cboКафедра", "Кафедра", "SELECT КодКафедры, Кафедра FROM [З07_Кафедры] ORDER BY Кафедра", 2, "0;4500", "1"
    CF_Param "cboДисциплина", "Дисциплина", "SELECT КодДисциплины, Дисциплина FROM [З07_Дисциплины] ORDER BY Дисциплина", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Преподаватели и их дисциплины", "F:З07_Преподаватели"
    CF_Btn 0, "Кафедры и их преподаватели", "F:З07_Кафедры"
    CF_Btn 0, "Дисциплины и их преподаватели", "F:З07_Дисциплины"
    CF_Btn 0, "Ученые степени", "F:З07_УченыеСтепени"
    CF_Btn 1, "Преподаватели (все сведения)", "Q:З07_Преподаватели подробно"
    CF_Btn 1, "Общее количество преподавателей", "Q:З07_Общее количество преподавателей"
    CF_Btn 1, "Количество по дисциплинам", "Q:З07_Количество по дисциплинам"
    CF_Btn 1, "Количество по ученым степеням", "Q:З07_Количество по ученым степеням"
    CF_Btn 1, "Дисциплины без преподавателей", "Q:З07_Дисциплины без преподавателей"
    CF_Btn 1, "Преподаватели дисциплины", "Q:З07_Преподаватели дисциплины|cboДисциплина"
    CF_Btn 1, "Преподаватели с учеными степенями", "Q:З07_Преподаватели с учеными степенями"
    CF_Btn 1, "Кафедры: количество преподавателей", "Q:З07_Кафедры и количество преподавателей"
    CF_Btn 1, "Преподаватели кафедры", "Q:З07_Преподаватели кафедры|cboКафедра"
    CF_Btn 1, "Сводная: кафедры и степени", "Q:З07_Сводная кафедры и степени"
    CF_Btn 2, "Кафедры и преподаватели", "R:З07_Преподаватели кафедр"
    CF_Btn 2, "Преподаватели по дисциплинам", "R:З07_Преподаватели по дисциплинам"
    CF_Btn 2, "Преподаватели с учеными степенями", "R:З07_Преподаватели с учеными степенями"
    CF_Btn 2, "Список преподавателей кафедры", "R:З07_Список преподавателей кафедры|cboКафедра"
    CF_End "З07_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЕ 08. База данных «Библиотека»
'==============================================================================
Private Sub BuildT08()
    Dim s As String
    Const FR As String = "[Forms]![З08_Управляющая]!"

    MakeTable "З08_Рубрики", "КодРубрики COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Рубрика TEXT(60) NOT NULL"
    MakeTable "З08_Книги", "КодКниги COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, ИнвНомер TEXT(20), Автор TEXT(100), Название TEXT(200) NOT NULL, " & _
        "КодРубрики LONG, Издательство TEXT(60), ГодИздания INTEGER, Страниц INTEGER, Цена CURRENCY"
    MakeTable "З08_Читатели", "КодЧитателя COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, НомерБилета TEXT(10), Фамилия TEXT(30) NOT NULL, Имя TEXT(30), " & _
        "Отчество TEXT(30), Адрес TEXT(100), Телефон TEXT(20)"
    MakeTable "З08_Выдачи", "КодВыдачи COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, КодКниги LONG NOT NULL, КодЧитателя LONG NOT NULL, " & _
        "ДатаВыдачи DATETIME, СрокВозврата DATETIME, ДатаВозврата DATETIME"
    MakeRel "З08_Рубрики", "КодРубрики", "З08_Книги", "КодРубрики"
    MakeRel "З08_Книги", "КодКниги", "З08_Выдачи", "КодКниги"
    MakeRel "З08_Читатели", "КодЧитателя", "З08_Выдачи", "КодЧитателя"
    Lookup "З08_Книги", "КодРубрики", "SELECT КодРубрики, Рубрика FROM [З08_Рубрики] ORDER BY Рубрика"
    Lookup "З08_Выдачи", "КодКниги", "SELECT КодКниги, ИнвНомер & ' ' & Автор & ' ' & Название FROM [З08_Книги] ORDER BY ИнвНомер", 2, "0;5000"
    Lookup "З08_Выдачи", "КодЧитателя", "SELECT КодЧитателя, НомерБилета & ' ' & Фамилия & ' ' & Имя FROM [З08_Читатели] ORDER BY Фамилия"

    InsRows "З08_Рубрики", "1,'Экономика'|2,'Юриспруденция'|3,'Психология'|4,'Информатика'|5,'Математика'|6,'Художественная литература'"
    s = "1,'И-0001','Мэнкью Н.','Принципы экономикс',1,'Питер',2018,672,1250"
    s = s & "|2,'И-0002','Мэнкью Н.','Принципы экономикс',1,'Питер',2018,672,1250"
    s = s & "|3,'И-0003','Липсиц И.В.','Экономика',1,'Юрайт',2020,452,890"
    s = s & "|4,'И-0004','Марченко М.Н.','Теория государства и права',2,'Проспект',2019,640,980"
    s = s & "|5,'И-0005','Сергеев А.П.','Гражданское право',2,'Проспект',2021,880,1450"
    s = s & "|6,'И-0006','Майерс Д.','Социальная психология',3,'Питер',2019,800,1350"
    s = s & "|7,'И-0007','Немов Р.С.','Общая психология',3,'Юрайт',2020,304,720"
    s = s & "|8,'И-0008','Кнут Д.','Искусство программирования. Том 1',4,'Вильямс',2017,720,2100"
    s = s & "|9,'И-0009','Дейт К.','Введение в системы баз данных',4,'Вильямс',2016,1328,2400"
    s = s & "|10,'И-0010','Дейт К.','Введение в системы баз данных',4,'Вильямс',2016,1328,2400"
    s = s & "|11,'И-0011','Кузин А.В.','Базы данных',4,'Академия',2021,320,650"
    s = s & "|12,'И-0012','Фихтенгольц Г.М.','Курс дифференциального и интегрального исчисления',5,'Лань',2019,608,1100"
    s = s & "|13,'И-0013','Демидович Б.П.','Сборник задач по математическому анализу',5,'Лань',2020,624,990"
    s = s & "|14,'И-0014','Толстой Л.Н.','Война и мир',6,'Эксмо',2018,1300,800"
    s = s & "|15,'И-0015','Булгаков М.А.','Мастер и Маргарита',6,'АСТ',2021,480,450"
    s = s & "|16,'И-0016','Булгаков М.А.','Собачье сердце',6,'АСТ',2020,160,300"
    s = s & "|17,'И-0017','Пушкин А.С.','Евгений Онегин',6,'Эксмо',2019,288,350"
    s = s & "|18,'И-0018','Кузин А.В.','Базы данных',4,'Академия',2021,320,650"
    InsRows "З08_Книги", s
    s = "1,'Ч-101','Алексеев','Павел','Игоревич','ул. Строителей, 4-10','+7 (916) 400-10-01'"
    s = s & "|2,'Ч-102','Богданова','Светлана','Юрьевна','ул. Мира, 12-7','+7 (916) 400-10-02'"
    s = s & "|3,'Ч-103','Веселов','Артём','Михайлович','ул. Лесная, 3-15','+7 (916) 400-10-03'"
    s = s & "|4,'Ч-104','Горбунова','Ирина','Васильевна','пр. Вернадского, 41-9','+7 (916) 400-10-04'"
    s = s & "|5,'Ч-105','Дорофеев','Никита','Андреевич','ул. Садовая, 7-2','+7 (916) 400-10-05'"
    s = s & "|6,'Ч-106','Ершов','Виктор','Петрович','ул. Гагарина, 19-44','+7 (916) 400-10-06'"
    s = s & "|7,'Ч-107','Журавлёва','Анна','Константиновна','ул. Школьная, 5-5','+7 (916) 400-10-07'"
    s = s & "|8,'Ч-108','Золотарёв','Семён','Ильич','ул. Полевая, 11-3','+7 (916) 400-10-08'"
    InsRows "З08_Читатели", s
    s = "1,1,1,{d-60},{d-46},{d-50}|2,8,1,{d-20},{d-6},Null|3,6,2,{d-10},{d+4},Null|4,14,3,{d-40},{d-26},{d-27}"
    s = s & "|5,15,3,{d-5},{d+9},Null|6,9,4,{d-15},{d-1},Null|7,4,5,{d-30},{d-16},{d-18}|8,5,5,{d-3},{d+11},Null"
    s = s & "|9,12,6,{d-25},{d-11},{d-12}|10,12,7,{d-8},{d+6},Null|11,17,8,{d-2},{d+12},Null|12,11,2,{d-12},{d+2},Null"
    s = s & "|13,3,6,{d-45},{d-31},{d-33}|14,7,4,{d-7},{d+7},Null"
    InsRows "З08_Выдачи", s

    s = "SELECT v.КодВыдачи, v.КодКниги, v.КодЧитателя, r.Фамилия & ' ' & r.Имя & ' ' & r.Отчество AS Читатель, r.НомерБилета, "
    s = s & "r.Телефон, v.ДатаВыдачи, v.СрокВозврата, IIf(v.СрокВозврата < Date(),'Просрочена','') AS Просрочка "
    s = s & "FROM [З08_Выдачи] AS v INNER JOIN [З08_Читатели] AS r ON v.КодЧитателя = r.КодЧитателя WHERE v.ДатаВозврата Is Null"
    MakeQuery "З08_Книги на руках", s
    s = "SELECT k.КодКниги, k.ИнвНомер, k.Автор, k.Название, k.КодРубрики, ru.Рубрика, k.Издательство, k.ГодИздания, k.Страниц, k.Цена, "
    s = s & "IIf(h.КодВыдачи Is Null,'В наличии','На руках') AS Наличие, h.КодЧитателя, h.Читатель, h.ДатаВыдачи, h.СрокВозврата, h.Просрочка "
    s = s & "FROM ([З08_Книги] AS k LEFT JOIN [З08_Рубрики] AS ru ON k.КодРубрики = ru.КодРубрики) "
    s = s & "LEFT JOIN [З08_Книги на руках] AS h ON k.КодКниги = h.КодКниги"
    MakeQuery "З08_Каталог книг", s
    MakeQuery "З08_Количество книг всего", _
        "SELECT Count(*) AS ВсегоКниг, Sum(IIf(Наличие='В наличии',1,0)) AS ВНаличии, Sum(IIf(Наличие='На руках',1,0)) AS НаРуках " & _
        "FROM [З08_Каталог книг]"
    MakeQuery "З08_Количество книг по рубрикам", _
        "SELECT ru.Рубрика, Count(k.КодКниги) AS Книг FROM [З08_Рубрики] AS ru LEFT JOIN [З08_Книги] AS k ON ru.КодРубрики = k.КодРубрики " & _
        "GROUP BY ru.Рубрика ORDER BY ru.Рубрика"
    MakeQuery "З08_Поиск книги", _
        "PARAMETERS " & FR & "[txtНазвание] Text ( 255 ); " & _
        "SELECT ИнвНомер, Автор, Название, ГодИздания, Наличие, Читатель, СрокВозврата FROM [З08_Каталог книг] " & _
        "WHERE Название Like '*' & " & FR & "[txtНазвание] & '*' ORDER BY Название, ИнвНомер"
    MakeQuery "З08_Книги рубрики", _
        "SELECT Рубрика, ИнвНомер, Автор, Название, Издательство, ГодИздания, Наличие FROM [З08_Каталог книг] " & _
        "WHERE КодРубрики = " & FR & "[cboРубрика] ORDER BY Автор, Название"
    MakeQuery "З08_Книги автора", _
        "SELECT Автор, ИнвНомер, Название, Издательство, ГодИздания, Наличие FROM [З08_Каталог книг] " & _
        "WHERE Автор = " & FR & "[cboАвтор] ORDER BY Название"
    MakeQuery "З08_Книги издательства", _
        "SELECT Издательство, ИнвНомер, Автор, Название, ГодИздания, Наличие FROM [З08_Каталог книг] " & _
        "WHERE Издательство = " & FR & "[cboИздательство] ORDER BY Автор, Название"
    MakeQuery "З08_Книги года издания", _
        "SELECT ГодИздания, ИнвНомер, Автор, Название, Издательство, Наличие FROM [З08_Каталог книг] " & _
        "WHERE ГодИздания = " & FR & "[cboГод] ORDER BY Автор, Название"
    MakeQuery "З08_Выданные книги", _
        "SELECT Читатель, КодЧитателя, ИнвНомер, Автор, Название, ДатаВыдачи, СрокВозврата, Просрочка FROM [З08_Каталог книг] " & _
        "WHERE Наличие = 'На руках'"
    MakeQuery "З08_Книги читателя", _
        "SELECT Читатель, ИнвНомер, Автор, Название, ДатаВыдачи, СрокВозврата, Просрочка FROM [З08_Выданные книги] " & _
        "WHERE КодЧитателя = " & FR & "[cboЧитатель] ORDER BY ДатаВыдачи"
    MakeQuery "З08_Просроченные книги", _
        "SELECT Читатель, ИнвНомер, Автор, Название, ДатаВыдачи, СрокВозврата FROM [З08_Выданные книги] " & _
        "WHERE Просрочка = 'Просрочена' ORDER BY СрокВозврата"

    MakeTableForm "З08_Рубрики", "З08_Рубрики", "Рубрики"
    MakeTableForm "З08_Книги", "З08_Книги", "Книги (библиотечный фонд)"
    MakeTableForm "З08_Читатели", "З08_Читатели", "Читатели"
    MakeTableForm "З08_Выдачи", "З08_Выдачи", "Выдача и возврат книг"
    MakeTableForm "З08_Книги подчиненная", "З08_Книги", "Книги", 2, "КодРубрики"
    MakeTableForm "З08_Выдачи подчиненная", "З08_Выдачи", "Выдачи", 2, "КодЧитателя"
    AddSubform "З08_Рубрики", "З08_Книги подчиненная", "КодРубрики", "КодРубрики", "Книги рубрики"
    AddSubform "З08_Читатели", "З08_Выдачи подчиненная", "КодЧитателя", "КодЧитателя", "Формуляр читателя (выдачи книг)"

    MakeReport "З08_Книги по рубрикам", "З08_Каталог книг", "Библиотечный фонд по рубрикам", "Рубрика", _
        "ИнвНомер;Автор;Название;ГодИздания;Наличие", "950;2000;4200;900;1300", "Автор;Название"
    MakeReport "З08_Книги у читателей", "З08_Выданные книги", "Книги на руках у читателей", "Читатель", _
        "ИнвНомер;Автор;Название;СрокВозврата;Просрочка", "950;2000;4000;1300;1300", "СрокВозврата"
    MakeReport "З08_Формуляр читателя", "З08_Книги читателя", "='Книги, числящиеся за читателем: ' & [Читатель]", "", _
        "ИнвНомер;Автор;Название;ДатаВыдачи;СрокВозврата;Просрочка", "950;1900;3200;1200;1300;1200", "ДатаВыдачи"
    MakeReport "З08_Книги рубрики", "З08_Книги рубрики", "='Книги рубрики: ' & [Рубрика]", "", _
        "ИнвНомер;Автор;Название;ГодИздания;Наличие", "950;2000;4200;900;1300", "Автор;Название"

    CF_Begin "Задание 08. Библиотека"
    CF_Param "txtНазвание", "Название книги (часть)", "", , , , "T"
    CF_Param "cboРубрика", "Рубрика", "SELECT КодРубрики, Рубрика FROM [З08_Рубрики] ORDER BY Рубрика", 2, "0;4500", "1"
    CF_Param "cboАвтор", "Автор", "SELECT DISTINCT Автор FROM [З08_Книги] ORDER BY Автор", 1, ""
    CF_Param "cboИздательство", "Издательство", "SELECT DISTINCT Издательство FROM [З08_Книги] ORDER BY Издательство", 1, ""
    CF_Param "cboГод", "Год издания", "SELECT DISTINCT ГодИздания FROM [З08_Книги] ORDER BY ГодИздания DESC", 1, ""
    CF_Param "cboЧитатель", "Читатель", "SELECT КодЧитателя, НомерБилета & ' ' & Фамилия & ' ' & Имя & ' ' & Отчество FROM [З08_Читатели] ORDER BY Фамилия", 2, "0;4500", "1"
    CF_Columns
    CF_Btn 0, "Книги", "F:З08_Книги"
    CF_Btn 0, "Рубрики и книги", "F:З08_Рубрики"
    CF_Btn 0, "Читатели и их формуляры", "F:З08_Читатели"
    CF_Btn 0, "Выдача и возврат книг", "F:З08_Выдачи"
    CF_Btn 1, "Каталог книг (наличие)", "Q:З08_Каталог книг"
    CF_Btn 1, "Количество книг всего", "Q:З08_Количество книг всего"
    CF_Btn 1, "Количество книг по рубрикам", "Q:З08_Количество книг по рубрикам"
    CF_Btn 1, "Есть ли книга? У кого на руках?", "Q:З08_Поиск книги|txtНазвание"
    CF_Btn 1, "Книги рубрики", "Q:З08_Книги рубрики|cboРубрика"
    CF_Btn 1, "Книги автора", "Q:З08_Книги автора|cboАвтор"
    CF_Btn 1, "Книги издательства", "Q:З08_Книги издательства|cboИздательство"
    CF_Btn 1, "Книги года издания", "Q:З08_Книги года издания|cboГод"
    CF_Btn 1, "Книги, числящиеся за читателем", "Q:З08_Книги читателя|cboЧитатель"
    CF_Btn 1, "Просроченные книги", "Q:З08_Просроченные книги"
    CF_Btn 2, "Библиотечный фонд по рубрикам", "R:З08_Книги по рубрикам"
    CF_Btn 2, "Книги на руках у читателей", "R:З08_Книги у читателей"
    CF_Btn 2, "Формуляр читателя", "R:З08_Формуляр читателя|cboЧитатель"
    CF_Btn 2, "Книги рубрики", "R:З08_Книги рубрики|cboРубрика"
    CF_End "З08_Управляющая"
End Sub

'==============================================================================
' ЗАДАНИЯ 09 и 10. «Издательский отдел» и «Информационный отдел»
' Структура одинаковая; в задании 10 учитываются электронные материалы
' (формат, объем, ссылка вместо страниц и тиража).
'==============================================================================
Private Sub BuildPub(ByVal P As String, ByVal electronic As Boolean)
    Dim s As String, FR As String, extra As String, ttl As String, kind As String
    Dim ed As Variant, a As Variant, au As Variant, i As Integer, j As Integer
    Dim types As Variant, fmt As Variant, t As Integer

    FR = "[Forms]![" & P & "Управляющая]!"
    If electronic Then
        ttl = "Задание 10. Информационный отдел"
        kind = "электронных изданий"
        types = Array("Электронный учебник", "Электронное пособие", "Видеолекции", "Практикум", "Тест")
        fmt = Array("PDF", "PDF", "MP4", "HTML", "HTML")
        extra = "Формат TEXT(20), ОбъемМБ DOUBLE, Ссылка TEXT(255)"
    Else
        ttl = "Задание 09. Издательский отдел"
        kind = "изданий"
        types = Array("Учебник", "Учебное пособие", "Конспект лекций", "Практикум", "Методические указания")
        extra = "Страниц INTEGER, Тираж INTEGER"
    End If

    MakeTable P & "Кафедры", "КодКафедры COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Кафедра TEXT(80) NOT NULL, Заведующий TEXT(60)"
    MakeTable P & "Преподаватели", "КодПреподавателя COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Фамилия TEXT(30) NOT NULL, Имя TEXT(30), " & _
        "Отчество TEXT(30), КодКафедры LONG, Должность TEXT(40)"
    MakeTable P & "ТипыИзданий", "КодТипа COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, ТипИздания TEXT(60) NOT NULL"
    MakeTable P & "Дисциплины", "КодДисциплины COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Дисциплина TEXT(80) NOT NULL"
    MakeTable P & "Издания", "КодИздания COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Название TEXT(200) NOT NULL, КодТипа LONG, " & _
        "КодДисциплины LONG, КодКафедры LONG, ГодИздания INTEGER, " & extra
    MakeTable P & "Авторство", "КодИздания LONG NOT NULL, КодПреподавателя LONG NOT NULL, " & _
        "CONSTRAINT PrimaryKey PRIMARY KEY (КодИздания, КодПреподавателя)"
    MakeRel P & "Кафедры", "КодКафедры", P & "Преподаватели", "КодКафедры"
    MakeRel P & "Кафедры", "КодКафедры", P & "Издания", "КодКафедры"
    MakeRel P & "ТипыИзданий", "КодТипа", P & "Издания", "КодТипа"
    MakeRel P & "Дисциплины", "КодДисциплины", P & "Издания", "КодДисциплины"
    MakeRel P & "Издания", "КодИздания", P & "Авторство", "КодИздания", True
    MakeRel P & "Преподаватели", "КодПреподавателя", P & "Авторство", "КодПреподавателя", True
    Lookup P & "Преподаватели", "КодКафедры", "SELECT КодКафедры, Кафедра FROM [" & P & "Кафедры] ORDER BY Кафедра"
    Lookup P & "Издания", "КодТипа", "SELECT КодТипа, ТипИздания FROM [" & P & "ТипыИзданий] ORDER BY ТипИздания"
    Lookup P & "Издания", "КодДисциплины", "SELECT КодДисциплины, Дисциплина FROM [" & P & "Дисциплины] ORDER BY Дисциплина"
    Lookup P & "Издания", "КодКафедры", "SELECT КодКафедры, Кафедра FROM [" & P & "Кафедры] ORDER BY Кафедра"
    Lookup P & "Авторство", "КодИздания", "SELECT КодИздания, Название FROM [" & P & "Издания] ORDER BY Название", 2, "0;5000"
    Lookup P & "Авторство", "КодПреподавателя", "SELECT КодПреподавателя, Фамилия & ' ' & Имя & ' ' & Отчество FROM [" & P & "Преподаватели] ORDER BY Фамилия"

    InsRows P & "Кафедры", "1,'Информатики','Орлов А.В.'|2,'Высшей математики','Лебедева Н.С.'|3,'Экономики','Громов П.И.'|4,'Социологии','Широкова Г.П.'"
    s = "1,'Орлов','Алексей','Владимирович',1,'Заведующий кафедрой'|2,'Белкин','Сергей','Юрьевич',1,'Доцент'"
    s = s & "|3,'Ковалёва','Инна','Петровна',1,'Старший преподаватель'|4,'Лебедева','Наталья','Сергеевна',2,'Профессор'"
    s = s & "|5,'Мухин','Андрей','Олегович',2,'Доцент'|6,'Громов','Пётр','Ильич',3,'Профессор'"
    s = s & "|7,'Сазонова','Елена','Викторовна',3,'Доцент'|8,'Тарасов','Игорь','Николаевич',3,'Старший преподаватель'"
    s = s & "|9,'Широкова','Галина','Павловна',4,'Доцент'|10,'Юдин','Роман','Васильевич',4,'Старший преподаватель'"
    InsRows P & "Преподаватели", s
    For i = 0 To 4
        InsRows P & "ТипыИзданий", (i + 1) & ",'" & types(i) & "'"
    Next i
    InsRows P & "Дисциплины", "1,'Информатика'|2,'Базы данных'|3,'Математика'|4,'Статистика'|5,'Экономика'|6,'Социология'"

    ' название; тип; дисциплина; кафедра; сколько лет назад; авторы
    ed = Array("Основы информатики;1;1;1;0;1,3", "Проектирование баз данных в MS Access;4;2;1;1;2", _
               "Базы данных;3;2;1;3;2,3", "Информационные технологии в экономике;2;1;1;6;1,7", _
               "Высшая математика. Часть 1;1;3;2;2;4", "Высшая математика. Часть 2;1;3;2;1;4,5", _
               "Сборник задач по математике;4;3;2;0;5", "Теория вероятностей и статистика;2;4;2;4;4", _
               "Прикладная статистика;4;4;3;2;7,8", "Экономическая теория;1;5;3;0;6", _
               "Микроэкономика;3;5;3;5;6,7", "Макроэкономика;2;5;3;1;8", _
               "Курсовая работа по экономике;5;5;3;0;7", "Социология;1;6;4;3;9", _
               "Социологические исследования;4;6;4;1;9,10", "История социологии;3;6;4;7;10", _
               "Методы социологии;5;6;4;2;10", "Программирование на VBA;4;1;1;2;3")
    For i = 0 To UBound(ed)
        a = Split(ed(i), ";")
        t = CInt(a(1))
        s = (i + 1) & ",'" & a(0) & "'," & a(1) & "," & a(2) & "," & a(3) & "," & (Year(Date) - CInt(a(4))) & ","
        If electronic Then
            s = s & "'" & fmt(t - 1) & "'," & Replace(CStr(Round(2 + (i + 1) * 1.7, 1)), ",", ".") & ",'https://edu.example.ru/e/" & (i + 1) & "'"
        Else
            s = s & (60 + ((i + 1) * 37) Mod 300) & "," & (100 * (1 + (i + 1) Mod 5))
        End If
        InsRows P & "Издания", s
        au = Split(a(5), ",")
        For j = 0 To UBound(au)
            InsRows P & "Авторство", (i + 1) & "," & au(j)
        Next j
    Next i

    ' --- запросы ---
    s = "SELECT i.КодИздания, i.Название, AuthorsList('" & P & "', i.КодИздания) AS Авторы, i.КодТипа, t.ТипИздания, "
    s = s & "i.КодДисциплины, d.Дисциплина, i.КодКафедры, k.Кафедра, i.ГодИздания, "
    If electronic Then
        s = s & "i.Формат, i.ОбъемМБ, i.Ссылка "
    Else
        s = s & "i.Страниц, i.Тираж "
    End If
    s = s & "FROM (([" & P & "Издания] AS i LEFT JOIN [" & P & "ТипыИзданий] AS t ON i.КодТипа = t.КодТипа) "
    s = s & "LEFT JOIN [" & P & "Дисциплины] AS d ON i.КодДисциплины = d.КодДисциплины) "
    s = s & "LEFT JOIN [" & P & "Кафедры] AS k ON i.КодКафедры = k.КодКафедры"
    MakeQuery P & "Издания подробно", s
    MakeQuery P & "Издания по кафедрам и дисциплинам", _
        "SELECT Кафедра, Дисциплина, Название, Авторы, ТипИздания, ГодИздания FROM [" & P & "Издания подробно] " & _
        "ORDER BY Кафедра, Дисциплина, ГодИздания DESC"
    MakeQuery P & "Количество изданий кафедр за год", _
        "SELECT Кафедра, Count(*) AS Изданий FROM [" & P & "Издания подробно] WHERE ГодИздания = " & FR & "[cboГод] " & _
        "GROUP BY Кафедра ORDER BY Кафедра"
    MakeQuery P & "Количество изданий университета за год", _
        "SELECT Count(*) AS ИзданийЗаГод FROM [" & P & "Издания] WHERE ГодИздания = " & FR & "[cboГод]"
    MakeQuery P & "Издания за последние 5 лет", _
        "SELECT Кафедра, ГодИздания, Название, Авторы, ТипИздания, Дисциплина FROM [" & P & "Издания подробно] " & _
        "WHERE ГодИздания >= Year(Date())-4 ORDER BY Кафедра, ГодИздания DESC"
    MakeQuery P & "Количество изданий кафедр за 5 лет", _
        "SELECT Кафедра, Count(*) AS Изданий FROM [" & P & "Издания подробно] WHERE ГодИздания >= Year(Date())-4 " & _
        "GROUP BY Кафедра ORDER BY Кафедра"
    MakeQuery P & "Количество изданий университета за 5 лет", _
        "SELECT Count(*) AS ИзданийЗа5Лет FROM [" & P & "Издания] WHERE ГодИздания >= Year(Date())-4"
    MakeQuery P & "Сводная кафедры по годам", _
        "TRANSFORM Count(КодИздания) SELECT Кафедра, Count(КодИздания) AS Всего FROM [" & P & "Издания подробно] " & _
        "WHERE ГодИздания >= Year(Date())-4 GROUP BY Кафедра PIVOT ГодИздания"
    MakeQuery P & "Издания по типам за период", _
        "SELECT ТипИздания, Название, Авторы, Дисциплина, Кафедра, ГодИздания FROM [" & P & "Издания подробно] " & _
        "WHERE ГодИздания Between " & FR & "[cboГодС] And " & FR & "[cboГодПо] ORDER BY ТипИздания, ГодИздания DESC"
    MakeQuery P & "Количество по типам за период", _
        "SELECT ТипИздания, Count(*) AS Изданий FROM [" & P & "Издания подробно] " & _
        "WHERE ГодИздания Between " & FR & "[cboГодС] And " & FR & "[cboГодПо] GROUP BY ТипИздания ORDER BY ТипИздания"
    s = "SELECT p.Фамилия & ' ' & p.Имя & ' ' & p.Отчество AS Автор, q.Название, q.Авторы, q.ТипИздания, q.Дисциплина, q.ГодИздания "
    s = s & "FROM ([" & P & "Авторство] AS a INNER JOIN [" & P & "Издания подробно] AS q ON a.КодИздания = q.КодИздания) "
    s = s & "INNER JOIN [" & P & "Преподаватели] AS p ON a.КодПреподавателя = p.КодПреподавателя "
    s = s & "WHERE a.КодПреподавателя = " & FR & "[cboАвтор] AND q.ГодИздания Between " & FR & "[cboГодС] And " & FR & "[cboГодПо] "
    s = s & "ORDER BY q.ГодИздания DESC"
    MakeQuery P & "Издания автора за период", s
    MakeQuery P & "Количество изданий автора за период", _
        "SELECT Автор, Count(*) AS Изданий FROM [" & P & "Издания автора за период] GROUP BY Автор"

    ' --- формы ---
    MakeTableForm P & "Кафедры", P & "Кафедры", "Кафедры"
    MakeTableForm P & "Преподаватели", P & "Преподаватели", "Преподаватели (авторы)"
    MakeTableForm P & "ТипыИзданий", P & "ТипыИзданий", "Типы учебных материалов"
    MakeTableForm P & "Дисциплины", P & "Дисциплины", "Учебные дисциплины"
    MakeTableForm P & "Издания", P & "Издания", IIf(electronic, "Электронные учебные материалы", "Изданные учебные материалы")
    MakeTableForm P & "Авторы издания", P & "Авторство", "Авторы", 2, "КодИздания"
    MakeTableForm P & "Издания автора", P & "Авторство", "Издания", 2, "КодПреподавателя"
    AddSubform P & "Издания", P & "Авторы издания", "КодИздания", "КодИздания", "Авторы"
    AddSubform P & "Преподаватели", P & "Издания автора", "КодПреподавателя", "КодПреподавателя", "Издания автора"

    ' --- отчеты ---
    MakeReport P & "Издания по кафедрам и дисциплинам", P & "Издания подробно", "Перечень " & kind & " по кафедрам и учебным дисциплинам", _
        "Кафедра;Дисциплина", "Название;Авторы;ТипИздания;ГодИздания", "3600;2600;2000;900", "-ГодИздания;Название"
    MakeReport P & "Издания за последние 5 лет", P & "Издания за последние 5 лет", "Количество " & kind & " кафедр и университета за последние 5 лет", _
        "Кафедра", "ГодИздания;Название;Авторы;ТипИздания", "900;3600;2600;2000", "-ГодИздания;Название"
    MakeReport P & "Издания по типам за период", P & "Издания по типам за период", _
        "='Перечень и количество " & kind & " по типам за ' & [Forms]![" & P & "Управляющая]![cboГодС] & ' - ' & [Forms]![" & P & "Управляющая]![cboГодПо] & ' гг.'", _
        "ТипИздания", "Название;Авторы;Дисциплина;ГодИздания", "3600;2600;2000;900", "-ГодИздания;Название"
    MakeReport P & "Издания автора за период", P & "Издания автора за период", _
        "='Издания автора: ' & [Автор] & ' (' & [Forms]![" & P & "Управляющая]![cboГодС] & ' - ' & [Forms]![" & P & "Управляющая]![cboГодПо] & ' гг.)'", _
        "", "Название;Авторы;ТипИздания;ГодИздания", "3600;2600;2000;900", "-ГодИздания;Название"

    ' --- управляющая форма ---
    s = "SELECT DISTINCT ГодИздания FROM [" & P & "Издания] ORDER BY ГодИздания DESC"
    CF_Begin ttl
    CF_Param "cboГод", "Год", s, 1, "", CStr(Year(Date))
    CF_Param "cboАвтор", "Автор", "SELECT КодПреподавателя, Фамилия & ' ' & Имя & ' ' & Отчество FROM [" & P & "Преподаватели] ORDER BY Фамилия", 2, "0;4500", "1"
    CF_Param "cboГодС", "Период: с года", s, 1, "", CStr(Year(Date) - 4)
    CF_Param "cboГодПо", "Период: по год", s, 1, "", CStr(Year(Date))
    CF_Columns
    CF_Btn 0, IIf(electronic, "Электронные материалы", "Издания"), "F:" & P & "Издания"
    CF_Btn 0, "Преподаватели (авторы)", "F:" & P & "Преподаватели"
    CF_Btn 0, "Кафедры", "F:" & P & "Кафедры"
    CF_Btn 0, "Типы учебных материалов", "F:" & P & "ТипыИзданий"
    CF_Btn 0, "Учебные дисциплины", "F:" & P & "Дисциплины"
    CF_Btn 1, "Издания (все сведения)", "Q:" & P & "Издания подробно"
    CF_Btn 1, "По кафедрам и дисциплинам", "Q:" & P & "Издания по кафедрам и дисциплинам"
    CF_Btn 1, "Количество: кафедры за год", "Q:" & P & "Количество изданий кафедр за год|cboГод"
    CF_Btn 1, "Количество: университет за год", "Q:" & P & "Количество изданий университета за год|cboГод"
    CF_Btn 1, "Количество: кафедры за 5 лет", "Q:" & P & "Количество изданий кафедр за 5 лет"
    CF_Btn 1, "Количество: университет за 5 лет", "Q:" & P & "Количество изданий университета за 5 лет"
    CF_Btn 1, "Сводная: кафедры по годам", "Q:" & P & "Сводная кафедры по годам"
    CF_Btn 1, "Перечень по типам за период", "Q:" & P & "Издания по типам за период|cboГодС|cboГодПо"
    CF_Btn 1, "Количество по типам за период", "Q:" & P & "Количество по типам за период|cboГодС|cboГодПо"
    CF_Btn 1, "Перечень изданий автора", "Q:" & P & "Издания автора за период|cboАвтор|cboГодС|cboГодПо"
    CF_Btn 1, "Количество изданий автора", "Q:" & P & "Количество изданий автора за период|cboАвтор|cboГодС|cboГодПо"
    CF_Btn 2, "По кафедрам и дисциплинам", "R:" & P & "Издания по кафедрам и дисциплинам"
    CF_Btn 2, "Издания за последние 5 лет", "R:" & P & "Издания за последние 5 лет"
    CF_Btn 2, "По типам за период", "R:" & P & "Издания по типам за период|cboГодС|cboГодПо"
    CF_Btn 2, "Издания автора за период", "R:" & P & "Издания автора за период|cboАвтор|cboГодС|cboГодПо"
    CF_End P & "Управляющая"
End Sub
