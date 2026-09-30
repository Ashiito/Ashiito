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
