' グローバル変数 '
Dim starting_staff_col As String
Dim end_staff_col As String
Dim total_staff_except_chief As Integer
Dim head_staff_except_chief As Integer
Dim starting_staff_num As Integer
Dim end_staff_num As Integer

' グローバル変数への代入 '
Function input_global_var()
	starting_staff_col = "B"
	end_staff_col = "O"
	total_staff_except_chief = 15
	head_staff_except_chief = 4
	starting_staff_num = 5
	End_staff_num = 18
End Function

' main関数 '
Sub スケジュール作成()
	Dim col As Integer
	Dim src_book As Workbook

	Set src_book = make_book()

	Application.ScreenUpdating = False

	For col = 4 To 34
		Call make_sheet(col)
	Next

	Call make_discharge_sheet

	Application.DisplayAlerts = False
	Sheets(1).Delete
	Application.DisplayAlerts = True
	Application.ScreenUpdating = True
	MsgBox "作成完了しました。"
	src_book.Close
End Sub

' 任意のエクセルファイルを作成 '
Function make_book () As Workbook

	Dim file_name As String, path As String, schedule As String
	Dim src_sheet As Worksheet
	Dim src_book As Workbook

	path = ""  ' スケジュールを作成するフォルダのパス

	file_name = InputBox("ファイル名を入力してください。")
	If StrPtr (file_name) = 0 Then End
	Set src_sheet = ActiveSheet
	Set src_book = ActiveWorkbook
	schedule = path & file_name & ".xlsm"
	FileCopy path & "原本.xlsm", schedule
	Workbooks.Open schedule
	src_sheet.Copy before:=Worksheets(1)
	Set make_book = src_book
End Function

' シート作成 '
Function make_sheet (ByVal col As Integer)

	Call input_global_var

	Dim day As String, weekday As String
	Dim roster As Worksheet
	Dim pos As Range
	Dim random_num As Integer

	day = Sheets(1).Cells(3, col).text
	weekday = Sheets(1).Cells(4, col).text

	Sheets("原本").Copy after:=Sheets(Sheets.Count)
	ActiveSheet.Name = day & "(" & weekday & ")"
	Cells(1, 2).Value = day & "日"

	If weekday = "日" Then
		Set pos = Range(starting_staff_col & "4:" & end_staff_col & "26")
		Call make (pos, "公休", 30, RGB(255. 162. 128))
		Exit Function
	Else
		Call judde_dayoff(col)
		Call judge_weekday(weekday)
		random_num = Int( (total_staff_except_chief - head_staff_except_chief + 1) * rand + head_staff_except_chief)
		Call make_box(Range(Cells(4, rangom num).Address), "スキャン", 18. RGB(255, 255, 255))
		Call make_box(Range("C4"), "日報", 18, RGB(255. 255. 255))
		Call make_box(Range("D4"), "看護MTG", 18, RGB(255. 255, 255))
	End If
End Function

' スタッフの休暇を判断 '
Function judge_dayoff (ByVal col As Integer)
	Call input_global_var

	Dim row As Integer
	Dim pos As Range
	Dim text As String

	For row = starting_staff_num To end_staff_num
		text = "公休"
		If Sheets(1).Cells(row, col).Value ="日" Or _
			Sheets(1).Cells(row, col).Value = "" Then GoTo continue:
		If Sheets(1).Cells(row, col).Value = "休" Then
			Set pos = Range(Cells(4, row - 3), Cells(26, row - 3))
		ElseIf Sheets(1).Cells(row, col).Value = "AM" Then
			Set pos = Range(Cells(14, row — 3), Cells(26, row -3))
		ElseIf Sheets(1).Cells(row, col).Value = "PM" Then
			Set pos = Range(Cells(4, row — 3), Cells(15, row -3))
		ElseIf Sheets(1).Cells(row, col).Value = "有" Then
			Set pos = Range(Cells(4, row — 3), Cells(26, row — 3))
			text = "有休"
		ElseIf Sheets(1).Cells(row, col).Value = "リ" Then
			Set pos = Range(Cells(4, row — 3), Cells(26, row -3))
			text = "リモート"
		ElseIf Sheets(1).Cells(row, col).Value = "夏" Then
			Set pos = Range (Cells(4, row — 3), Cells(26, row-3))
			text = "夏休"
		Else
			Set pos = Range (Cells(4, row — 3), Cells(26, row -3))
			text  = "不明"
			MsgBox (Sheets(1).Cells(3, col).Value & "日に不明な休みを検知しました。")
		End If
		Call make_box(pos, text, 18, RGB(255, 162. 128))
continue:
	Next
End Function

' 曜日ごとの予定を入力 '
Function judge_weekday(ByVal weekday As Str ing)
	Dim random_num As Integer

	If weekday = "月" Then
		Call make_box(Range(starting_staff_col $ "12:" & end_staff_col & "12"), "チームMTG", 28, RGB(255, 255,255))
		Call make_box(Range("B22:B25"), "退院支援" & vbCrLf & "カンファ", 18, RGB(255, 255, 255))
		Call make_box(Range("D22:D25"), "退院支援" & vbCrLf & "カンファ", 18, RGB(255, 255, 255))
		Call make_box(Range("J22:J25"), "退院支援" & vbCrLf & "カンファ", 18, RGB(255, 255, 255))
		Call make_box(Range("E14"),	"装具TEL" , 18,	RGB(255, 255, 255))
	ElseIf weekday = "水" Then
		Call make_box(Range(starting_staff_col & "26:" & end_staff_col & "26", "整形カンファ", 28, RGB(255, 255, 255))
		random_num = Int((total_staff_except_chief - head_staff_except_chief + 1) * rand + head_staff_except_chief)
		Call make_box(Range(Cells(25, random_num).Address), "洗濯",18, RGB(255, 255, 255))
	ElseIf weekday =  "金" Then
		Call make_box(Range("E25") , "装具TEL", 18, RGB(255, 255, 255))
	ElseIf weekday = "土" Then
		random_num = Int((total_staff_except_chief — head_staff_except_chief + 1) * rand + head_staff_except_chief)
		Call make_box(Range(Cells(26, random_num).Address), "洗濯" 18, RGB (255, 255, 255) )
	End If
End Function

' 退院シートの作成 '
Function make_discharge_sheet()
	Worksheets.Add after:=Worksheets(Worksheets.Count)
	ActiveSheet.Name = "退院"
	Range("A1:D1").Interior.Color = RGB(211. 211, 21 1)
	Range("A1").Value = "退院患者"
	Range("B1").Value = "退院日"
	Range("C1").Value = "月単位"
	Range("D1").Value = "担当者"
End Function

' 図形作成 '
Function make_box(ByVal pos As Range, text As String, text_size As Integer, bg_color As Long)

	With ActiveSheet.Shapes.AddShape(msoShapeReGtangle, 0, 0, pos.Width, pos.Height)
		.Fill.ForeColor.RGB = bg_color
		.Line.ForeColor.RGB = RGB(0, 0, 0)
		With .TextFrame
			HorizontalAligment = xlHAlignCenter
			VerticalAligment = xlVAlignCenter
			Characters.text = text
			Characters.Font.Size = text_size
			Characters.Font.Color = RGB(0, 0, 0)
		End With
		.Top = pos.Top
		.Left = pos.Left
	End With
End Function