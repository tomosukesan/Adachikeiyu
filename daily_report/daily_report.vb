' 構造体 '
Type s_data
	unit As Integer start As Date
	fin  As Date patient As String
End TypeA

' グローバル変数 '
Dim starting_staff_num As Integer
Dim end_staff_num As Integer

'グローバル変数への代入関数'
Function input_staff_num()
	starting_staff_num = 2
	End_staff_num = 15
End Function

' main関数 '
Sub 実行()
	Dim staff As String
	Dim i As Integer, num As Integer, total_unit As Integer
	Dim pos As Range

	Call make_sheet

	For i = 2 To 15
		staff = Cells(3, i).Value
		Set pos = Range(Cells(3, i).Address)
		num = count_intervention(staff)

		If Not num = 0 Then
			total unit = get_data(staff, num, pos)
			Cells(27, i).Value = total_unit
		End If
	Next
End Sub

' "貼付け"シートから担当者の介入回数を数える '
Function count_intervention(ByVal staff As String) As Integer
	Dim count As Integer, i As Integer

	i = 2
	Do While Not IsEmpty(Worksheets("貼付け").Range("R" & i) Value)
		If Worksheets("貼付け").Range("E" & i).Value = 0 Then
			Worksheets("貼付け").Rows(i).Delete shift:=xlUp
			GoTo Continue:
		End If
		If Worksheets("貼付け").Range("R" & i).Value = staff Then count = count + 1
		i = i + 1
Continue:
	Loop
	count_intervention = count
End Function

' 介入の詳細を取得 '
Function get_data (ByVal staff As String, do_num As Integer, pos As Range) As Integer
	Dim i As Integer, trans_num As Integer, total_unit As Integer
	Dim data() As s_data
	ReDim data (do_num — 1)

	Do While Not IsEmpty (Worksheets("貼付け").Range("R" & i).Value)
		If Worksheets("貼付け").Range ("R" & i).Value = staff Then
			data(trans_num).start = Worksheets("貼付け").Range("C" & i).Value
			data(trans_num).fin = Worksheets("貼付け").Range("D" & i).Value
			data(trans_num).unit = Worksheets("貼付け").Range("E" & i).Value
			data(trans_num).patient = Worksheets("貼付け").Range("H" & i).Value
			total_unit = total_unit + data(trans_num).unit
			trans_num = trans_num + 1
			Worksheets("貼付け").Rows(i).Delete shift:=xlUp
		Else
			i = i + 1
		End If
		If trans_num = do_num Then Exit Do
	Loop
	Call make_box(data, staff, pos, do_num)
	get_data = total_unit
End Function

' 図形作成 '
Function make_box (ByRef data() As s_data, ByVal staff As String, x_pos As Range, do_num As Integer)

	Dim y_pos As Range
	Dim start As String, fin As String
	Dim i As Integer

	For i = i To do_num - 1
		start = Format (data(i).start, "Short Time")
		fin = Format (data(i).fin, "Short Time")
		Set y_pos = judge_start(data(i).start)

		With ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 135, data(i).unit * 60)
			.Fill.ForeColor.RGB = RGB(255, 255, 255)
			With .TextFrame
				.HorizontalAlignment = xlHAlignCenter
				.VerticalAlignment = xlVAlignCenter
				.Characters.Text = data(i).patient & vbCrLf & start & " - " & fin
				.Characters.Font.Size = 16
				.Characters.Font.Color = RGB (0, 0, 0)
			End With
		.Top = y_pos.Top
		.Left = x_pos.Left
		End With
	Next
End Function

' 開始時間の判断 '
Function judge_start(ByVal start As Date) As Range
	Dim i As Integer
	For i = 4 To 26
		If DateAdd("n", 8, CDate(Range("A" & i).Value)) >= start Then
			Set judge_start = Range ("A" & i)
			Exit Function
		End If
	Next
End Function

' シートの作成 '
Function make_sheet()
	Dim new_sheet As String
	Dim today As String
	Dim day_of_week As String
	Dim i As Integer, total_unit As Integer

	new_sheet = InputBox("作成したい日にちを、数字のみ、入力してください")
	If new_sheet = "" Then End

	Worksheets("原本").Copy after :=Worksheets(Worksheets.count)
	ActiveSheet.Name = new_sheet & "日"
	Range ("B1").Value = Range("B1").Value & new_sheet & "日"

	today = Range("B1").Value
	day_of_week = WeekdayName(Weekday(today), True)

	If day of week = "月" Then
		For i = 2 To 15
			Range(Cells(28, i).Address).FormulaR1C1 = "=IF(R27C" & i & "= """", 0, R27C" & i & ")"
		Next
	Else
		For i = 2 To 15
			total unit = ActiveSheet.Previous.Cells(28, i).Value
			Range(Cells(28, i).Address).FormulaR1C1 = "=R27C" & i & "+" & total_unit
		Next
	End If
End Function