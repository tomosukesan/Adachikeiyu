' グローバル変数 '
Dim place of creation As Integer
Dim a_team As Integer
Dim b_team As Integer
Dim starting row num As Integer
Dim end_row_num As Integer

' グローバル変数への代入関数 '
Sub global_input_num()
	place_of_creation = 78
	a_team = 150
	b_team = 1500
End Sub

' グローバル変数への代入関数 '
Sub input_row_num()
	starting_row_num = 36
	end_row_num = 116
End Sub

' 作成ボタン：図形作成 '
Sub make_box()
	On Error GoTo error_msg		' エラーハンドリング
	Dim txt As String			' セルにある文字列を格納する変数
	Dim i As Integer, x As Integer, selection_num As Integer
	Dim background_color		' セルの背景色を格納

	Call global_input_num		' グローバル変数の値を代入
	i = Selection(1).Row					' 選択されたセルの最初の行の数値を格納
	selection_num = Selection.Rows.count	' 選択されたセルの最後の行の数値を格納
	If IsEmpty(Range("C" & i).Value) Then	' 空白セルの時は無視する
		Exit Sub
	End If
	background_color = Range("A" & i).Interior.Color
	For I = i To i + selection_num — 1		' 選択された行の情報を１つずつ処理していく
		txt = Range ("B" & i).Value & vbLf & Range("C" & i).Value
		If Not IsEmpty(Range("D" & i).Value) Then
			txt = txt & vbLf & Range("D" & i).Value
		End If
		If Not IsEmpty(Range("E" & i).Value) Then
			txt = txt & vbLf & Range ("E" & i).Value
		End If

		If i < place_of_creation Then		' 図形の作成場所を変数xに代入
			x = a_team
		Else
			x = b_team
		End If

		' 作成する図形の情報を与える
		With ActiveSheet.Shapes.AddShape(msoShapeRectangle, x. 200, 135, 115)
			.Fill.ForeColor.RGB = background_color
			.Line.ForeColor.RGB = RGB(0, 0, 0)
			With .TextFrame
				.Characters.Text = txt
				.Characters.Font.Size = 18
				.Characters.Font.Color = black
				.HorizontalAlignment = xlCenter
				.VerticalAlignment = xlCenter
			End With
		End With
	Next
	Exit Sub
error_msg:
     MsgBox  "エラーが発生しました。" & vbCrLf & "選択箇所を確認し、再度お試しください。"
End Sub

' 更新ボタン：表への出力 '
Sub display_count()
	Application.ScreenUpdating = False
	Dim count As Integer, num As Integer
	Dim box As Shape

	Range("A34").Select
	Call input_row_num
	For num = starting_row_num To end_row_num
		If IsEmpty(Range("C" & num).Value) Then
			GoTo continue:
		End If
		For Each box In ActiveSheet.Shapes
			If box.Type = 1 And_
				InStr (box.TextFrame.Characters.Text, Range("C" & num).Value) > 0 And _
				Not InStr(box.TextFrame.Characters.Text, "同席") > 0 And _
				Not InStr(box.TextFrame.Characters.Text, "IC") > 0 Then
					count = count + 1
					box.Select False
			End If
		Next
		Range("J" & num) = count
		If Not count = O Then
			Call check_time (count, num)
			Call check_room_num(count, num)
		Else
			Range("I" & num).Value = ""
			Range("J" & num).Value = 0
			Range("K" & num).Value = 0
		End If
		'Call count_total_unit(num, Range("C" & num).Value)
		count = 0
		Range("A34").Select
continue:
	Next
	MsgBox ("更新完了しました")
End Sub

' 部屋番号の確認と変更 '
Function check_room_num(ByVal count As Integer, num As Integer)
	Dim i As Integer, beginning As Integer
	Dim txt As String

	For i = 1 To count
		txt = Selection.ShapeRange(i).TextFrame.Characters.Text
		If InStr(txt, Range("B" & num).Value) = 0 Then
			beginning = InStr(txt, Range("C" & num).Value)
			Selection.ShapeRange(i).TextFrame.Characters.Text = Range("B" & num).Value & vbCrLf & Mid(txt, beginning)
		End if
	Next
End Function

' 時間の重複確認 '
Function check_time(ByVal count As Integer, ByVal num As Integer)
	Dim i As Integer, end_time As Integer, ok_count As Integer
	Dim sort_array() As Integer
	ReDim sort_array(count)

	Call sort_num(count, sort_array)

	For i = i To count - 1
		end_time = Selection.Shaperange(sort_array(i)).Top + Selection.ShapeRange(sort_array(i)).Height
		If end_time + 100 < Selection.ShapeRange(sort_array(i + 1)).Top Then
			ok_count = ok_count + 1
		ElseIf InStr(Selection.ShapeRange(sort_array(i)).TextFrame.Characters.Text, "術前単位") > 0 Then
			ok_count = ok_count + 1
		Elseif InStr(Selection.ShapeRange(sort_array(i + 1)).TextFrame.Characters.Text, "術前単位") > 0 Then
			ok_count = ok_count + 1
		End if
	Next
	If count - 1 = ok_count Then
		Range("I" & num).Value = "●"
	Else
		Range("I" & num).Value = "NG"
	End If

	Call count_unit(count, num)
End Function

' 時間順にソート '
Function sort_num(ByVal count As Integer, ByRef sort_array() As Integer)
	Dim i As Integer, j As Integer, tmp As Integer

	For i = 1 To count
		sort_array(i) = i
	Next
	For i = 1 To count -1
		For j = i + 1 to count
			If Selection.ShapeRange(sort_array(i)).Top > Selection.ShapeRange(sort_array(j)).Top Then
				tmp = sort_array(i)
				sort_array(i) = sort_array(j)
				sort_array(j) = tmp
			End if
		Next
	Next
End Function

' 日単位数の計算 '
Function count_unit(ByVal count As Integer, num As Integer)
	Dim i As Integer, unit As Integer, total_unit As Integer

	Range("K" & num).Value = ""
	For i = 1 To count - 1
		If Selection.ShapeRange(i).Height < 90 Then
			unit = 1
		ElseIf Selection.ShapeRange(i).Height < 135 Then
			unit = 2
		ElseIf Selection.ShapeRange(i).Height < 190 Then
			unit = 3
		ElseIf Selection.ShapeRange(i).Height < 255 Then
			unit = 4
		Else
			MsgBox Range("C" & num).Value & vbCrLf & "手動での単位入力が必要です"
			Exit Function
	Next
	Range("K" & num).Value = total_unit
End Function

' 月単位数の計算 '  : 下記関数、機能停止(2023/02/07)
Function count_total_unit(ByVal cel_num As Integer, search_name As String)
	Dim ref_num As Integer, i As Integer, j As Integer
	Dim pre_day_total_unit As Integer, today_unit As Integer
	Dim pre_sheet As Worksheet

	If ActiveSheet.Previous Is Nothing Or InStr(ActiveSheet.Previous.Name, "(0)") > 0 Then
		Range("L" & cel_num).Value = Range("K" & cel_num).Value
		Exit Function
	End If
	If InStr (ActiveSheet.Previous.Name, "日") > 0 Then ActiveSheet.Previous.Select

	If search_name = ActiveSheet.Previous.Range("C" & cel_num) Then
		pre_day_total_unit = ActiveSheet.Previous.Range("L" & cel_num).Value
	End If

	If InStr(ActiveSheet.Name, "日") > 0 Then ActiveSheet.Next.Select

	today_unit = Range("K" & cel_num).Value
	Range("L" & cel num).Value = pre_day_total_unit + today_unit
End Function

' 削除ボタン：図形の削除 '
Sub delete_box()
	On Error GoTo error_msg
	Dim i As Integer, selection_num As Integer
	Dim box As Shape
 	i = Selection(1).Row
	selection_num = Selection.Rows.count

	For i = i To i + selection_num — 1
		If IsEmpty (Range("C" & i).Value) Then GoTo continue:
		For Each box In ActiveSheet.Shapes
			If box.Type = 1 And InStr(box.TextFrame.Characters.Text, Range("C" & i).Value) > 0 Then
				box.Delete
				Range ("I" & i).Value = ""
				Range ("J" & i).Value = 0
				Range ("L" & i).Value = Range("L" & i).Value - Range ("K" & i).Value
				Range ("K" & i).Value = 0
			End If
		Next
continue:
	Next
	Exit Sub
error_msg:
	MsgBox  "エラーが発生しました。" & vbCrLf & "選択箇所を確認し、再度お試しください。"
End Sub

' 退院ボタン：図形と表データの削除 '
Sub discharge()
	On Error GoTo error_msg
	Dim delete_num, ent_day, staff, patient As String
	Dim num, month_unit, i As Integer
	Dim current_ sheet As Worksheet

	Set current_ sheet = ActiveSheet
	num = Selection.Row
	staff = Range("A" & num).MergeArea(1, 1).Value
	patient = Range("C" & num).Value
	month_unit = Range("L" & num).Value
	ent_day = Range("B1").MergeArea(1 , 1).Value

	If patient = "" Then Exit Sub

	delete_num = InputBox("何日分の退院処理をしますか？")  & vbCrLf &  "日曜日も含め、数字のみ、入力してください。"

	If StrPtr(delete_num) = 0 Then
		Exit Sub
	ElseIf Not delete_num = 0 Then
		Call erase_box_and_info(num, patient, CInt(delete_num))
	End If

	Worksheets(Worksheets.count).Select
	If Not ActiveSheet.Name = "退院" Then
		MsgBox "「退院」シートの後ろに余分なシートが存在しています。" & vbCrLf & "余分なシートであれば削除してください。"
		Exit Sub
	End If
	Do While Not IsEmpty(Cells(i, 1).Value)
		i = i + 1
	Loop

	Cells(i, 1)	= patient
	Cells(i, 2) = ent_day
	Cells(i, 3) = month_unit
	Cells(i, 4) = staff

	current_sheet.Select
	Exit Sub
error_msg:
	MsgBox   "エラーが発生しました"	& vbCrLf & "選択箇所を確認し、再度お試しください。"
End Sub

' 図形と表データの削除 '
Function erase_box_and_info(ByVal row_num As Integer, patient As String, delete_num As Integer)
	Dim col _num, sheet_num As Integer
	Dim box As Shape

	For sheet_num = 1 To delete_num
		Worksheets(ActiveSheet.Index + 1).Select
		If ActiveSheet.Index = Worksheets.count Then
			Exit Function
		End If
		For Each box In ActiveSheet.Shapes
			If box. Type = 1 And InStr(box.TextFrame.Characters.Text, Range("C" & row_num).Value) > 0 Then
				box. Delete
			End If
		Next
		For col num = 2 To 12
			Cells(row_num, col_num) Value = ""
			If col_num = 6 Then col_num = col num + 2
		Next
	Next
End Function

' 介入時間表示 '
Sub time_announce()
	Application.ScreenUpdating = False
	' 名前(C?~)で図形を検索し、選択状態へ（更新ボタンの重複チェックを参考に）
	' 図形をソート（sort_num(count, sort_array())関数を参考に）
	' shape_top変数に.Topを代入し、その値に応じて開始時間を【介入時間(N?)】へ記載
	Dim count As Integer, num As Integer
	Dim box As Shape


	Range("A34").Select
	Call input_row_num
	For num = starting_row_num To end_row_num
		If IsEmpty(Range("C" & num).Value) Then
			GoTo continue:
		Else
			Range("N" & num).Value = ""
		End If
		For Each box In ActiveSheet.Shapes
			' 下記、displayと全く同じなので関数分割しても良いか
			If box.Type = 1 And_
				InStr (box.TextFrame.Characters.Text, Range("C" & num).Value) > 0 And _
				Not InStr(box.TextFrame.Characters.Text, "同席") > 0 And _
				Not InStr(box.TextFrame.Characters.Text, "IC") > 0 Then
					count = count + 1
					box.Select False
			End If
		Next
		If Not count = 0 Then
			Dim sort_array() As Integer
			ReDim sort_array(count)
			Call sort_num(count, sort_array)
			Call judge_time(num, count, sort_array)
		End If
		Range("A34").Select
continue:
	Next
	MsgBox ("介入時間を更新しました")
End Sub

' 図形の位置から介入開始時間を判断 '
Function judge_time(ByVal num As Integer, ByVal count As Integer, ByRef sort_array() As Integer)
	Dim shape_top As Integer, i As Integer
	Dim start As String

	For i = i To count -1
		shape_top = Selection.Shaperange(sort_array(i)).Top
		Select Case shape_top
			Case Is < ?
				start = " 9:05 "
			Case Is < ?
				start = " 9:25 "
			Case Is < ?
				start = " 9:45 "
			Case Is < ?
				start = " 10:05 "
			Case Is < ?
				start = " 10:25 "
			Case Is < ?
				start = " 10:45 "
			Case Is < ?
				start = " 11:05 "
			Case Is < ?
				start = " 11:25 "
			Case Is < ?
				start = " 11:45 "
			Case Is < ?
				start = " 13:05 "
			Case Is < ?
				start = " 13:25 "
			Case Is < ?
				start = " 13:45 "
			Case Is < ?
				start = " 14:05 "
			Case Is < ?
				start = " 14:25 "
			Case Is < ?
				start = " 14:45 "
			Case Is < ?
				start = " 15:05 "
			Case Is < ?
				start = " 15:25 "
			Case Is < ?
				start = " 15:45 "
			Case Is < ?
				start = " 16:05 "
			Case Is < ?
				start = " 16:25 "
			Case Is < ?
				start = " 16:45 "
		End Select
		Cells(num, 14).Value = Cells(num, 14).Value & start
	Next
End Function