Dim place of creation As Integer
Dim a_team As Integer
Dim b_team As Integer
Dim starting row num As Integer
Dim end_row_num As Integer

Sub global_input_num()
	place_of_creation = 78
	a_team = 150
	b_team = 1500
End Sub

Sub input_row_num()
	starting_row_num = 36
	end_row_num = 116
End Sub

Sub make_box()
	On Error GoTo error_msg
	Dim txt As String
	Dim i As Integer, x As Integer, selection_num As Integer
	Dim background_color

	i = Selection(1).Row
	selection_num = Selection.Rows.count
	If IsEmpty(Range("C" & i).Value) Then
		Exit Sub
	End If
	background_color = Range("A" & i).Interior.Color
	For I = i To i + selection_num — 1
		txt = Range ("B" & i).Value & vbLf & Range("C" & i).Value
		If Not IsEmpty(Range("D" & i).Value) Then
			txt = txt & vbLf & Range("D" & i).Value
		End If
		If Not IsEmpty(Range("E" & i).Value) Then
			txt = txt & vbLf & Range ("E" & i).Value
		End If

		Call global_input_num

		If i < place_of_creation Then
			x = a_team
		Else
			x = b_team
		End If

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
				Instr (box.TextFrame.Characters.Text, Range("C" & num).Value) > 0 And _
				Not Instr(box.TextFrame.Characters.Text, "同席") > 0 And _
				Not Instr(box.TextFrame.Characters.Text, "IC") > 0 Then
					count = count - 1
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
		Call count_total_unit(num, Range("C" & num).Value)
		count = 0
		Range("A34").Select
continue:
	Next
	MsgBox ("更新完了しました")
End Sub

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

Function check_time(ByVal count As Integer, ByVal num As Integer)
	Dim i As Integer, end_time As Integer, ok_count As Integer
	Dim sort_array() As Integer
	ReDim sort_array(count)
	' Dim boxes() As Variant
	' ReDim boxes(count — 1, 1)
	' boxes = sort_boxes (count)

	Call sort_num(count, sort_array)

	For i = i To count - 1
		' end_time = boxes(i, 0) + boxes(i, 1)
		end_time = Selection.Shaperange(sort_array(i)).Top + Selection.ShapeRange(sort_array(i)).Height
		If end_time + 100 < Selection.ShapeRange(sort_array(i + 1)).Top Then
			ok_count = ok_count + 1
		ElseIf InStr(Selection.ShapeRange(sort_array(i)).TextFrame.Characters.Text, "術前単位") > 0 Then
			ok_count = ok_count + 1
		Elseif InStr(Selection.ShapeRange(sort_array(i + 1)).TextFrame.Characters.Text, "術前単位") > 0 Then
			ok_count = ok_count + 1
		End if
		' If end_time + 100 < boxes(i + 1, 0) Then	' 100はセル2つ分の高さ
		' If end_time + 100 < boxes(i + 1, 0) Or _ InStr(Selection.box.TextFrame.Characters.Text, "術前単位") > 0 Then
		' 	ok count = ok_count + 1
		' End If
	Next
	If count = ok_count + 1 Then
		Range("I" & num).Value = "●"
	Else
		Range("I" & num).Value = "NG"
	End If

	Call count_unit(count, num)
	' Call count_unit(count, num, boxes())
End Function

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

Function sort_boxes(ByVal count As Integer) As Variant
	Dim box As Shape, i As Integer, j As Integer, tmp As Integer
	Dim boxes()
	ReDim boxes(count —	1, 1)

	For Each box In Selection.ShapeRange
		boxes(i, 0) = box.Top
		boxes(i , 1) = box.Height
		i = i + 1
	Next
	i = 0
	For i = i To count - 2
		For j = i + 1 To count - 1
			If boxes(i, 0) > boxes(j. 0) Then
			tmp = boxes(i, 0)
			boxes(i, 0) = boxes(j, 0)
			boxes (j . 0) tmp
			tmp = boxes(i , l) boxes ( i , boxes (j . l) boxes (j .
			End If

		Next
	Next
	sort_boxes = boxes()
End Function

Function count_unit(ByVal count As Integer, num As Integer, boxes ())

Dim i As Integer, unit As Integer, total_unit As Integer



Range ("K" & num) . Value =

For i = 0 To count — 1

If boxes ( i , l) < 90 Then
 unit =1

Elself boxes ( i , uni unit = 2

l) < 135 Then

El self boxes(i, uni t = 3

l) < 190 Then

El self boxes ( i .

l ) < 255 Then

unit = 4 El se

MsgBox Range ("C" & num) . Value & vbCrLf & “手動での入力が必要です。”

Exit Function End If tota I _un i t = unit + total_unit

Next

Range ("K" & num). Va lue = total_unit

End Function

Function count_tota I _un it (ByVal cel_num As Integer, search_name As Str ing)

Dim ref_num As Integer, i As Integer, j As Integer

Dim  pre_day_total_unit As Integer, today_unit As Integer

Dim pre_sheet As Worksheet

If ActiveSheet. Previous I s Noth ing Or InStr (ActiveSheet. Previous. Name, “(0)” > O Then Range ("L" & cel_num). Value — Range ("K" & cel_num) . Value

Exit Function

End If

If InStr (Acti veSheet. Previous. Name, " a ") > 0 Then ActiveSheet. Previous. Select

If search_name = ActiveSheet. Previous. Range ("C" & cel_num) Then pre_day total_unit = ActiveSheet. Previous. Range ("L" & cel_num) . Value

End If

If InStr (ActiveSheet. Name, " a ") > 0 Then ActiveSheet. Next. Select

today_unit - Range ("K" & cel_num) . Value

Range ("L" & cel num) . Va lue = pre_day_tota l_unit + today_unit

End Function

Sub discharge ( )

On Error GoTo error_msg

Dim delete_num, ent_day, staff, patient As String

Dim num, month_unit, i As Integer Dim current_ sheet As Worksheet

Set current_ sheet = Act i veSheet num = Selecti on. Row staff = Range ("A" & num) . MergeArea (1, 1) . Va lue patient = Range ("C" & num) . Value month_unit — Range ("L" & num) . Value ent_day — Range ( 'F BI MergeArea (1 , 1) . Value

godulet - 4

	If patient =	Then Exit Sub

“数字のみ入力してください。”

delete num = InputBox(“何日分の退院処理をしますか？”)  & vbCrLf & " a gæ a t, ab,

If StrPtr (delete_num) = 0 Then Exit Sub

Elself Not delete num = 0 Then

CalI erase_box_and_info (num, pati ent, Clnt (delete_num) ) End If

Worksheets gWorksheets. count) . Se I ect

If Not ActiveSheet.Name = “退院” Then

	MsgBox “「退院」シートの後ろに余分なシートが存在しています。” & vbCrLf & “余分なシートであれば削除してください。”

Exit Sub

End If

Do Whi le Not IsEmpty(Cel I s ( i . 1) . Va lue)

I = i + 1

Loop

Cel Is ( i ,

= patient

cel Is ( i . 2)

= ent_day

cel Is ( i . 3)

= month_unit

Cel Is ( i . 4)

= staff

current_sheet Se I ect

Exit Sub error_msg.

	MsgBox   “エラーが発生しました”	& vbCrLf & “選択箇所を確認し、再度お試しください。”

End Sub

Function erase_box_and_info (ByVal row_num As Integer, patient As Str ing, delete_num As Integer)

Dim col _num, sheet_num As Integer Dim box As Shape

For sheet_num = 1 To delete_num

Worksheets (ActiveSheet. Index + 1) . Select

If Acti veSheet. Index = Worksheets. count Then Exit Function

End If

For Each box In ActiveSheet. Shapes

If box. Type = 1 And InStr (box. TextFrame. Characters. Text, Range ("C" & row_num) Value) > 0 Then box. Delete

End If

Next

For col num = 2 To 12

Cells (row_num, col_num) Value —

If col_num = 6 Then col_num = col num + 2

Next Next

End Function
