Attribute VB_Name = "fileNameComp"
Public Sub fileNameWrite()
    Dim st As Worksheet
    Dim i As Integer
    Dim folderPath As String
    Dim lastRowNum(1) As Long
    Dim maxRowNum As Long
    Dim rng As Range
    Set st = ThisWorkbook.Sheets("compare")
    '�V�[�g�N���A
    For i = 2 To 3
        lastRowNum(i - 2) = st.Cells(st.Rows.Count, i).End(xlUp).row
    Next i
    If lastRowNum(0) < lastRowNum(1) Then
        maxRowNum = lastRowNum(1)
    Else
        maxRowNum = lastRowNum(0)
    End If
    If maxRowNum < 4 Then maxRowNum = 4
    Set rng = st.Range(st.Cells(4, 2), st.Cells(maxRowNum, 3))
    Call sheetClr(rng)
    '�t�@�C���������o��
    For i = 2 To 3
        folderPath = st.Cells(3, i).Value
        Call GetFileNames(folderPath, i)
    Next i
    MsgBox "�t�@�C�����������o���܂���"
End Sub

Private Sub GetFileNames(folderPath As String, rowNum As Integer)
    Dim fileName As String
    Dim row As Long
    
    ' �t�H���_�p�X�̖����� "\" ���Ȃ��ꍇ�͒ǉ�
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' �����ݒ�
    fileName = Dir(folderPath) ' �ŏ��̃t�@�C�����擾
    row = 4 ' �o�͊J�n�s�i��: �V�[�g��1�s�ځj
    
    ' �t�@�C�������擾���ăV�[�g�ɏo��
    Do While fileName <> ""
        fileName = LCaseChange(fileName)
        Cells(row, rowNum).Value = fileName ' A��Ƀt�@�C�������o��
        row = row + 1
        fileName = Dir ' ���̃t�@�C�����擾
    Loop
End Sub

Private Function LCaseChange(fileName As String) As String
    '�g���q���������ɕϊ�
    Dim kakucho As String
    Dim c_kakucho As String
    
    kakucho = Right(fileName, Len(fileName) - InStr(fileName, "."))
    Debug.Print kakucho
    c_kakucho = LCase(kakucho)
    LCaseChange = Left(fileName, InStr(fileName, ".") - 1) + "." + c_kakucho
End Function

Private Function sheetClr(rng As Range)
   rng.Interior.ColorIndex = 2
   rng.Clear
End Function

Public Sub fileComp()
    'B���C��̃t�@�C�����r
    Dim lastRowNum(1) As Long
    Dim i As Integer
    Dim j As Integer
    Dim sameflag As Boolean
    Dim st As Worksheet
    Set st = ThisWorkbook.Sheets("compare")
    For i = 2 To 3
        lastRowNum(i - 2) = st.Cells(st.Rows.Count, i).End(xlUp).row
    Next i
    For i = 4 To lastRowNum(1)
        st.Range(st.Cells(i, 2), st.Cells(i, 3)).Interior.ColorIndex = 2
        st.Range(st.Cells(i, 3), st.Cells(i, 3)).Interior.ColorIndex = 6
    Next i
    For i = 4 To lastRowNum(0)
        findval = st.Cells(i, 2).Value
        sameflag = False
        For j = 4 To lastRowNum(1)
            If st.Cells(j, 3).Value = findval Then
                sameflag = True
                st.Range(st.Cells(j, 3), st.Cells(j, 3)).Interior.ColorIndex = 2
                Exit For
            End If
        Next j
        If sameflag = False Then
            '�����t�@�C�������Ȃ��ꍇ�̓Z���ɐF�Â�����
            st.Range(st.Cells(i, 2), st.Cells(i, 2)).Interior.ColorIndex = 6
        End If
    Next i
    MsgBox "�����I�����܂���"
End Sub


