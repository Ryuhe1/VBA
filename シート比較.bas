Attribute VB_Name = "Module30" 
Option Explicit

Sub ������r�ł̍��ق̖����ƃ��X�g��()

Dim motob As Workbook
Dim wbl As Workbook
Dim sakib As Workbook
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ds As Worksheet
Dim cll As Range
Dim c12 As Range
Dim snl As String 
Dim sn2 As String
Dim dc As Long
Dim rtn As Long
Dim rtn2 As Long
Dim i As Long
Dim j As Long
Dim arr()
Dim mt As Variant
Dim sk As Variant
Dim f1
Dim f2
Dim p1
Dim p2

Const ros As Long = 3 '�����\�X�^�[�g�s�i�[�L�q�͂��܂�s��
Const lis As Long = 2 '�����\�X�^�[�g��i�[�L�q�͂��܂��

Set wb1 = ThisWorkbook '���̃u�b�N

Call ���ʑO����(wb1, motob, sakib, ws1, ws2, ds, cl1, cl2, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

ws1.Activate
Call �����N�폜(sakib)

''******���Ԍv���͂��܂�**************
'    Dim start_time As Double
'    Dim fin_time As Double
'    start_time = Timer
''***********************************
'******�����������Z�b�g�i�J�n�j**************
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xCalculationManual
'******************************************

'�C���f�b�N�X�̕\��
With ds.Cells(ros, lis)
    .Offset(-2, 0) = "�u�b�N"
    .Offset(-1, 0) = "�V�[�g"
    .Offset(-2, 1) = mt  '�t�@�C�����\��
    .Offset(-2, 2) = sk
    .Offset(-1, 1) = sn1 '�V�[�g���\��
    .Offset(-1, 2) = sn2 
    .Offset(0, 0) = "�ΏۃZ��"
    .Offset(0, 1) = "���Ƃ̎�"
    .Offset(0, 2) = "�����̎�"
End With

ds.Activate
Call �\�̐��`(ds, lis, ros)

dc = ros

'�V�[�g�Ƀe�[�u��������ꍇ�̑I��
Call �͈͂ɕϊ�����(rtn2, cl1, cl2, ws1, ws2, ds, dc, lis, wb1)

'�����e�[�u���łȂ��ꍇ�Ɉȉ��̏���
If rtn2 <> vbOK Then
'�e�V�[�g�̍��ق���Z���̔��菈��
    Set cl1 = ws1.UsedRange
    Set cl2 = ws2.Range(cl1.Address)
    Dim ursr As Long
    ursr = ws1.UsedRange.Row - 1
    Dim ursc As Long
    ursc = ws1.UsedRange.Column - 1
    Dim arys1()
    Dim arys2()
    Dim arys3 As Variant
    Dim aryex As Variant
    arys1 = cl1.Formula  '���ƃV�[�g�̎g�p�Z���͈͐�����S�Ĕz��ɑ��
    arys2 = cl2.Formula  '�����V�[�g�̎g�p�Z���͈͐�����S�Ĕz��ɑ��

    '���I�z��̂P�n�܂�ɍĒ�`
    ReDim arys3(1 To 2, 1 To dc)
    ReDim aryex(1 To 2, 1 To 2)

' ���ƃV�[�g�̃Z�������[�v����
    For i = 1 To UBound(arys1, 1)
        For j = 1 To UBound(arys1, 2)
    '�Q�̃V�[�g�̐������قȂ�ꍇ�Z���ɐF��t����
            If arys1(i, j) <> arys2(i, j) Then
        '�G���[�����������ꍇ�͈قȂ鐔���Ȃ̂ŃZ���ɐF��t����
                ws1.Cells(i, j).Offset(ursr, ursc).Interior.Color = rgbYellow
                ws2.Cells(i, j).Offset(ursr, ursc).Interior.Color = rgbGold
        '�����̂���Z���ԍ����o��
                ds.Cells(dc, lis).Offset(1, 0).Formula = ws1.Cells(i, j).Offset(ursr, ursc).Address(False, False)
        '�����N��t���i65,530���ȉ��̏ꍇ�̂݁j
                If dc < 65531 Then
                    ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc, lis).Offset(1), Address:="", _
                    SubAddress:=" '����'!" & ws1.Cells(i, j).Offset(ursr, ursc).Address(False, False)
                End If
        '�e�V�[�g�̍�����񎟌��z��ɑ���i������r�\�j
                ReDim Preserve arys3(1 To 2, 1 To dc)   '�񎟌��̂ݒǉ��\�i�z��̐����j
                    arys3(1, dc - ros + 1) = arys1(i, j)
                    arys3(2, dc - ros + 1) = arys2(i, j)
        '�����V�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
                ws1.Cells(i, j).Offset(ursr, ursc).AddComment
                ws1.Cells(i, j).Offset(ursr, ursc).Comm
                ent.Text Text:="<����>: " & arys2(i, j) & vbCrLf & "<����>: " & arys1(i, j)
        '���ƃV�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
                ws2.Cells(i, j).Offset(ursr, ursc).AddComment
                ws2.Cells(i, j).Offset(ursr, ursc).Comment.Text Text:="<����>: " & arys1(i, j) & vbCrLf & "<����>: " & arys2(i, j)
            dc = dc + 1
            End If
        Next
    Application.StatusBar = dc & "�s�ڂ̏��������Ă��܂�..."
    Next
    Application.StatusBar = False
            
    Erase arr
    Erase arys1
    Erase arys2

    '������r�\�]�L�p�z��̍s����ꂩ��
    Dim ro: Dim col
    ReDim aryex(1 To UBound(arys3, 2), 1 To UBound(arys3, 1))
    For ro = 1 To UBound(arys3, 1)
        For col = 1 To UBound(arys3, 2)
            aryex(col, ro) = arys3(ro, col)
        Next
    Next
    DoEvents
    '������r�\�ɍs�����ւ��ςݔz���\�t��
    ds.Cells(ros + 1, lis + 1).Resize(UBound(aryex, 1), UBound(aryex, 2)).FormulaLocal = aryex

    Erase arys3
    Erase aryex
End If  
'�����e�[�u���łȂ��ꍇ�̏����I���

'������r�\�̌r���̋L��
ds.Cells(ros, lis). CurrentRegion.Borders.LineStyle = xlContinuous
'�����\���ɕύX
ds.Activate
ActiveWindow.DisplayFormulas = True

'�����t�@�C���R�s�[���{�̏ꍇ�̓R�s�[�����t�@�C������ăt�H���_���J���i�蓮�폜�j
If sakib.Name Like "���R�s�[��*" Then
    MsgBox "�t�@�C�����������̂��߃��l�[���R�s�[���܂����B�Y���t�H���_���J���܂��B���R�s�[���Ŏn�܂閼�O�̃t�@�C���͕K�v�Ȃ��̂ō폜���Ă��������B"
    Dim pth
    pth = Left(p1, Len(p1) - 1)
    Shell "C:\windows\explorer.exe " & pth & "\", vbNormalFocus  '�Y���t�H���_���J��
End If

ws1.Activate '���ƃV�[�g���A�N�e�B�u��

'��r���Ƃ����u�b�N�����
    On Error Resume Next
    Application.DisplayAlerts = False
        sakib.Close savechanges:=False
        motob.Close savechanges:=False
    Application.DisplayAlerts = True

Set wb1 = Nothing
Set ws1 = Nothing 
Set ws2 = Nothing
Set ds = Nothing
Set cl1 = Nothing
Set cl2 = Nothing

'******�����������Z�b�g�i�I���j**************
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xCalculationAutomatic
'******************************************
''******���Ԍv�������**************
'    fin_time = Timer
' MsgBox "�o�ߎ��ԁF" & fin_time - start_time
''***********************************

'�����̐���\��
MsgBox dc - ros & "�̍��������o����܂����B�����͊e�V�[�g�ɐF�t���R�����g�L�q�A����сu�����v �V�[�g�Ɉꗗ�\�����܂��B"
If dc > 65531 Then MsgBox "Excel�̎d�l�ɂ��A65530���ȍ~�̍����\�̃V�[�g�����N�͕t������Ă��܂���B"

End Sub


Public Function ShowSelect SheetDialog()
'// �V�[�g�I���_�C�A���O��\�� 
'// �߂�l: �I�����ꂽ�V�[�g�I�u�W�F�N�g

    Dim shsele As Worksheet
    Application.ScreenUpdating = False 
    Set shsele = ActiveSheet
    With CommandBars.Add(Temporary:=True) 
        .Controls.Add(ID:=957)
        .Execute
        .Delete
    End With
    Set ShowSelectSheetDialog = ActiveSheet
    shsele.Select
    Application.ScreenUpdating = True

End Function


Sub ���ʑO����(wb1 as Workbook, motob as Workbook, sakib as Workbook, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, cl1 as Range, cl2 as Range, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

'���O�ɔ�r����V�[�g���폜
For Each ws1 In wbl.Sheets
    If ws1.Name = "����" Or ws1.Name = "����" Then
        rtn = MsgBox("��r�V�[�g�Ɠ������O�̃V�[�g������܂��B�폜���܂���?", vbYesNo + vbQuestion + vbDefaultButton2,"�폜or�R�s�[")
            Application.DisplayAlerts = False '�A���[�g OFF 
                Select Case rtn '�����ꂽ�{�^���̊m�F
                Case vbYes
                    ws1.Delete
                Case vbNo
                    ws1.Activate
                    ActiveSheet.Copy after:=Sheets(Sheets.Count) '�R�s�[�쐬
                    ws1.Delete
                End Select
            Application.DisplayAlerts = True '�A���[�g ON
    End If 
Next ws1

'���O�ɔ�r��\������V�[�g���폜
For Each ds In wbl.Sheets
    If ds. Name Then
        rtn = MsgBox("���łɍ����V�[�g������܂��B�폜���܂���?", vbYesNo + vbQuestion + vbDefaultButton2,"�폜or�R�s�[")
            Application.DisplayAlerts = False '�A���[�g OFF 
                Select Case rtn'�����ꂽ�{�^���̊m�F
                Case vbYes
                    ds. Delete
                Case vbNo
                    ds. Activate 
                    ActiveSheet. Copy after:=Sheets(Sheets.Count) '�R�s�[�쐬
                    ds. Delete
                End Select
            Application.DisplayAlerts = True '�A���[�g ON
    End If 
Next ds

'���ƃV�[�g�ɔ�r���V�[�g�R�s�[
'��r���t�@�C���w��̂��߂̃_�C�A���O�\��
With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "��r�� �u���Ɓv �t�@�C����I�����Ă�������"
    .Filters.Clear
    .Filters.Add Description:="Excel�t�@�C���h, Extensions:="*.xlsx"
    .Filters.Add Description:="Excel�}�N���L���h, Extensions:="*.xlsm"
    .Filters.Add Description:="CSV�t�@�C��", Extensions:="*.csv"
    .InitialFileName = wb1.Path & "\"
    .AllowMultiSelect = False
    If .Show = True Then
        mt = .SelectedItems (1)
    Else
        MsgBox "�������͂���܂���ł����B�I�����܂��B"
        End
    End If
End With

'���̓t�@�C����ǂݎ���p�ŊJ���I�u�W�F�N�g�� 
Application.DisplayAlerts = False '�A���[�gOFF
Set motob Workbooks.Open(Filename:=mt, Update Links:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
Application.DisplayAlerts = True '�A���[�g ON
    If motob Is Nothing Then
        MsgBox "�t�@�C�����J���܂���B�I�����܂�"
        Exit Sub
    End If
    motob. Activate

Call �V�[�g�ꗗ��z��Ɋi�[(arr)

'�V�[�g�I���_�C�A���O�\��
If Not UBound(arr) = 1 Then '�z��Ɋi�[���ꂽ�V�[�g����1�łȂ���Έȉ����s
    Dim Sh As Worksheet '�V�[�g�I���I�u�W�F�N�g��
    Set Sh ShowSelectSheetDialog() '�V�[�g�I�� Function�Ăяo��
    If Not Sh Is Nothing Then
        MsgBox Sh.Name & "���I������܂����B�V�[�g���R�s�[���܂�", vbInformation
    Sh.Activate
    Else
        MsgBox "�L�����Z������܂���", vbExclamation
        Exit Sub
    End If
    sn1 = Sh.Name '�I�������V�[�g���i�[ 
    Set Sh Nothing '�I�u�W�F�N�g�J��
Else
    sn1 = arr(1) '�z��1�ڂ̒l���
End If
DoEvents

Application.DisplayAlerts = False '�A���[�g OFF
motob. Worksheets(sn1). Copy before:=wb1.Sheets(1) '�V�[�g�R�s�[ 
Application.DisplayAlerts = True '�A���[�gON

ActiveSheet Name = "����" 
Set ws1 wb1.Worksheets("����")

' �����V�[�g�ɔ�r��V�[�g�R�s�[
'��r��t�@�C���w��̂��߂̃_�C�A���O�\��
With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "��r�� �u�����v �t�@�C����I�����Ă�������"
    .Filters.Clear
    .Filters.Add Description:="Excel�t�@�C���h, Extensions:="*.xlsx"
    .Filters.Add Description:="Excel�}�N���L���h, Extensions:="*.xlsm"
    .Filters.Add Description:="CSV�t�@�C��", Extensions:="*.csv"
    .InitialFileName = wb1.Path & "\"
    .AllowMultiSelect = False
    If .Show = True Then
        sk = .SelectedItems (1)
    Else
        MsgBox "�������͂���܂���ł����B�I�����܂��B"
        End
    End If
End With
' ���̓t�@�C����ǂݎ���p�ŊJ���I�u�W�F�N�g��
Application.DisplayAlerts = False
Set sakib Workbooks.Open(Filename:=sk, Update Links:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
Application.DisplayAlerts = True '�A���[�g ON 

' �����u�b�N��Excel�̃G���[���Ńt�@�C�����J���Ȃ��ꍇ�I��
    If sakib Is Nothing Then
        MsgBox "�t�@�C�����J���܂���B�I�����܂��B"
        Exit Sub
    End If
sakib. Activate


'�t�@�C�����������̂��t���p�X�Ŋm�F���A�قȂ�Γ����V�[�g�����m�F����
If Not motob Is sakib Then
'�������O�̃V�[�g������
    Dim A 
    For Each A In sakib.Sheets
'�V�[�g���������ꍇ�A���ƃV�[�g�Ɠ����V�[�g����I�� 
        If A.Name sn1 Then
            sn2 = sn1
        End If
    Next A
Else
'�����V�[�g�����Ȃ��ꍇ
    If sn2 = "" Then 
'�V�[�g�I���_�C�A���O�\��
    Dim SS As Worksheet '�V�[�g�I���I�u�W�F�N�g�� 
    Set SS ShowSelectSheetDialog() '�V�[�g�I�� Function�Ăяo��
        If Not SS Is Nothing Then MsgBox SS. Name & "���I������܂����B�V�[�g���R�s�[���܂�", vbInformation
            SS. Activate
        Else
            MsgBox "�L�����Z������܂���", vbExclamation 
            Exit Sub
        End If
    sn2 = SS. Name '�I�������V�[�g���i�[ 
        Set SS = Nothing '�I�u�W�F�N�g�J��
    End If 
End If

    Application.DisplayAlerts = False '�A���[�g OFF 
        sakib.Worksheets(sn2).Copy after:=wb1.Sheets("����") '���ƃV�[�g�̌��ɃR�s�[
    Application.DisplayAlerts = True ' �A���[�gON
        ActiveSheet.Name = "����"
    Set ws2 = wb1.Worksheets("����")

'��r���ʂ�\������V�[�g���쐬
    Worksheets.Add.Name = "����"
    ActiveSheet.Move before:=Sheets (1) '��ԍ��� 
    Set ds = wb1.Worksheets("����")
        ds.Activate

End Sub


Sub �����N�폜(sakib)

    Dim strLinks As Variant
    Dim k As Long
    Dim Ln As Long

    strLinks = ActiveWorkbook.LinkSources(Type:=xLinkTypeExcelLinks)
    If IsArray(strLinks) Then
        For k = 1 To UBound(strLinks)
            On Error Resume Next
            ActiveWorkbook.ChangeLink _
                Name:=strLinks(k), _
                NewName:=sakib.Name,
                Type:=xLinkTypeExcelLinks
            On Error GoTo 0
        Next k
    End If
    If IsArray(strLinks) Then Erase strLinks

End Sub


Sub �����t�@�C�����l�[��(sk, motob, sakib, f1, p1, beforecpysakib)

    Dim FSO As Object
    Dim fileFullPath As String
    Dim copyFileFullPath As String

    Set beforecpysakib = sk

        fileFullPath = sk
        copyFileFullPath = p1 & "���R�s�[��" & f1

        Set FSO = CreateObject("Scripting.FileSystemObject")

        Call FSO.CopyFile(Source:=fileFullPath, _
                          Destination:=copyFileFullPath, _
                          OverWriteFiles:=False)
        Set FSO = Nothing

    Application.DisplayAlerts = False
    Set sakib = Workbooks.Open(Filename:=copyFileFullPath, UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    Application.DisplayAlerts = True

    If sakib Is Nothing Then
        MsgBox "�t�@�C�����J���܂���B�I�����܂��B"
        End
    End If

End Sub


Sub �V�[�g�ꗗ��z��Ɋi�[(arr)

    ReDim arr(1 To Sheets.Count)
    Dim p
    For p = 1 To Sheets.Count
        arr(p) = Sheets(p).Name
    Next p

End Sub


Sub �͈͂ɕϊ�����(rtn2, cl1 as Range, cl2 as Range, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, dc, lis, wb1)

    Dim ws As Worksheet
    For Each ws In Workbooks
        If ws.ListObjects.Count > 0 Then
            rtn2 = MsgBox("�V�[�g�Ƀe�[�u�����܂܂�Ă��܂��B�������x���Ȃ邱�Ƃ�����܂��B�����܂����H", _
                    vbOKCancel + vbQuestion + vbDefaultButton1)
            Select Case rtn2
                Case vbOK
                    wb1.Activate
                    Call �Z�����ڏ���(cl1, cl2, ws1, ws2, ds, dc, lis)
                    Exit Sub
                Case vbOKCancel
                    MsgBox "�L�����Z������܂����B�I�����܂��B", vbExclamation
                    End
            End Select
        End If
    Next

End Sub

Sub �Z�����ڏ���(cl1 as Range, cl2 as Range, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, dc, lis)

' ���ƃV�[�g�̃Z�������[�v����
For Each cl1 In ws1.UsedRange
'�����V�[�g�̑Ή�����Z�����擾
    Set cl2 = ws2.Range (cl1.Address)
 '�Z���̒l���r
    If cl1.Formula <> cl2.Formula Then
      cl1.Interior.Color = rgbYellow '���F�h��Ԃ�
      cl2.Interior.Color = rgbGold '�S�[���h�h��Ԃ�
    '�����ڍׂ��o��
        ds.Cells(dc, lis).Offset(1, 0).Formula = Cells(cl1.Row, cl1.Column).Address(False, False) 
    On Error GoTo errorlabel
        ds.Cells(dc, lis).Offset(1, 1).Formula = cl1.Formula
        ds.Cells(dc, lis).Offset(1, 2).Formula = cl2.Formula
    '�����N��t���i65,530���ȉ��̏ꍇ�̂݁j
        If dc < 65531 Then
            ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc + 1, lis), Address:="", SubAddress:=" '����'!" & Cells(cl1.Row, cl1.Column).Address(False, False)
        End If
    '�����V�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
        ws1.Cells(cl1.Row, cl1.Column).AddComment
        ws1.Cells(cl1.Row, cl1.Column).Comment.Text Text:="<����>: " & cl2.Formula & vbCrLf & "<����>: " & cl1.Formula
    '���ƃV�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
        ws2.Cells(cl2.Row, cl2.Column).AddComment
        ws2.Cells(cl2.Row, cl2.Column).Comment.Text Text:="<����>: " & cl1.Formula & vbCrLf & "<����>: " & cl2.Formula
    dc = dc + 1
    Application.StatusBar = dc & "�s�ڂ̏��������Ă��܂�..."
    End If
Next cl1
    Application.StatusBar = False
Exit Sub '�G���[������΂����ŗ��E

'�G���[���͂����܂ŃX�L�b�v
errorlabel:
    Select Case err.Number
        Case 1004
            With ds.Cells(dc, lis).Offset(1, 1)
                .NumberFormatLocal = "@"
                .Formula = cl1.Formula
            End With
            With ds.Cells(dc, lis).Offset(1, 2)
                .NumberFormatLocal = "@"
                .Formula = cl2.Formula
            End With
            err.Clear
        Case Else
            MsgBox "�\�����ʃG���[���������܂���", vbInformation
    End Select
Resume Next '�G���[�������������̃R�[�h���珈���p��

End Sub


Sub �\�̐��`(ds as Worksheet, lis, ros)

    ds.Activate
        Range(Columns(lis + 1), Columns (lis + 2)).ColumnWidth = 30 '�񕝐ݒ�
        Range(Cells(ros, lis), Cells(ros, lis + 2)).Interior.Color = rgbYellow '�C���f�b�N�X���F�h��Ԃ�
        Cells(ros, lis +2).Interior.Color = rgbGold
        Range(Cells(ros - 2, lis + 1), Cells(ros - 2, lis + 2)).WrapText = True '�܂�Ԃ��ĕ\��

End Sub


Sub �l��r�ł̍��ق̖����ƃ��X�g��()

Dim motob As Workbook
Dim wbl As Workbook
Dim sakib As Workbook
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ds As Worksheet
Dim cll As Range
Dim c12 As Range
Dim snl As String 
Dim sn2 As String
Dim dc As Long
Dim rtn As Long
Dim rtn2 As Long
Dim i As Long
Dim j As Long
Dim arr()
Dim mt As Variant
Dim sk As Variant
Dim f1
Dim f2
Dim p1
Dim p2

Const ros As Long = 3 '�����\�X�^�[�g�s�i�[�L�q�͂��܂�s��
Const lis As Long = 2 '�����\�X�^�[�g��i�[�L�q�͂��܂��

Set wb1 = ThisWorkbook '���̃u�b�N

Call ���ʑO����(wb1, motob, sakib, ws1, ws2, ds, cl1, cl2, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

'******�����������Z�b�g�i�J�n�j**************
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xCalculationManual
'******************************************

'�C���f�b�N�X�̕\��
With ds.Cells(ros, lis)
    .Offset(-2, 0) = "�u�b�N"
    .Offset(-1, 0) =  "�V�[�g"
    .Offset(-2, 1) = mt
    .Offset(-2, 2) = sk
    .Offset(-1, 1) = sn1 '�V�[�g���\��
    .Offset(-1, 2) = sn2 
    .Offset(0, 0) = "�ΏۃZ��"
    .Offset(0, 1) = "���Ƃ̎�"
    .Offset(0, 2) = "�����̎�"
End With

ds.Activate
Call �\�̐��^(ds, lis, ros)

dc = ros

' ���ƃV�[�g�̃Z�������[�v����
For Each cl1 In ws1.UsedRange
'�����V�[�g�̑Ή�����Z�����擾
    Set cl2 = ws2.Range (cl1.Address)

On Error Resume Next
 '�Z���̒l���r
    If cl1.Value <> cl2.Value Then
      cl1.Interior.Color = rgbYellow '���F�h��Ԃ�
      cl2.Interior.Color = rgbGold '�S�[���h�h��Ԃ�
    '�����ڍׂ��o��
        ds.Cells(dc, lis).Offset(1, 0).Value = Cells(cl1.Row, cl1.Column).Address(False, False) 
        ds.Cells(dc, lis).Offset(1, 1).Value = cl1.Value 
        ds.Cells(dc, lis).Offset(1, 2).Value = cl2.Value
    '�����N��t���i65,530���ȉ��̏ꍇ�̂݁j
        If dc < 65531 Then
            ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc + 1, lis), Address:="", SubAddress:=" '����'!" & Cells(cl1.Row, cl1.Column).Address(False, False)
        End If
    '�����V�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
        ws1.Cells(cl1.Row, cl1.Column).AddComment
        ws1.Cells(cl1.Row, cl1.Column).Comment.Text Text:="<����>: " & cl2 & vbCrLf & "<����>: " & cl1
    '���ƃV�[�g�ƒl���قȂ�ꍇ�̍����R�����g�\�� 
        ws2.Cells(cl2.Row, cl2.Column).AddComment
        ws2.Cells(cl2.Row, cl2.Column).Comment.Text Text:="<����>: " & cl1 & vbCrLf & "<����>: " & cl2
    dc = dc + 1
    Application.StatusBar = dc & "�s�ڂ̏��������Ă��܂�..."
    End If
Next cl1
    Application.StatusBar = False
On Error GoTo 0

'������r�\�̌r���̋L��
    ds.Cells(ros, lis).CurrentRegion.Borders.LineStyle = xlContinuous

'�����t�@�C���R�s�[���{�̏ꍇ�R�s�[�����t�@�C������ăt�H���_���J���i�蓮�폜�j
If sakib.Name Like "���R�s�[��*" Then
'�R�s�[�t�@�C���̎蓮�폜�\��
    MsgBox "�t�@�C�����������̂��߃��l�[���R�s�[���܂����B�Y���t�H���_���J���܂��B���R�s�[���Ŏn�܂閼�O�̃t�@�C���͕K�v�Ȃ��̂ō폜���Ă��������B"

    Dim pth
    pth = Left(p1, Len(p1) - 1)
    Shell "C:\windows\explorer.exe " & pth & "\", vbNormalFocus  '�Y���t�H���_���J��

End If
ws1.Activate '���ƃV�[�g���A�N�e�B�u��

'��r���Ƃ����u�b�N�����
    On Error Resume Next
    Application.DisplayAlerts = False
    sakib.Close savechanges:=False
    motob.Close savechanges:=False
    Application.DisplayAlerts = True

Set wb1 = Nothing
Set ws1 = Nothing 
Set ws2 = Nothing
Set cl1 = Nothing
Set cl2 = Nothing
Erase arr

'******�����������Z�b�g�i�I���j**************
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xCalculationAutomatic
'******************************************

'�����̐���\��
MsgBox dc - ros & "�̍��������o����܂����B�����͊e�V�[�g�ɐF�t���R�����g�L�q�A����сu�����v �V�[�g�Ɉꗗ�\�����܂��B"

End Sub