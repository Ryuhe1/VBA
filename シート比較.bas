Attribute VB_Name = "Module30" 
Option Explicit

Sub 数式比較での差異の明示とリスト化()

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

Const ros As Long = 3 '差分表スタート行格納記述はじまり行数
Const lis As Long = 2 '差分表スタート列格納記述はじまり列数

Set wb1 = ThisWorkbook 'このブック

Call 共通前処理(wb1, motob, sakib, ws1, ws2, ds, cl1, cl2, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

ws1.Activate
Call リンク削除(sakib)

''******時間計測はじまり**************
'    Dim start_time As Double
'    Dim fin_time As Double
'    start_time = Timer
''***********************************
'******処理高速化セット（開始）**************
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xCalculationManual
'******************************************

'インデックスの表示
With ds.Cells(ros, lis)
    .Offset(-2, 0) = "ブック"
    .Offset(-1, 0) = "シート"
    .Offset(-2, 1) = mt  'ファイル名表示
    .Offset(-2, 2) = sk
    .Offset(-1, 1) = sn1 'シート名表示
    .Offset(-1, 2) = sn2 
    .Offset(0, 0) = "対象セル"
    .Offset(0, 1) = "もとの式"
    .Offset(0, 2) = "さきの式"
End With

ds.Activate
Call 表の整形(ds, lis, ros)

dc = ros

'シートにテーブルがある場合の選択
Call 範囲に変換判定(rtn2, cl1, cl2, ws1, ws2, ds, dc, lis, wb1)

'★★テーブルでない場合に以下の処理
If rtn2 <> vbOK Then
'各シートの差異ありセルの判定処理
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
    arys1 = cl1.Formula  'もとシートの使用セル範囲数式を全て配列に代入
    arys2 = cl2.Formula  'さきシートの使用セル範囲数式を全て配列に代入

    '動的配列の１始まりに再定義
    ReDim arys3(1 To 2, 1 To dc)
    ReDim aryex(1 To 2, 1 To 2)

' もとシートのセルをループ処理
    For i = 1 To UBound(arys1, 1)
        For j = 1 To UBound(arys1, 2)
    '２つのシートの数式が異なる場合セルに色を付ける
            If arys1(i, j) <> arys2(i, j) Then
        'エラーが発生した場合は異なる数式なのでセルに色を付ける
                ws1.Cells(i, j).Offset(ursr, ursc).Interior.Color = rgbYellow
                ws2.Cells(i, j).Offset(ursr, ursc).Interior.Color = rgbGold
        '差分のあるセル番号を出力
                ds.Cells(dc, lis).Offset(1, 0).Formula = ws1.Cells(i, j).Offset(ursr, ursc).Address(False, False)
        'リンクを付加（65,530件以下の場合のみ）
                If dc < 65531 Then
                    ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc, lis).Offset(1), Address:="", _
                    SubAddress:=" 'もと'!" & ws1.Cells(i, j).Offset(ursr, ursc).Address(False, False)
                End If
        '各シートの差分を二次元配列に代入（差分比較表）
                ReDim Preserve arys3(1 To 2, 1 To dc)   '列次元のみ追加可能（配列の制限）
                    arys3(1, dc - ros + 1) = arys1(i, j)
                    arys3(2, dc - ros + 1) = arys2(i, j)
        'さぎシートと値が異なる場合の差分コメント表示 
                ws1.Cells(i, j).Offset(ursr, ursc).AddComment
                ws1.Cells(i, j).Offset(ursr, ursc).Comm
                ent.Text Text:="<さき>: " & arys2(i, j) & vbCrLf & "<もと>: " & arys1(i, j)
        'もとシートと値が異なる場合の差分コメント表示 
                ws2.Cells(i, j).Offset(ursr, ursc).AddComment
                ws2.Cells(i, j).Offset(ursr, ursc).Comment.Text Text:="<もと>: " & arys1(i, j) & vbCrLf & "<さき>: " & arys2(i, j)
            dc = dc + 1
            End If
        Next
    Application.StatusBar = dc & "行目の処理をしています..."
    Next
    Application.StatusBar = False
            
    Erase arr
    Erase arys1
    Erase arys2

    '差分比較表転記用配列の行列入れかえ
    Dim ro: Dim col
    ReDim aryex(1 To UBound(arys3, 2), 1 To UBound(arys3, 1))
    For ro = 1 To UBound(arys3, 1)
        For col = 1 To UBound(arys3, 2)
            aryex(col, ro) = arys3(ro, col)
        Next
    Next
    DoEvents
    '差分比較表に行列入れ替え済み配列を貼付け
    ds.Cells(ros + 1, lis + 1).Resize(UBound(aryex, 1), UBound(aryex, 2)).FormulaLocal = aryex

    Erase arys3
    Erase aryex
End If  
'★★テーブルでない場合の処理終わり

'差分比較表の罫線の記入
ds.Cells(ros, lis). CurrentRegion.Borders.LineStyle = xlContinuous
'数式表示に変更
ds.Activate
ActiveWindow.DisplayFormulas = True

'同名ファイルコピー実施の場合はコピーしたファイルを閉じてフォルダを開く（手動削除）
If sakib.Name Like "★コピー★*" Then
    MsgBox "ファイル名が同じのためリネームコピーしました。該当フォルダを開きます。★コピー★で始まる名前のファイルは必要ないので削除してください。"
    Dim pth
    pth = Left(p1, Len(p1) - 1)
    Shell "C:\windows\explorer.exe " & pth & "\", vbNormalFocus  '該当フォルダを開く
End If

ws1.Activate 'もとシートをアクティブ化

'比較もとさきブックを閉じる
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

'******処理高速化セット（終了）**************
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xCalculationAutomatic
'******************************************
''******時間計測おわり**************
'    fin_time = Timer
' MsgBox "経過時間：" & fin_time - start_time
''***********************************

'差分の数を表示
MsgBox dc - ros & "個の差分が検出されました。差分は各シートに色付きコメント記述、および「差分」 シートに一覧表示します。"
If dc > 65531 Then MsgBox "Excelの仕様により、65530件以降の差分表のシートリンクは付加されていません。"

End Sub


Public Function ShowSelect SheetDialog()
'// シート選択ダイアログを表示 
'// 戻り値: 選択されたシートオブジェクト

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


Sub 共通前処理(wb1 as Workbook, motob as Workbook, sakib as Workbook, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, cl1 as Range, cl2 as Range, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

'事前に比較するシートを削除
For Each ws1 In wbl.Sheets
    If ws1.Name = "もと" Or ws1.Name = "さき" Then
        rtn = MsgBox("比較シートと同じ名前のシートがあります。削除しますか?", vbYesNo + vbQuestion + vbDefaultButton2,"削除orコピー")
            Application.DisplayAlerts = False 'アラート OFF 
                Select Case rtn '押されたボタンの確認
                Case vbYes
                    ws1.Delete
                Case vbNo
                    ws1.Activate
                    ActiveSheet.Copy after:=Sheets(Sheets.Count) 'コピー作成
                    ws1.Delete
                End Select
            Application.DisplayAlerts = True 'アラート ON
    End If 
Next ws1

'事前に比較を表示するシートを削除
For Each ds In wbl.Sheets
    If ds. Name Then
        rtn = MsgBox("すでに差分シートがあります。削除しますか?", vbYesNo + vbQuestion + vbDefaultButton2,"削除orコピー")
            Application.DisplayAlerts = False 'アラート OFF 
                Select Case rtn'押されたボタンの確認
                Case vbYes
                    ds. Delete
                Case vbNo
                    ds. Activate 
                    ActiveSheet. Copy after:=Sheets(Sheets.Count) 'コピー作成
                    ds. Delete
                End Select
            Application.DisplayAlerts = True 'アラート ON
    End If 
Next ds

'もとシートに比較元シートコピー
'比較元ファイル指定のためのダイアログ表示
With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "比較元 「もと」 ファイルを選択してください"
    .Filters.Clear
    .Filters.Add Description:="Excelファイル”, Extensions:="*.xlsx"
    .Filters.Add Description:="Excelマクロ有効”, Extensions:="*.xlsm"
    .Filters.Add Description:="CSVファイル", Extensions:="*.csv"
    .InitialFileName = wb1.Path & "\"
    .AllowMultiSelect = False
    If .Show = True Then
        mt = .SelectedItems (1)
    Else
        MsgBox "何も入力されませんでした。終了します。"
        End
    End If
End With

'入力ファイルを読み取り専用で開きオブジェクト化 
Application.DisplayAlerts = False 'アラートOFF
Set motob Workbooks.Open(Filename:=mt, Update Links:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
Application.DisplayAlerts = True 'アラート ON
    If motob Is Nothing Then
        MsgBox "ファイルを開けません。終了します"
        Exit Sub
    End If
    motob. Activate

Call シート一覧を配列に格納(arr)

'シート選択ダイアログ表示
If Not UBound(arr) = 1 Then '配列に格納されたシート数が1でなければ以下実行
    Dim Sh As Worksheet 'シート選択オブジェクト化
    Set Sh ShowSelectSheetDialog() 'シート選択 Function呼び出し
    If Not Sh Is Nothing Then
        MsgBox Sh.Name & "が選択されました。シートをコピーします", vbInformation
    Sh.Activate
    Else
        MsgBox "キャンセルされました", vbExclamation
        Exit Sub
    End If
    sn1 = Sh.Name '選択したシート名格納 
    Set Sh Nothing 'オブジェクト開放
Else
    sn1 = arr(1) '配列1個目の値代入
End If
DoEvents

Application.DisplayAlerts = False 'アラート OFF
motob. Worksheets(sn1). Copy before:=wb1.Sheets(1) 'シートコピー 
Application.DisplayAlerts = True 'アラートON

ActiveSheet Name = "もと" 
Set ws1 wb1.Worksheets("もと")

' さきシートに比較先シートコピー
'比較先ファイル指定のためのダイアログ表示
With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "比較先 「さき」 ファイルを選択してください"
    .Filters.Clear
    .Filters.Add Description:="Excelファイル”, Extensions:="*.xlsx"
    .Filters.Add Description:="Excelマクロ有効”, Extensions:="*.xlsm"
    .Filters.Add Description:="CSVファイル", Extensions:="*.csv"
    .InitialFileName = wb1.Path & "\"
    .AllowMultiSelect = False
    If .Show = True Then
        sk = .SelectedItems (1)
    Else
        MsgBox "何も入力されませんでした。終了します。"
        End
    End If
End With
' 入力ファイルを読み取り専用で開きオブジェクト化
Application.DisplayAlerts = False
Set sakib Workbooks.Open(Filename:=sk, Update Links:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
Application.DisplayAlerts = True 'アラート ON 

' 同名ブックのExcelのエラー等でファイルが開かない場合終了
    If sakib Is Nothing Then
        MsgBox "ファイルを開けません。終了します。"
        Exit Sub
    End If
sakib. Activate


'ファイルが同じものかフルパスで確認し、異なれば同じシート名を確認する
If Not motob Is sakib Then
'同じ名前のシートを検索
    Dim A 
    For Each A In sakib.Sheets
'シート名が同じ場合、もとシートと同じシート名を選択 
        If A.Name sn1 Then
            sn2 = sn1
        End If
    Next A
Else
'同じシート名がない場合
    If sn2 = "" Then 
'シート選択ダイアログ表示
    Dim SS As Worksheet 'シート選択オブジェクト化 
    Set SS ShowSelectSheetDialog() 'シート選択 Function呼び出し
        If Not SS Is Nothing Then MsgBox SS. Name & "が選択されました。シートをコピーします", vbInformation
            SS. Activate
        Else
            MsgBox "キャンセルされました", vbExclamation 
            Exit Sub
        End If
    sn2 = SS. Name '選択したシート名格納 
        Set SS = Nothing 'オブジェクト開放
    End If 
End If

    Application.DisplayAlerts = False 'アラート OFF 
        sakib.Worksheets(sn2).Copy after:=wb1.Sheets("もと") 'もとシートの後ろにコピー
    Application.DisplayAlerts = True ' アラートON
        ActiveSheet.Name = "さき"
    Set ws2 = wb1.Worksheets("さき")

'比較結果を表示するシートを作成
    Worksheets.Add.Name = "差分"
    ActiveSheet.Move before:=Sheets (1) '一番左に 
    Set ds = wb1.Worksheets("差分")
        ds.Activate

End Sub


Sub リンク削除(sakib)

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


Sub 同名ファイルリネーム(sk, motob, sakib, f1, p1, beforecpysakib)

    Dim FSO As Object
    Dim fileFullPath As String
    Dim copyFileFullPath As String

    Set beforecpysakib = sk

        fileFullPath = sk
        copyFileFullPath = p1 & "★コピー★" & f1

        Set FSO = CreateObject("Scripting.FileSystemObject")

        Call FSO.CopyFile(Source:=fileFullPath, _
                          Destination:=copyFileFullPath, _
                          OverWriteFiles:=False)
        Set FSO = Nothing

    Application.DisplayAlerts = False
    Set sakib = Workbooks.Open(Filename:=copyFileFullPath, UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    Application.DisplayAlerts = True

    If sakib Is Nothing Then
        MsgBox "ファイルを開けません。終了します。"
        End
    End If

End Sub


Sub シート一覧を配列に格納(arr)

    ReDim arr(1 To Sheets.Count)
    Dim p
    For p = 1 To Sheets.Count
        arr(p) = Sheets(p).Name
    Next p

End Sub


Sub 範囲に変換判定(rtn2, cl1 as Range, cl2 as Range, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, dc, lis, wb1)

    Dim ws As Worksheet
    For Each ws In Workbooks
        If ws.ListObjects.Count > 0 Then
            rtn2 = MsgBox("シートにテーブルが含まれています。処理が遅くなることがあります。続けますか？", _
                    vbOKCancel + vbQuestion + vbDefaultButton1)
            Select Case rtn2
                Case vbOK
                    wb1.Activate
                    Call セル直接処理(cl1, cl2, ws1, ws2, ds, dc, lis)
                    Exit Sub
                Case vbOKCancel
                    MsgBox "キャンセルされました。終了します。", vbExclamation
                    End
            End Select
        End If
    Next

End Sub

Sub セル直接処理(cl1 as Range, cl2 as Range, ws1 as Worksheet, ws2 as Worksheet, ds as Worksheet, dc, lis)

' もとシートのセルをループ処理
For Each cl1 In ws1.UsedRange
'さきシートの対応するセルを取得
    Set cl2 = ws2.Range (cl1.Address)
 'セルの値を比較
    If cl1.Formula <> cl2.Formula Then
      cl1.Interior.Color = rgbYellow '黄色塗りつぶし
      cl2.Interior.Color = rgbGold 'ゴールド塗りつぶし
    '差分詳細を出力
        ds.Cells(dc, lis).Offset(1, 0).Formula = Cells(cl1.Row, cl1.Column).Address(False, False) 
    On Error GoTo errorlabel
        ds.Cells(dc, lis).Offset(1, 1).Formula = cl1.Formula
        ds.Cells(dc, lis).Offset(1, 2).Formula = cl2.Formula
    'リンクを付加（65,530件以下の場合のみ）
        If dc < 65531 Then
            ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc + 1, lis), Address:="", SubAddress:=" 'もと'!" & Cells(cl1.Row, cl1.Column).Address(False, False)
        End If
    'さぎシートと値が異なる場合の差分コメント表示 
        ws1.Cells(cl1.Row, cl1.Column).AddComment
        ws1.Cells(cl1.Row, cl1.Column).Comment.Text Text:="<さき>: " & cl2.Formula & vbCrLf & "<もと>: " & cl1.Formula
    'もとシートと値が異なる場合の差分コメント表示 
        ws2.Cells(cl2.Row, cl2.Column).AddComment
        ws2.Cells(cl2.Row, cl2.Column).Comment.Text Text:="<もと>: " & cl1.Formula & vbCrLf & "<さき>: " & cl2.Formula
    dc = dc + 1
    Application.StatusBar = dc & "行目の処理をしています..."
    End If
Next cl1
    Application.StatusBar = False
Exit Sub 'エラー無ければここで離脱

'エラー時はここまでスキップ
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
            MsgBox "予期せぬエラーが発生しました", vbInformation
    End Select
Resume Next 'エラーが発生した次のコードから処理継続

End Sub


Sub 表の整形(ds as Worksheet, lis, ros)

    ds.Activate
        Range(Columns(lis + 1), Columns (lis + 2)).ColumnWidth = 30 '列幅設定
        Range(Cells(ros, lis), Cells(ros, lis + 2)).Interior.Color = rgbYellow 'インデックス黄色塗りつぶし
        Cells(ros, lis +2).Interior.Color = rgbGold
        Range(Cells(ros - 2, lis + 1), Cells(ros - 2, lis + 2)).WrapText = True '折り返して表示

End Sub


Sub 値比較での差異の明示とリスト化()

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

Const ros As Long = 3 '差分表スタート行格納記述はじまり行数
Const lis As Long = 2 '差分表スタート列格納記述はじまり列数

Set wb1 = ThisWorkbook 'このブック

Call 共通前処理(wb1, motob, sakib, ws1, ws2, ds, cl1, cl2, dc, rtn, arr, sn1, sn2, beforecpysakib, mt, sk, p1, f1, p2, f2)

'******処理高速化セット（開始）**************
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xCalculationManual
'******************************************

'インデックスの表示
With ds.Cells(ros, lis)
    .Offset(-2, 0) = "ブック"
    .Offset(-1, 0) =  "シート"
    .Offset(-2, 1) = mt
    .Offset(-2, 2) = sk
    .Offset(-1, 1) = sn1 'シート名表示
    .Offset(-1, 2) = sn2 
    .Offset(0, 0) = "対象セル"
    .Offset(0, 1) = "もとの式"
    .Offset(0, 2) = "さきの式"
End With

ds.Activate
Call 表の成型(ds, lis, ros)

dc = ros

' もとシートのセルをループ処理
For Each cl1 In ws1.UsedRange
'さきシートの対応するセルを取得
    Set cl2 = ws2.Range (cl1.Address)

On Error Resume Next
 'セルの値を比較
    If cl1.Value <> cl2.Value Then
      cl1.Interior.Color = rgbYellow '黄色塗りつぶし
      cl2.Interior.Color = rgbGold 'ゴールド塗りつぶし
    '差分詳細を出力
        ds.Cells(dc, lis).Offset(1, 0).Value = Cells(cl1.Row, cl1.Column).Address(False, False) 
        ds.Cells(dc, lis).Offset(1, 1).Value = cl1.Value 
        ds.Cells(dc, lis).Offset(1, 2).Value = cl2.Value
    'リンクを付加（65,530件以下の場合のみ）
        If dc < 65531 Then
            ActiveSheet.Hyperlinks.Add Anchor:= Cells(dc + 1, lis), Address:="", SubAddress:=" 'もと'!" & Cells(cl1.Row, cl1.Column).Address(False, False)
        End If
    'さぎシートと値が異なる場合の差分コメント表示 
        ws1.Cells(cl1.Row, cl1.Column).AddComment
        ws1.Cells(cl1.Row, cl1.Column).Comment.Text Text:="<さき>: " & cl2 & vbCrLf & "<もと>: " & cl1
    'もとシートと値が異なる場合の差分コメント表示 
        ws2.Cells(cl2.Row, cl2.Column).AddComment
        ws2.Cells(cl2.Row, cl2.Column).Comment.Text Text:="<もと>: " & cl1 & vbCrLf & "<さき>: " & cl2
    dc = dc + 1
    Application.StatusBar = dc & "行目の処理をしています..."
    End If
Next cl1
    Application.StatusBar = False
On Error GoTo 0

'差分比較表の罫線の記入
    ds.Cells(ros, lis).CurrentRegion.Borders.LineStyle = xlContinuous

'同名ファイルコピー実施の場合コピーしたファイルを閉じてフォルダを開く（手動削除）
If sakib.Name Like "★コピー★*" Then
'コピーファイルの手動削除表示
    MsgBox "ファイル名が同じのためリネームコピーしました。該当フォルダを開きます。★コピー★で始まる名前のファイルは必要ないので削除してください。"

    Dim pth
    pth = Left(p1, Len(p1) - 1)
    Shell "C:\windows\explorer.exe " & pth & "\", vbNormalFocus  '該当フォルダを開く

End If
ws1.Activate 'もとシートをアクティブ化

'比較もとさきブックを閉じる
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

'******処理高速化セット（終了）**************
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xCalculationAutomatic
'******************************************

'差分の数を表示
MsgBox dc - ros & "個の差分が検出されました。差分は各シートに色付きコメント記述、および「差分」 シートに一覧表示します。"

End Sub