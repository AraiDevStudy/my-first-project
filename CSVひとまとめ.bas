Attribute VB_Name = "Module1"

Sub JointSheets()
    Dim buf As String
    Dim xlsxFile As String
    Dim newWbName As String
    Dim tmpSheet As Object
   
'    新しいワークブックを作成、新しいWBがアクティブになる
    Workbooks.Add
    newWbName = ActiveWorkbook.Name
   
   
'    指定したディレクトリ内にあるxlsxファイルをすべて列挙
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            buf = .SelectedItems(1)
            xlsxFile = Dir(buf & "\*.csv")
        End If
    End With
   
'    xlsxファイルの数だけループ
    Do While xlsxFile <> ""
   
        '    csvファイルを開く
        Workbooks.Open buf & "\" & xlsxFile
       
        '    ブックを不可視にする
        ActiveWindow.Visible = False
       
        '    ブック内のシートの数だけループ
        For Each tmpSheet In Workbooks(xlsxFile).Sheets
           
            '   列幅を自動調節する
            tmpSheet.Columns.EntireColumn.AutoFit
       
            '    新しいブックに開いたブックのシートをコピーする
            tmpSheet.Copy After:=Workbooks(newWbName).Sheets(1)
        Next tmpSheet
       
        '    csvファイルを閉じる
        Workbooks(xlsxFile).Close SaveChanges:=False
       
        '    次のxlsxファイルに移る
        xlsxFile = Dir()
    Loop
   
    'Sheet1は削除する
    Application.DisplayAlerts = False ' メッセージを非表示
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True  ' メッセージを表示
   
End Sub

Sub JointSheets2(buf As String)
   
    Dim xlsxFile As String
    Dim newWbName As String
    Dim tmpSheet As Object
   
'    新しいワークブックを作成、新しいWBがアクティブになる
    Workbooks.Add
    newWbName = ActiveWorkbook.Name
   
   
'    指定したディレクトリ内にあるxlsxファイルをすべて列挙
    xlsxFile = Dir(buf & "\*.csv")
   
'    xlsxファイルの数だけループ
    Do While xlsxFile <> ""
   
        '    csvファイルを開く
        Workbooks.Open buf & "\" & xlsxFile
       
        '    ブックを不可視にする
        ActiveWindow.Visible = False
       
        '    ブック内のシートの数だけループ
        For Each tmpSheet In Workbooks(xlsxFile).Sheets
           
            '   列幅を自動調節する
            tmpSheet.Columns.EntireColumn.AutoFit
       
            '    新しいブックに開いたブックのシートをコピーする
            tmpSheet.Copy After:=Workbooks(newWbName).Sheets(1)
        Next tmpSheet
       
        '    csvファイルを閉じる
        Workbooks(xlsxFile).Close SaveChanges:=False
       
        '    次のxlsxファイルに移る
        xlsxFile = Dir()
    Loop
   
    'Sheet1は削除する
    Application.DisplayAlerts = False ' メッセージを非表示
    If Sheets.Count > 1 Then
      Sheets("Sheet1").Delete
    End If
    Application.DisplayAlerts = True  ' メッセージを表示
   
    'ブックを保存する。
    Workbooks(newWbName).SaveAs buf & "\" & newWbName & ".xlsx"
    'ブックを閉じる
    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=True
   
End Sub


Sub 一括実行()
    Dim i As Integer
    Dim foldername(8) As String
    foldername(0) = "①"
    foldername(1) = "②"
    foldername(2) = "③"
    foldername(3) = "④"
    foldername(4) = "⑤"
    foldername(5) = "⑥"
    foldername(6) = "⑦"
    foldername(7) = "⑧"
   
   
    For i = 0 To 7
   
      JointSheets2 foldername(i)
   
    Next i
   
   
End Sub
