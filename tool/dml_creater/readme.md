```
Option Base 0
Option Explicit

Sub BuildDML()
Dim x As Integer
Dim y As Integer

    Dim tableName As String
    tableName = Cells(1, 2).value
    
    Dim fileName As String
    fileName = Cells(2, 2).value
    
    ' 確認ダイアログ----------------------------------
    Dim ret As Integer
    ret = MsgBox(fileName & "を作成しますか？", vbYesNo)
    
    If ret = vbNo Then
        Exit Sub
    End If
    
    '項目名取得--------------------------------------
    Dim columnCount As Integer
    Dim columnName() As String

    columnCount = Cells(4, Columns.Count).End(xlToLeft).Column
    ReDim columnName(columnCount)
    For x = 1 To columnCount
        columnName(x) = Cells(4, x).value
    Next x
    
    '項目名をDML化-----------------------------------
    Dim outputStr As String
    outputStr = outputStr + "INSERT INTO " & tableName & "(" & Chr(10)
    For x = 1 To columnCount
        outputStr = outputStr & " " & columnName(x)
        If x = columnCount Then
            outputStr = outputStr & Chr(10)
        Else
            outputStr = outputStr & "," & Chr(10)
        End If
    Next x
    outputStr = outputStr & ") VALUES" & Chr(10)
    
    '値をDML化--------------------------------------
    Dim rowCount As Integer
    
    rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
    For y = 5 To rowCount
        outputStr = outputStr & "(" & Chr(10)
        For x = 1 To columnCount
            outputStr = outputStr & " "
            Dim value
            Dim formatLocal
            value = Cells(y, x)
            formatLocal = Cells(y, x).NumberFormatLocal
            Select Case VarType(value)
            Case 2, 5 ' Integer, Double
                If formatLocal = "@" Then
                    outputStr = outputStr & "'" & value & "'"
                Else
                    outputStr = outputStr & value
                End If
            Case 11 'Boolean
                If value Then
                    outputStr = outputStr & "true"
                Else
                    outputStr = outputStr & "false"
                End If
            Case Else
                If LCase(value) = "null" Then
                    outputStr = outputStr & "null"
                Else
                    Dim convertValue As String
                    convertValue = Replace(value, "'", "''")
                    If InStr(convertValue, vbLf) > 0 Then
                        outputStr = outputStr & "E'" & Replace(convertValue, vbLf, "n") & "'"""
                    Else
                        outputStr = outputStr & "'" & convertValue & "'"
                    End If
                End If
            End Select
            If x = columnCount Then
                outputStr = outputStr & Chr(10)
            Else
                outputStr = outputStr & "," & Chr(10)
            End If
        Next x
        If y = rowCount Then
            outputStr = outputStr & ");" & Chr(10)
        Else
            outputStr = outputStr & ")," & Chr(10)
        End If
    Next y
    
    OutputFile fileName, outputStr
    
    MsgBox (fileName & "を作成しました")
End Sub

' ファイル出力

Sub OutputFile(fileName As String, outputStr As String)
Dim fileAccessGranted As Boolean
Dim filePermissionCandidates
Dim filePath As String
    filePath = "/Users/oyamakohei/Downloads/" + fileName
    filePermissionCandidates = Array(filePath)
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    If Not fileAccessGranted Then
        MsgBox ("アクセスを許可してください")
        Return
    End If
    
    If Dir(filePath) <> "" Then
        Kill filePath
    End If
    
    Open filePath For Binary Access Write As #1
    PutUTF8String 1, outputStr
    Close #1
    
End Sub

' 文字列をUTF-8でPutする
' 引数：
'   fileNum：Openステートメントで指定したファイル番号
'   str：出力する文字列
' 備考
'   ファイルはOpen〜For Binaryで開かれていること

Sub PutUTF8String(ByVal fileNum As Integer, ByRef str As String)
    Dim byteUTF8() As Byte  ' 文字をUTF-8エンコードしたものを格納する
    Dim c, d As String  ' 変換する文字
    Dim i, w, v As Integer  ' ループカウンタと文字コード取得用
    Dim u As Long  ' Unicode化した文字コード
    
    i = 1
    Do While i <= Len(str)
        c = Mid(str, i, 1)
        w = AscW(c)
        
        If w >= &HD800 And w < &HDBFF Then
            ' サロゲートペア上位の場合はカウントを進めて下位も取得する
            i = i + 1
            d = Mid(str, i, 1)
            v = AscW(d)
            
             ' サロゲートペアのデコード
             u = &H10000 + ((w And &HFFFF&) - &HD800&) * &H400& + ((v And &HFFFF&) - &HDC00&)
        Else
            ' 符号ありで表現された文字コードを符号なし表現へ
            u = w And &HFFFF&
        End If
        
        byteUTF8() = Unicode2UTF8(u)
        
        Put #fileNum, , byteUTF8
        
        i = i + 1
    Loop
End Sub

' UnicodeをUTF-8にエンコードする
' 引数：
'    u：Unicode文字コード
'戻り値：
'    UTF-8エンコードした文字のByte列

Function Unicode2UTF8(u As Long) As Byte()
    Dim byteUTF8() As Byte
    
    Select Case u
        Case Is < &H80&
            ReDim byteUTF8(0)
            byteUTF8(0) = CByte(u)
        Case Is < &H800&
            ReDim byteUTF8(1)
            byteUTF8(0) = CByte(((u And &H7F0&) / 64) + 192)
            byteUTF8(1) = CByte((u And &H3F&) + 128)
        Case Is < &H10000
            ReDim byteUTF8(2)
            byteUTF8(0) = CByte(((u And &HF000&) / 4096) + 224)
            byteUTF8(1) = CByte(((u And &HFC0&) / 64) + 128)
            byteUTF8(2) = CByte((u And &H3F&) + 128)
        Case Is < &H200000
            ReDim byteUTF8(3)
            byteUTF8(0) = CByte(((u And &H1C0000) / 262144) + 240)
            byteUTF8(1) = CByte(((u And &H3F000) / 4096) + 128)
            byteUTF8(2) = CByte(((u And &HFC0&) / 64) + 128)
            byteUTF8(3) = CByte((u And &H3F&) + 128)
        Case Else
            ' UTF-8で5バイト以上になる範囲はエラーの代わりに1バイトの0を返す
            ReDim byteUTF8(0)
            byteUTF8(0) = 0
    End Select
    
    Unicode2UTF8 = byteUTF8()
End Function

```