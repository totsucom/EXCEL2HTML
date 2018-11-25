'エクセルファイルをhtmlに頑張って変換する
'PHPから実行されます
'
'最初の引数　エクセルファイルのパス
'2個目の引数 オプション
'  -f   htmlヘッダも含めて出力
'
'エクセルファイルと同一フォルダにエクセルと同一ファイル名のcssファイルがが生成され
'コンソールにhtmlが出力される。ファイルに出力したい場合は > test.html 等のように出力先を変更する
'対象のシートはアクティブなシートが選択される

Imports NPOI.SS.UserModel
Imports NPOI.HSSF.UserModel
Imports NPOI.XSSF.UserModel
Imports System.Web

Module Module1
    Dim WB As IWorkbook = Nothing
    Dim WS As ISheet = Nothing
    Dim styles As New Hashtable

    '-fオプションで出力するテスト用のhtmlヘッダーとフッター
    Sub OutputHeader(styleSheetPath As String)
        Console.WriteLine("<!DOCTYPE html>")
        Console.WriteLine("<html>")
        Console.WriteLine("<head>")
        Console.WriteLine("<title>Excel2Html</title>")
        Console.WriteLine("<meta charset=""Shift_jis"">")
        If styleSheetPath.Length > 0 Then Console.WriteLine("<link rel = ""stylesheet"" href=""" & styleSheetPath & """>")
        Console.WriteLine("</head>")
        Console.WriteLine("<body>")
    End Sub
    Sub OutputFooter()
        Console.WriteLine("</body>")
        Console.WriteLine("</html>")
    End Sub

    Sub Main()

        Dim bHeader As Boolean = False
        For i As Integer = 1 To My.Application.CommandLineArgs.Count - 1
            Dim opt As String = My.Application.CommandLineArgs(i)
            Select Case opt
                Case "-f" : bHeader = True '-f オプション。HTML出力にヘッダを付ける
            End Select
        Next

        If My.Application.CommandLineArgs.Count = 0 Then
            If bHeader Then OutputHeader(String.Empty)
            Console.WriteLine("<p style=""background-color:red;color:white"">Excel2Html: 引数にエクセルファイルが指定されていません</p>")
            If bHeader Then OutputFooter()
            Exit Sub
        End If

        'エクセルパスを取得、存在チェック
        Dim excelPath As String = My.Application.CommandLineArgs(0)
        If Not IO.File.Exists(excelPath) Then
            If bHeader Then OutputHeader(String.Empty)
            Console.Write("<p style=""background-color:red;color:white"">Excel2Html: エクセルファイルが見つかりません<br>")
            Console.Write(HttpUtility.HtmlEncode(excelPath))
            Console.WriteLine("</p>")
            If bHeader Then OutputFooter()
            Exit Sub
        End If

        'CSSファイルパスを作成（有無を言わさず上書きする）
        Dim styleSheetPath As String = excelPath.Substring(0, excelPath.Length - IO.Path.GetExtension(excelPath).Length) & ".css"

        Try
            WB = WorkbookFactory.Create(excelPath) 'ブックを開く
            Debug.Print("Book opened")

            WS = WB.GetSheetAt(WB.ActiveSheetIndex) 'シート取得
            Debug.Print("WorkSheet opened")

            '使用しているカラムの範囲を取得
            Dim firstCellNum As Integer = Integer.MaxValue
            Dim lastCellNum As Integer = Integer.MinValue
            For rowIndex As Integer = WS.FirstRowNum To WS.LastRowNum
                Dim row As IRow = WS.GetRow(rowIndex) '行取得
                If row IsNot Nothing Then
                    If row.FirstCellNum < firstCellNum Then
                        firstCellNum = row.FirstCellNum
                    End If
                    If row.LastCellNum > lastCellNum Then
                        lastCellNum = row.LastCellNum
                    End If
                End If
            Next
            If firstCellNum > lastCellNum Then
                If bHeader Then OutputHeader(styleSheetPath)
                Console.Write("<p style=""background-color:red;color:white"">Excel2Html: シートが空です<br>")
                Console.Write(HttpUtility.HtmlEncode(WS.SheetName))
                Console.WriteLine("</p>")
                If bHeader Then OutputFooter()
                Exit Sub
            End If

            Dim df As New DataFormatter
            Dim crateHelper As ICreationHelper = WB.GetCreationHelper()
            Dim fe As IFormulaEvaluator = crateHelper.CreateFormulaEvaluator()

            Dim ignoreCells As New List(Of CellArea)

            'htmlのテーブルに変換
            If bHeader Then OutputHeader(styleSheetPath)
            Console.WriteLine("<table class=""excel-sheet"">")

            '適当にセル幅再現のためダミー行を入れる
            Console.Write("<thead><tr>")
            For colIndex As Integer = firstCellNum To lastCellNum
                Console.Write("<th style=""width:{0}px""></th>", (WS.GetColumnWidthInPixels(colIndex) * 1.3).ToString("F0"))
            Next
            Console.WriteLine("</tr></thead>")

            Console.WriteLine("<tbody>")
            For rowIndex As Integer = WS.FirstRowNum To WS.LastRowNum
                Dim row As IRow = WS.GetRow(rowIndex) '行取得

                If row IsNot Nothing Then
                    Dim bFirstCell As Boolean = True
                    Console.Write("<tr>")
                    For colIndex As Integer = firstCellNum To lastCellNum

                        '結合されたセルかどうか調べる
                        Dim ignore As Boolean = False
                        For Each ca As CellArea In ignoreCells
                            If ca.InRange(colIndex, rowIndex) Then
                                ignore = True
                                Exit For
                            End If
                        Next

                        If Not ignore Then
                            '結合されてないセルのみ処理
                            Dim cell As ICell = row.GetCell(colIndex)

                            Console.Write("<td")
                            If bFirstCell Then
                                '最初のセルには行の高さを適当に設定
                                Console.Write(" style=""height:{0}px""", (row.Height / 10.0).ToString("F0"))
                                bFirstCell = False
                            End If

                            If cell IsNot Nothing Then

                                Dim rightCell As ICell = Nothing
                                Dim bottomCell As ICell = Nothing

                                Dim columns As Integer = 0, rows As Integer = 0
                                If GetCellSize(cell, columns, rows) Then

                                    'このセルが結合されたセルだった場合
                                    If columns > 1 Then
                                        Console.Write(" colspan=""{0}""", columns)

                                        'スタイル参照用に右端のセルを取得
                                        rightCell = row.GetCell(colIndex + columns - 1)
                                    End If
                                    If rows > 1 Then
                                        Console.Write(" rowspan=""{0}""", rows)

                                        'スタイル参照用に下端のセルを取得
                                        Dim bottomRow As IRow = WS.GetRow(rowIndex + rows - 1)
                                        If bottomRow IsNot Nothing Then
                                            bottomCell = row.GetCell(colIndex)
                                        End If
                                    End If

                                    If columns > 1 OrElse rows > 1 Then
                                        'このセルが結合されたセルだった場合、このセル以外については
                                        '処理しないようにするため、ignoreCellsに結合セル範囲を記憶
                                        ignoreCells.Add(New CellArea(colIndex, rowIndex, columns, rows))
                                    End If
                                End If

                                'セルスタイルからCSSを生成、クラス名を設定
                                Dim rightCellStyle As ICellStyle
                                If rightCell Is Nothing Then
                                    rightCellStyle = cell.CellStyle
                                Else
                                    rightCellStyle = rightCell.CellStyle
                                End If
                                Dim bottomCellStyle As ICellStyle
                                If bottomCell Is Nothing Then
                                    bottomCellStyle = cell.CellStyle
                                Else
                                    bottomCellStyle = bottomCell.CellStyle
                                End If
                                Console.Write(" class=""{0}"">", CellStyle2ClassName(cell.CellStyle, rightCellStyle, bottomCellStyle, cell.CellType))

                                'パラメータをここで出力するとセル内に見えるのでデバッグに便利
                                'Console.Write("Index" & cell.CellStyle.Index.ToString() & "<br>")
                                'Console.Write("Indention" & cell.CellStyle.Indention.ToString() & "<br>")
                                'Console.Write("FontIndex" & cell.CellStyle.FontIndex.ToString() & "<br>")


                                'Console.Write(BuiltinFormats.GetBuiltinFormat(cell.CellStyle.DataFormat) & "<br>")

                                If (cell.CellType = CellType.Numeric OrElse cell.CellType = CellType.Formula) AndAlso IsContainDateFormat(cell.CellStyle.GetDataFormatString) Then
                                    '日付型はカルチャは効いているものの年月日順がおかしいので自前で行う
                                    FormatDateValue(cell.NumericCellValue, cell.CellStyle.GetDataFormatString)

                                Else
                                    '書式に応じたデフォルトの表示
                                    Console.Write(HttpUtility.HtmlEncode(df.FormatCellValue(cell, fe)))
                                End If

                            Else
                                Console.Write(">"c)
                            End If
                            Console.Write("</td>")
                        End If
                    Next
                    Console.WriteLine("</tr>")
                Else
                    'rowがNothngのとき
                    Dim bFirstCell As Boolean = True
                    Console.Write("<tr>")
                    For colIndex As Integer = firstCellNum To lastCellNum

                        '結合されたセルかどうか調べる
                        Dim ignore As Boolean = False
                        For Each ca As CellArea In ignoreCells
                            If ca.InRange(colIndex, rowIndex) Then
                                ignore = True
                                Exit For
                            End If
                        Next

                        If Not ignore Then
                            '結合されてないセルのみ処理

                            Console.Write("<td")
                            If bFirstCell Then
                                'rowが存在しないので行の高さを適当に設定
                                Console.Write(" style=""height:{0}px""", (270 / 10.0).ToString("F0"))
                                bFirstCell = False
                            End If
                            Console.Write("></td>")
                        End If
                    Next
                    Console.WriteLine("</tr>")
                End If
            Next
            Console.WriteLine("</tbody>")
            Console.WriteLine("</table>")

            'CSSファイル書き出し
            Try
                Using textFile As New IO.StreamWriter(styleSheetPath, False, Text.Encoding.Default)
                    textFile.WriteLine("table.excel-sheet { border-collapse: collapse; table-layout: fixed; }")
                    For Each className As String In styles.Keys
                        textFile.Write(".")
                        textFile.Write(className)
                        textFile.Write(" { ")
                        textFile.Write(styles(className))
                        textFile.WriteLine("}")
                    Next
                End Using
            Catch ex As Exception
                'CSS出力中の例外
                Console.WriteLine("<p style=""background-color:red;color:white"">Excel2Html: Exception CSSファイルの出力を失敗しました。<br>")
                Console.Write(HttpUtility.HtmlEncode(styleSheetPath))
                Console.WriteLine("<br>")
                Console.Write(HttpUtility.HtmlEncode(ex.Message))
                Console.WriteLine("</p>")
            End Try

        Catch ex As Exception
            'HTML作成中の例外
            Console.WriteLine("<p style=""background-color:red;color:white"">Excel2Html: Exception HTMLへの変換を失敗しました。<br>")
            Console.Write(excelPath)
            Console.WriteLine("<br>")
            Console.Write(HttpUtility.UrlEncode(ex.Message))
            Console.WriteLine("</p>")
        End Try

        If bHeader Then OutputFooter()
    End Sub

    'セルのフォーマット文字が日時を表すものか？
    Function IsContainDateFormat(format As String)
        Return format.IndexOf("yyyy") >= 0 OrElse format.IndexOf("mm") >= 0 OrElse format.IndexOf("dd") >= 0 OrElse format.IndexOf("hh") >= 0 OrElse format.IndexOf("ss") >= 0 OrElse format.IndexOf("aaa") >= 0
    End Function

    '日付フォーマット文字列を処理する
    Sub FormatDateValue(v As Double, format As String)
        ';区切りでフォーマットを分ける
        Dim formats As New List(Of String)
        Dim sb As New Text.StringBuilder
        Dim bRiteral As Boolean = False
        Dim bInKakko As Boolean = False
        For Each c As String In format
            If bRiteral Then
                sb.Append(c)
                bRiteral = False
            ElseIf c = "\" Then '←こいつのせいでsplit()は使えない
                sb.Append(c)
                bRiteral = True
            ElseIf c = "[" Then
                bInKakko = True
            ElseIf c = "]" Then
                bInKakko = False
            ElseIf c = ";" Then
                formats.Add(sb.ToString())
                sb.Clear()
            Else
                If Not bInKakko Then sb.Append(c)
            End If
        Next
        formats.Add(sb.ToString())
        For formatIndex As Integer = 0 To formats.Count - 1
            Dim f As String = formats(formatIndex)
            Dim bAMPM As Boolean = (f.IndexOf("AM/PM") >= 0)
            sb.Clear()
            Try
                Dim i0 As Integer = 0
                Dim i As Integer = f.IndexOf("\")
                While i >= 0
                    If i - i0 > 0 Then
                        sb.Append(_formatDateValue(v, f.Substring(i0, i - i0), bAMPM))
                    End If
                    If i + 1 < f.Length Then
                        sb.Append(f.Chars(i + 1))
                    End If
                    i0 = i + 2
                    i = f.IndexOf("\", i0)
                End While
                If f.Length > i0 Then
                    sb.Append(_formatDateValue(v, f.Substring(i0), bAMPM))
                End If
                '何事もなければ最初のフォーマットの結果を使用
                Console.Write(HttpUtility.HtmlEncode(sb.ToString()))
            Catch ex As Exception
            End Try
        Next
    End Sub
    Function _formatDateValue(v As Double, ByRef format As String, bAMPM As Boolean) As String
        If format.IndexOf("@"c) >= 0 OrElse format.IndexOf("General") >= 0 Then Return v.ToString()

        Dim d As Date = New Date(1899, 12, 30).AddDays(v)
        Dim f As String = format.Replace("mmmm", "MMMM") '〇月
        f = f.Replace("mm", "$$") '秒
        f = f.Replace("m", "M")
        f = f.Replace("$$", "mm")
        f = f.Replace("dddd", "$$$$")
        f = f.Replace("dd", "d日")
        f = f.Replace("$$$$", "dddd") '〇曜日
        f = f.Replace("yyyy", "$$$$")
        f = f.Replace("yy", "yyyy")
        f = f.Replace("$$$$", "yyyy年")
        If bAMPM Then
            f = f.Replace("AM/PM", "tt") '午前午後
        Else
            f = f.Replace("h", "H") '24h
        End If
        'Console.Write("""{0}""", f)
        Dim x As String = d.ToString(f)
        Return x
    End Function

    'セルのサイズを計算する
    '結合されていたら１以上を返す
    Function GetCellSize(cell As ICell, ByRef columns As Integer, ByRef rows As Integer) As Boolean
        Dim WS As ISheet = cell.Sheet
        For i As Integer = 0 To WS.NumMergedRegions - 1
            With WS.GetMergedRegion(i)
                If .FirstColumn = cell.ColumnIndex AndAlso .FirstRow = cell.RowIndex Then
                    columns = .LastColumn - .FirstColumn + 1
                    rows = .LastRow - .FirstRow + 1
                    Return True
                End If
            End With
        Next
        columns = 1
        rows = 1
        Return False
    End Function

    '与えられたセルスタイルからCSSを生成し、そのクラス名を返す
    'CSSはstylesに保存される
    Function CellStyle2ClassName(cs As ICellStyle, right_cs As ICellStyle, bottom_cs As ICellStyle, ct As CellType) As String

        'CSSを定義するクラス名を生成
        Dim styleClassName = "S"c & cs.Index.ToString()
        Dim fontClassName = "F"c & cs.FontIndex.ToString()

        If Not styles.Contains(styleClassName) Then
            '同名のスタイルを定義したクラス名が無いので作成
            Dim sb As New Text.StringBuilder

            'テキストアライメント
            sb.Append("text-align:")
            Dim ali As String = Alignment2StyleString(cs.Alignment)
            If ali.Length > 0 Then
                sb.Append(ali)
            ElseIf ct = CellType.Numeric OrElse ct = CellType.Formula Then
                sb.Append("right")
            Else
                sb.Append("left")
            End If
            sb.Append(";"c)
            sb.Append("vertical-align:")
            sb.Append(VerticalAlignment2StyleString(cs.VerticalAlignment))
            sb.Append(";"c)

            '罫線
            sb.Append("border-left:")
            sb.Append(BorderStyle2StyleString(cs.BorderLeft))
            sb.Append(" "c)
            sb.Append(ColorIndex2StyleString(cs.LeftBorderColor))
            sb.Append(";"c)
            sb.Append("border-top:")
            sb.Append(BorderStyle2StyleString(cs.BorderTop))
            sb.Append(" "c)
            sb.Append(ColorIndex2StyleString(cs.TopBorderColor))
            sb.Append(";"c)
            sb.Append("border-right:")
            sb.Append(BorderStyle2StyleString(right_cs.BorderRight))
            sb.Append(" "c)
            sb.Append(ColorIndex2StyleString(right_cs.RightBorderColor))
            sb.Append(";"c)
            sb.Append("border-bottom:")
            sb.Append(BorderStyle2StyleString(bottom_cs.BorderBottom))
            sb.Append(" "c)
            sb.Append(ColorIndex2StyleString(bottom_cs.BottomBorderColor))
            sb.Append(";"c)

            '折り返し
            If cs.WrapText Then sb.Append("overflow-wrap:break-word;")

            '背景色
            If cs.FillPattern = FillPattern.SolidForeground Then
                'べた塗りの場合はForegroundColorを使う
                sb.Append("background-color:#")
                sb.Append(BitConverter.ToString(cs.FillForegroundColorColor.RGB).Replace("-", String.Empty))
                sb.Append(";"c)
            Else
                'べた塗り以外は対応しない
            End If

            '生成したCSSをクラス名とともに保存
            styles.Add(styleClassName, sb.ToString())
        End If

        If Not styles.Contains(fontClassName) Then
            '同名のフォントを定義したクラス名が無いので作成
            Dim sb As New Text.StringBuilder

            '文字の大きさ。適当に大小だけ辻褄を合わせる
            sb.Append("font-size:")
            sb.Append((cs.GetFont(WB).FontHeight / 180).ToString("F1"))
            sb.Append("em;")

            With cs.GetFont(WB)
                '文字の装飾
                If .IsBold Then sb.Append("font-weight:bold;")
                If .IsItalic Then sb.Append("font-style:italic;")
                If .Underline OrElse .IsStrikeout Then
                    sb.Append("text-decoration:")
                    If .Underline Then sb.Append("underline")
                    If .IsStrikeout Then sb.Append("line-through")
                    sb.Append(";"c)
                End If

                '文字色
                sb.Append("color:")
                Dim col As String = ColorIndex2StyleString(.Color)
                If col.Length > 0 Then
                    sb.Append(col)
                Else
                    '不明
                    sb.Append("black")
                End If
                sb.Append(";"c)
            End With

            '生成したCSSをクラス名とともに保存
            styles.Add(fontClassName, sb.ToString())
        End If

        Return styleClassName & " "c & fontClassName
    End Function

    'カラーインデックスからCSS色 #000000 を返す
    Function ColorIndex2StyleString(index As Short) As String

        'カラーインデックスからRGB値を得るためのハッシュテーブルを取得
        Static hssfColorHash As Hashtable = Nothing
        If hssfColorHash Is Nothing Then
            hssfColorHash = NPOI.HSSF.Util.HSSFColor.GetIndexHash()
        End If

        Dim hssfColor As NPOI.HSSF.Util.HSSFColor = hssfColorHash(CInt(index)) '重要 CInt()
        If hssfColor IsNot Nothing Then '文字色
            Return "#"c & BitConverter.ToString(hssfColor.RGB).Replace("-", String.Empty)
        Else
            Return ""
        End If
    End Function

    '対応するテキストアライメントをCSS文字列で返す
    Function Alignment2StyleString(ha As HorizontalAlignment) As String
        Select Case ha
            Case HorizontalAlignment.General : Return ""
            Case HorizontalAlignment.Center : Return "center"
            Case HorizontalAlignment.CenterSelection : Return "center"
            Case HorizontalAlignment.Distributed : Return "match-parent"
            Case HorizontalAlignment.Fill : Return "justify-all"
            Case HorizontalAlignment.General : Return "start"
            Case HorizontalAlignment.Justify : Return "justify"
            Case HorizontalAlignment.Left : Return "left"
            Case HorizontalAlignment.Right : Return "right"
        End Select
        Return "start"
    End Function
    Function VerticalAlignment2StyleString(va As VerticalAlignment) As String
        Select Case va
            Case VerticalAlignment.Bottom : Return "bottom"
            Case VerticalAlignment.Center : Return "middle"
            Case VerticalAlignment.Top : Return "top"
        End Select
        Return "baseline"
    End Function

    '対応する罫線のスタイルをCSS文字列で返す
    Function BorderStyle2StyleString(bs As BorderStyle) As String
        Select Case bs
            Case BorderStyle.DashDot : Return "1px dashed"
            Case BorderStyle.DashDotDot : Return "1px dotted"
            Case BorderStyle.Dashed : Return "1px dashed"
            Case BorderStyle.Dotted : Return "1px dotted"
            Case BorderStyle.Double : Return "1px double"
            Case BorderStyle.Hair : Return "1px solid"
            Case BorderStyle.Medium : Return "2px solid"
            Case BorderStyle.MediumDashDot : Return "2px dashed"
            Case BorderStyle.MediumDashDotDot : Return "2px dotted"
            Case BorderStyle.MediumDashed : Return "2px dashed"
            Case BorderStyle.None : Return "none"
            Case BorderStyle.SlantedDashDot : Return "2px solid"
            Case BorderStyle.Thick : Return "3px solid"
            Case BorderStyle.Thick : Return "3px solid"
            Case BorderStyle.Thin : Return "1px solid"
        End Select
        Return "none"
    End Function

End Module

'結合されたセル範囲を格納するだけの簡単なクラス
'Rectangleクラスでもよかったんですが
Class CellArea
    Public column As Integer '0 or more
    Public row As Integer '0 or more
    Public width As Integer '1 or more
    Public height As Integer '1 or more
    Sub New()
        column = 0
        row = 0
        width = 0
        height = 0
    End Sub
    Sub New(c As Integer, r As Integer, w As Integer, h As Integer)
        column = c
        row = r
        width = w
        height = h
    End Sub
    Public Function InRange(c As Integer, r As Integer) As Boolean
        Return (c >= column AndAlso c < column + width AndAlso r >= row AndAlso r < row + height)
    End Function
End Class
