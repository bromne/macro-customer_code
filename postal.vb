Option Compare Database

Public Const FIND_NUM As String = "壱弐参〇一二三四五六七八九十百千"
Public Const FIND_TEXT As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Function CreateCustomerBarCode(strZip As String, ByVal strAddress As String) As String

    Dim strRetZip       As String
    Dim strRetAddress   As String
    Dim strChar         As String
    Dim blnConvFlg      As Boolean   'ハイフンに置き換え済みフラグ
    Dim intLoop         As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim intCount        As Integer
    Dim intPos          As Integer
    Dim intTextLength   As Integer
    Dim intTargetPos    As Integer
    Dim strFindKeyword  As String
    Dim strTargetText   As String
    Dim strConvertText  As String
    Dim strTemp()       As String
    Dim strResult       As String
    Dim varConvertNum   As Variant
    Dim varDeleteText   As Variant
    Dim intFindPos      As Integer
    Dim objRegExp       As RegExp   '参照設定要(Microsoft VBScript Regular Expressions 5.5)
    Dim objMatchCollect As MatchCollection
    Dim objMatch        As Match
    Dim strTarget()     As String
    Dim intMatchPos()   As Integer
   
    varConvertNum = Array("1", "2", "3", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "100", "1000")
   
    varDeleteText = Array("&", "＆", "/", "／", "･", "・", ".", "．")
   
    '郵便番号にハイフンがあったら除外します
    strRetZip = Replace(strZip, "-", "", , , vbTextCompare)
   
    '住所の全角文字を半角に変換します
    strAddress = StrConv(strAddress, vbNarrow)
     
    '住所に含まれる記号を除外します
    For intLoop = 1 To UBound(varDeleteText)
      strAddress = Replace(strAddress, varDeleteText(intLoop), "", , , vbTextCompare)
    Next intLoop

    '２つ以上連続したアルファベットを「-」へ置換する
    Set objRegExp = New RegExp
    '正規表現で２つ以上連続したアルファベットを検索し、「-」へ置換する。
    objRegExp.Pattern = "[a-zA-Z][a-zA-Z]+"
    objRegExp.Global = True '複数該当する場合に対応
    strAddress = objRegExp.Replace(strAddress, "-")
   
    '抜き出しの補足ルール
    'http://www.post.japanpost.jp/zipcode/zipmanual/p19.html
   
    '1. 漢数字が下記の特定文字の前にある場合は抜き出し対象とし、算用数字に変換して抜き出します。
    '特定文字群(9種類) "丁目"　 "丁 "　"番地" 　"番"　 "号" 　"地割" 　"線"　 "の" 　"ノ"
   
    '１つずつ見ていく
    For i = 1 To 9
        strFindKeyword = ""
        Select Case i
            Case 1
                strFindKeyword = "丁目"
            Case 2
                strFindKeyword = "丁"
            Case 3
                strFindKeyword = "番地"
            Case 4
                strFindKeyword = "番"
            Case 5
                strFindKeyword = "号"
            Case 6
                strFindKeyword = "地割"
            Case 7
                strFindKeyword = "線"
            Case 8
                strFindKeyword = "の"
            Case 9
                strFindKeyword = "ノ"
        End Select
       
        '当該キーワードが見つかるか？
        intPos = InStr(strAddress, strFindKeyword)
       
        '見つかったら変換処理を行う。
        If intPos > 0 Then
            '見つかったところから左へ１つずつ移動し漢数字じゃなくなったら終了
            For intCount = intPos To 2 Step -1
                strTargetText = Mid$(strAddress, (intCount - 1), 1)
                intTargetPos = InStr(FIND_NUM, strTargetText)
               
                '漢数字なら継続
                If intTargetPos > 0 Then
                Else
                    '漢数字以外ならループを抜ける
                    Exit For
                End If
            Next intCount
           
            strResult = ""
           
            'カウントが開始位置と同一なら漢数字なしと判定して何もしない。
            If intPos = intCount Then
            Else
                strConvertText = Mid$(strAddress, intCount, (intPos - intCount))
                intTextLength = Len(strConvertText)
                ReDim strTemp(intTextLength)
                For j = 1 To intTextLength
                    intTargetPos = InStr(FIND_NUM, Mid$(strConvertText, j, 1))
                    strTemp(j) = varConvertNum(intTargetPos - 1)
                Next j
               
                strResult = "0"
               
                '計算して数値を求める
                For j = 1 To intTextLength
                    If (j = 1) Then
                        strResult = strTemp(j)
                    Else
                        '２桁以上なら前の桁と掛け算して足す
                        If Len(strTemp(j)) >= 2 Then
                            'すでに代入されている値が２桁以上なら足す
                            If Len(strResult) >= 2 Then
                                strResult = CStr(CInt(strResult) + (CInt(strTemp(j - 1)) * CInt(strTemp(j))))
                            Else
                                '代入されている数字が１桁なら、そのまま代入
                                strResult = CStr(CInt(strTemp(j - 1)) * CInt(strTemp(j)))
                            End If
                        Else
                            '１つ前が２桁以上か？
                            If Len(strTemp(j - 1)) >= 2 Then
                                '最終桁だったら足す
                                If j = intTextLength Then
                                    strResult = CStr(CInt(strResult) + CInt(strTemp(j)))
                                Else
                                    'それ以外だったら何もしないで次にまわす
                                End If
                            Else
                                '１桁なら結合する
                                strResult = strResult & strTemp(j)
                            End If
                        End If
                    End If
                Next j
            End If
            '変換した結果があれば処理する
            If Len(strResult) = 0 Then
            Else
                strAddress = Left$(strAddress, (intCount - 1)) & strResult & Mid$(strAddress, intPos)
            End If
        End If
    Next i
  
    '住所から数字部分だけを取り出して"("で区切ります
    strRetAddress = ""
   
    blnConvFlg = True '先頭の数字以外の文字は無視する設定
   
    For intLoop = 1 To Len(strAddress)
      '１文字取り出し
      strChar = Mid$(strAddress, intLoop, 1)
     
        If IsNumeric(strChar) Then
           
            '数字のとき
            strRetAddress = strRetAddress & strChar
            blnConvFlg = False
         
        Else
       
            'ハイフンに置き換え済みでないとき
            If Not blnConvFlg Then
                'アルファベットだったらそのままくっつける
                intFindPos = InStr(FIND_TEXT, strChar)
                If intFindPos > 0 Then
                    If intFindPos = 6 Then  'Fだったら特別扱い->"-"に置き換え
                        strRetAddress = strRetAddress & "-"
                    Else
                        If Right$(strRetAddress, 1) = "-" Then
                            strRetAddress = strRetAddress & strChar
                        Else
                            strRetAddress = strRetAddress & "-" & strChar & "-"
                        End If
                    End If
                Else
                    strRetAddress = strRetAddress & "-"
                End If
                blnConvFlg = True
            Else

                'アルファベットだったらそのままくっつける
                intFindPos = InStr(FIND_TEXT, strChar)
                If intFindPos > 0 Then
                    If Right$(strRetAddress, 1) = "-" Then
                        strRetAddress = strRetAddress & strChar
                    Else
                        strRetAddress = strRetAddress & "-" & strChar & "-"
                    End If
                Else
                    strRetAddress = strRetAddress & "-"
                End If
                blnConvFlg = True
            End If
        End If
    Next intLoop
   
    '最終処理
   
    '連続したハイフンは１つする
    objRegExp.Pattern = "\-\-+"
    objRegExp.Global = True '複数該当する場合に対応
    strRetAddress = objRegExp.Replace(strRetAddress, "-")
   
    '最後と先頭のハイフンを除去します
   
    If Left$(strRetAddress, 1) = "-" Then
      strRetAddress = Mid$(strRetAddress, 2)
    End If
   
    If Right$(strRetAddress, 1) = "-" Then
      strRetAddress = Left$(strRetAddress, Len(strRetAddress) - 1)
    End If
   
    '（アルファベットの前後の-(ハイフン)は取り除きます）
    objRegExp.Pattern = "-[a-zA-Z]-"
    objRegExp.Global = True '複数該当する場合に対応
   
    'パターンマッチするものをCollectionへ格納
    Set objMatchCollect = objRegExp.Execute(strRetAddress)
    '１つ以上あれば以降の処理を継続
    If objMatchCollect.Count > 0 Then
        '配列を再定義
        ReDim strTarget(objMatchCollect.Count)
        ReDim intMatchPos(objMatchCollect.Count)
       
        '複数該当に対応する
        For i = 0 To (objMatchCollect.Count - 1)
            'Collectionの情報を展開
            Set objMatch = objMatchCollect.Item(i)
            '-(ハイフン)をスペースへ変換して格納
            strTarget(i + 1) = Replace(objMatch.Value, "-", " ", , , vbTextCompare)
            '文字列が見つかった場所を格納
            intMatchPos(i + 1) = objMatch.FirstIndex
        Next i
       
        '実際に置き換える
        For i = 1 To objMatchCollect.Count
            strRetAddress = Left$(strRetAddress, intMatchPos(i)) & strTarget(i) & Mid$(strRetAddress, (intMatchPos(i) + Len(strTarget(i)) + 1))
        Next i
        '文字列中の空白を削って終了
        strRetAddress = Replace(strRetAddress, " ", "", , , vbTextCompare)
       
    End If

    '変換された郵便番号と住所を結合して返します
    CreateCustomerBarCode = strRetZip & strRetAddress

End Function