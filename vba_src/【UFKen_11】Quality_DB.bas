Attribute VB_Name = "【UFKen_11】Quality_DB"
Option Explicit
Public GradeArray As Variant
'<<<<<<<<<<DB検索用>>>>>>>>>>>>>>>>
Public Function FalseSearch_S(TempArr As Variant, Temp As Long) As Variant '欠陥数検索(グレード表)
'Q-S-QA-SRF-128_別表1の検索結果
Dim i As Long, a As Long, I_count As Long, n As Long, Amatch As Long, nCount As Long

    For i = LBound(TempArr, 2) To UBound(TempArr, 2)
      If TempArr(1, i) <> "管理番号" And TempArr(1, i) <> "パターン" And _
       TempArr(1, i) <> "備考" And TempArr(1, i) <> "更新日" Then n = n + 1
    Next i
    
    '条件(O,S,K,A) *大きさ 触感 数 素洗い
    ReDim TempSize(1 To 5, 1 To n)

    '異常名抜き取り　'条件、ﾄﾋﾞｺﾐ、ｷｽﾞ、突上、ﾄﾓｽﾞﾚ、汚れ、油より、ﾛｰﾙｷｽﾞ、ﾛｰﾙﾏｰｸ、ｶﾐｺﾐ、切断面、ｴｯﾁﾞ
    For i = LBound(TempArr, 2) To UBound(TempArr, 2)
      If TempArr(1, i) <> "管理番号" And TempArr(1, i) <> "パターン" And _
      TempArr(1, i) <> "備考" And TempArr(1, i) <> "更新日" Then
        nCount = nCount + 1
        TempSize(1, nCount) = TempArr(1, i)
      End If
    Next i
    
    Dim C_パターン As Long, C_条件 As Long
    For i = LBound(TempArr, 2) To UBound(TempArr, 2)
      Select Case TempArr(1, i)
      Case "パターン":    C_パターン = i
      Case "条件":        C_条件 = i
      End Select
    Next i
    
    n = 0
    '異常グレード抜き取り
    For i = LBound(TempArr, 1) To UBound(TempArr, 1)
      If val(TempArr(i, C_パターン)) = val(Temp) Then 'パターンNO一致
      n = n + 1
        For a = LBound(TempSize, 2) To UBound(TempSize, 2)
          For I_count = LBound(TempArr, 2) To UBound(TempArr, 2) '品質基準グレード表のフィールド
            If TempArr(1, I_count) = TempSize(1, a) Then '異常名一致
'              TempSize(n + 1, 1) = TempArr(i, C_条件) 'O,S,K,A
              TempSize(n + 1, a) = TempArr(i, I_count) '格納
            End If
           Next I_count
        Next a
      End If
    Next i
FalseSearch_S = TempSize
'Stop
End Function
Public Function FalseSearch_L(TempArr As Variant, Resul As Variant) As Variant '欠陥数検索
'TempArr=Q-S-QA-124_別表1,Resul=欠陥名と検索結果ｸﾞﾚｰﾄﾞの配列
Dim i As Long, a As Long, I_count As Long, n As Long, Amatch As Long

ReDim TempSize(1 To UBound(Resul, 1) - 1, 1 To 3) '欠陥名、欠陥許容値、グレードの配列にする UBound(Resul, 1) - 1で管理番号分を除去
Dim C_欠陥名 As Long
If Not IsEmpty(TempArr) Then
  For i = 1 To UBound(TempArr, 2)
    Select Case TempArr(1, i)
    Case "欠陥名": C_欠陥名 = i
    End Select
  Next i
End If

For i = LBound(Resul, 1) To UBound(Resul, 1)
  For a = LBound(TempArr, 1) To UBound(TempArr, 1)
    If TempArr(a, C_欠陥名) = Resul(i, 1) Then  '欠陥名一致
      For I_count = LBound(TempArr, 2) To UBound(TempArr, 2) 'フィールド名検索
        If Resul(i, 2) = TempArr(1, I_count) Then '欠陥グレード一致
          Amatch = 1 + Amatch
          TempSize(Amatch, 1) = TempArr(a, C_欠陥名) '(Resul(i, 1)) '欠陥名
          TempSize(Amatch, 2) = TempArr(a, I_count) 'グレード内容
          TempSize(Amatch, 3) = TempArr(LBound(TempArr, 1), I_count) 'グレード
        End If
      Next I_count
    End If
  Next a
Next i

FalseSearch_L = TempSize

End Function

'DBから条件抽出　グレード表から欠陥数を表示させる
Public Function DBdisplay外観_H(ロット As String, Optional TempPath As String = "") As Variant 'SIKALOTNOW.accdb
Dim target As String: target = "SIKALOTNOW.accdb"
If TempPath = "" Then
 TempPath = MainPath
Else
 TempPath = NextPath
End If

Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

'パスセット
Dim FilePath As String: FilePath = TempPath & target
If TempPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    TempPath = NextPath
    FilePath = TempPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_DBdisplay", vbCritical
      Exit Function
    End If
  End If
End If

Dim TempH As String: TempH = ""
If TempPath = NextPath Then TempH = UF1.DBflag.Text
Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", TempH) '2022.07_引数追加  UF1.DBflag.text


Dim i As Long, a As Long, n As Long, I_count As Long
Dim C_オーダー板厚 As Long, C_オーダー幅 As Long, C_オーダー丈 As Long, _
C_製造厚 As Long, C_品質グレード_表面処理 As Long, C_納入先名 As Long, C_VC As Long, C_用途名 As Long, _
C_用途コード As Long, C_オーダー番号 As Long, C_幅 As Long, C_丈 As Long, C_材質 As Long, C_調質 As Long
'------------------------------------------------------------------------------------------------------------------------
'フォールド検索
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "ｵｰﾀﾞｰ板厚":              C_オーダー板厚 = i
    Case "ｵｰﾀﾞｰ板幅":              C_オーダー幅 = i
    Case "ｵｰﾀﾞｰ板丈":              C_オーダー丈 = i
    
    Case "品質ｸﾞﾚｰﾄﾞ_表面処理":    C_品質グレード_表面処理 = i
    Case "納入先名":               C_納入先名 = i
    Case "VC":                     C_VC = i
    Case "用途ｺｰﾄﾞ":               C_用途コード = i
    Case "用途名":                 C_用途名 = i
    Case "ｵｰﾀﾞｰ番号":              C_オーダー番号 = i
    Case "製造材質":               C_材質 = i
    Case "製造調質":               C_調質 = i
    
    Case "製造板厚":               C_製造厚 = i
    Case "製造板幅":               C_幅 = i
    Case "製造板丈":               C_丈 = i
    End Select
  Next i
ElseIf IsEmpty(HiAcc) Then
  Exit Function
End If
'------------------------------------------------------------------------------------------------------------------------
'ロット検索
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", ロット, "Record", TempH) 'UF1.DBflag.text
If IsEmpty(LotAcc) Then
  If ResultT = False Then MsgBox target & "に該当ロットがありません" & "_DBdisplay外観", vbCritical
  Exit Function
End If
'------------------------------------------------------------------------------------------------------------------------
'変数セット
Dim オーダー厚 As Single:  オーダー厚 = LotAcc(1, C_オーダー板厚)
Dim オーダー幅 As Single:  オーダー幅 = LotAcc(1, C_オーダー幅)
Dim オーダー丈 As Single:  オーダー丈 = LotAcc(1, C_オーダー丈)

Dim 製造厚 As Single:      製造厚 = LotAcc(1, C_製造厚)
Dim 製造幅 As Single:      製造幅 = LotAcc(1, C_幅)
Dim 製造丈 As Single:      製造丈 = LotAcc(1, C_丈)

Dim グレード As String:    グレード = LotAcc(1, C_品質グレード_表面処理)
Dim 納入 As String:        納入 = LotAcc(1, C_納入先名)
Dim 用途 As String:        用途 = LotAcc(1, C_用途コード)
Dim 用途名 As String:      用途名 = LotAcc(1, C_用途名)
Dim 受注番号 As String:    受注番号 = LotAcc(1, C_オーダー番号)
Dim 材質 As String:        材質 = LotAcc(1, C_材質)
Dim 調質 As String:        調質 = LotAcc(1, C_調質)
'ﾌｫｰﾑ転記
With UF品質基準
  .ATU = Format(製造厚, "0.000")
  .HABA = Format(製造幅, "0.0")
  .TAKE = Format(製造丈, "0.0")
  .Zai = 材質
  .TYO = 調質
  .YoutC = 用途
  .Noun = 納入
End With

'------------------------------------------------------------------------------------------------------------------------
'厚板か薄板か
Dim AtuFlag As Boolean: AtuFlag = False

'If オーダー厚 <= 8# Then AtuFlag = True　8mm以下ならばTrue　　If AtuFlag = False Then '厚板になっていた
If オーダー厚 >= 8# Then AtuFlag = True '*****不等号の向きとTrue、Falseがあべこべだった＿2026.03.27_修正
'------------------------------------------------------------------------------------------------------------------------

Dim Target2 As String:       Target2 = "品質基準.accdb"
Dim TempPath2 As String:     TempPath2 = QualityPath
Dim TableName As String, TABLENAME2 As String '厚用途で変化
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String, SQLfilter2 As String

FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant, ArrQuality2 As Variant
Dim FieldsQuality As Variant, FieldsQuality2 As Variant
Dim TempSize_L, Arry_L, TempSize_S, Arry_S As Long

n = 0
Dim matchFlag As Boolean
'------------------------------------------------------------------------------------------------------------------------
'規格分岐
'Q-S-QA-124_別表1     ***厚板外観基準*****行と列を入れ替えたものでしか対応できない
'Q-S-QA-124_別表2     ***厚板品質グレード
'------------------------------------------------------------------------------------------------------------------------
'<<<外観>>>
If AtuFlag = True Then '厚板
'QSQA124-2
    TableName = "Q_S_QA_124_別表2"
    TABLENAME2 = "Q_S_QA_124_別表1"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter) 'グレード一覧フィールド
    If IsEmpty(FieldsQuality) Then Exit Function
    
    FieldsQuality2 = GetFieldsArr(FilePath, TABLENAME2, fieldsfilter) '品質基準一覧フィールド
    If IsEmpty(FieldsQuality2) Then Exit Function
    
    Dim R_材質 As Long, R_用途コード As Long, R_板厚下 As Long, R_板厚上 As Long, R_厚 As Long, _
    R_飛込み As Long, R_キズ As Long, R_鋳造剥離AL圧込 As Long, R_AL付着跡 As Long, R_汚れボタ落ち As Long, R_カミコミ跡ショック跡 As Long, _
    R_備考 As Long, R_更新日 As Long, R_番号 As Long, R_納入先 As Long, R_板幅下 As Long, R_板幅上 As Long, R_調質 As Long, _
    R_板幅 As Long, R_板丈下 As Long, R_板丈上 As Long, R_板丈 As Long
    
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      Select Case FieldsQuality(1, i)
      Case "材質":                  R_材質 = i
      Case "用途コード":            R_用途コード = i
      Case "板厚下":                R_板厚下 = i
      Case "板厚上":                R_板厚上 = i
      Case "厚":                    R_厚 = i
      Case "飛込み":                R_飛込み = i
      Case "キズ":                  R_キズ = i
      Case "鋳造剥離AL圧込":        R_鋳造剥離AL圧込 = i
      Case "AL付着跡":              R_AL付着跡 = i
      Case "汚れボタ落ち":          R_汚れボタ落ち = i
      Case "カミコミ跡ショック跡":  R_カミコミ跡ショック跡 = i
      Case "備考":                  R_備考 = i
      Case "更新日":                R_更新日 = i
      Case "管理番号":              R_番号 = i
      
      Case "調質":                  R_調質 = i
      Case "納入先":                R_納入先 = i
      Case "板幅下":                R_板幅下 = i
      Case "板幅上":                R_板幅上 = i
      Case "幅":                    R_板幅 = i
      Case "板丈下":                R_板丈下 = i
      Case "板丈上":                R_板丈上 = i
      Case "丈":                    R_板丈 = i
      End Select
    Next i

    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      If FieldsQuality(1, i) <> "管理番号" And FieldsQuality(1, i) <> "材質" And FieldsQuality(1, i) <> "用途コード" _
      And FieldsQuality(1, i) <> "板厚下" And FieldsQuality(1, i) <> "板厚上" And FieldsQuality(1, i) <> "厚" _
      And FieldsQuality(1, i) <> "更新日" And FieldsQuality(1, i) <> "備考" And FieldsQuality(1, i) <> "調質" _
      And FieldsQuality(1, i) <> "納入先" And FieldsQuality(1, i) <> "板幅下" And FieldsQuality(1, i) <> "板幅上" _
      And FieldsQuality(1, i) <> "幅" And FieldsQuality(1, i) <> "板丈下" And FieldsQuality(1, i) <> "板丈上" _
      And FieldsQuality(1, i) <> "丈" Then n = n + 1
    Next i
    ReDim Arry_L(1 To n + 1, 1 To 2) '欠陥名とグレードの配列　グレードは後で入れる

    '************異常名用配列*************
    '異常名のみを抜き取る  条件を除いているのでそれ以外は異常名として扱う
    n = 0
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      If FieldsQuality(1, i) <> "管理番号" And FieldsQuality(1, i) <> "材質" And FieldsQuality(1, i) <> "用途コード" _
      And FieldsQuality(1, i) <> "板厚下" And FieldsQuality(1, i) <> "板厚上" And FieldsQuality(1, i) <> "厚" _
      And FieldsQuality(1, i) <> "更新日" And FieldsQuality(1, i) <> "備考" And FieldsQuality(1, i) <> "調質" _
      And FieldsQuality(1, i) <> "納入先" And FieldsQuality(1, i) <> "板幅下" And FieldsQuality(1, i) <> "板幅上" _
      And FieldsQuality(1, i) <> "幅" And FieldsQuality(1, i) <> "板丈下" And FieldsQuality(1, i) <> "板丈上" _
      And FieldsQuality(1, i) <> "丈" Then
        n = n + 1
        Arry_L(n, 1) = FieldsQuality(1, i) '欠陥名
      End If
    Next i

    SQLfilter = "[用途コード] LIKE " & "'*" & 用途 & "*'"
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途が一致した配列
    If IsEmpty(ArrQuality) Then Exit Function
    
    Dim JoinQuality: JoinQuality = Call_MergeArray_Row(FieldsQuality, ArrQuality) '別表2配列結合

    SQLfilter2 = ""
    ArrQuality2 = GetRecordsArr(FilePath, TABLENAME2, SQLfilter2) '品質基準一覧
    If IsEmpty(ArrQuality2) Then Exit Function
    
    Dim JoinArr: JoinArr = Call_MergeArray_Row(FieldsQuality2, ArrQuality2) '別表1配列結合
    
'--------------------------------------------------------------------------------------------------------------
  Dim Conditions(1 To 6) As Boolean ' 各条件の結果を格納する変数
  
  Dim MatchingArray() As Long ' マッチングしたインデックスを格納する配列
  ReDim MatchingArray(1 To UBound(ArrQuality, 1)) ' 配列のサイズを設定
  Dim SearchS As String '参照元用
  n = 0
  
  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
'  If ArrQuality(i, R_番号) = 13 Then 'Test
'  Stop
'  End If
    Conditions(1) = CheckCondition(i, R_材質, 材質, ArrQuality)  ' 製造材質のチェック
    Conditions(2) = Conditions(1) And CheckCondition(i, R_材質, 材質, ArrQuality) ' 製造調質のチェック
    Conditions(3) = Conditions(2) And CheckThickness(i, R_板厚下, R_板厚上, R_厚, 製造厚, ArrQuality) '2026.03.27_オーダー厚から製造厚参照に修正
    Conditions(4) = Conditions(3) And CheckWidth(i, R_板幅下, R_板幅上, R_板幅, オーダー幅, ArrQuality) '幅のチェック　現状では未チェック状態　テーブルがカラ
    Conditions(5) = Conditions(4) And CheckWidth(i, R_板丈下, R_板丈上, R_板丈, オーダー丈, ArrQuality) '丈のチェック
    Conditions(6) = Conditions(5) And CheckCondition(i, R_納入先, 納入, ArrQuality) ' 納入先のチェック
  
   ' 条件が全て満たされた場合の処理
    If Conditions(1) And Conditions(2) And Conditions(3) _
    And Conditions(4) And Conditions(5) And Conditions(6) Then
      n = n + 1
      MatchingArray(n) = ArrQuality(i, R_番号)
    End If
  Next i

  If n = 0 Then
    Exit Function 'ヒットなし
  End If
  
  ' MatchingArrayから0を削除する
  Dim NewArray() As Long
  ReDim NewArray(1 To n)
  For i = 1 To n
      NewArray(i) = MatchingArray(i)
  Next i
  
  ''公差格納
  'Dim Kousa()
  'ReDim Kousa(1 To UBound(NewArray, 1), 1 To 10)
  Dim Joi_count As Long
  For a = LBound(NewArray, 1) To UBound(NewArray, 1) '条件一致数分
    For i = LBound(JoinQuality, 1) To UBound(JoinQuality, 1)
      If (JoinQuality(i, R_番号) = NewArray(a)) Then '管理番号一致
      Arry_L(UBound(Arry_L, 1), 1) = "管理番号"
      Arry_L(UBound(Arry_L, 1), 2) = NewArray(a)
        For I_count = LBound(Arry_L, 1) To UBound(Arry_L, 1)
          For Joi_count = LBound(JoinQuality, 2) To UBound(JoinQuality, 2)
            If Arry_L(I_count, 1) = JoinQuality(1, Joi_count) Then '欠陥名一致
              Arry_L(I_count, 2) = JoinQuality(i, Joi_count) 'グレード格納
            End If
          Next Joi_count
        Next I_count
      End If
    Next i
  Next a

'---------------------------------------------------------------------------------------------------------------
    '材質の文字数を調べ、4文字以上の材質の場合は右から3文字を抜き取る
    '***DB側の材質が52S　探す側がF52Sの場合あいまい検索で引っ掛けれない
'    If Len(材質) >= 4 Then 材質 = Right(材質, 3)
'Dim TempSingleU As Single, TempSingleS As Single
'    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
'      If CheckCondition(i, R_材質, 材質, ArrQuality) Then ' 製造材質のチェック
''      If ArrQuality(i, R_材質) Like "*" & 材質 & "*" Then  '探す側に*
''        If Not 納入 Like "*ﾅﾒｶﾜｱﾙﾐ*" Then '滑川フラグのないもの
'          TempSingleU = Val(ArrQuality(i, R_板厚上))
'          TempSingleS = Val(ArrQuality(i, R_板厚下))
'          If ((TempSingleS <= 製造厚) And (TempSingleU >= 製造厚)) And (ArrQuality(i, R_ナメカワフラグ) = "") Then  '厚が収まる範囲を探す
'              Arry_L(1, 2) = ArrQuality(i, R_飛込み)
'              Arry_L(2, 2) = ArrQuality(i, R_キズ)
'              Arry_L(3, 2) = ArrQuality(i, R_鋳造剥離AL圧込)
'              Arry_L(4, 2) = ArrQuality(i, R_AL付着跡)
'              Arry_L(5, 2) = ArrQuality(i, R_汚れボタ落ち)
'              Arry_L(6, 2) = ArrQuality(i, R_カミコミ跡ショック跡)
'              Arry_L(7, 2) = ArrQuality(i, R_番号)
'              matchFlag = True
'              Exit For
'            If matchFlag Then Exit For
''          End If
'        Else
''          If 納入 Like "*ﾅﾒｶﾜｱﾙﾐ*" Then ' "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ ｲﾀｺｺｳｼﾞﾖｳ" Then '"ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Nou = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ" Or Oku = "ﾅﾒｶﾜｱﾙﾐ(ｶ ｲﾀｺｺｳｼﾞﾖｳ"
'            If ((TempSingleS <= 製造厚) And (TempSingleU >= 製造厚)) And (ArrQuality(i, R_ナメカワフラグ) <> "") Then
'                Arry_L(1, 2) = ArrQuality(i, R_飛込み)
'                Arry_L(2, 2) = ArrQuality(i, R_キズ)
'                Arry_L(3, 2) = ArrQuality(i, R_鋳造剥離AL圧込)
'                Arry_L(4, 2) = ArrQuality(i, R_AL付着跡)
'                Arry_L(5, 2) = ArrQuality(i, R_汚れボタ落ち)
'                Arry_L(6, 2) = ArrQuality(i, R_カミコミ跡ショック跡)
'                Arry_L(7, 2) = ArrQuality(i, R_番号)
'                matchFlag = True
'                Exit For
'              If matchFlag Then Exit For
'            End If
''          End If
'          If matchFlag = True Then Exit For
'        End If
'      If matchFlag Then Exit For
'      End If
'    If matchFlag Then Exit For
'    Next i

    'DBヒットNo
'    UF品質基準.G_Frame.Caption = "品質ｸﾞﾚｰﾄﾞ & 基準"
'    UF品質基準.G_Frame.Caption = UF品質基準.G_Frame.Caption & "_" & Arry_L(7, 2)
    TempSize_L = FalseSearch_L(JoinArr, Arry_L) '欠陥数検索
    If IsEmpty(TempSize_L) Then Exit Function
    
    UF品質基準.AbnoName.Tag = "stop" 'チェンジイベント停止
    Dim Temp As String:      Temp = AddPagesArrN(UF品質基準.TabAb, TempSize_L) 'ページ追加
    UF品質基準.AbnoName = "「異常グレード確認一覧」" & vbCrLf & Temp
    
    SQLfilter = ""
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途が一致した配列
    
    If Not IsEmpty(ArrQuality) Then UF品質基準.KousinL1.Caption = ArrQuality(1, R_更新日)
    Dim Settab As String: Settab = TempSize_L(LBound(TempSize_L, 1), LBound(TempSize_L, 2)) '初めに追加したTabCaption
    Call UF_DbDisplay(TempSize_L, Settab)
    UF品質基準.AbnoName.Tag = ""
    
    DBdisplay外観_H = TempSize_L
'Stop
End If
End Function
Public Sub UF_DbDisplay(TempArr As Variant, tabCaption As String) 'DBdisplay外観_H用
If IsEmptyArray(TempArr) Then Exit Sub
If IsEmpty(TempArr) Then Exit Sub
Dim i As Long, a As Long
Dim n As Long: n = LBound(TempArr, 2)
Dim k As Long: k = LBound(TempArr, 2) + 1
Dim j As Long: j = UBound(TempArr, 2)

'欠陥名、欠陥許容値、グレードの配列
  With UF品質基準
    For i = LBound(TempArr, 1) To UBound(TempArr, 1)
    ',を改行とする
      Dim tMP As String:       tMP = TempArr(i, k)
      Dim FinalText As String: FinalText = ""
'      For a = 1 To Len(Tmp) ' コンマがある場所で改行を挿入
'        If Mid(Tmp, a, 1) = "," Then
'          FinalText = FinalText & vbCrLf
'        Else
'          FinalText = FinalText & Mid(Tmp, n, 1)
'        End If
'      Next a
      If tabCaption = TempArr(i, n) Then
        .AbnoLabel = TempArr(i, n)
        .gradeTB = TempArr(i, j)
        .AbnoTB = TempArr(i, k)
      End If
      
'      Select Case TempArr(i, a)
'      Case "飛込み"
'        .ToG = TempArr(i, 3)
'        .Tobi = FinalText 'TempArr(i, 2)
'      Case "キズ"
'        .KiG = TempArr(i, 3)
'        .Kizu = FinalText 'TempArr(i, 2)
'      Case "鋳造剥離AL圧込"
'        .HaG = TempArr(i, 3)
'        .Haku = FinalText 'TempArr(i, 2)
'      Case "AL付着跡"
'        .ALG = TempArr(i, 3)
'        .ALhu = FinalText 'TempArr(i, 2)
'      Case "汚れボタ落ち"
'        .BoG = TempArr(i, 3)
'        .Bota = FinalText 'TempArr(i, 2)
'      Case "カミコミ跡ショック跡"
'        .KaG = TempArr(i, 3)
'        .Kami = FinalText 'TempArr(i, 2)
'      End Select
    Next i
  End With
End Sub

Public Sub DBdisplay外観_S(ロット As String, OutName As String, Optional TempPath As String = "") '押したボタン名を引数にする
Dim target As String: target = "SIKALOTNOW.accdb"
If TempPath = "" Then
 TempPath = MainPath
Else
 TempPath = NextPath
End If
'パスセット
Dim FilePath As String: FilePath = TempPath & target
If TempPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    TempPath = NextPath
    FilePath = TempPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      MsgBox "日稲:" & target & "がありません" & "_ReadCheck_DBdisplay", vbCritical
      Exit Sub
    End If
  End If
End If

UF品質基準.L2.Visible = True
UF品質基準.Repaint
Dim TempH As String: TempH = ""
If TempPath = NextPath Then TempH = UF1.DBflag.Text
Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", TempH) '2022.07_引数追加  UF1.DBflag.text

Dim C_オーダー板厚 As Long, C_製造厚 As Long, C_品質グレード_表面処理 As Long, C_納入先名 As Long, C_VC As Long, C_用途名 As Long, _
C_用途コード As Long, C_オーダー番号 As Long, C_幅 As Long, C_丈 As Long, C_材質 As Long, i As Long, a As Long, n As Long, I_count As Long, C_備考 As Long, _
C_パターン As Long, C_厚下 As Long, C_厚上 As Long, C_厚 As Long, C_パターン1 As Long, C_パターン2 As Long, C_条件 As Long, C_厚フラグ As Long, C_板厚1 As Long, C_板厚下 As Long, C_板厚上 As Long, _
C_トビコミ As Long, C_キズスリキズ As Long, C_突上げ突下げ As Long, C_共ズレ As Long, C_汚れ汚れ圧込み As Long, C_油寄りレベラースジ As Long, _
C_ロールキズ As Long, C_ロールマーク As Long, C_噛み込み跡 As Long, C_切断面異常 As Long, C_エッジ変色油残り跡油焼け, C_表面処理コード As Long, _
C_グレード As Long, C_飛込み As Long, C_キズ As Long, C_鋳造剥離AL圧込 As Long, C_AL付着跡 As Long, C_汚れボタ落ち As Long, C_カミコミ跡ショック跡 As Long, C_調質 As Long, C_更新日 As Long

'------------------------------------------------------------------------------------------------------------------------
'フォールド検索
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "ｵｰﾀﾞｰ板厚":              C_オーダー板厚 = i
    Case "製造板厚":               C_製造厚 = i
    Case "品質ｸﾞﾚｰﾄﾞ_表面処理":    C_品質グレード_表面処理 = i
    Case "納入先名":               C_納入先名 = i
    Case "VC":                     C_VC = i
    Case "用途ｺｰﾄﾞ":               C_用途コード = i
    Case "用途名":                 C_用途名 = i
    Case "ｵｰﾀﾞｰ番号":              C_オーダー番号 = i
    Case "製造材質":               C_材質 = i
    Case "製造板幅":               C_幅 = i
    Case "製造板丈":               C_丈 = i
    Case "製造調質":               C_調質 = i
    End Select
  Next i
ElseIf IsEmpty(HiAcc) Then
  UF品質基準.L2.Visible = False
  UF品質基準.Repaint
  Exit Sub
End If
'------------------------------------------------------------------------------------------------------------------------
Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

'ロット検索
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", ロット, "Record", TempH) 'UF1.DBflag.text
If IsEmpty(LotAcc) Then
  If ResultT = False Then MsgBox "SIKALOTNOWに該当ロットがありません" & "_DBdisplay外観", vbCritical
  UF品質基準.L2.Visible = False
  Exit Sub
End If
'------------------------------------------------------------------------------------------------------------------------
'変数セット
Dim オーダー厚 As Single:  オーダー厚 = LotAcc(1, C_オーダー板厚)
Dim 製造厚 As Single:      製造厚 = LotAcc(1, C_製造厚)
Dim 製造幅 As Single:      製造幅 = LotAcc(1, C_幅)
Dim 製造丈 As Single:      製造丈 = LotAcc(1, C_丈)
Dim グレード As String:    グレード = LotAcc(1, C_品質グレード_表面処理)
Dim 納入 As String:        納入 = LotAcc(1, C_納入先名)
Dim 用途 As String:        用途 = LotAcc(1, C_用途コード)
Dim 用途名 As String:      用途名 = LotAcc(1, C_用途名)
Dim 受注番号 As String:    受注番号 = LotAcc(1, C_オーダー番号)
Dim 材質 As String:        材質 = LotAcc(1, C_材質)
Dim 調質 As String:        調質 = LotAcc(1, C_調質)
'ﾌｫｰﾑ転記
With UF品質基準
  .Uatu = Format(製造厚, "0.000")
  .Uhaba = Format(製造幅, "0.0")
  .Utake = Format(製造丈, "0.0")
  .Uzai = 材質
  .Utyou = 調質
  .UYoucode = 用途
  .Uhyou = グレード
End With
'------------------------------------------------------------------------------------------------------------------------
'厚板か薄板か
Dim AtuFlag As Boolean: AtuFlag = False
If オーダー厚 < 8# Then AtuFlag = True  '7.9mm以下ならばTrue　=削除　以上から未満へ_2026.03.27
'------------------------------------------------------------------------------------------------------------------------

Dim Target2 As String:    Target2 = "品質基準.accdb"
Dim TempPath2 As String:  TempPath2 = QualityPath
Dim TableName As String, TABLENAME2 As String '厚用途で変化
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String, SQLfilter2 As String

FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant, ArrQuality2 As Variant
Dim FieldsQuality As Variant, FieldsQuality2 As Variant
Dim TempSize_L, Arry_L, TempSize_S, Arry_S As Long

n = 0
Dim matchFlag As Boolean
'------------------------------------------------------------------------------------------------------------------------
'規格分岐
'Q-S-QA-124_別表1     ***厚板外観基準
'Q-S-QA-124_別表2     ***厚板品質グレード

If AtuFlag Then '薄板
'Q-S-QA-SR-128_別表1
    TableName = "Q_S_QA_SRF_128_別表1"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter) 'パターン判別表フィールド
    If IsEmpty(FieldsQuality) Then Exit Sub
    
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      Select Case FieldsQuality(1, i)
      Case "表面処理コード":   C_表面処理コード = i
      Case "材質":             C_材質 = i
      Case "用途名":           C_用途名 = i
      Case "用途コード":       C_用途コード = i
      Case "厚フラグ":         C_厚フラグ = i
      Case "板厚1":            C_板厚1 = i
      Case "パターン1":        C_パターン1 = i
      Case "パターン2":        C_パターン2 = i
      Case "備考":             C_備考 = i
      Case "更新日":           C_更新日 = i
      End Select
    Next i

    '変数をフィルターに含む記述がむずい件***********"[残仕掛設備ｺｰｽ] LIKE '*KEN*'"  "[ｵｰﾀﾞｰ板丈]<>0 AND [残仕掛設備ｺｰｽ] LIKE '*KEN*' AND  [BOX設計_設備名] = 'KEN' "
    '"[用途コード] LIKE " & "'*" & 用途 & "*' AND [材質] LIKE " & " '*" & 材質 & "*'"
    SQLfilter = "[用途コード] LIKE " & "'*" & 用途 & "*'"
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途コードが一致した配列
    If IsEmpty(ArrQuality) Then
      UF品質基準.L2.Visible = False
      UF品質基準.Upater = "該当用途ｺｰﾄﾞなし"
      Exit Sub
    End If
'Q-S-QA-SR-128_別表2
    TABLENAME2 = "Q_S_QA_SRF_128_別表2"
    FieldsQuality2 = GetFieldsArr(FilePath, TABLENAME2, fieldsfilter) '品質基準(グレード表)フィールド
    If IsEmpty(FieldsQuality2) Then
      Exit Sub
      UF品質基準.L2.Visible = False
    End If
    
    SQLfilter2 = ""
    ArrQuality2 = GetRecordsArr(FilePath, TABLENAME2, SQLfilter2) '品質基準(グレード表)一覧
    If IsEmpty(ArrQuality2) Then
      Exit Sub
      UF品質基準.L2.Visible = False
    End If
    
    Dim JoinArr: JoinArr = Call_MergeArray_Row(FieldsQuality2, ArrQuality2) '配列結合
    
    matchFlag = False

'材質の文字数を調べ、4文字以上の材質の場合は右から3文字を抜き取る
'***DB側の材質が52S　調べる側がF52Sの場合あいまい検索で引っ掛けれない
'If Len(材質) >= 4 Then 材質 = Right(材質, 3)

    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
      If CheckCondition(i, C_材質, 材質, ArrQuality) Then ' 製造材質のチェック
'      If ArrQuality(i, C_材質) Like "*" & 材質 & "*" Then  '探す側に*
  '      MsgBox "ArrQuality:" & ArrQuality(i, C_材質) & "_" & ロット & "材質:" & 材質
          If ArrQuality(i, C_表面処理コード) Like "*" & グレード & "*" Then
  '          MsgBox "ArrQuality:" & ArrQuality(i, C_表面処理コード) & "_" & ロット & "グレード:" & 材質
              If ArrQuality(i, C_厚フラグ) <> "" Then '厚フラグありの場合はパターン1かパターン2(板厚条件あり)
                If val(ArrQuality(i, C_板厚1)) >= val(オーダー厚) Then '条件板厚以下はパターン1
                  Arry_S = ArrQuality(i, C_パターン1)
                  matchFlag = True
                  Exit For
                ElseIf val(ArrQuality(i, C_板厚1)) < val(オーダー厚) Then '条件板厚超えはパターン2
                  Arry_S = ArrQuality(i, C_パターン2)
                  matchFlag = True
                  Exit For
                End If
              Else '厚フラグなしの場合はパターン1
                Arry_S = ArrQuality(i, C_パターン1)
                matchFlag = True
                Exit For
              End If
          End If
      If matchFlag Then Exit For
      End If
    If matchFlag Then Exit For
    Next i
  UF品質基準.Upater = Arry_S
  TempSize_S = FalseSearch_S(JoinArr, Arry_S) '欠陥数検索(グレード表)の配列
  
  If Arry_S = 0 Then
    UF品質基準.Upater = "該当ﾊﾟﾀｰﾝなし"
    UF品質基準.L2.Visible = False
    Exit Sub
  End If
  
  Call 基準反映(OutName, TempSize_S)
  
  If Not IsEmpty(ArrQuality) Then UF品質基準.KousinL4.Caption = ArrQuality(1, C_更新日)
  
  UF品質基準.L2.Visible = False
  If Not (UF品質基準.Sansyou1.Text Like "* Q-S-QA-SRF-128*") Then UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " Q-S-QA-SRF-128"
  UF品質基準.Repaint
Else
  UF品質基準.Upater = "該当板厚なし"
  UF品質基準.L2.Visible = False
End If
End Sub
Public Sub 基準反映(Temp As String, TempArr As Variant) 'ボタンのcaption、グレード配列
Dim i As Long, a As Long, I_count As Long, n As Long, Amatch As Boolean: Amatch = False
With UF品質基準
.SelectO = ""
.SelectS = ""
.SelectK = ""
.SelectA = ""
  For i = LBound(TempArr, 2) To UBound(TempArr, 2) '異常名検索
    If TempArr(1, i) = Temp Then '押したボタンと異常名の一致
      For a = 2 To UBound(TempArr, 1)
        If TempArr(a, 1) = "O" Then .SelectO = TempArr(a, i)
        If TempArr(a, 1) = "S" Then .SelectS = TempArr(a, i)
        If TempArr(a, 1) = "K" Then .SelectK = TempArr(a, i)
        If TempArr(a, 1) = "A" Then .SelectA = TempArr(a, i)
      Next a
      Amatch = True
      Exit For
    End If
  If Amatch Then Exit For
  Next i
End With
End Sub

'DBから条件抽出して表示させる
Public Sub DBdisplay歪(ロット As String, Optional TempPath As String = "")  'SIKALOTNOW.accdb

Dim target As String: target = "SIKALOTNOW.accdb"
If TempPath = "" Then
 TempPath = MainPath
Else
 TempPath = NextPath
End If

'パスセット
Dim FilePath As String: FilePath = TempPath & target
If TempPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    TempPath = NextPath
    FilePath = TempPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
       MsgBox "日稲:" & target & "がありません" & "_ReadCheck_DBdisplay", vbCritical
      Exit Sub
    End If
  End If
End If

Dim TempH As String: TempH = ""
If TempPath = NextPath Then TempH = UF1.DBflag.Text

Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", TempH) '2022.07_引数追加  UF1.DBflag.text

Dim C_オーダー板厚 As Long, C_製造厚 As Long, C_品質グレード_表面処理 As Long, C_納入先名 As Long, C_VC As Long, C_用途名 As Long, _
C_用途コード As Long, C_オーダー番号 As Long, C_幅 As Long, C_丈 As Long, C_材質 As Long, i As Long, a As Long, n As Long, I_count As Long, _
VcFlag As Boolean: VcFlag = False
Dim C_表面処理コード As Long, C_コメント As Long, C_備考 As Long, C_調質 As Long
Dim C_パターン As Long, C_厚下 As Long, C_厚上 As Long, C_厚 As Long, C_パターン1 As Long, C_パターン2 As Long, C_条件 As Long, C_厚フラグ As Long, C_板厚1 As Long, C_板厚下 As Long, C_板厚上 As Long
Dim C_1m当たり平坦度上限 As Long, C_平坦度上限値参照元 As Long, C_品質グレード_フラットネス As Long, C_納入フラグ As Long, C_オーダー板幅, C_オーダー板丈
Dim C_オーダー板厚T As Long, C_製造材質 As Long, C_製造調質 As Long, C_幅公差 As Long, C_丈公差 As Long, C_板幅下 As Long, C_板幅上 As Long, C_板幅 As Long, C_板丈 As Long
Dim Amatch As Long, C_フラットネス品質等級 As Long

Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")
Dim arr() As String, Result As String
'------------------------------------------------------------------------------------------------------------------------
'フォールド検索
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "ｵｰﾀﾞｰ板厚":              C_オーダー板厚 = i
    Case "ｵｰﾀﾞｰ板幅":              C_オーダー板幅 = i
    Case "ｵｰﾀﾞｰ板丈":              C_オーダー板丈 = i
    Case "製造板厚":               C_製造厚 = i
    Case "品質ｸﾞﾚｰﾄﾞ_表面処理":    C_品質グレード_表面処理 = i
    Case "納入先名":               C_納入先名 = i
    Case "VC":                     C_VC = i
    Case "用途ｺｰﾄﾞ":               C_用途コード = i
    Case "用途名":                 C_用途名 = i
    Case "ｵｰﾀﾞｰ番号":              C_オーダー番号 = i
    Case "製造材質":               C_材質 = i
    Case "製造板幅":               C_幅 = i
    Case "製造板丈":               C_丈 = i
    Case "製造調質":               C_調質 = i
    Case "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ":      C_品質グレード_フラットネス = i
    End Select
  Next i
ElseIf IsEmpty(HiAcc) Then
  Exit Sub
End If
'------------------------------------------------------------------------------------------------------------------------
Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

'ロット検索
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", ロット, "Record", TempH) 'UF1.DBflag.text
If IsEmpty(LotAcc) Then
  If ResultT = False Then MsgBox "SIKALOTNOWに該当ロットがありません" & "_DBdisplay歪", vbCritical
  Exit Sub
End If
'------------------------------------------------------------------------------------------------------------------------
'変数セット
Dim オーダー厚 As Single:    オーダー厚 = LotAcc(1, C_オーダー板厚)
Dim オーダー板幅 As Single:  オーダー板幅 = LotAcc(1, C_オーダー板幅)
Dim オーダー板丈 As Single:  オーダー板丈 = LotAcc(1, C_オーダー板丈)
Dim 製造厚 As Single:        製造厚 = LotAcc(1, C_製造厚)
Dim 製造幅 As Single:        製造幅 = LotAcc(1, C_幅)
Dim 製造丈 As Single:        製造丈 = LotAcc(1, C_丈)
Dim グレード As String:      グレード = LotAcc(1, C_品質グレード_表面処理)
Dim フラットネス As Long:    フラットネス = LotAcc(1, C_品質グレード_フラットネス)
Dim 納入 As String:          納入 = LotAcc(1, C_納入先名)
Dim 用途 As String:          用途 = LotAcc(1, C_用途コード) 'コード
Dim 用途名 As String:        用途名 = LotAcc(1, C_用途名) '名
Dim 受注番号 As String:      受注番号 = LotAcc(1, C_オーダー番号)
Dim 材質 As String:          材質 = LotAcc(1, C_材質)
Dim 調質 As String:          調質 = LotAcc(1, C_調質)
'ﾌｫｰﾑ転記
With UF品質基準
  .Uatu = Format(製造厚, "0.000")
  .Uhaba = Format(製造幅, "0.0")
  .Utake = Format(製造丈, "0.0")
  .Uzai = 材質
  .Utyou = 調質
  .UYoucode = 用途
  .Uhyou = グレード
  .Noun = 納入 '白銅がフラグに追加_2023.7.3
  .Noun.BackColor = &H80000000
End With
'------------------------------------------------------------------------------------------------------------------------
'厚板か薄板か
Dim AtuFlag As Boolean: AtuFlag = False
If オーダー厚 <= 6# Then AtuFlag = True '6mm以下ならばTrue
'VC材フラグ
If val(LotAcc(1, C_VC)) = 1 Then VcFlag = True
'引数セット
'幅,丈,ゲージ,耳,中
Dim ReturnSize As Single
'Select Case GaugeWidthHeight
'Case "幅":     ReturnSize = 製造幅
'Case "丈":     ReturnSize = 製造丈
'End Select
'------------------------------------------------------------------------------------------------------------------------
If VcFlag Then
'VC一覧ｵｰﾀﾞｰとSIKALOTNOWのｵｰﾀﾞｰの一致チェック
    Dim VcTarget As String: VcTarget = "SIKAODRNOW.accdb"
    Dim VCPath As String: VCPath = TempPath & VcTarget
    Dim TB As String: TB = "仕掛"
    Dim VcFields As String: VcFields = ""
    
    Dim VC: VC = GetFieldsArr(VCPath, TB, VcFields)
    If IsEmpty(VC) Then GoTo LineSkip
    
    Dim R_VCA As Long, R_VCB As Long, R_Deliver As Long, R_Packing As Long, R_Vcodr As Long
    For i = LBound(VC, 2) To UBound(VC, 2) 'ﾌｨｰﾙﾄﾞ名
      Select Case VC(1, i)
      Case "受注番号":    R_Vcodr = i
      Case "VC_表":       R_VCA = i
      Case "VC_裏":       R_VCB = i
      Case "納入先名称":  R_Deliver = i
      Case "包装仕様NO":  R_Packing = i
      End Select
    Next i
    
    Dim Vcfilter As String: Vcfilter = "[受注番号]='" & 受注番号 & "'"
    
    Dim VCR: VCR = GetRecordsArr(VCPath, TB, Vcfilter)
    If IsEmpty(VCR) Then GoTo LineSkip
    
    Dim VCA As String: VCA = VCR(1, R_VCA)
    Dim VCB As String: VCB = VCR(1, R_VCB)
    'QA-022-01で使用
    Dim useFlag As Boolean: useFlag = False
    If (VCA = "208PRNF" Or VCA = "V325NNF" Or VCA = "V311PNF") Or _
    (VCB = "208PRNF" Or VCB = "V325NNF" Or VCB = "V311PNF") Then useFlag = True
End If
'------------------------------------------------------------------------------------------------------------------------

LineSkip:

Dim Target2 As String:    Target2 = "品質基準.accdb"
Dim TempPath2 As String:  TempPath2 = QualityPath   '"C:\Users\xfa-ngykonpo\Desktop\DB\品質\"'C:\Users\xfa-ngyinazawa\Desktop\DB\品質
Dim TableName As String '厚用途で変化
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String

FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant
Dim FieldsQuality As Variant
Dim TempSize

n = 0
Dim matchFlag As Boolean

Dim SitaSingele As Single, UeSingele As Single, hsitasingele As Single, huesingele As Single
'------------------------------------------------------------------------------------------------------------------------
'規格分岐

'QA-022-01　　　　　　***VC製品平坦度
'Q-S-QA-ALL-122       ***厚板平坦度'2023.7.3_修正変更　Q-S-QA-ALL-122からQ-S-QA-DIM-122_A *2023.7.14_材質列追加
'Q-S-QA-DIM-122_B     ***鉄道ボディ平坦度(Q-S-QA-ALL-129)
'Q-S-QA-DIM-122_O     ***O材限定
'Q-S-QA-DIM-122_D     ***HX8材
'Q-S-QA-124_別表2     ***厚板品質グレード
'Q-S-QA-124_別表1     ***厚板外観基準

'すべてQ-S-QA-DIM-122に統合A,B,C,D_2023.8.30
'------------------------------------------------------------------------------------------------------------------------
'<<<歪>>>
If AtuFlag Then  '耳中歪_4級指定*製品規格
'ReturnSize = 製造丈 'ここに関しては指定
'4級
  UF品質基準.Toukyu = "4級"
    TableName = "耳歪4級"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(FieldsQuality) Then GoTo SkipA
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      Select Case FieldsQuality(1, i)
      Case "平坦度上限値参照元": C_平坦度上限値参照元 = i
      Case "用途コード":         C_用途コード = i
      Case "板厚下":             C_板厚下 = i
      Case "板厚上":             C_板厚上 = i
      Case "備考":               C_備考 = i
      End Select
    Next i
    
    SQLfilter = "" '"[用途コード] LIKE " & "'*" & 用途 & "*'"'_2023.9.14*表示を増やすためフィルター停止
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途が一致した配列
    If IsEmpty(ArrQuality) Then GoTo SkipA
    
    UF品質基準.Mimi1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
    If UF品質基準.Mimi1 = 0 Then UF品質基準.Mimi1 = ""
    
    TableName = "中歪4級"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(FieldsQuality) Then GoTo SkipA
    
    SQLfilter = "" '"[用途コード] LIKE " & "'*" & 用途 & "*'"'_2023.9.14*表示を増やすためフィルター停止
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途が一致した配列
    If IsEmpty(ArrQuality) Then GoTo SkipA
    
    UF品質基準.Naka1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
    If UF品質基準.Naka1 = 0 Then UF品質基準.Naka1 = ""
    UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)"
    '薄板3級用
    If フラットネス = 3 Then
        TableName = "耳歪3級"
        SQLfilter = ""
        ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
        If IsEmpty(ArrQuality) Then GoTo SkipA
        UF品質基準.Mimi2 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
        If UF品質基準.Mimi2 = 0 Then UF品質基準.Mimi2 = ""
    
        TableName = "中歪3級"
        SQLfilter = ""
        ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
        If IsEmpty(ArrQuality) Then GoTo SkipA
        UF品質基準.Naka2 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
        If UF品質基準.Naka2 = 0 Then UF品質基準.Naka2 = ""
     UF品質基準.Toukyu = "3級"
     UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " ﾌﾗｯﾄﾈｽ製品規格中耳歪(ｵｰﾀﾞｰ丈判定)"
    End If
    
    arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
    For i = LBound(arr) To UBound(arr)
      Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
    Next i
    Result = Join(Dict.Keys, " ")
    UF品質基準.Sansyou1 = Result
End If
'------------------------------------------------------------------------------------------------------------------------
SkipA:

If useFlag Then 'VC指定'ゲージ有り
    'QA02201
    TableName = "QA_022_01"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(FieldsQuality) Then Exit Sub
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      Select Case FieldsQuality(1, i)
      Case "VC":                    C_VC = i
      Case "1m当たり平坦度上限":    C_1m当たり平坦度上限 = i
      Case "板厚下":                C_板厚下 = i
      Case "板厚上":                C_板厚上 = i
      Case "厚":                    C_厚 = i
      Case "備考":                  C_備考 = i
      End Select
    Next i

    matchFlag = False
    SQLfilter = ""
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
    If IsEmpty(ArrQuality) Then GoTo skipO
    
    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    SitaSingele = val(ArrQuality(i, C_板厚下))
    UeSingele = val(ArrQuality(i, C_板厚上))
    '厚範囲は調整してあるので両方以上以下で良い
      If (SitaSingele <= 製造厚) And (UeSingele >= 製造厚) Then '厚が収まる範囲を探す
          UF品質基準.MG1 = Format(ArrQuality(i, C_1m当たり平坦度上限), "0.0")
          matchFlag = True
          Exit For
      If matchFlag Then Exit For
      End If
    If matchFlag Then Exit For
    Next i
    
    UF品質基準.HaHiz1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板幅, オーダー厚), "0.0")
    If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
    UF品質基準.TaHiz1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
    If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
    
    UF品質基準.VcFlag.Caption = "あり"
    UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " QA-022-01"
    
    arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
    For i = LBound(arr) To UBound(arr)
      Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
    Next i
    Result = Join(Dict.Keys, " ")
    UF品質基準.Sansyou1 = Result
    Set Dict = Nothing
    Exit Sub
End If
'------------------------------------------------------------------------------------------------------------------------
'2024.2.1_用途コードが変わった為修正
  TableName = "Q_S_QA_DIM_122_B" ' "Q_S_QA_ALL_129"***************テーブル変更_2023.9.8
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(FieldsQuality) Then GoTo skipO

  For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
    Select Case FieldsQuality(1, i)
    Case "材質":         C_材質 = i
    Case "調質":         C_調質 = i
    Case "用途コード":   C_用途コード = i
    Case "板厚下":       C_板厚下 = i
    Case "板厚上":       C_板厚上 = i
    Case "厚":           C_厚 = i
    Case "板幅下":       C_板幅下 = i
    Case "板幅上":       C_板幅上 = i
    Case "幅":           C_幅 = i
    Case "板幅":         C_板幅 = i
    Case "板丈":         C_板丈 = i
    Case "備考":         C_備考 = i
    End Select
  Next i
SQLfilter = "[用途コード] LIKE " & "'*" & 用途 & "*'"
ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途が一致した配列

If IsEmpty(ArrQuality) Then GoTo skip122A 'なければ用途が鉄道ボディー以外へスキップ

Dim 検索変数 As String: 検索変数 = 用途 '******(K21S限定?)"ﾃﾂﾄﾞｳﾎﾞﾃﾞ-"
Dim 配列() As String: 配列 = Split(ArrQuality(1, C_用途コード), ",")
Dim VariableFound  As Boolean: VariableFound = False

For i = LBound(配列) To UBound(配列)
  If Trim(配列(i)) = 検索変数 Then
    VariableFound = True
    Exit For
  End If
Next i

Select Case True
Case VariableFound '"K21S", "K21B", "K217", "K218", "K221", "K222"
''中耳の規格なし
    TableName = "Q_S_QA_DIM_122_B" ' "Q_S_QA_ALL_129"***************テーブル変更_2023.9.8
'    FieldsQuality = GetFieldsArr(filePath, TableName, fieldsfilter)
'    If IsEmpty(FieldsQuality) Then GoTo skipO
'
'    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
'      Select Case FieldsQuality(1, i)
'      Case "材質":         C_材質 = i
'      Case "調質":         C_調質 = i
'      Case "用途コード":   C_用途コード = i
'      Case "板厚下":       C_板厚下 = i
'      Case "板厚上":       C_板厚上 = i
'      Case "厚":           C_厚 = i
'      Case "板幅下":       C_板幅下 = i
'      Case "板幅上":       C_板幅上 = i
'      Case "幅":           C_幅 = i
'      Case "板幅":         C_板幅 = i
'      Case "板丈":         C_板丈 = i
'      Case "備考":         C_備考 = i
'      End Select
'    Next i
    
      '***DB側の材質が52S　探す側がF52Sの場合あいまい検索で引っ掛けれない
      Dim Ittiflag As Boolean: Ittiflag = False
      If Len(材質) >= 4 Then 材質 = Right(材質, 3)
      SQLfilter = "[材質] LIKE " & "'*" & 材質 & "*' AND  [調質] LIKE " & " '*" & 調質 & "*'" '2023.7_調質追加　*完全一致　AND  [調質] = '" & 調質 & "'"
      ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '材質と調質が一致した配列
      
    'Stop
        
      If IsEmpty(ArrQuality) Then Exit Sub
    
      ReDim ArrQ_T(1 To UBound(ArrQuality, 1), 1 To UBound(ArrQuality, 2))
      n = 0
    
      For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
      SitaSingele = val(ArrQuality(i, C_板厚下))
      UeSingele = val(ArrQuality(i, C_板厚上))
      hsitasingele = val(ArrQuality(i, C_板幅下))
      huesingele = val(ArrQuality(i, C_板幅上))

        If SitaSingele <= (オーダー厚) And UeSingele >= (オーダー厚) Then
          If hsitasingele <= (オーダー板幅) And huesingele >= (オーダー板幅) Then '先に幅が収まる範囲を探す
          'LoopArray不要
            UF品質基準.HaHiz1 = Format(ArrQuality(i, C_板幅), "0.0")
            If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
            UF品質基準.TaHiz1 = Format(ArrQuality(i, C_板丈), "0.0")
            If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
            UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " Q-S-QA-DIM-122_B"
            
            '中耳歪不問_2023.9.05
            UF品質基準.Mimi1 = "不問"
            UF品質基準.Naka1 = "不問"
            UF品質基準.Mimi2 = "不問"
            UF品質基準.Naka2 = "不問"
            
            Ittiflag = True
          End If
        End If
      If Ittiflag Then Exit For
      Next i
    
    arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
    For i = LBound(arr) To UBound(arr)
      Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
    Next i
    Result = Join(Dict.Keys, " ")
    UF品質基準.Sansyou1 = Result
    Set Dict = Nothing
    Exit Sub
'------------------------------------------------------------------------------------------------------------------------
Case Else '"ﾃﾂﾄﾞｳﾎﾞﾃﾞ-"以外'ゲージ有り

skip122A:

    TableName = "Q_S_QA_DIM_122_A"
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(FieldsQuality) Then GoTo skipO
    'QSQADIM122
    For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
      Select Case FieldsQuality(1, i)
      Case "平坦度上限値参照元":   C_平坦度上限値参照元 = i
      Case "用途コード":           C_用途コード = i
      Case "板厚下":               C_板厚下 = i
      Case "板厚上":               C_板厚上 = i
      Case "厚":                   C_厚 = i
      Case "1m当たり平坦度上限":   C_1m当たり平坦度上限 = i
      Case "備考":                 C_備考 = i
      Case "納入フラグ":           C_納入フラグ = i
      Case "材質":                 C_材質 = i
      End Select
    Next i

    matchFlag = False
    '変数をフィルターに含む記述がむずい件***********"[残仕掛設備ｺｰｽ] LIKE '*KEN*'"
    SQLfilter = "[用途コード] LIKE " & "'*" & 用途 & "*' AND  [調質] LIKE " & " '*" & 調質 & "*'"  '2023.7_調質追加
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '用途と調質が一致した配列
    
    If IsEmpty(ArrQuality) Then GoTo skipO
    
    ReDim ArrQ_H(1 To UBound(ArrQuality, 1), 1 To UBound(ArrQuality, 2)) '白銅で分岐したとき用
    n = 0
    Dim HakuFlag As Boolean: HakuFlag = False
'2023.7.3_規格変更有のため修正**納入先による分岐が必要になった
'納入が白銅の時は板厚の範囲が重複する箇所がある、そのため用途が一致している配列ArrQualityから厚一致まで探す(用途一致だけで引数配列を渡すと正しく判定できない)

    '材質の文字数を調べ、4文字以上の材質の場合は右から3文字を抜き取る_2023.7.14_材質列追加
    '***DB側の材質が52S　探す側がF52Sの場合あいまい検索で引っ掛けれない
    If Len(材質) >= 4 Then 材質 = Right(材質, 3)

    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    SitaSingele = val(ArrQuality(i, C_板厚下))
    UeSingele = val(ArrQuality(i, C_板厚上))
      If ArrQuality(i, C_材質) Like "*" & 材質 & "*" Then  '探す側に* 文字の前後に何かしらの文字列があってもTrue
        If (SitaSingele <= 製造厚) And (UeSingele >= 製造厚) Then  '先に厚が収まる範囲を探す
          If (ArrQuality(i, C_納入フラグ) = "白銅") And (納入 Like "*ﾊｸﾄﾞｳ*") Then '納入先がﾊｸﾄﾞｳかつ白銅フラグのある板厚範囲ならば配列に入れる
            n = n + 1
            For a = LBound(ArrQuality, 2) To UBound(ArrQuality, 2)
              ArrQ_H(n, a) = ArrQuality(i, a)
            Next a
            UF品質基準.MG1 = Format(ArrQuality(i, C_1m当たり平坦度上限), "0.0")
            UF品質基準.Noun.BackColor = RGB(175, 15, 5)
            HakuFlag = True
          End If
        If HakuFlag = True Then Exit For
        End If
      End If
    If HakuFlag Then Exit For
    Next i
    
    If HakuFlag Then GoTo skipF 'フラグありの白銅作成済**
    
    'スキップなしということは白銅フラグのある板厚範囲ではなかったということなのでフラグなしの範囲で公差配列を作る
    
    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    SitaSingele = val(ArrQuality(i, C_板厚下))
    UeSingele = val(ArrQuality(i, C_板厚上))
      If ArrQuality(i, C_材質) Like "*" & 材質 & "*" Then   '探す側に*
        If (SitaSingele <= 製造厚) And (UeSingele >= 製造厚) Then  '先に厚が収まる範囲を探す
          If (ArrQuality(i, C_納入フラグ) <> "白銅") Then 'ﾊｸﾄﾞｳﾌﾗｸﾞなし
            n = n + 1
            For a = LBound(ArrQuality, 2) To UBound(ArrQuality, 2)
              ArrQ_H(n, a) = ArrQuality(i, a)
            Next a
            UF品質基準.MG1 = Format(ArrQuality(i, C_1m当たり平坦度上限), "0.0")
            If Not HakuFlag Then UF品質基準.Noun.BackColor = RGB(255, 255, 255)
          End If
        End If
      End If
    Next i
    
skipF:
    
    UF品質基準.HaHiz1 = Format(LoopArray(ArrQ_H, FieldsQuality, 製造幅, 製造厚), "0.0")
    UF品質基準.TaHiz1 = Format(LoopArray(ArrQ_H, FieldsQuality, 製造丈, 製造厚), "0.0")
    If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
    If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
    
    If UF品質基準.HaHiz1 = "" And UF品質基準.TaHiz1 = "" Then GoTo skipO 'ヒットなしの場合次のDB検索
    
    UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " Q-S-QA-DIM-122"
    
    arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
    For i = LBound(arr) To UBound(arr)
      Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
    Next i
    Result = Join(Dict.Keys, " ")
    UF品質基準.Sansyou1 = Result
    
    Set Dict = Nothing
    Exit Sub
End Select


'後日O材限定の歪基準追加　ﾌﾛｰ図で進めるとQ-S-QA-DIM-122⇒Q-S-QA-ALL-129⇒Q-S-QA-DIM-122_O
'--------------------------------上記でヒットしなかった場合のみ-----------------------------------------
skipO:
  TableName = "Q_S_QA_DIM_122_C" '2024.2.1_O削除
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(FieldsQuality) Then Exit Sub
  'QSQADIM122_O
  For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
    Select Case FieldsQuality(1, i)
    Case "用途名":         C_用途名 = i
    Case "製造材質":       C_製造材質 = i
    Case "用途コード":     C_用途コード = i
    Case "板厚下":         C_板厚下 = i
    Case "板厚上":         C_板厚上 = i
    Case "オーダー板厚T":  C_オーダー板厚T = i
    Case "オーダー板幅":   C_オーダー板幅 = i
    Case "幅公差":         C_幅公差 = i
    Case "丈公差":         C_丈公差 = i
    Case "備考":           C_備考 = i
    Case "納入先名":       C_納入先名 = i
    Case "調質":           C_調質 = i
    Case "パターン":       C_パターン = i
    End Select
  Next i

  SQLfilter = "[調質] LIKE " & "'*" & 調質 & "*'" 'O材限定
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) 'O材配列
  
  If IsEmpty(ArrQuality) Then GoTo skipHX8

  If Len(材質) >= 4 Then 材質 = Right(材質, 3)
  
  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    If ArrQuality(i, C_製造材質) Like "*" & 材質 & "*" Then  '探す側に* 文字の前後に何かしらの文字列があってもTrue
      If ArrQuality(i, C_用途コード) Like "*" & 用途 & "*" Then
        If 納入 Like "*" & ArrQuality(i, C_納入先名) & "*" Then '基準書側の納入がSIKALOTの納入に含まれているか
        SitaSingele = val(ArrQuality(i, C_板厚下))
        UeSingele = val(ArrQuality(i, C_板厚上))
          If SitaSingele <= val(オーダー厚) And UeSingele >= val(オーダー厚) Then  'オーダー厚参照の指定
            If ArrQuality(i, C_パターン) <> "" Then '幅指定あり
              Select Case ArrQuality(i, C_パターン)
              Case 1
                If ArrQuality(i, C_オーダー板幅) >= オーダー板幅 Then
                 UF品質基準.HaHiz1 = Format(ArrQuality(i, C_幅公差), "0.0")
                 UF品質基準.TaHiz1 = Format(ArrQuality(i, C_丈公差), "0.0")
                 UF品質基準.Noun.BackColor = RGB(255, 255, 153)
                End If
              Case 2
                If ArrQuality(i, C_オーダー板幅) <= オーダー板幅 Then
                 UF品質基準.HaHiz1 = Format(ArrQuality(i, C_幅公差), "0.0")
                 UF品質基準.TaHiz1 = Format(ArrQuality(i, C_丈公差), "0.0")
                 UF品質基準.Noun.BackColor = RGB(255, 255, 153)
                End If
              End Select
            Else
              UF品質基準.HaHiz1 = Format(ArrQuality(i, C_幅公差), "0.0")
              UF品質基準.TaHiz1 = Format(ArrQuality(i, C_丈公差), "0.0")
              UF品質基準.Noun.BackColor = RGB(255, 255, 153)
            End If
          End If
        End If
      End If
    End If
  Next i
  If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
  If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
  
  '中耳歪不問_2023.9.05
  UF品質基準.Mimi1 = "不問"
  UF品質基準.Naka1 = "不問"
  UF品質基準.Mimi2 = "不問"
  UF品質基準.Naka2 = "不問"
  
  UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " Q-S-QA-DIM-122_O"

  arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
  For i = LBound(arr) To UBound(arr)
    Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
  Next i
  Result = Join(Dict.Keys, " ")
  UF品質基準.Sansyou1 = Result
  Set Dict = Nothing
  Exit Sub
'-------------------------------------------------------------------------
skipHX8:
  TableName = "Q_S_QA_DIM_122_D"
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(FieldsQuality) Then Exit Sub
'Q-S-QA-DIM-122_D     ***HX8材
  For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
    Select Case FieldsQuality(1, i)
    Case "用途名":         C_用途名 = i
    Case "製造材質":       C_製造材質 = i
    Case "製造調質":       C_製造調質 = i
    Case "用途コード":     C_用途コード = i
    Case "オーダー板厚":   C_オーダー板厚 = i
    Case "幅公差":         C_幅公差 = i
    Case "丈公差":         C_丈公差 = i
    Case "備考":           C_備考 = i
    Case "納入先名":       C_納入先名 = i
    Case "フラットネス品質等級": C_フラットネス品質等級 = i
    End Select
  Next i
  
  If Len(材質) >= 4 Then 材質 = Right(材質, 3)
  SQLfilter = "[製造材質] LIKE " & "'*" & 材質 & "*'"
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '材質が一致した配列
  
  If IsEmpty(ArrQuality) Then GoTo skipother
  
  Dim 調質_A As String: 調質_A = Left(調質, 1)
  Dim 調質_K As String: 調質_K = Right(調質, 1)

  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    If ArrQuality(i, C_製造調質) Like 調質_A & "*" & 調質_K Then 'c*t 前後が一致
      If ArrQuality(i, C_用途コード) Like "*" & 用途 & "*" Then '完全一致*************************2024.2.1_部分一致に修正
        If 納入 Like "*" & ArrQuality(i, C_納入先名) & "*" Then '基準書側の納入がSIKALOTの納入に含まれているか
          If val(ArrQuality(i, C_オーダー板厚)) = val(オーダー厚) Then '完全一致
            If ArrQuality(i, C_フラットネス品質等級) = "" Then
            'LoopArray不要
              UF品質基準.HaHiz1 = Format(ArrQuality(i, C_幅公差), "0.0")
              If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
              UF品質基準.TaHiz1 = Format(ArrQuality(i, C_丈公差), "0.0")
              If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
              UF品質基準.Noun.BackColor = RGB(0, 102, 204)
              UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " Q-S-QA-DIM-122_D"
              Ittiflag = True
            ElseIf ArrQuality(i, C_フラットネス品質等級) <> "" Then
              SQLfilter = ""
              If フラットネス = 4 Then
                TableName = "ソリ4級"
                FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
                If IsEmpty(FieldsQuality) Then Exit Sub
                ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
                If IsEmpty(ArrQuality) Then Exit Sub
                UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " 製品規格_ソリ4級"
              ElseIf フラットネス = 3 Then
                TableName = "ソリ3級"
                FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
                If IsEmpty(FieldsQuality) Then Exit Sub
                ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
                If IsEmpty(ArrQuality) Then Exit Sub
                UF品質基準.Sansyou1 = UF品質基準.Sansyou1.Text & " 製品規格_ソリ3級"
              End If
                UF品質基準.HaHiz1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板幅, オーダー厚), "0.0")
                UF品質基準.TaHiz1 = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
                UF品質基準.Noun.BackColor = RGB(0, 102, 204)
            End If
            End If
        End If
      End If
    End If
  If Ittiflag Then Exit For
  Next i
  If UF品質基準.HaHiz1 = 0 Then UF品質基準.HaHiz1 = ""
  If UF品質基準.TaHiz1 = 0 Then UF品質基準.TaHiz1 = ""
  '中耳歪不問_2023.9.05
  UF品質基準.Mimi1 = "不問"
  UF品質基準.Naka1 = "不問"
  UF品質基準.Mimi2 = "不問"
  UF品質基準.Naka2 = "不問"
  
  arr = Split(UF品質基準.Sansyou1.Text, " ") ' 文字列を" "で分割して配列に格納
  For i = LBound(arr) To UBound(arr)
    Dict(arr(i)) = 0 ' 辞書オブジェクトを使用して重複を削除
  Next i
  Result = Join(Dict.Keys, " ")
  UF品質基準.Sansyou1 = Result
  Set Dict = Nothing
  Exit Sub
'-------------------------------------------------------------------------
skipother: '未実装



End Sub

Public Function LoopArray(TempA As Variant, tempB As Variant, ReturnSize As Single, ATU As Single) As Single
'A:品質基準レコード,　B:品質基準フィールド, 幅か丈の寸法,　板厚
Dim i As Long, a As Long, I_count As Long, n As Long, Amatch As Long, C_板厚下 As Long, C_板厚上 As Long
Dim TempSize
Dim matchFlag As Boolean: matchFlag = False

    For i = LBound(tempB, 2) To UBound(tempB, 2)
      Select Case tempB(1, i)
      Case "板厚下":   C_板厚下 = i
      Case "板厚上":   C_板厚上 = i
      End Select
      
      If IsNumeric((tempB(1, i))) Then n = n + 1
      
    Next i
    ReDim TempSize(1 To n)
    n = 0

    '************サイズ用配列*************
    For i = LBound(tempB, 2) To UBound(tempB, 2)
      If IsNumeric((tempB(1, i))) Then
        n = n + 1
        TempSize(n) = (val(tempB(1, i))) 'フィールド名が数字の時配列に入れる
      End If
    Next i
    Dim AtuUnd As Single, AtuOve As Single
    For i = LBound(TempA, 1) To UBound(TempA, 1)
      AtuUnd = val(TempA(i, C_板厚下))
      AtuOve = val(TempA(i, C_板厚上))
    '厚範囲は調整してあるので両方以上以下で良い
      If (AtuUnd <= ATU) And (AtuOve >= ATU) Then '厚が収まる範囲を探す
          For a = UBound(TempSize, 1) To LBound(TempSize, 1) Step -1 '一番右のフィールドとその隣のフィールドの範囲に収まっているかのチェック(大きい順)
            n = a - 1
              If n = 0 Then '0<=製造幅<=500
                If TempSize(a) >= ReturnSize And val(0) < ReturnSize Then '500mm以下の時
                    For I_count = LBound(tempB, 2) To UBound(tempB, 2)
                        If val(tempB(1, I_count)) = TempSize(LBound(TempSize, 1)) Then  'サイズ用配列で一番小さいもの
                            Amatch = I_count
                            LoopArray = Format(TempA(i, Amatch), "0.0")
                            matchFlag = True
                            Exit For
                        End If
                    Next I_count
                End If
              If matchFlag Then Exit For
              End If
    '幅丈範囲はフィールド名で調整できないので、以上未満を使う
    If n = 0 Then GoTo skipErr 'DBにないサイズ(4500mm超え)_2023.8.22修正
                  If TempSize(a) >= ReturnSize And TempSize(n) < ReturnSize Then '最大値フィールド-1(その隣)の範囲内かのチェック
skipErr:
If n = 0 Then a = UBound(TempSize, 1) 'DBにないサイズ(4500mm超え)の時は一番大きいサイズの公差を確認_2023.8.22
                    For I_count = LBound(tempB, 2) To UBound(tempB, 2)
                      If IsNumeric(tempB(1, I_count)) Then 'フィールド名が数字
                        If val(tempB(1, I_count)) = val(TempSize(a)) Then 'フィールド名と配列に入れた物が一致したとき
                            Amatch = I_count
                            LoopArray = Format(TempA(i, Amatch), "0.0")
                            matchFlag = True
                            Exit For
                        End If
                      End If
                    Next I_count
                  End If
          If matchFlag Then Exit For
          Next a
      End If
    If matchFlag Then Exit For
    Next i
End Function
Public Function BhSearch(tempZ As String, TempT As String) As Long 'BH下限値検索
Dim target As String:       target = "品質基準.accdb"
Dim TempPath As String:     TempPath = QualityPath
Dim TableName As String:    TableName = "Q_S_KE_ALL_053"
Dim fieldsfilter As String: fieldsfilter = ""
Dim FilePath As String:     FilePath = TempPath & target
Dim FieldsQuality As Variant, ArrQuality As Variant
Dim C_材質 As Long, C_調質 As Long, C_バーコール下限値 As Long, C_引張強 As Long, i As Long
'------------------------------------------------------------------------------------------------------------------------
'Q-S-KE-ALL-053　　　　　　***操業基準バーコール硬度計による測定
'------------------------------------------------------------------------------------------------------------------------
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter) 'フィールド
  If IsEmpty(FieldsQuality) Then Exit Function

  For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
    Select Case FieldsQuality(1, i)
    Case "材質":             C_材質 = i
    Case "調質":             C_調質 = i
    Case "バーコール下限値": C_バーコール下限値 = i
    Case "引張強さ(Mpa)":    C_引張強 = i
    End Select
  Next i
  
  '材質の文字数を調べ、4文字以上の材質の場合は右から3文字を抜き取る
  '***DB側の材質が52S　探す側がF52Sの場合あいまい検索で引っ掛けれない
  If Len(tempZ) >= 4 Then tempZ = Right(tempZ, 3)
  Dim SQLfilter As String: SQLfilter = "[材質] LIKE " & "'*" & tempZ & "*'"
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '材質が一致した配列
  If IsEmpty(ArrQuality) Then Exit Function
  
  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    If ArrQuality(i, C_調質) = TempT Then '探す
      BhSearch = ArrQuality(i, C_バーコール下限値)
    End If
  Next i
  '一致がない場合は0を返す
End Function
Public Function DBdisplayラテラル(ロット As String, TargetText As String, Optional TempPath As String = "") As Single ', 納入Hiki As String, 取引Hiki As String
'2023.9.5_作業指示書QL-QA-023-05に伴い作成
'2024.8.5_Q-S-QA-DIM-179に統一？引数追加_個別で検索でよかったものが引き当て数分ループで探すことになったので
Dim target As String: target = "SIKALOTNOW.accdb"
If TempPath = "" Then
 TempPath = MainPath
Else
 TempPath = NextPath
End If

Dim ResultT As Boolean: ResultT = IsOver12Hours '2024.9.3

'パスセット
Dim FilePath As String: FilePath = TempPath & target
If TempPath = MainPath Then
  If ReadCheck(FilePath, 1) = False Or TimeDesignation Or FileDesignation = False Then   '1：ファイル,2：フォルダ
    TempPath = NextPath
    FilePath = TempPath & target
    If ReadCheck(NextPath & target, 1) = False Then '梱包課ﾌｧｲﾙを使う*1日更新
      If ResultT = False Then MsgBox "日稲:" & target & "がありません" & "_ReadCheck_DBdisplay", vbCritical
      Exit Function
    End If
  End If
End If

Dim TempH As String ': TempH = "a"
If TempPath = NextPath Then TempH = UF1.DBflag.Text

Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", TempH) '2022.07_引数追加  UF1.DBflag.text

Dim C_オーダー板厚 As Long, C_オーダー板幅 As Long, C_オーダー板丈 As Long, C_製造厚 As Long, C_納入先名 As Long, _
C_VC As Long, C_用途コード As Long, C_用途名 As Long, C_オーダー番号 As Long, C_材質 As Long, C_幅 As Long, C_丈 As Long, C_調質 As Long

Dim C_板厚下 As Long, C_板厚上 As Long, C_板幅下 As Long, C_板幅上 As Long, C_厚 As Long, C_備考 As Long, C_番号 As Long
Dim C_C_納入先名 As Long, C_取引先名 As Long
'------------------------------------------------------------------------------------------------------------------------
'フォールド検索
Dim i As Long, a As Long, n As Long
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "ｵｰﾀﾞｰ板厚":              C_オーダー板厚 = i
    Case "ｵｰﾀﾞｰ板幅":              C_オーダー板幅 = i
    Case "ｵｰﾀﾞｰ板丈":              C_オーダー板丈 = i
    Case "製造板厚":               C_製造厚 = i
    Case "納入先名":               C_納入先名 = i
    Case "VC":                     C_VC = i
    Case "用途ｺｰﾄﾞ":               C_用途コード = i
    Case "用途名":                 C_用途名 = i
    Case "ｵｰﾀﾞｰ番号":              C_オーダー番号 = i
    Case "製造材質":               C_材質 = i
    Case "製造板幅":               C_幅 = i
    Case "製造板丈":               C_丈 = i
    Case "製造調質":               C_調質 = i
    End Select
  Next i
ElseIf IsEmpty(HiAcc) Then
  Exit Function
End If
'------------------------------------------------------------------------------------------------------------------------
'ロット検索
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", ロット, "Record", TempH) 'UF1.DBflag.text
If IsEmpty(LotAcc) Then
  If ResultT = False Then MsgBox "SIKALOTNOWに該当ロットがありません" & "_DBdisplayﾗﾃﾗﾙ", vbCritical
  Exit Function
End If
'------------------------------------------------------------------------------------------------------------------------
'変数セット
Dim オーダー厚 As Single:    オーダー厚 = LotAcc(1, C_オーダー板厚)
Dim オーダー板幅 As Single:  オーダー板幅 = LotAcc(1, C_オーダー板幅)
Dim オーダー板丈 As Single:  オーダー板丈 = LotAcc(1, C_オーダー板丈)
Dim 製造厚 As Single:        製造厚 = LotAcc(1, C_製造厚)
Dim 製造幅 As Single:        製造幅 = LotAcc(1, C_幅)
Dim 製造丈 As Single:        製造丈 = LotAcc(1, C_丈)
Dim 納入 As String:          納入 = LotAcc(1, C_納入先名)
Dim 用途 As String:          用途 = LotAcc(1, C_用途コード) 'コード
Dim 用途名 As String:        用途名 = LotAcc(1, C_用途名) '名
Dim 受注番号 As String:      受注番号 = LotAcc(1, C_オーダー番号)
Dim 材質 As String:          材質 = LotAcc(1, C_材質)
Dim 調質 As String:          調質 = LotAcc(1, C_調質)

'------------------------------------------------------------------------------------------------
Dim Target2 As String:    Target2 = "品質基準.accdb"
Dim TempPath2 As String:  TempPath2 = QualityPath
Dim TableName As String '厚用途で変化
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String

FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant
Dim FieldsQuality As Variant
Dim TempSize

'2024.8.5_引数追加　n
If TargetText = "日車" Then
  TableName = "Q_S_QA_DIM_179_日車" '日車
ElseIf TargetText = "白銅滑川" Then
  TableName = "Q_S_QA_DIM_179_白銅滑川" '白銅滑川
End If

FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)

If IsEmpty(FieldsQuality) Then Exit Function

'Q_L_QA_023_05*****Q_S_QA_DIM_179へ
'引当コード取引コードはまぁ・・・・一緒の意味なんでスルーします
For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
  Select Case FieldsQuality(1, i)
  Case "材質":         C_材質 = i
  Case "調質":         C_調質 = i
  '2024.8.20
  Case "納入先名":     C_納入先名 = i
  Case "取引先名":     C_取引先名 = i
  
  Case "用途コード":   C_用途コード = i
  Case "板厚下":       C_板厚下 = i
  Case "板厚上":       C_板厚上 = i
  Case "厚":           C_厚 = i
  Case "板幅下":       C_板幅下 = i
  Case "板幅上":       C_板幅上 = i
  Case "幅":           C_幅 = i
  Case "備考":         C_備考 = i
  Case "備考":         C_備考 = i
  Case "管理番号":     C_番号 = i
  End Select
Next i

  '***DB側の材質が52S　探す側がF52Sの場合あいまい検索で引っ掛けれない
  Dim Ittiflag As Boolean: Ittiflag = False
'  If Len(材質) >= 4 Then 材質 = Right(材質, 3)
  SQLfilter = "" '"[材質] LIKE " & "'*" & 材質 & "*'"  '*********読み方変更_2024.5.27
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
  
  Dim MatchingArray() As Long ' マッチングしたインデックスを格納する配列
  ReDim MatchingArray(1 To UBound(ArrQuality, 1)) ' 配列のサイズを設定
  
  Dim Conditions(1 To 7) As Boolean ' 各条件の結果
  Dim j As Long
  n = 0
  
  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
'If ArrQuality(i, C_番号) = 4 Then 'Test
'Stop
'End If
  '2024.8.27_オーダー厚から製造厚に変更
    Conditions(1) = CheckCondition(i, C_材質, 材質, ArrQuality)  ' 製造材質のチェック
    Conditions(2) = Conditions(1) And CheckThickness(i, C_板厚下, C_板厚上, C_厚, 製造厚, ArrQuality) ' 厚さのチェック
    Conditions(3) = Conditions(2) And CheckWidth(i, C_板幅下, C_板幅上, C_幅, 製造幅, ArrQuality) ' 幅のチェック
    '2024.8.5
    'この条件は別のテーブルでチェックしているから不要
'    Conditions(4) = Conditions(3) And CheckWidth(i, C_納入先名, 納入Hiki, ArrQuality) '納入
'    Conditions(5) = Conditions(4) And CheckWidth(i, C_取引先名, 取引Hiki, ArrQuality) '取引
    
   ' 条件が全て満たされた場合の処理
    If Conditions(1) And Conditions(2) And Conditions(3) Then 'And Conditions(4) And Conditions(5)
      n = n + 1
      MatchingArray(n) = ArrQuality(i, C_番号)
    End If
  Next i
  
  If n = 0 Then
  '  MsgBox "データベース上にデータがありません", vbInformation
    Exit Function 'ヒットなし
  End If
  
  ' MatchingArrayから0を削除する
  Dim NewArray() As Long
  ReDim NewArray(1 To n)
  For i = 1 To n
    NewArray(i) = MatchingArray(i)
  Next i
  '格納
  Dim HitArry()
  ReDim HitArry(1 To UBound(NewArray, 1), 1 To UBound(ArrQuality, 2))
  For a = LBound(NewArray, 1) To UBound(NewArray, 1)
    For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
      For j = LBound(ArrQuality, 2) To UBound(ArrQuality, 2)
        If ArrQuality(i, C_番号) = NewArray(a) Then
          HitArry(a, j) = ArrQuality(i, j)
        End If
      Next j
    Next i
  Next a
  
'Stop
    
'  If IsEmpty(ArrQuality) Then Exit Function
'
'  ReDim ArrQ_H(1 To UBound(ArrQuality, 1), 1 To UBound(ArrQuality, 2))
'  n = 0
'
'  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
'    If ((ArrQuality(i, C_板厚下) <= オーダー厚) And (ArrQuality(i, C_板厚上) >= オーダー厚)) Then
'      If ((ArrQuality(i, C_板幅下) <= オーダー板幅) And (ArrQuality(i, C_板幅上) >= オーダー板幅)) Then  '先に幅が収まる範囲を探す
'            n = n + 1
'            For a = LBound(ArrQuality, 2) To UBound(ArrQuality, 2)
'              ArrQ_H(n, a) = ArrQuality(i, a) 'ヒットするものだけ入れる
'            Next a
'        Ittiflag = True
'      End If
'    End If
'  If Ittiflag Then Exit For
'  Next i
  
DBdisplayラテラル = Format(LoopArray(HitArry, FieldsQuality, 製造丈, 製造厚), "0.0") '2024.8.27_オーダーから製造に変更
End Function
