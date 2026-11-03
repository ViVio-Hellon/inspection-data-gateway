Attribute VB_Name = "【UFKen_11】QP_DB"
Option Explicit
Public Const C_TARGET_HOSONOS As String = "1P0001,1P0104,1P0110,1P0115,1P0116,1P0118,1P0119,1P0120" '蓋アングルの絡む包装仕様書No
' --- キャッシュ用モジュール変数 ---
Private m_PackQuality()   As Variant  ' ArrQuality のキャッシュ
Private m_CacheOrdNo      As String   ' キャッシュ対象のオーダーNo
' --- 列インデックス個別変数 ---
Private m_C_番号          As Long
Private m_C_包装仕様書     As Long
Private m_C_用途コード     As Long
Private m_C_材質           As Long
Private m_C_調質           As Long
Private m_C_板厚下         As Long
Private m_C_板厚上         As Long
Private m_C_板厚           As Long
Private m_C_板幅下         As Long
Private m_C_板幅上         As Long
Private m_C_板幅           As Long
Private m_C_板丈下         As Long
Private m_C_板丈上         As Long
Private m_C_板丈           As Long
Private m_C_使用保護材     As Long
Private m_C_備考           As Long

'========================================
Private Const TBL_H As String = "外注条件"
Private Const TBL_D As String = "外注条件明細"

'--- 外注条件：列名リスト内の位置
Private Const H_NO As Long = 0, H_NAME As Long = 1, H_RET As Long = 2
Private Const H_PRI As Long = 3, H_VALID As Long = 4

'--- 外注条件明細：列名リスト内の位置
Private Const D_HNO As Long = 0, D_CNO As Long = 1, D_TGT As Long = 2
Private Const D_MTD As Long = 3, D_VAL As Long = 4, D_VALID As Long = 5

Private m_Head As Variant          'GetRecordsArr の結果 (行, 列)
Private m_Det As Variant
Private m_HeadIdx() As Long        '上の位置 → 実際の列番号
Private m_DetIdx() As Long
Private m_Loaded As Boolean
Private m_LoadFailed As Boolean

'========================================
Private Const TBL_G As String = "ゲージ判定条件"

'--- 列名リスト内の位置
Private Const G_VALID As Long = 0
Private Const G_W_GT As Long = 1       '重量_超
Private Const G_L_GE As Long = 2       '丈_以上
Private Const G_L_LT As Long = 3       '丈_未満
Private Const G_L_GT As Long = 4       '丈_超
Private Const G_H_LT As Long = 5       '幅_未満
Private Const G_ZAI As Long = 6        '材質
Private Const G_TYU As Long = 7        '調質
Private Const G_CRS As Long = 8        'コース含む

Private m_Rule As Variant              'GetRecordsArr の結果 (行, 列)
Private m_GaugeIdx() As Long
Private m_GaugeLoaded As Boolean
Private m_GaugeLoadFailed As Boolean

'キャシュクリア
Public Sub ClearMaterialsCache()
    Erase m_PackQuality
    m_CacheOrdNo = ""
End Sub

Public Function DBQP(ロットNo As String, 検査No As String, delivery As String, Good As String, Optional TempPath As String = "", Optional Execution As String = "") As Variant
'LotNo,KenNo,納入先,取引先,パス,フォーム反映有無**2024.7.245_取引追加
'どの条件にもかからない場合
'まず Kousa(1 To 1, 1 To 9) を初期化（エラー防止）
'製造調質が "O" → 別表C記載なしO材として判定不要で返す
'製造調質が H_8 (例: H18, H28, H38, H48...) → 別表D記載なしHx8材として判定不要で返す
'それ以外 → 等級（フラットネス3級/4級）で公差を算出する処理へ続く
'Exit Function（空/Empty返し）するわけではなく、フラットネス品質等級（3級/4級）にフォールバックして公差を算出して返す

Dim i As Long, n As Long, a As Long
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
      MsgBox "日稲:" & target & "がありません" & "_ReadCheck_DBQP", vbCritical
      Exit Function
    End If
  End If
End If
Dim TempH As String: TempH = ""
'If TempPath = NextPath Then
'  If Test.Visible Then
'    TempH = "A"
'  Else
'    TempH = UF1.DBflag.Text
'  End If
'End If

Dim HiAcc: HiAcc = Arrbox("SIKALOTNOW.accdb", "板", "Fields", TempH) '2022.07_引数追加  UF1.DBflag.text

Dim C_品質グレード_表面処理 As Long, C_納入先名 As Long, C_納入先コード As Long, C_取引先名 As Long, C_取引先コード As Long, C_用途コード As Long, C_用途名 As Long _
, C_製造厚 As Long, C_製造幅 As Long, C_製造丈 As Long, C_製造材質 As Long, C_調質 As Long, C_送り As Long, C_送り先コード As Long _
, C_オーダー板厚 As Long, C_オーダー板幅 As Long, C_オーダー板丈 As Long, C_オーダー材質 As Long, C_オーダー調質 As Long _
, C_品質グレード_フラットネス As Long, C_横割数 As Long, C_中歪_高 As Long, C_耳歪_高 As Long, C_巾方向高 As Long, C_丈方向高 As Long, C_Course As Long

'------------------------------------------------------------------------------------------------------------------------
'フォールド検索
If Not IsEmpty(HiAcc) Then
  For i = 1 To UBound(HiAcc, 2)
    Select Case HiAcc(1, i)
    Case "品質ｸﾞﾚｰﾄﾞ_表面処理":    C_品質グレード_表面処理 = i
    Case "納入先名":               C_納入先名 = i
    Case "納入先ｺｰﾄﾞ":             C_納入先コード = i
    
    Case "送り先名":               C_送り = i '2025.7.23
    Case "送り先ｺｰﾄﾞ":             C_送り先コード = i '2025.7.23
    
    Case "取引先名":               C_取引先名 = i '2024.7.24
    Case "取引先ｺｰﾄﾞ":             C_取引先コード = i
    
    Case "用途ｺｰﾄﾞ":               C_用途コード = i
    Case "用途名":                 C_用途名 = i
    
    Case "製造板厚":               C_製造厚 = i
    Case "製造板幅":               C_製造幅 = i
    Case "製造板丈":               C_製造丈 = i
    Case "製造材質":               C_製造材質 = i
    Case "製造調質":               C_調質 = i
    
    Case "ｵｰﾀﾞｰ板厚":              C_オーダー板厚 = i
    Case "ｵｰﾀﾞｰ板幅":              C_オーダー板幅 = i
    Case "ｵｰﾀﾞｰ板丈":              C_オーダー板丈 = i
    Case "ｵｰﾀﾞｰ材質":              C_オーダー材質 = i
    Case "ｵｰﾀﾞｰ調質":              C_オーダー調質 = i
    
    Case "実績_設備ｺｰｽ":           C_Course = i
    
    Case "品質ｸﾞﾚｰﾄﾞ_ﾌﾗｯﾄﾈｽ":      C_品質グレード_フラットネス = i
    Case "BOX設計_横割数(板)":     C_横割数 = i
    
    Case "指示_中歪_高さ":         C_中歪_高 = i
    Case "指示_耳歪_高さ":         C_耳歪_高 = i
    Case "指示_そり巾_方向高さ":   C_巾方向高 = i '製板カードに表示されないメモ書き程度の数値_2024.12.26　不要
    Case "指示_そり丈_方向高さ":   C_丈方向高 = i
    End Select
  Next i
ElseIf IsEmpty(HiAcc) Then
  Exit Function
End If


'ロット検索
Dim LotAcc: LotAcc = Arrbox("SIKALOTNOW.accdb", ロットNo, "Record", TempH) 'UF1.DBflag.text
If IsEmpty(LotAcc) Then
  MsgBox "SIKALOTNOWに該当ロットがありません" & "_DBQP", vbCritical
  Exit Function
End If
'検査No
Dim KenA As String: KenA = Left(検査No, 1)
If KenA = "0" Or KenA = "" Then
'Q-S-QA-ALL-088_検査番号の付け方に基づき作る
Dim InputString As String: InputString = LotAcc(1, C_Course)
Dim Result_C As String:    Result_C = FindPriority(InputString) 'コース最終圧延検索
  If Result_C <> "" Then KenA = Result_C

  If Result_C = "" Then
    MsgBox "検査Noがないファイルです" & "_DBQP", vbCritical
    Exit Function
  End If
End If
'------------------------------------------------------------------------------------------------------------------------
'変数セット
Dim オーダー厚 As Single:    オーダー厚 = LotAcc(1, C_オーダー板厚)
Dim オーダー板幅 As Single:  オーダー板幅 = LotAcc(1, C_オーダー板幅)
Dim オーダー板丈 As Single:  オーダー板丈 = LotAcc(1, C_オーダー板丈)
Dim 製造厚 As Single:        製造厚 = LotAcc(1, C_製造厚)
Dim 製造幅 As Single:        製造幅 = LotAcc(1, C_製造幅)
Dim 製造丈 As Single:        製造丈 = LotAcc(1, C_製造丈)
Dim フラットネス As Long:    フラットネス = LotAcc(1, C_品質グレード_フラットネス)
Dim グレード As String:      グレード = LotAcc(1, C_品質グレード_表面処理)
Dim 納入 As String:          納入 = LotAcc(1, C_納入先名)
Dim 納入コード As String:    納入コード = LotAcc(1, C_納入先コード)

Dim 送り先 As String:        送り先 = LotAcc(1, C_送り) '2025.7.23
Dim 送先コード As String:    送先コード = LotAcc(1, C_送り先コード) '2025.7.23

Dim 取引 As String:          取引 = LotAcc(1, C_取引先名) '2024.7.24
Dim 取引コード As String:    取引コード = LotAcc(1, C_取引先コード)

Dim 用途 As String:          用途 = LotAcc(1, C_用途コード) 'コード
Dim 用途名 As String:        用途名 = LotAcc(1, C_用途名) '名
Dim 製造材質 As String:      製造材質 = LotAcc(1, C_製造材質)
Dim 製造調質 As String:      製造調質 = LotAcc(1, C_調質)
Dim オーダー材質 As String:  オーダー材質 = LotAcc(1, C_オーダー材質)
Dim オーダー調質 As String:  オーダー調質 = LotAcc(1, C_オーダー調質)

'必ず指示があるわけではないので文字列宣言
Dim 横割数 As String:        横割数 = LotAcc(1, C_横割数)
Dim 中歪 As String:          中歪 = LotAcc(1, C_中歪_高) 'カード左上
Dim 耳歪 As String:          耳歪 = LotAcc(1, C_耳歪_高)
Dim 巾歪 As String:          巾歪 = LotAcc(1, C_巾方向高) 'カード右下(ｺﾒﾝﾄ欄)
Dim 丈歪 As String:          丈歪 = LotAcc(1, C_丈方向高)

'--------------------------------------------------------------------
Dim Target2 As String:    Target2 = "品質基準.accdb"
Dim TempPath2 As String:  TempPath2 = QualityPath  '"C:\Users\xfa-ngykonpo\Desktop\DB\品質\"
Dim TableName As String
Dim fieldsfilter As String:  fieldsfilter = ""
Dim SQLfilter As String

'2026.05.12_板丈下,板丈上追加
Dim C_平坦度上限値参照元 As Long, C_板厚下 As Long, C_板厚上 As Long, C_1m当たり平坦度上限 As Long _
, C_納入先 As Long, C_取引先 As Long, C_巾取数幅分割数のこと As Long, C_検査番号の頭 As Long _
, C_板幅下 As Long, C_板幅上 As Long, C_板丈下 As Long, C_板丈上 As Long _
, C_コメント As Long, C_番号, C_製造調質 As Long, C_中耳 As Long, C_フラグ As Long, C_更新日 As Long _
, C_送り先 As Long

FilePath = TempPath2 & Target2 '再定義
Dim ArrQuality As Variant
Dim FieldsQuality As Variant

TableName = "平坦度一覧QP追加"
FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)

If IsEmpty(FieldsQuality) Then Exit Function

'平坦度一覧QP追加
For i = LBound(FieldsQuality, 2) To UBound(FieldsQuality, 2)
  Select Case FieldsQuality(1, i)
  Case "平坦度上限値参照元":    C_平坦度上限値参照元 = i
  Case "用途コード":            C_用途コード = i
  Case "板厚下":                C_板厚下 = i
  Case "板厚上":                C_板厚上 = i
  Case "1m当たり平坦度上限":    C_1m当たり平坦度上限 = i
  Case "納入先":                C_納入先 = i
  Case "納入先コード":          C_納入先コード = i
  
  Case "送り先":                C_送り先 = i '2025.7.23
  Case "送り先コード":          C_送り先コード = i '2025.7.23
  
  Case "取引先":                C_取引先 = i '2024.7.24
  Case "取引先コード":          C_取引先コード = i
  
  Case "巾取数幅分割数のこと":  C_巾取数幅分割数のこと = i
  Case "検査番号の頭":          C_検査番号の頭 = i
  
  Case "板幅下":                C_板幅下 = i
  Case "板幅上":                C_板幅上 = i
  Case "板丈下":                C_板丈下 = i '2026.05.12
  Case "板丈上":                C_板丈上 = i '2026.05.12
  
  Case "オーダー材質":          C_オーダー材質 = i
  Case "オーダー調質":          C_オーダー調質 = i
  Case "製造調質":              C_製造調質 = i
  Case "製造材質":              C_製造材質 = i
  Case "用途名":                C_用途名 = i
  Case "表示コメント":          C_コメント = i
  Case "管理番号":              C_番号 = i
  Case "オーダー板厚":          C_オーダー板厚 = i
  Case "オーダー板幅":          C_オーダー板幅 = i
  Case "オーダー板丈":          C_オーダー板丈 = i
  
  Case "中歪・耳歪":            C_中耳 = i
  Case "フラグ":                C_フラグ = i
  Case "更新日":                C_更新日 = i
  End Select
Next i

Dim 厚下 As Single, 厚上 As Single, 幅下 As Single, 幅上 As Single
Dim TmpString As String, SearchString As String, flag(1 To 18) As Boolean
'フル検索用
'SQLfilter = ""
'ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)

'検索一致レコードと空白フィールド
SQLfilter = "[検査番号の頭] LIKE " & "'*" & KenA & "*' OR [検査番号の頭] = " & " '" & "" & "'"
ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter) '検査番号の頭が一致した配列

If IsEmpty(ArrQuality) Then Exit Function

Dim UpdateDay As String: UpdateDay = ArrQuality(1, C_更新日)

Dim Conditions(1 To 15) As Boolean  ' 各条件の結果を格納する変数****_2026.05.12 配列サイズ拡張

Dim MatchingArray() As Long ' マッチングしたインデックスを格納する配列
ReDim MatchingArray(1 To UBound(ArrQuality, 1)) ' 配列のサイズを設定

'幅丈のインデックス
Dim HitHaba As Variant: HitHaba = FindIndices(FieldsQuality, "H", オーダー板幅)
Dim HitTake As Variant: HitTake = FindIndices(FieldsQuality, "T", オーダー板丈)
Dim SoriHaba As Variant, SoriTake As Variant

'Stop
Dim SearchS As String '参照元用
n = 0

For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
'If ArrQuality(i, C_番号) = 50 Then 'Test
'Stop
'End If
    Conditions(1) = CheckCondition(i, C_用途コード, 用途, ArrQuality) ' 用途コードのチェック
    Conditions(2) = Conditions(1) And CheckCondition(i, C_製造材質, 製造材質, ArrQuality) '製造材質のチェック
    Conditions(3) = Conditions(2) And CheckCondition(i, C_オーダー材質, オーダー材質, ArrQuality) 'オーダー材質のチェック
    Conditions(4) = Conditions(3) And CheckCondition(i, C_製造調質, 製造調質, ArrQuality) '製造調質のチェック
    Conditions(5) = Conditions(4) And CheckCondition(i, C_オーダー調質, オーダー調質, ArrQuality) ' オーダー調質のチェック
    Conditions(6) = Conditions(5) And CheckThickness(i, C_板厚下, C_板厚上, C_オーダー板厚, オーダー厚, ArrQuality) '厚さのチェック
    Conditions(7) = Conditions(6) And CheckWidth(i, C_板幅下, C_板幅上, C_オーダー板幅, オーダー板幅, ArrQuality) '幅のチェック
    Conditions(8) = Conditions(7) And CheckWidth(i, C_板丈下, C_板丈上, C_オーダー板丈, オーダー板丈, ArrQuality) '丈のチェック_2026.05.12 条件追加
    Conditions(9) = Conditions(8) And CheckCondition(i, C_納入先, delivery, ArrQuality) '納入先のチェック
    Conditions(10) = Conditions(9) And CheckCondition(i, C_納入先コード, 納入コード, ArrQuality) '納入先コードのチェック
    
    Conditions(11) = Conditions(10) And CheckCondition(i, C_取引先, Good, ArrQuality) '2024.7.24'取引先のチェック
    Conditions(12) = Conditions(11) And CheckCondition(i, C_取引先コード, 取引コード, ArrQuality) '2024.7.24'取引先コードのチェック
    
    Conditions(13) = Conditions(12) And CheckCondition(i, C_送り先, 送り先, ArrQuality) '2025.7.23'送り先のチェック
    Conditions(14) = Conditions(13) And CheckCondition(i, C_送り先コード, 送先コード, ArrQuality) '2025.7.23'送り先コードのチェック
    
    Conditions(15) = Conditions(14) And CheckSize(i, C_巾取数幅分割数のこと, 横割数, ArrQuality) ' 巾取数幅分割数のことのチェック
  
   ' 条件が全て満たされた場合の処理
    If Conditions(1) And Conditions(2) And Conditions(3) And Conditions(4) And Conditions(5) And _
    Conditions(6) And Conditions(7) And Conditions(8) And Conditions(9) And Conditions(10) _
    And Conditions(11) And Conditions(12) And Conditions(13) Then
      n = n + 1
      MatchingArray(n) = ArrQuality(i, C_番号)
    End If
Next i

If n = 0 Then
'  MsgBox "データベース上にデータがありません", vbInformation
  GoTo SkipQP '----------------------------------------------------------フラットネス等級公差へ
  Exit Function 'ヒットなし
End If

' MatchingArrayから0を削除する
Dim NewArray() As Long
ReDim NewArray(1 To n)
For i = 1 To n
    NewArray(i) = MatchingArray(i)
Next i

'納入その他の優先があるかのチェック 基本的に番号が小さいものが優先だが下記3つは最優先
'取引先、取引先コード追加_2024.7.24
'送り先、送り先コード追加_2025.7.23
Dim PriorityFlag As Boolean, PriorityNo As Long, CountM As Long
For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
  For a = LBound(NewArray, 1) To UBound(NewArray, 1)
    If NewArray(a) <> 0 Then
      CountM = NewArray(a)
      If ArrQuality(i, C_番号) = CountM Then
        If ArrQuality(i, C_納入先) <> "" Or ArrQuality(i, C_納入先コード) <> "" _
        Or ArrQuality(i, C_取引先) <> "" Or ArrQuality(i, C_取引先コード) <> "" _
        Or ArrQuality(i, C_送り先) <> "" Or ArrQuality(i, C_送り先コード) <> "" _
        Or ArrQuality(i, C_巾取数幅分割数のこと) <> "" Then
          PriorityNo = CountM
          PriorityFlag = True
          Exit For
        End If
      '複数ヒットした場合に
      If PriorityFlag Then Exit For
      End If
    If PriorityFlag Then Exit For
    End If
    If PriorityFlag Then Exit For
  Next a
If PriorityFlag Then Exit For
Next i

If PriorityFlag Then NewArray(LBound(NewArray, 1)) = PriorityNo '「空白項目でないものが優先」 上書きする

Dim PositionFlag As Long, PositionTemp As String

'公差格納
Dim Kousa()
ReDim Kousa(1 To UBound(NewArray, 1), 1 To 9) 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ

For a = LBound(NewArray, 1) To UBound(NewArray, 1)
  For i = LBound(ArrQuality, 1) To UBound(ArrQuality, 1)
    If (ArrQuality(i, C_番号) = NewArray(a)) Then    '納入先がある場合は優先
      Kousa(a, 1) = ArrQuality(i, C_1m当たり平坦度上限) 'ゲージ
      Kousa(a, 2) = ArrQuality(i, HitHaba(1)) '幅歪
      Kousa(a, 3) = ArrQuality(i, HitTake(1)) '丈歪
      Kousa(a, 4) = ArrQuality(i, C_中耳) '中歪耳歪フラグ　中歪
      Kousa(a, 5) = "" 'フラグがあれば後で耳歪をいれる
      Kousa(a, 6) = ArrQuality(i, C_コメント) 'コメント
      Kousa(a, 7) = ArrQuality(i, C_平坦度上限値参照元) '参照元
      Kousa(a, 8) = ArrQuality(i, C_番号) '番号
      Kousa(a, 9) = ArrQuality(i, C_フラグ) 'フラグ
    End If
  Next i
Next a
'-----------------------------------------------------DB下部コメントが必要なものの処理

Dim FlCount As Long: FlCount = UBound(Kousa, 1)
'重複可能。コメントのみ表示　フラグ列にこのワードがある
Dim Kakunou() As Long, Count As Long, KakuC As Long
If FlCount >= 2 Then '格納した公差が2つ以上ある
  For i = LBound(Kousa, 1) To UBound(Kousa, 1)
    If Kousa(i, 9) <> "" Then 'フラグの有無
      PositionTemp = Kousa(i, 9)
      PositionFlag = InStr(1, PositionTemp, "重複可") '注意コメントを重ねる(重複可がある時だけ)
      If PositionFlag > 0 Then
        Count = Count + 1 '"重複可"の出現回数
        ReDim Preserve Kakunou(1 To Count) ' 配列のサイズを増やす
        Kakunou(Count) = i ' 位置を格納する
      End If
    End If
 Next i
End If

If Count <> 0 Then ' "重複可"あり
  For i = LBound(Kakunou, 1) To UBound(Kakunou, 1)
    KakuC = Kakunou(i)
'    If Kousa(LBound(Kousa, 1), 6) <> Kousa(KakuC, 6) Then '同じコメントではない
    If InStr(1, Kousa(LBound(Kousa, 1), 6), Kousa(KakuC, 6)) = 0 Then ' Kousa(LBound(Kousa, 1), 6) に Kousa(KakuC, 6) が含まれていない場合
      Kousa(LBound(Kousa, 1), 6) = Kousa(LBound(Kousa, 1), 6) & vbCrLf & Kousa(KakuC, 6) 'コメント重ね
    End If
  Next i
End If


'DBサイズ越があった場合***********2024.4.11　不要DB内コメントで対応する
'別表A,B許可あり　C,Dは幅丈の条件がそもそもない　QPはﾋｽﾞﾐ･ｿﾘ_ﾊﾊﾞ_1.2mmｲｶ/M,ﾋｽﾞﾐ･ｿﾘ_ﾀｹ_1.2mmｲｶ/Mの時の対応⇒削除するとのこと
'*************技術確認が必要なもの確認***幅1800超え　かつ別表Bを参照するもの_2024.4.5************************
If HitHaba(UBound(HitHaba)) = 9999 Then
  Dim SourceTemp As String:  SourceTemp = Kousa(LBound(Kousa, 1), 7) '参照元
  If (InStr(SourceTemp, "別表A") > 0) Then
    Kousa(LBound(Kousa, 1), 6) = Kousa(LBound(Kousa, 1), 6) & vbCrLf & "3054mm越えは3050mmと同様の公差"
  End If
End If

'-------------------------------------------------------------------------------------------------

'データベース上にデータがあるが等級で公差を見る必要がある
Dim MinimumValue As Long: If Not IsEmpty(Kousa) Then MinimumValue = LBound(Kousa, 1)
If IsEmpty(Kousa) Then MinimumValue = 1

'基準書＞QP、マスター＞製品規格
'中耳LoopArray　ソリFindIndices

'等級チェック
If フラットネス <> 3 And フラットネス <> 4 Then '3でも4でもない
  MsgBox "歪等級が3,4以外です", vbCritical
  Exit Function
End If
'中歪耳歪
Dim TableN As String, TableM As String, Position As Long, UpFlag As Boolean, DowFlag As Boolean, TablFlag As Boolean
If Not IsEmpty(Kousa) Then
  '特別条件
  If Kousa(MinimumValue, 4) = "不問" Or Kousa(MinimumValue, 4) = "不問とする。" Then
'    GoTo skipEnd   '公差確認不要
    GoTo SkipNoQuestionsAsked '2024.8.8_フラットネス品質等級の場合にスキップしてしまうため修正
  End If
  
  If Kousa(MinimumValue, 4) = 0 Or Kousa(MinimumValue, 4) = "無き事" Then
    If MinimumValue = 0 Then MinimumValue = 1
    GoTo SkipEndA '公差確認不要
  End If
  
  If Kousa(MinimumValue, 4) = "" Then GoTo SkipT '検査要
  
  SQLfilter = ""
  'テーブル設定
  If Kousa(MinimumValue, 4) <> "" Then
    Dim Elements() As String: Elements = Split(Kousa(MinimumValue, 4), ",")
    Dim has3Grade As Long, has4Grade As Long, Result As Single, tMP As String
    Dim WhatUp As Variant, WhatDow As Variant
    
'    For i = LBound(Elements) To UBound(Elements)
       '各要素に対して処理
       tMP = Kousa(MinimumValue, 4) 'Splitなし
       
       WhatUp = ExtractNums(tMP, "以上") 'Tmpに指定文字列があるかチェックしあればその文字列から数値をぬきとる
       WhatDow = ExtractNums(tMP, "以下")
       If WhatUp <> "" And WhatUp <> " " Then UpFlag = True   '指定文字列がなければ""が返ってくる
       If WhatDow <> "" And WhatUp <> " " Then DowFlag = True
       
       If UpFlag Then Result = val(ExtractNums(tMP, "以上"))
       If DowFlag Then Result = val(ExtractNums(tMP, "以下"))
       
       has3Grade = InStr(tMP, "3級")
       has4Grade = InStr(tMP, "4級")
    
       If has3Grade > 0 And (UpFlag Or DowFlag) Then ' 3級かつ以上以下が含まれる
         If CompareStrings(tMP, オーダー厚, Result) Then '指定板厚の範囲内か
           TableN = "中歪3級"
           TableM = "耳歪3級"
           SearchS = "Q-E-001_品質規格:等級3"
           TablFlag = True
         End If
       ElseIf has4Grade > 0 And (UpFlag Or DowFlag) Then  '4級かつ以上以下が含まれる
         If CompareStrings(tMP, オーダー厚, Result) Then '指定板厚の範囲内か
           TableN = "中歪4級"
           TableM = "耳歪4級"
           SearchS = "Q-E-001_品質規格:等級4"
           TablFlag = True
         End If
       ElseIf has3Grade > 0 And (UBound(Elements) = 0) Then '3級しかない
         TableN = "中歪3級"
         TableM = "耳歪3級"
         SearchS = "Q-E-001_品質規格:等級3"
         TablFlag = True
       ElseIf has4Grade > 0 And (UBound(Elements) = 0) Then '4級しかない
         TableN = "中歪4級"
         TableM = "耳歪4級"
         SearchS = "Q-E-001_品質規格:等級4"
         TablFlag = True
       End If
'    Next i
    
    If TablFlag = False Then Kousa(MinimumValue, 4) = "" 'フラグはあるが条件から外れた場合
    
     '以上以下を含まない(直接公差記載)
    If UpFlag = False And DowFlag = False Then '以上以下含まずmmもない　かつ　数値
      Dim mimiS As Variant, nakaS As Variant
      nakaS = ExtractNums(tMP, "中") 'Tmpに指定文字列があるかチェックしあればその文字列から数値をぬきとる　,区切り
      mimiS = ExtractNums(tMP, "耳")
      If nakaS <> "" Then nakaS = val(nakaS)
      If mimiS <> "" Then mimiS = val(mimiS)
      Kousa(MinimumValue, 4) = nakaS
      Kousa(MinimumValue, 5) = mimiS
      GoTo skipEnd
    End If
    
    '以上以下は含むが等級指示がない
    If has3Grade = 0 And has4Grade = 0 Then
      If CompareStrings(tMP, オーダー厚, Result) Then
        If IsEmpty(Kousa) Then ReDim Kousa(1 To 1, 1 To 10) 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号
        '例　3mm以下　品質等級　の場合3mm以下の確認後等級で見る必要がある
        
        If n = 0 Then MinimumValue = 1
        Select Case フラットネス
        Case 3
          TableN = "中歪3級"
          TableM = "耳歪3級"
          TableName = "ソリ3級"
          SearchS = "Q-E-001_品質規格:等級3"
        Case 4
          TableN = "中歪4級"
          TableM = "耳歪4級"
          TableName = "ソリ4級"
          SearchS = "Q-E-001_品質規格:等級4"
        End Select
        '中歪
        SQLfilter = ""
        FieldsQuality = GetFieldsArr(FilePath, TableN, fieldsfilter)
        If IsEmpty(FieldsQuality) Then Exit Function
        ArrQuality = GetRecordsArr(FilePath, TableN, SQLfilter)
        If IsEmpty(ArrQuality) Then Exit Function
        Kousa(MinimumValue, 4) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0") '丈の長さで公差算出
        '耳歪
        FieldsQuality = GetFieldsArr(FilePath, TableM, fieldsfilter)
        If IsEmpty(FieldsQuality) Then Exit Function
        ArrQuality = GetRecordsArr(FilePath, TableM, SQLfilter)
        If IsEmpty(ArrQuality) Then Exit Function
        Kousa(MinimumValue, 5) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
        
      End If
    End If
  End If 'Kousa(MinimumValue, 4) <> ""
  
  'テーブル=""の場合中耳の公差不要
  If TableN = "" Then GoTo SkipT
  
  '中歪から
    FieldsQuality = GetFieldsArr(FilePath, TableN, fieldsfilter)
    If IsEmpty(FieldsQuality) Then Exit Function
    ArrQuality = GetRecordsArr(FilePath, TableN, SQLfilter)
    If IsEmpty(ArrQuality) Then Exit Function
    Kousa(MinimumValue, 4) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0") '丈の長さで公差算出
  '耳歪
    FieldsQuality = GetFieldsArr(FilePath, TableM, fieldsfilter)
    If IsEmpty(FieldsQuality) Then Exit Function
    ArrQuality = GetRecordsArr(FilePath, TableM, SQLfilter)
    If IsEmpty(ArrQuality) Then Exit Function
    Kousa(MinimumValue, 5) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
    
    If InStr(Kousa(MinimumValue, 7), SearchS) = 0 Then Kousa(MinimumValue, 7) = Kousa(MinimumValue, 7) & " " & SearchS
    
End If 'Not IsEmpty(Kousa)

SkipT:
'耳中検査不要
'ソリの公差
'TableName = "平坦度一覧QP追加"
'FieldsQuality = GetFieldsArr(filePath, TableName, fieldsfilter)

'---------------------------------------------------------------------------------------------'公差確認不要(判定不要)_判定不要なだけで0ではない
SkipNoQuestionsAsked:


If Not IsEmpty(Kousa) Then
  If Kousa(MinimumValue, 2) = "フラットネス品質等級" Or Kousa(MinimumValue, 3) = "フラットネス品質等級" Then '等級での歪公差
  '等級チェック
  If フラットネス <> 3 And フラットネス <> 4 Then '3でも4でもない
    MsgBox "歪等級が3,4以外です", vbCritical
    GoTo skipEnd
  End If
  'ソリ
  Select Case フラットネス
  Case 3
    TableName = "ソリ3級"
    SearchS = "Q-E-001_品質規格:等級3"
  Case 4
    TableName = "ソリ4級"
    SearchS = "Q-E-001_品質規格:等級4"
  End Select
  SQLfilter = ""
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(FieldsQuality) Then Exit Function
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(ArrQuality) Then Exit Function
  
  '幅丈のインデックス
  SoriHaba = FindIndices(FieldsQuality, "H", オーダー板幅)
  SoriTake = FindIndices(FieldsQuality, "T", オーダー板丈)
  
  Kousa(MinimumValue, 2) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriHaba(1)), "0.0")
  Kousa(MinimumValue, 3) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriTake(1)), "0.0")
  If InStr(Kousa(MinimumValue, 7), SearchS) = 0 Then Kousa(MinimumValue, 7) = Kousa(MinimumValue, 7) & " " & SearchS
  End If
End If

'Stop

GoTo skipEnd
'Kousa(MinimumValue, 4) = "不問"
'[[[[DB上にデータなし]]]]等級で公差をだす
'Kousa(MinimumValue, 5) はn=0でスキップしてきているので使用できない

SkipQP:
'DBで合う条件がない
If n = 0 Then ReDim Kousa(1 To 1, 1 To 9) 'kousa配列がないと後々エラーになる

'製造調質のチェック_______DB上にデータなしかつO,Hx8のチェック
If 製造調質 = "O" Then
'  DBQP = "O材一致なし"
  Kousa(1, 6) = "別表Cに記載のないO材_判定不要"
  Kousa(1, 7) = "Q-S-QA-DIM-122_別表C"
  DBQP = Kousa
  If Execution = "" Then
    If UF1.Visible Then
      UF1.HizCome.Caption = "別表Cに記載のないO材_判定不要"
      UF1.DistortionFrame.Caption = "参照元:" & "Q-S-QA-DIM-122_別表C"
    End If
  End If
  Exit Function
End If
If Len(製造調質) = 3 And Left(製造調質, 1) = "H" And Right(製造調質, 1) = "8" Then
  Kousa(1, 6) = "別表Dに記載のないHx8材_判定不要"
  Kousa(1, 7) = "Q-S-QA-DIM-122_別表D"
  DBQP = Kousa
  If Execution = "" Then If UF1.Visible Then UF1.HizCome.Caption = "別表Dに記載のないHx8材_判定不要"
  Exit Function
End If

'---------------------------------------------------------------------------------------------
'基準書の条件にない　かつ　マスター(QPではない、技術がいじる数値)にない
'平坦度一覧QP追加 ＞ SIKALOTNOWではあるが平坦度一覧QP追加が中耳歪の条件をすべて網羅していない平坦度一覧QP追加に条件記載があればそれが最優先、空白であれば等級検索　SIKALOTNOWに数値があればSIKALOTNOW　＞　等級
  If IsEmpty(Kousa) Then ReDim Kousa(1 To 1, 1 To 9) 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号
  
  If n = 0 Then MinimumValue = 1
  
  Select Case フラットネス
  Case 3
    TableN = "中歪3級"
    TableM = "耳歪3級"
    TableName = "ソリ3級"
    SearchS = "Q-E-001_品質規格:等級3"
  Case 4
    TableN = "中歪4級"
    TableM = "耳歪4級"
    TableName = "ソリ4級"
    SearchS = "Q-E-001_品質規格:等級4"
  End Select

    'ソリ
    SQLfilter = ""
    FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(FieldsQuality) Then Exit Function
    ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
    If IsEmpty(ArrQuality) Then Exit Function
    
    '幅丈のインデックス
    SoriHaba = FindIndices(FieldsQuality, "H", オーダー板幅)
'    Dim countCol: countCol = SoriHaba(1) 'テスト用
'    MsgBox FieldsQuality(1, countCol)
    SoriTake = FindIndices(FieldsQuality, "T", オーダー板丈)
'    Dim countCol2: countCol2 = SoriTake(1)
'    MsgBox FieldsQuality(1, countCol2)
    
    Kousa(MinimumValue, 2) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriHaba(1)), "0.0")
    Kousa(MinimumValue, 3) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriTake(1)), "0.0")
    If InStr(Kousa(MinimumValue, 7), SearchS) = 0 Then Kousa(MinimumValue, 7) = Kousa(MinimumValue, 7) & " " & SearchS
            
            
If 6# >= オーダー厚 Then '中耳の基本条件
  If (中歪 = "") Or (耳歪 = "") Then
      '中歪
      SQLfilter = ""
      FieldsQuality = GetFieldsArr(FilePath, TableN, fieldsfilter)
      If IsEmpty(FieldsQuality) Then Exit Function
      ArrQuality = GetRecordsArr(FilePath, TableN, SQLfilter)
      If IsEmpty(ArrQuality) Then Exit Function
      Kousa(MinimumValue, 4) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0") '丈の長さで公差算出
      '耳歪
      FieldsQuality = GetFieldsArr(FilePath, TableM, fieldsfilter)
      If IsEmpty(FieldsQuality) Then Exit Function
      ArrQuality = GetRecordsArr(FilePath, TableM, SQLfilter)
      If IsEmpty(ArrQuality) Then Exit Function
      Kousa(MinimumValue, 5) = Format(LoopArray(ArrQuality, FieldsQuality, オーダー板丈, オーダー厚), "0.0")
      
      GoTo skipEnd
  Else
    GoTo skipEnd
  End If
Else '中　耳　見る必要がないサイズ
  If MinimumValue = 0 Then MinimumValue = 1
  
  Kousa(MinimumValue, 4) = ""
  Kousa(MinimumValue, 5) = ""
  
  
  '等級チェック
  If フラットネス <> 3 And フラットネス <> 4 Then '3でも4でもない
    MsgBox "歪等級が3,4以外です", vbCritical
    GoTo skipEnd
  End If
  'ソリ
  Select Case フラットネス
  Case 3
    TableName = "ソリ3級"
    SearchS = "Q-E-001_品質規格:等級3"
  Case 4
    TableName = "ソリ4級"
    SearchS = "Q-E-001_品質規格:等級4"
  End Select
  SQLfilter = ""
  FieldsQuality = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(FieldsQuality) Then Exit Function
  ArrQuality = GetRecordsArr(FilePath, TableName, SQLfilter)
  If IsEmpty(ArrQuality) Then Exit Function
  
  '幅丈のインデックス
  SoriHaba = FindIndices(FieldsQuality, "H", オーダー板幅)
  SoriTake = FindIndices(FieldsQuality, "T", オーダー板丈)
  
  Kousa(MinimumValue, 2) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriHaba(1)), "0.0")
  Kousa(MinimumValue, 3) = Format(SoriArray(ArrQuality, FieldsQuality, オーダー厚, オーダー板幅, オーダー板丈, SoriTake(1)), "0.0")
  If InStr(Kousa(MinimumValue, 7), SearchS) = 0 Then Kousa(MinimumValue, 7) = Kousa(MinimumValue, 7) & " " & SearchS

GoTo skipEnd
End If
'---------------------------------------------------------------------------------------------'公差確認不要(歪無き事)_歪があったらダメなので0
SkipEndA:
Kousa(MinimumValue, 4) = 0
Kousa(MinimumValue, 5) = 0
'---------------------------------------------------------------------------------------------
'Stop

'フォームに反映
skipEnd:
If MinimumValue = 0 Then MinimumValue = 1
' Kousa(MinimumValue, 4) = "不問"
'不問とする　をけす
If Kousa(MinimumValue, 1) = "不問とする。" Then Kousa(MinimumValue, 1) = ""
If Kousa(MinimumValue, 4) = "不問とする。" Then Kousa(MinimumValue, 4) = ""
If Kousa(MinimumValue, 4) = "不問" Then Kousa(MinimumValue, 4) = ""
If Kousa(MinimumValue, 5) = "不問" Then Kousa(MinimumValue, 5) = ""
'基準書の条件にない　かつ　マスター(QPではない、技術がいじる数値 SIKALOT)にある
If Kousa(MinimumValue, 4) = "" And 中歪 <> "" Then Kousa(MinimumValue, 4) = 中歪 '中歪
If Kousa(MinimumValue, 5) = "" And 耳歪 <> "" Then Kousa(MinimumValue, 5) = 耳歪 '耳歪

DBQP = Kousa 'Format(Kousa(MinimumValue, 1), "0.0") & "," & Format(Kousa(MinimumValue, 2), "0.0") & "," & Format(Kousa(MinimumValue, 3), "0.0") & "," & Format(Kousa(MinimumValue, 4), "0.0") & "," & Format(Kousa(MinimumValue, 5), "0.0") & "," & Kousa(MinimumValue, 6) & "," & Kousa(MinimumValue, 7) & "," & Kousa(MinimumValue, 8)

If Execution <> "" Then Exit Function 'フォームに反映なし Test用 & vbCrLf & _

'反映前に数値比較がいる_2024.4.24
With UF品質基準
  .Uatu = Format(製造厚, "0.000")
  .Uhaba = Format(製造幅, "0.0")
  .Utake = Format(製造丈, "0.0")
  .Uzai = 製造材質
  .Utyou = 製造調質
  .UYoucode = 用途
  .Uhyou = グレード
  .Noun = delivery
'ﾘｾｯﾄ
  .KousinL2.Caption = ""
  .KousinL3.Caption = ""
  .KousinL2.Caption = UpdateDay
  .KousinL3.Caption = UpdateDay
End With
End Function

'区切られた文字列 Tmp から変数 List を探し、Tmp 内の文字列の条件に基づいて True または False を返す関数,Trueで条件通過
Function FindInList(tMP As String, List As String, Optional BlankFlag As String = "") As Boolean
'Tmp=探される　List=探す
Dim TmpArray() As String
Dim Item As Variant
Dim ExcludedItems As String
Dim TempLength As Long

If BlankFlag = "" Then
  ' Tmpが空文字列の場合はTrueを返す
  ' BlankFlagに何か渡した場合はこのチェックをスキップ
  If tMP = "" Then
    FindInList = True
    Exit Function
  End If
End If

'------------------------------------------------------------------------3文字以上の材質があったので修正_2024.8.20
TmpArray = Split(tMP, ",") 'Tmpを`,`で区切って配列にする
Dim i_fl As Long
For i_fl = LBound(TmpArray) To UBound(TmpArray)
    TmpArray(i_fl) = Trim(CStr(TmpArray(i_fl))) 'スペース削除（ミス入力対策）
Next i_fl

' 配列をループしてListを探す、False条件はExit Functionでぬける
  For Each Item In TmpArray

    If Left(Item, 1) = "≠" Then
      '≠あり（括弧なし）
      Dim CompareString As String: CompareString = Mid(Item, 2) '≠をとる

      If CompareString <> List Then
        FindInList = True   '不一致なのでTrue、後続条件でさらに判定継続
      ElseIf CompareString = List Then
        FindInList = False
        Exit Function
      End If

      '≠かつ*
      If Left(CompareString, 1) = "*" Then
        Dim MidItem_a As String
        Dim LeftCount_a As Long
        Dim LeftItem_a As String
        If Right(CompareString, 1) = "*" Then  '≠*文字列* 部分一致で除外
          MidItem_a = Mid(CompareString, 2)
          LeftCount_a = Len(MidItem_a) - 1
          LeftItem_a = Left(MidItem_a, LeftCount_a)
          If InStr(List, LeftItem_a) > 0 Then
            FindInList = False
            Exit Function
          End If
        Else                                    '≠*文字列 後方一致で除外
          MidItem_a = Mid(CompareString, 2)
          TempLength = Len(MidItem_a)
          Dim RightList_a As String: RightList_a = Right(List, TempLength)
          If RightList_a = MidItem_a Then
            FindInList = False
            Exit Function
          End If
        End If
      ElseIf Right(CompareString, 1) = "*" Then  '≠文字列* 前方一致で除外
        LeftCount_a = Len(CompareString) - 1
        LeftItem_a = Left(CompareString, LeftCount_a)
        TempLength = Len(LeftItem_a)
        Dim LeftList_a As String: LeftList_a = Left(List, TempLength)
        If LeftList_a = LeftItem_a Then
          FindInList = False
          Exit Function
        End If
      Else                                        '≠文字列 完全一致（再チェック）
        If CompareString = List Then
          FindInList = False
          Exit Function
        End If
      End If

    '----------------------------------------------------------------------
    ' 以下は肯定条件（≠なし）
    ' [後勝ちOR評価]
    ' 一致してもExit Functionしない。
    ' 後続に≠条件がある場合はそちらが優先してFalseに上書きされる。
    ' 例: tmp="SUS304,≠SUS304", List="SUS304"
    '   SUS304でTrueになるが≠SUS304でFalseに上書きされて終了
    '----------------------------------------------------------------------
    ElseIf Left(Item, 1) = "*" Then '頭に*がついている場合
      Dim MidItem As String
      Dim LeftCount As Long
      Dim LeftItem As String
      If Right(Item, 1) = "*" Then  '*文字列* 部分一致
        MidItem = Mid(Item, 2) '左*除去
        LeftCount = Len(MidItem) - 1
        LeftItem = Left(MidItem, LeftCount) '右*除去
        If InStr(List, LeftItem) > 0 Then FindInList = True  '後勝ちOR評価
      Else                          '*文字列 後方一致
        MidItem = Mid(Item, 2)
        TempLength = Len(MidItem)
        Dim RightList As String: RightList = Right(List, TempLength)
        If RightList = MidItem Then FindInList = True         '後勝ちOR評価
      End If
    ElseIf Right(Item, 1) = "*" Then  '文字列* 前方一致
      LeftCount = Len(Item) - 1
      LeftItem = Left(Item, LeftCount)
      TempLength = Len(LeftItem)
      Dim LeftList As String: LeftList = Left(List, TempLength)
      If LeftList = LeftItem Then FindInList = True           '後勝ちOR評価
    Else                              '文字列 完全一致
      If Item = List Then FindInList = True                   '後勝ちOR評価
    End If

  Next Item
End Function

' 条件をチェックする関数
Function CheckCondition(row As Long, Column As Long, Criteria As String, ArrQuality As Variant, Optional BlankFlag As String = "") As Boolean
Dim TmpString As String: TmpString = ArrQuality(row, Column)

Dim OutputString As String
If TmpString <> "" Then OutputString = StrConv(TmpString, vbNarrow) '納入先の文字列を半角に変換

OutputString = Replace(OutputString, "ｰ", "-") 'さらに置換

CheckCondition = FindInList(OutputString, Criteria, BlankFlag)
End Function

' 厚さの範囲をチェックする関数
Function CheckThickness(row As Long, Col1 As Long, Col2 As Long, Col3 As Long, OrderThickness As Variant, ArrQuality As Variant) As Boolean
' --- すべて空白なら、無条件でTrueを返して終了 ---
If ArrQuality(row, Col1) = "" And _
   ArrQuality(row, Col2) = "" And _
   ArrQuality(row, Col3) = "" Then
    CheckThickness = True
    Exit Function
End If

Dim LowerThickness As Single: LowerThickness = val(ArrQuality(row, Col1))
Dim UpperThickness As Single: UpperThickness = val(ArrQuality(row, Col2))

'(ArrQuality(Row, Col3) = "") Or 条件から削除_2026.05.13
If (LowerThickness <= val(OrderThickness) And (UpperThickness >= val(OrderThickness))) Then
  CheckThickness = True 'オーダー板厚が空白か、範囲内の場合は True を返す
End If
End Function

' 幅、丈の範囲をチェックする関数
Function CheckWidth(row As Long, Col1 As Long, Col2 As Long, Col3 As Long, OrderWidth As Variant, ArrQuality As Variant) As Boolean
' --- すべて空白なら、無条件でTrueを返して終了 ---
If ArrQuality(row, Col1) = "" And _
   ArrQuality(row, Col2) = "" And _
   ArrQuality(row, Col3) = "" Then
    CheckWidth = True
    Exit Function
End If

Dim LowerWidth As Single: LowerWidth = val(ArrQuality(row, Col1))
Dim UpperWidth As Single: UpperWidth = val(ArrQuality(row, Col2))

'(ArrQuality(Row, Col3) = "") Or 条件から削除_2026.05.13
If (LowerWidth <= val(OrderWidth) And (UpperWidth >= val(OrderWidth))) Then
  CheckWidth = True  'オーダー板幅が空白か、範囲内の場合は True を返す
End If
End Function

' 範囲をチェックする関数　範囲指定のみのやつ
Function CheckWithin(row As Long, Col1 As Long, Col2 As Long, Order As Variant, ArrQuality As Variant) As Boolean
Dim LowerWidth As Single: LowerWidth = val(ArrQuality(row, Col1))
Dim UpperWidth As Single: UpperWidth = val(ArrQuality(row, Col2))

If (LowerWidth <= val(Order) And UpperWidth >= val(Order)) Then CheckWithin = True  '範囲内の場合は True を返す
End Function

'指定数より大きいか
Function CheckSize(row As Long, Col1 As Long, OrderSize As Variant, ArrQuality As Variant) As Boolean
If (ArrQuality(row, Col1) = "") Then
  CheckSize = True
  Exit Function
End If
Dim LowSize As Long: LowSize = val(ArrQuality(row, Col1))
If (val(OrderSize) >= LowSize) Then CheckSize = True
End Function

'引数 配列 Arry 内の文字列に対して、左から1文字目が引数 WidthLength の場合に、引数 WidthLength を除いて数字だった場合は配列 Temp に格納し、
'Temp 内の数字を小さい順に並び変え、引数 ReturnSize を Temp 内の数字の大きい順で比較し範囲に収まるArryのインデックスを返す
Function FindIndices(Arry As Variant, WidthLength As String, ReturnSize As Single) As Variant
Dim Temp() As Double, TempArry(1 To 2) As Variant
Dim i As Long, a As Long, j As Long, n As Long, tempIndex As Long, Amatch As Long
Dim Result As String, matchFlag As Boolean

' Temp配列に数字だけを抽出して格納
For i = LBound(Arry, 2) To UBound(Arry, 2)
  If Left(Arry(LBound(Arry, 2), i), 1) = WidthLength And IsNumeric(Mid(Arry(LBound(Arry, 2), i), 2, Len(Arry(LBound(Arry, 2), i)) - 1)) Then
    ReDim Preserve Temp(j)
    Temp(j) = CDbl(Mid(Arry(LBound(Arry, 2), i), 2, Len(Arry(LBound(Arry, 2), i)) - 1))
    j = j + 1
  End If
Next i

' Temp配列を昇順に並び替え 0配列
For i = LBound(Temp, 1) To UBound(Temp, 1) - 1
  For j = i + 1 To UBound(Temp)
    If Temp(i) > Temp(j) Then
      tempIndex = Temp(i)
      Temp(i) = Temp(j)
      Temp(j) = tempIndex
    End If
  Next j
Next i

'Stop
'---------------------------------------------------------------------------どことどこの間に収まるかを示す
  For a = UBound(Temp, 1) To LBound(Temp, 1) Step -1 '一番右のフィールドとその隣のフィールドの範囲に収まっているかのチェック(大きい順)
    n = a - 1
      If n = 0 Then '0<=製造幅<=500
        If Temp(n) >= ReturnSize And val(0) < ReturnSize Then '500mm以下の時
          Amatch = n
          matchFlag = True
          Exit For
        ElseIf Temp(a) >= ReturnSize And Temp(n) < ReturnSize Then
          Amatch = a
          matchFlag = True
          Exit For
        End If
      Else
      
      If n = -1 Then 'DB範囲外のサイズが来ている状態
        Amatch = UBound(Temp)  '一番大きいサイズのインデックスにする
        TempArry(1) = Amatch
        TempArry(2) = 9999 'サイズオーバーフラグ
        GoTo skipMax
      End If
      
        If Temp(a) >= ReturnSize And Temp(n) < ReturnSize Then
          Amatch = a
          matchFlag = True
          Exit For
        End If
      End If
  Next a
  
skipMax:
Dim MatchSize As String: MatchSize = WidthLength & Temp(Amatch) '元に戻して元フィールドから同サイズをさがす
For i = LBound(Arry, 2) To UBound(Arry, 2)
  If Arry(LBound(Arry, 2), i) = MatchSize Then Amatch = i
Next i

TempArry(1) = Amatch
FindIndices = TempArry
End Function

Function ExtractNumber(str As String) As Variant '以上以下の文字列チェック後,数字抽出
Dim Position As Long, i As Long
Dim Char As String, Result As String
Dim IsNumericFound As Boolean, HasDecimalPoint As Boolean

'以上以下が含まれない文字列の数字は不要なのでチェックが必要
Position = InStr(str, "以上") '以上が含まれるかチェック
If Position > 0 Then
Else ' 以下が含まれるかチェック
  Position = InStr(str, "以下")
  If Position > 0 Then
  Else ' どちらも含まれていない場合は空文字列を返す
    ExtractNumber = ""
    Exit Function
  End If
End If

' 数字を抽出
Result = ""
IsNumericFound = False
HasDecimalPoint = False
For i = 1 To Len(str)
  Char = Mid(str, i, 1)
  If IsNumeric(Char) Then
    Result = Result & Char
    IsNumericFound = True
  ElseIf Char = "." And Not HasDecimalPoint Then '小数点がまだ出てきていない場合は追加
    Result = Result & Char
    HasDecimalPoint = True
    IsNumericFound = True
  ElseIf IsNumericFound Then ' 数字が見つかった後、非数字が出てきたらループを終了
    Exit For
  End If
Next i

If Result = "" Then ' 数字が抽出されなかった場合は空文字列を返す
  ExtractNumber = ""
Else
  ExtractNumber = CDbl(Result) '数字を返す
End If
End Function

Function CompareStrings(tMP As String, Lot厚 As Single, DB厚 As Single) As Boolean '引数比較
If InStr(tMP, "以上") > 0 Then ' "以上" の場合の比較
  If (Lot厚 >= DB厚) Then CompareStrings = True
ElseIf InStr(tMP, "以下") > 0 Then ' "以下" の場合の比較
  If (Lot厚 <= DB厚) Then CompareStrings = True
Else
  CompareStrings = False
End If
End Function

Function IsHalfWidth(ByVal S As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(S)
        If AscW(Mid(S, i, 1)) > &HFF Then
            IsHalfWidth = False
            Exit Function
        End If
    Next i
    IsHalfWidth = True
End Function

'与えられた文字列 Tmp が,で区切られた複数の部分文字列から構成されている場合、それぞれの文字列に対してConditionを含むかどうかをチェック
Function ExtractNums(tMP As String, Condition As String) As String
Dim parts() As String
Dim Result As String
Dim i As Long

parts = Split(tMP, ",") ' 文字列を","で分割

For i = LBound(parts) To UBound(parts)
  Dim num As String
  num = ExtractNumFromPart(parts(i), Condition) ' 部分文字列から数値を抽出
  If num <> "" Then
      Result = Result & num & ", "
  End If
Next i

If Len(Result) > 0 Then
  ' 最後のカンマとスペースを削除して返す
  ExtractNums = Left(Result, Len(Result) - 2)
Else
  ExtractNums = ""
End If
End Function

' 文字列から数値を抽出
Function ExtractNumFromPart(part As String, Condition As String) As String
Dim Position As Long, i As Long
Dim Char As String, Result As String
Dim IsNumericFound As Boolean, HasDecimalPoint As Boolean

Position = InStr(part, Condition) ' Conditionが含まれるかチェック
If Position > 0 Then
Else
  ExtractNumFromPart = "" ' Conditionが含まれない場合、空文字列を返す
  Exit Function
End If

' 数字を抽出
Result = ""
IsNumericFound = False
HasDecimalPoint = False
For i = 1 To Len(part)
  Char = Mid(part, i, 1)
  If IsNumeric(Char) Then
    Result = Result & Char
    IsNumericFound = True
  ElseIf Char = "." And Not HasDecimalPoint Then ' 小数点がまだ出てきていない場合は追加
    Result = Result & Char
    HasDecimalPoint = True
    IsNumericFound = True
  ElseIf IsNumericFound Then ' 数字が見つかった後、非数字が出てきたらループを終了
    Exit For
  End If
Next i

If Result = "" Then ' 数字が抽出されなかった場合は空文字列を返す
  ExtractNumFromPart = ""
Else
  ExtractNumFromPart = Result ' 数字を返す
End If
End Function

Public Function DbReturnItem(Nou As String, Tori As String, YoutC As String, Zai As String, Tyou As String, ATU As Single, HABA As Single, TAKE As Single) As String 'DBから出力テキストを返す
'納入、取引、用途コード、材質、調質、厚、幅、丈
'2024.8.5_引数に取引追加
Dim FPath As String:         FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
Dim TableName As String:     TableName = "ラテフリ直角"
Dim target As String:        target = "品質基準.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim PackF():                 PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)


UF1.HURIS = "" 'リセット

''オーダー厚のセット_2024.5.14
'ラテフリ直角は製造厚_2024.8.27
'If SikaSetAtu <> 0 Then
'  If SikaSetAtu <> ATU Then ATU = SikaSetAtu
'End If

If IsEmpty(PackF) Then
  FPath = PackingPath
  FilePath = FPath & target
  PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(PackF) Then Exit Function
End If

Dim C_納入先 As Long, C_用途コード As Long, C_製造材質 As Long, C_製造調質 As Long, C_取引先 As Long, _
C_厚下 As Long, C_厚上 As Long, C_幅下 As Long, C_幅上 As Long, i As Long, _
 C_丈下 As Long, C_丈上 As Long, C_フリー反り As Long, C_ラテラルボー As Long, C_直角度 As Long, _
 C_製造板厚 As Long, C_製造板幅 As Long, C_製造板丈 As Long, C_番号 As Long, C_公差 As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "管理番号":     C_番号 = i
  Case "納入先":       C_納入先 = i
  
  Case "取引先":       C_取引先 = i '2024.8.5
  
  Case "用途コード":   C_用途コード = i
  Case "製造材質":     C_製造材質 = i
  Case "製造調質":     C_製造調質 = i
  Case "厚下":         C_厚下 = i
  Case "厚上":         C_厚上 = i
  Case "幅下":         C_幅下 = i
  Case "幅上":         C_幅上 = i
  Case "丈下":         C_丈下 = i
  Case "丈上":         C_丈上 = i
  Case "フリー反り":   C_フリー反り = i
  Case "ラテラルボー": C_ラテラルボー = i
  Case "直角度":       C_直角度 = i
  
  Case "製造板厚":     C_製造板厚 = i
  Case "製造板幅":     C_製造板幅 = i
  Case "製造板丈":     C_製造板丈 = i
  Case "公差":         C_公差 = i
  End Select
Next i

Dim SQLfilter As String: SQLfilter = "" '"[納入先] LIKE " & "'*" & Nou & "*'" 'DB側には長い文字列を入れておくこと
'ﾊｸﾄﾞｳ(ｶ)の登録からﾊｸﾄﾞｳ(ｶ)ﾌｸｼﾏｼﾞｷﾞﾖｳｼﾖは探せない
'2024.8.20_フィルター停止
Dim TempArr As Variant:  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)

If IsEmpty(TempArr) Then
  DbReturnItem = ""
  Exit Function
End If


'引当コード取引コードはまぁ・・・・一緒の意味なんでスルーします

Dim MatchingArray() As Long ' マッチングしたインデックスを格納する配列
ReDim MatchingArray(1 To UBound(TempArr, 1)) ' 配列のサイズを設定

Dim Conditions(1 To 8) As Boolean ' 各条件の結果
Dim j As Long, a As Long, n As Long: n = 0

For i = LBound(TempArr, 1) To UBound(TempArr, 1)
'If TempArr(i, C_番号) = 5 Then 'Test
'Stop
'End If
  Conditions(1) = CheckCondition(i, C_納入先, Nou, TempArr) '納入先のチェック_2024.8.5

  Conditions(2) = Conditions(1) And CheckCondition(i, C_取引先, Tori, TempArr)  '取引先のチェック_2024.8.5

  Conditions(3) = Conditions(2) And CheckCondition(i, C_用途コード, YoutC, TempArr)  ' 用途コードのチェック
  Conditions(4) = Conditions(3) And CheckCondition(i, C_製造材質, Zai, TempArr) ' 製造材質のチェック
  Conditions(5) = Conditions(4) And CheckCondition(i, C_製造調質, Tyou, TempArr) ' 製造調質のチェック
  Conditions(6) = Conditions(5) And CheckThickness(i, C_厚下, C_厚上, C_製造板厚, ATU, TempArr) ' 厚さのチェック　カラム名オーダー板厚にしてしまっているが実際の参照対象は製造なので操業基準と一致
  Conditions(7) = Conditions(6) And CheckWidth(i, C_幅下, C_幅上, C_製造板幅, HABA, TempArr) ' 幅のチェック
  Conditions(8) = Conditions(7) And CheckWidth(i, C_丈下, C_丈上, C_製造板丈, TAKE, TempArr) ' 丈のチェック
  
 ' 条件が全て満たされた場合の処理
  If Conditions(1) And Conditions(2) And Conditions(3) _
  And Conditions(4) And Conditions(5) And Conditions(6) And Conditions(7) And Conditions(8) Then
    n = n + 1
    MatchingArray(n) = TempArr(i, C_番号)
  End If
Next i

If n = 0 Then '2024.5.16_納入一致だが不要の場合
  If Nou Like "*ﾆﾎﾝｼﾔﾘﾖｳｾｲｿﾞｳ*" Then '2024.5.15_引当の表記のままで行う
  UF1.Sumisei.Caption = Nou
    If Not (UF1.Additional.Caption Like "*ﾗﾃﾗﾙﾎﾞ-追記公差あり*") Then '追記公差なし
      If Conditions(4) = False Then '2025.6.4_追加
        UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_製造材質"
        UF1.LateralText = ""
      ElseIf Conditions(4) And Conditions(5) = False Then
        UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ幅"
        UF1.LateralText = ""
      ElseIf Conditions(4) And Conditions(6) = False Then
        UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ丈"
        UF1.LateralText = ""
      ElseIf Conditions(4) = False Then
        UF1.Additional.Caption = "ﾗﾃﾗﾙﾎﾞｰ検査不要_ｻｲｽﾞ厚"
        UF1.LateralText = ""
      End If
    End If
  End If
  '2024.8.5
  '2024.8.27_等級公差が初めに書き込まれている、それを優先度の高い公差で上書きする仕組みだが、この場合上書きする数字がないため等級公差が表示されたままになってしまう
  If Tori Like "ﾅﾒｶﾜｱﾙﾐ(ｶ*" Then
    UF1.Sumisei.Caption = Nou
    If Nou Like "*K.T.N. CO. LTD*" Then
      UF1.Additional.Caption = "K.T.N ﾗﾃﾗﾙﾎﾞｰ検査不要"
      UF1.LateralText = ""
    End If
    If Nou Like "ﾆﾎﾝﾊﾂｼﾞﾖｳ*" Then
      UF1.Additional.Caption = "ﾆﾎﾝﾊﾂｼﾞﾖｳ ﾗﾃﾗﾙﾎﾞｰ検査不要"
      UF1.LateralText = ""
    End If
  End If
  Exit Function
End If


' MatchingArrayから0を削除する
Dim NewArray() As Long
ReDim NewArray(1 To n)
For i = 1 To n
  NewArray(i) = MatchingArray(i)
Next i
'格納
Dim HitArry()
ReDim HitArry(1 To UBound(NewArray, 1), 1 To UBound(TempArr, 2))
For a = LBound(NewArray, 1) To UBound(NewArray, 1)
  For i = LBound(TempArr, 1) To UBound(TempArr, 1)
    For j = LBound(TempArr, 2) To UBound(TempArr, 2)
      If TempArr(i, C_番号) = NewArray(a) Then
        HitArry(a, j) = TempArr(i, j)
      End If
    Next j
  Next i
Next a

Dim flag(1 To 4) As Boolean

For i = LBound(HitArry, 1) To UBound(HitArry, 1)
  If HitArry(i, C_フリー反り) = "要" Then
     DbReturnItem = DbReturnItem & "," & "フリー反り"
     
    '2026.03.05_フリー反り公差
     UF1.HURIS = HitArry(i, C_公差)
  End If
  If HitArry(i, C_ラテラルボー) = "要" Then DbReturnItem = DbReturnItem & "," & "ラテラルボー"
  If HitArry(i, C_直角度) = "要" Then DbReturnItem = DbReturnItem & "," & "直角度"
  
Next i

'Stop

'公差表示関係はここで実行、描写はLSelect
If UFMitsubishi.Visible = False Then
  With UF1
  '---------------Q-S-QA-DIM-179_板材曲がり(ラテラルボー)検査基準
    If InStr(DbReturnItem, "ラテラルボー") > 0 Then
      .Sumisei.Caption = Nou
      
      If Not (.Additional.Caption Like "*ﾗﾃﾗﾙﾎﾞ-追記公差あり*") Then '追記公差が優先 含まれていたら抜ける
        '引当表記そのままで行う_2024.5.15
        If IsNumeric(HitArry(1, C_公差)) Then '公差特別指示あり
          .LateralText = Format(HitArry(1, C_公差), "0.0")
        Else '公差に文字列がある
          If HitArry(1, C_公差) = "別表B" Then '今のところ日車のみ_2024.5.15******************'2024.8.5_別表公差が追加された
            .LateralText = Format(DBdisplayラテラル(.Lot, "日車", .DBflag.Text), "0.0") '別テーブル　日車
            .LateralText.BackStyle = 1
            .LateralText.BackColor = RGB(255, 153, 0)
          ElseIf HitArry(1, C_公差) = "別表A" Then '2024.8.5_別表公差が追加された
            .LateralText = Format(DBdisplayラテラル(.Lot, "白銅滑川", .DBflag.Text), "0.0") '別テーブル　白銅滑川
            .LateralText.BackStyle = 1
            .LateralText.BackColor = RGB(255, 153, 0)
          ElseIf HitArry(1, C_公差) = "等級" Then
          End If
        End If
      End If
'      Select Case True
'      Case flag(1) '公差等級
'        .Sumisei.Caption = "住友精密"
'      Case flag(2) '公差特別指示
'        .Sumisei.Caption = "白銅"
'        .LateralText = Format(HitArry(1, C_公差), "0.0")
'        .LateralText.BackStyle = 1
'        .LateralText.BackColor = RGB(255, 153, 0)
'      Case flag(3)  '公差別表
'        .Sumisei.Caption = "日本車両"
'        .LateralText = Format(DBdisplayラテラル(.Lot, .DBflag.text), "0.0")
'        .LateralText.BackStyle = 1
'        .LateralText.BackColor = RGB(255, 153, 0)
'      Case flag(4) '公差特別指示
'        .Sumisei.Caption = "滑川"
'        .LateralText = Format(HitArry(1, C_公差), "0.0")
'        .LateralText.BackStyle = 1
'        .LateralText.BackColor = RGB(255, 153, 0)
'      End Select
        
    End If
    
    '---------------Q-S-QA-DIM-022_フリー反り測定基準’Q-S-CH-DIM-022-R03_2026.06.05
    If InStr(DbReturnItem, "フリー反り") > 0 Then
    Dim kousaF As Long: kousaF = LBound(HitArry, 1)
      .HizCome.Tag = HitArry(kousaF, C_公差) & "mm以下"
      .flatnesText.BackStyle = 1
      .flatnesText.BackColor = RGB(255, 153, 0)
      .flatnesText = "ﾌﾘｰｿﾘ要"
      .LGS.Caption = "ﾌﾘｰｿﾘ"
      .LGU.Caption = "ﾌﾘｰｿﾘ"
      .LGC.Caption = "ﾌﾘｰｿﾘ"
      .LGD.Caption = "ﾌﾘｰｿﾘ"
      .LGE.Caption = "ﾌﾘｰｿﾘ"
      .LGF.Caption = "ﾌﾘｰｿﾘ"
      Call CBAdd(ATU)
      
      UF1.HURIS = HitArry(kousaF, C_公差)
      UF1.HizCome.Caption = "ﾌﾘｰ反り" & UF1.HizCome.Tag 'フリー反りがあった場合だけHizCome差し替え
    End If
  End With
End If

If UFMitsubishi.Visible Then
  If WsCheck("全ヵ所検査") = False Then Exit Function
  Dim ZenSh As Worksheet
  If UFMitsubishi.WorkFlag.Caption = "P1" Then Set ZenSh = ThisWorkbook.Sheets("全ヵ所検査")
  If UFMitsubishi.WorkFlag.Caption = "P_H494" Then Set ZenSh = ThisWorkbook.Sheets("H494検査")
  
  ZenSh.Unprotect 'シート保護解除

  With UFMitsubishi
    If InStr(DbReturnItem, "ラテラルボー") > 0 Then
      ZenSh.Cells(1, 15) = Nou & Zai
      '引当表記そのままで行う_2024.5.15
      If IsNumeric(HitArry(1, C_公差)) Then '公差特別指示あり
        ZenSh.Cells(1, 13) = Format(HitArry(1, C_公差), "0.0")
      Else '公差に文字列がある
        If HitArry(1, C_公差) = "別表B" Then '今のところ日車のみ_2024.5.15'2024.8.5_別表公差が追加された
          ZenSh.Cells(1, 13) = Format(DBdisplayラテラル(.LotBox, "日車"), "0.0")  '別テーブル 日車
        ElseIf HitArry(1, C_公差) = "別表A" Then '今のところ日車のみ_2024.5.15
          ZenSh.Cells(1, 13) = Format(DBdisplayラテラル(.LotBox, "白銅滑川"), "0.0")   '別テーブル　白銅滑川
        ElseIf HitArry(1, C_公差) = "等級" Then
        End If
      End If
    
'      Select Case True
'      Case flag(1) '公差等級
'        ZenSh.Cells(1, 15) = "住友精密" & ZAI
'      Case flag(2) '公差特別指示
'        ZenSh.Cells(1, 13) = Format(HitArry(1, C_公差), "0.0")
'        ZenSh.Cells(1, 15) = "白銅" & ZAI
'      Case flag(3)  '公差別表
'        ZenSh.Cells(1, 13) = Format(DBdisplayラテラル(.LotBox), "0.0")
'        ZenSh.Cells(1, 15) = "日本車両" & ZAI
'      Case flag(4) '公差特別指示
'        ZenSh.Cells(1, 13) = Format(HitArry(1, C_公差), "0.0")
'        ZenSh.Cells(1, 15) = "滑川" & ZAI
'      End Select
    End If
    '---------------Q-S-QA-DIM-022_フリー反り測定基準
    If InStr(DbReturnItem, "フリー反り") > 0 Then
      ZenSh.Cells(6, 29) = "ﾌﾘｰｿﾘ"
    End If
  End With
End If
End Function

Public Function SoriArray(TempA As Variant, tempB As Variant, ATU As Single, HabSize As Single, TakSize As Single, FindIndices As Variant) As Single
If FindIndices = 0 Then Exit Function
'基準レコード,基準テーブル, 厚,幅,丈,カラムインデックス*****インデックスを使い分けて幅丈公差を返す
'範囲内に収まらないければ0が返る
Dim i As Long, a As Long, C_板厚下 As Long, C_板厚上 As Long, C_板丈下 As Long, C_板丈上 As Long, C_板丈 As Long
Dim matchFlag As Boolean: matchFlag = False
Dim SitaSingele As Single, UeSingele As Single, TsitaSingele As Single, TueSingele As Single

For i = LBound(tempB, 2) To UBound(tempB, 2)
  Select Case tempB(1, i)
  Case "板厚下":   C_板厚下 = i
  Case "板厚上":   C_板厚上 = i
'  Case "板丈下":   C_板丈下 = i
'  Case "板丈上":   C_板丈上 = i
'  Case "板丈":     C_板丈 = i
  End Select
Next i

For i = LBound(TempA, 1) To UBound(TempA, 1)
  SitaSingele = val(TempA(i, C_板厚下))
  UeSingele = val(TempA(i, C_板厚上))
'  If TempA(i, C_板丈) <> "" Then TsitaSingele = Val(TempA(i, C_板丈下))
'  If TempA(i, C_板丈) <> "" Then TueSingele = Val(TempA(i, C_板丈上))
  If (SitaSingele <= ATU) And (UeSingele >= ATU) Then '厚が収まる範囲を探す
'    If TempA(i, C_板丈) = "" Then
      SoriArray = Format(TempA(i, FindIndices), "0.0")
      matchFlag = True
      Exit For
'    ElseIf TempA(i, C_板丈) <> "" Then
'      If ((TsitaSingele <= TakSize) And (TueSingele >= TakSize)) Then '丈が収まる範囲を探す
'        SoriArray = Format(TempA(i, FindIndices), "0.0")
'        matchFlag = True
'        Exit For
'      End If
'    End If
'  If matchFlag Then Exit For
  End If
If matchFlag Then Exit For
Next i

End Function
'Public Function MaterialsUpperPart(OrdNo As String) As String '緩衝材チェック**使用ﾀｲﾐﾝｸﾞはUF1起動時のコンボボックス変更後
''オーダーNo検索
'If WsCheck("Lot情報履歴") = False Then Exit Function '-------------------------ｼｰﾄなし処理停止
'If OrdNo = "" Then Exit Function
'Dim HisSH As Worksheet:   Set HisSH = ThisWorkbook.Sheets("Lot情報履歴")
'Dim HisRan As Range:      Set HisRan = Range(HisSH.Cells(1, 1), HisSH.Cells(UR(HisSH, 1), UC(HisSH, 1)))
'Dim Hist(): Hist = HisRan
'Dim i As Long, a As Long
'Dim R_Ord As Long, R_Packing As Long, R_Lot As Long, R_Hik As Long, R_厚 As Long, R_巾 As Long, R_丈 As Long, R_材質 As Long, R_調質 As Long, R_用途コード As Long, R_合紙 As Long
'For i = 1 To UBound(Hist, 2) 'ﾌｨｰﾙﾄﾞ名
'  Select Case Hist(1, i)
'  Case "ﾛｯﾄNo":        R_Lot = i
'  Case "ｵｰﾀﾞｰNo":      R_Ord = i
'  Case "引当NO":       R_Hik = i
'  Case "包装仕様NO":   R_Packing = i
'  Case "厚":           R_厚 = i
'  Case "巾":           R_巾 = i
'  Case "丈":           R_丈 = i
'  Case "材質":         R_材質 = i
'  Case "調質":         R_調質 = i
'  Case "用途ｺｰﾄﾞ":     R_用途コード = i
'  Case "合紙":         R_合紙 = i
'  End Select
'Next i
'
'For i = LBound(Hist, 1) To UBound(Hist, 1) '1次元ﾁｪｯｸ
'  If Hist(i, R_Ord) = OrdNo Then a = i
'Next i
'
'If a = 0 Then Exit Function
'
''包装仕様書Noが0001～0112までのとき***梱包指定が一般標準***2022.02.04時点での仕様書内容
'Dim CheckNo As Long:    CheckNo = Val(Right(Hist(a, R_Packing), 4))
'Dim CheckBuf As String: CheckBuf = Left(Hist(a, R_Packing), 2)
'
'If 121 > CheckNo And CheckBuf = "1P" Then '指定の数字以下の時
'    'チェックに必要なデータ収集
'    Dim You As String: You = Hist(a, R_用途コード)
'    Dim Zai As String: Zai = Hist(a, R_材質)
'    Dim Tyo As String: Tyo = Hist(a, R_調質)
'    Dim Atu As Single: Atu = Hist(a, R_厚)
'    Dim Hab As Single: Hab = Hist(a, R_巾)
'    Dim Tak As Single: Tak = Hist(a, R_丈)
'
'    Dim UpperPart As String
'
'    Select Case True
'    Case You Like "T23?" 'パネル
'      UpperPart = "ザラ板"
'    Case You Like "K21?" '鉄道ﾎﾞﾃﾞｨｰ
'      If Atu <= 3# Then UpperPart = "ザラ板"
'      If Atu > 3# Then UpperPart = "アングル"
'    End Select
'
'    Select Case True
'    Case Atu <= 1#: UpperPart = "ザラ板"
'    End Select
'
'    Select Case True
'    Case Zai = "NK01" Or Zai Like "*995*" Or Zai Like "*3S*" Or Zai Like "*2ST*"  'ｱﾙﾐ系'材質
'      Select Case True '調質
'      Case Tyo = "O"
'        If Atu <= 5# Then UpperPart = "ザラ板"
'        If Atu > 5# Then UpperPart = "アングル"
'      Case Tyo Like "H2?"
'        If Atu <= 3# Then UpperPart = "ザラ板"
'        If Atu > 3# Then UpperPart = "アングル"
'      End Select
'    Case Zai = "5N"
'        UpperPart = "ザラ板" '2023.01.11
'    Case Else '合金系'材質
'      Select Case True '調質
'      Case Tyo = "O"
'        If Atu <= 3# Then UpperPart = "ザラ板"
'        If Atu > 3# Then UpperPart = "アングル"
'      Case Tyo Like "H2?" Or Tyo Like "H3?"
'        If Atu <= 1# Then UpperPart = "ザラ板"
'        '条件分岐追記
'        If Atu > 1# Then
'        '2023.2.24_包装仕様書の用途板厚についての優先度的に＊＊合金かつテツドウボディーの時に限り条件分岐あり＊＊
'          Select Case True
'          Case You Like "K21?" '鉄道ﾎﾞﾃﾞｨｰ
'            UpperPart = "ザラ板"
'          Case Else
'            UpperPart = "アングル"
'          End Select
'        End If
'      End Select
'    End Select
'
'    Select Case True
'    Case (Hab >= 1219# And Hab <= 1222.5) And (Tak >= 2438# And Tak <= 2442.5)
'      If Atu < 50# Then UpperPart = "ザラ板"
'      If Atu >= 50# Then UpperPart = "アングル"
'    End Select
'    MaterialsUpperPart = UpperPart
'
'    If Hist(a, R_Packing) = "1P0120" Then
'      MaterialsUpperPart = UpperPart & vbCrLf & "ﾍﾟｯﾄﾊﾞﾝﾄﾞ"
'    Else
'      If Atu <= 6# And UpperPart = "アングル" Then MaterialsUpperPart = UpperPart & vbCrLf & "当板必要"
'    End If
'
'    If Hist(a, R_Packing) = "1P0115" Or Hist(a, R_Packing) = "1P0116" Then
'      If Atu <= 6# And UpperPart = "アングル" Then
'        MaterialsUpperPart = UpperPart & vbCrLf & "ﾌﾟﾛﾃｯｸﾎﾞｰﾄﾞ" & vbCrLf & "当板必要"
'      ElseIf Atu > 6# And UpperPart = "アングル" Then
'        MaterialsUpperPart = UpperPart & vbCrLf & "ﾌﾟﾛﾃｯｸﾎﾞｰﾄﾞ"
'      End If
'    End If
'    If Hist(a, R_Packing) = "1P0118" Or Hist(a, R_Packing) = "1P0119" Or Hist(a, R_Packing) = "1P1189" Then MaterialsUpperPart = UpperPart & vbCrLf & "縦ﾊﾞﾝﾄﾞ"
'
'End If
'End Function

'----------------------------------------------------------------
' 包装仕様書Noが対象リストに含まれるか判定する関数
' C_TARGET_HOSONOS はカンマ区切りで列挙した対象Noの定数
'----------------------------------------------------------------
Public Function IsTargetHosoNo(ByVal HosoNo As String) As Boolean
    Dim arrTargets() As String
    arrTargets = Split(C_TARGET_HOSONOS, ",")

    Dim i As Long
    For i = LBound(arrTargets) To UBound(arrTargets)
        If Trim(arrTargets(i)) = HosoNo Then
            IsTargetHosoNo = True
            Exit Function
        End If
    Next i

    IsTargetHosoNo = False
End Function

Private Function EnsureCache(OrdNo As String) As Boolean
    If m_CacheOrdNo = OrdNo And Not IsEmpty(m_PackQuality) Then
        EnsureCache = True: Exit Function   ' キャッシュ有効
    End If

    Dim FilePath  As String: FilePath = PackingPath & "梱包資材マスタ.accdb"
    Dim TableName As String: TableName = "梱包保護材"

    Dim PackF(): PackF = GetFieldsArr(FilePath, TableName, "")
    If IsEmpty(PackF) Then Exit Function

    Dim i As Long
    For i = LBound(PackF, 2) To UBound(PackF, 2)
        Select Case PackF(1, i)
            Case "管理番号":   m_C_番号 = i
            Case "包装仕様書": m_C_包装仕様書 = i
            Case "用途コード": m_C_用途コード = i
            Case "材質":       m_C_材質 = i
            Case "調質":       m_C_調質 = i
            Case "板厚下":     m_C_板厚下 = i
            Case "板厚上":     m_C_板厚上 = i
            Case "板厚":       m_C_板厚 = i
            Case "板幅下":     m_C_板幅下 = i
            Case "板幅上":     m_C_板幅上 = i
            Case "板幅":       m_C_板幅 = i
            Case "板丈下":     m_C_板丈下 = i
            Case "板丈上":     m_C_板丈上 = i
            Case "板丈":       m_C_板丈 = i
            Case "使用保護材": m_C_使用保護材 = i
            Case "備考":       m_C_備考 = i
        End Select
    Next i

    m_PackQuality = GetRecordsArr(FilePath, TableName, "")
'    Stop
    If IsEmpty(m_PackQuality) Then Exit Function

    m_CacheOrdNo = OrdNo
    EnsureCache = True
End Function

Public Function MaterialsUpperPart(OrdNo As String, Optional CheckBikou As Boolean = False) As String '緩衝材チェック**使用ﾀｲﾐﾝｸﾞはUF1起動時のコンボボックス変更後
    'オーダーNo検索
    If WsCheck("Lot情報履歴") = False Then Exit Function '-------------------------ｼｰﾄなし処理停止
    If OrdNo = "" Then Exit Function

    Dim HistorySH As Worksheet: Set HistorySH = ThisWorkbook.Sheets("Lot情報履歴")
    Dim LR As Long:             LR = UR(HistorySH, 1)
    Dim LC As Long:             LC = UC(HistorySH, 1)
    Dim HistoryR As Range:      Set HistoryR = HistorySH.Range(HistorySH.Cells(1, 1), HistorySH.Cells(LR, LC))
    Dim TempArry():             TempArry = HistoryR

    If IsEmpty(TempArry) Then Exit Function

    Dim i As Long, a As Long
    Dim R_Ord As Long, R_Packing As Long, R_Lot As Long, R_Hik As Long, _
        R_厚 As Long, R_巾 As Long, R_丈 As Long, R_材質 As Long, R_調質 As Long, R_用途コード As Long
        
    For i = 1 To UBound(TempArry, 2) 'ﾌｨｰﾙﾄﾞ名
        Select Case TempArry(1, i)
            Case "ﾛｯﾄNo":        R_Lot = i
            Case "ｵｰﾀﾞｰNo":      R_Ord = i
            Case "引当NO":       R_Hik = i
            Case "包装仕様NO":   R_Packing = i
            Case "厚":           R_厚 = i
            Case "巾":           R_巾 = i
            Case "丈":           R_丈 = i
            Case "材質":         R_材質 = i
            Case "調質":         R_調質 = i
            Case "用途ｺｰﾄﾞ":     R_用途コード = i
        End Select
    Next i

    Dim originalStr As String: originalStr = OrdNo
    Dim cleanStr As String:    cleanStr = Replace(originalStr, "'", "")

    For i = LBound(TempArry, 1) To UBound(TempArry, 1) '1次元ﾁｪｯｸ
        If TempArry(i, R_Ord) = cleanStr Then
          a = i
        Exit For
        End If
    Next i

    If a = 0 Then Exit Function

    Dim AvatarPacNo As String: AvatarPacNo = TempArry(a, R_Packing)

    '包装仕様書Noが0001～0112までのとき***梱包指定が一般標準***2022.02.04時点での仕様書内容
'    Dim CheckNo As Long:    CheckNo = val(Right(AvatarPacNo, 4))
'    Dim CheckBuf As String: CheckBuf = Left(AvatarPacNo, 2)
'    If 121 > CheckNo And CheckBuf = "1P" Then '指定の数字以下の時

    '2026.06.09指定の包装仕様書Noだけを検索に変更
    If IsTargetHosoNo(AvatarPacNo) Then '対象包装仕様書Noチェック
        'チェックに必要なデータ収集
        Dim You As String:  You = TempArry(a, R_用途コード)
        Dim Zai As String:  Zai = TempArry(a, R_材質)
        Dim TYO As String:  TYO = TempArry(a, R_調質)
        Dim ATU As Single:  ATU = TempArry(a, R_厚)
        Dim Hab As Single:  Hab = TempArry(a, R_巾)
        Dim Tak As Single:  Tak = TempArry(a, R_丈)
        Dim Pack As String: Pack = TempArry(a, R_Packing)
    Else
        Exit Function
    End If
    
    '2026.06.09
    'EnsureCache 内で m_CacheOrdNo = cleanStr と同一かチェックし、一致していれば即リターン
    If Not EnsureCache(cleanStr) Then Exit Function
        
    Dim Conditions(1 To 7) As Boolean ' 各条件の結果を格納する変数
    Dim IsMatched As Boolean:  IsMatched = False ' ヒットしたかどうかの判定フラグ

    ' --- マスタ判定ループ ---
    For i = LBound(m_PackQuality, 1) To UBound(m_PackQuality, 1)


'        If m_PackQuality(i, m_C_番号) = "16" Then 'Test
'        Stop
'        End If


        ' 備考フィルタ
        Dim IsBikou As Boolean
        IsBikou = (m_PackQuality(i, m_C_備考) = "当板用")
        If IsBikou <> CheckBikou Then GoTo NextRow

        Conditions(1) = CheckCondition(i, m_C_包装仕様書, Pack, m_PackQuality)
        Conditions(2) = Conditions(1) And CheckCondition(i, m_C_材質, Zai, m_PackQuality)
        Conditions(3) = Conditions(2) And CheckCondition(i, m_C_調質, TYO, m_PackQuality)
        Conditions(4) = Conditions(3) And CheckCondition(i, m_C_用途コード, You, m_PackQuality)
        Conditions(5) = Conditions(4) And CheckThickness(i, m_C_板厚下, m_C_板厚上, m_C_板厚, ATU, m_PackQuality)
        Conditions(6) = Conditions(5) And CheckWidth(i, m_C_板幅下, m_C_板幅上, m_C_板幅, Hab, m_PackQuality)
        Conditions(7) = Conditions(6) And CheckWidth(i, m_C_板丈下, m_C_板丈上, m_C_板丈, Tak, m_PackQuality)

        If Conditions(7) Then
            MaterialsUpperPart = m_PackQuality(i, m_C_使用保護材)
            IsMatched = True
            Debug.Print "ヒット行 管理番号=" & m_PackQuality(i, m_C_番号) '& " / 配列i=" & i  '←デバッグ用：管理番号確認
            Exit For
        End If

NextRow:
    Next i

    ' --- 一致しなかった場合の処理 ---
    If Not IsMatched Then
        MaterialsUpperPart = "一致なし"
    End If

End Function

'直の開始終了時間の配列吐出し
Public Function ShiftTime() As Variant 'ShiftTime(ShiftT As String, SorE As String) As String  SorE　開始　終了
Dim FPath As String:         FPath = QualityPath
Dim TableName As String:     TableName = "時間用"
Dim target As String:        target = "伝送用ファイル.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim PackF():                 PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
Dim i As Long, a As Long
If IsEmpty(PackF) Then
  FPath = PackingPath
  FilePath = FPath & target
  PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(PackF) Then Exit Function
End If
Dim C_開始 As Long, C_終了 As Long, C_直 As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "開始":  C_開始 = i
  Case "終了":  C_終了 = i
  Case "直":    C_直 = i
  End Select
Next i
Dim SQLfilter As String: SQLfilter = "" '"[直]='" & ShiftT & "'"
Dim TempArr As Variant:  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(TempArr) Then
  ShiftTime = ""
  Exit Function
End If
Dim Temp(1 To 4, 1 To 3) As Variant
For i = 1 To 4
  For a = 1 To 3
    Temp(i, 1) = TempArr(i, C_直)
    Temp(i, 2) = TempArr(i, C_開始)
    Temp(i, 3) = TempArr(i, C_終了)
  Next a
Next i

ShiftTime = Temp

'Select Case SorE
'Case "S": ShiftTime = TempArr(1, C_開始)
'Case "E": ShiftTime = TempArr(1, C_終了)
'End Select
End Function

Public Function StraightEdge(LineName As String) As Variant
Dim FPath As String:         FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
Dim TableName As String:     TableName = "校正ストレートエッヂ"
Dim target As String:        target = "伝送用ファイル.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim PackF():                 PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
Dim i As Long, a As Long
If IsEmpty(PackF) Then
  FPath = PackingPath
  FilePath = FPath & target
  PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(PackF) Then Exit Function
End If
Dim C_ライン名 As Long, C_長さ下 As Long, C_長さ上 As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "ライン名":  C_ライン名 = i
  Case "長さ下":    C_長さ下 = i
  Case "長さ上":    C_長さ上 = i
  End Select
Next i
Dim SQLfilter As String: SQLfilter = "[ライン名]='" & LineName & "'"
Dim TempArr As Variant:  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)
If IsEmpty(TempArr) Then
  StraightEdge = ""
  Exit Function
End If
Dim Temp(1 To 2, 1 To 1) As Variant
Temp(1, 1) = TempArr(LBound(TempArr, 1), C_長さ下)
Temp(2, 1) = TempArr(UBound(TempArr, 1), C_長さ上)

StraightEdge = Temp
End Function
'Lotについている引当で納入先だけで配列作成(いくつついているか知りたい)
'2024.7.24_取引先追加
Public Function HikiArrCreate(LotNo As String, Optional n As String = "") As Variant '1次元
If LotNo = "" Then Exit Function
Dim TableName As String:       TableName = "仕掛"
Dim FPath As String:           FPath = MainPath
Dim Target_Hiki As String:     Target_Hiki = "SIKAHIKINOW.accdb"
Dim FilePath_Hiki As String:   FilePath_Hiki = FPath & Target_Hiki
Dim fieldsfilter As String:    fieldsfilter = ""
Dim Hikifields As Variant:     Hikifields = GetFieldsArr(FilePath_Hiki, TableName, fieldsfilter)
If IsEmpty(Hikifields) Then
  FPath = NextPath
  FilePath_Hiki = FPath & Target_Hiki
  Hikifields = GetFieldsArr(FilePath_Hiki, TableName, fieldsfilter)
  If IsEmpty(Hikifields) Then Exit Function
End If

Dim i As Long, H_受注_納入先名称 As Long, H_受注_取引先名称 As Long
For i = LBound(Hikifields, 2) To UBound(Hikifields, 2)
  Select Case Hikifields(1, i)
  Case "受注_納入先名称": H_受注_納入先名称 = i
  Case "受注_取引先名称": H_受注_取引先名称 = i
  End Select
Next i

Dim SQLfilter As String:   SQLfilter = "[ﾛｯﾄ番号]='" & LotNo & "'" 'ﾛｯﾄNo検索
Dim TempHiki As Variant:   TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter)
If IsEmpty(TempHiki) Then
  FPath = NextPath
  FilePath_Hiki = FPath & Target_Hiki
  TempHiki = GetRecordsArr(FilePath_Hiki, TableName, SQLfilter)
  If IsEmpty(TempHiki) Then Exit Function
End If

Dim HikiOrder As Variant, H_count As Long: H_count = UBound(TempHiki, 1)
ReDim HikiOrder(1 To H_count, 1 To 2)
For i = LBound(TempHiki, 1) To UBound(TempHiki, 1)
'  If n = "納入" Then HikiOrder(i) = TempHiki(i, H_受注_納入先名称)
'  If n = "取引" Then HikiOrder(i) = TempHiki(i, H_受注_取引先名称) '2024.7.24
'2024.7.24
   HikiOrder(i, 1) = TempHiki(i, H_受注_納入先名称)
   HikiOrder(i, 2) = TempHiki(i, H_受注_取引先名称)
Next i
HikiArrCreate = HikiOrder
End Function

Public Function DistArray(LotNo As String, 検査No As String, Optional TempPath As String = "", Optional Execution As String = "") As Variant  '複数引き当ての場合に歪公差が複数あるかをチェックする
Dim TempD As Variant:   TempD = HikiArrCreate(LotNo) '"納入","取引"_2024.7.24

If IsEmpty(TempD) Then Exit Function '2024.9.5_エラー対策

Dim i As Long
Dim Result() As Variant
ReDim Result(LBound(TempD, 1) To UBound(TempD, 1))
For i = LBound(TempD, 1) To UBound(TempD, 1) '納入だけ抜き取る
  Result(i) = TempD(i, 1)
Next i

Dim LotPath As String, KenPath As String
If IsEmptyArray(TempD) Then
  DistArray = Array() ' 空の配列を返す
  MsgBox "引当なし"
  Exit Function
End If

Dim a As Long, 納入 As String, 取引 As String
Dim j As Long: j = UBound(TempD)
ReDim WidthLength(1 To j) As String
Dim FormerArr As Variant, TopFormerArr As Variant

If AllElementsSame(Result) = False Then '違う納入がある　1次元比較
  For i = LBound(TempD) To UBound(TempD) '納入数分　公差を格納
    LotPath = LotNo 'ループ中に引数消える...
    KenPath = 検査No
    納入 = TempD(i, 1)
    取引 = TempD(i, 2)
    
    FormerArr = DBQP(LotPath, KenPath, 納入, 取引, TempPath, Execution) 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
    If i = 1 Then TopFormerArr = FormerArr '初めにかかった引き当て分を残す
    If IsEmptyArray(FormerArr) = False Then
      WidthLength(i) = FormerArr(1, 2) & "_" & FormerArr(1, 3) '全公差抜き取り
    End If
  Next i
  
  Dim MinW_Value As Double, MinL_Value As Double
  Dim WidthArray() As Variant, LengthArray() As Variant
  Dim Widthnumeric() As Double, Lengthnumeric() As Double
  
  If AllElementsSame(WidthLength) = False Then '違う公差がある
    SplitWidthLength WidthLength, WidthArray, LengthArray '分解
    ReDim Widthnumeric(1 To UBound(WidthLength)) '文字列の場合MinMaxが使用できないので数値変換
    ReDim Lengthnumeric(1 To UBound(LengthArray))
    For a = LBound(WidthArray) To UBound(WidthArray)
      Widthnumeric(a) = val(WidthArray(a))
    Next
    For a = LBound(LengthArray) To UBound(LengthArray)
      Lengthnumeric(a) = val(LengthArray(a))
    Next
    MinW_Value = WorksheetFunction.Min(Widthnumeric)
    MinL_Value = WorksheetFunction.Min(Lengthnumeric)
    
'    '引当数分0チェック　0の表示はいらない次にある数値を表示する
    If MinW_Value = 0 Then MinW_Value = SecondSmallest(Widthnumeric)
    If MinL_Value = 0 Then MinL_Value = SecondSmallest(Lengthnumeric)
'    Stop
    
    TopFormerArr(LBound(TopFormerArr, 1), 2) = MinW_Value '一番厳しい公差で上書き
    TopFormerArr(LBound(TopFormerArr, 1), 3) = MinL_Value
    TopFormerArr(LBound(TopFormerArr, 1), 6) = TopFormerArr(LBound(TopFormerArr, 1), 6) & vbCrLf & "納入先による公差違い有,厳しい公差で上書き"
  End If
Else '違う納入がない　1引き当て
  
  納入 = TempD(1, 1)
  取引 = TempD(1, 2)

  TopFormerArr = DBQP(LotNo, 検査No, 納入, 取引, TempPath, Execution) 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号、フラグ
End If
DistArray = TopFormerArr
'Stop
End Function
'歪反映
Public Sub FormReflectDist(Kousa As Variant, Optional Execution As String = "")
Dim MinimumValue As Long: MinimumValue = LBound(Kousa, 1)
Dim ws As Worksheet, RngRock As Range

If UF品質基準.Visible = False And UFMitsubishi.Visible = False Then  '品質基準フォームを起動していない
  If IsNumeric(Kousa(MinimumValue, 1)) Then UF1.MG = Format(Kousa(MinimumValue, 1), "0.0")
  If IsNumeric(Kousa(MinimumValue, 1)) = False Then UF1.MG = "" 'Kousa(MinimumValue, 1)
  
  If UF1.MG = 0 Then UF1.MG = ""
  
  If IsNumeric(Kousa(MinimumValue, 2)) Then UF1.HaHiz = Format(Kousa(MinimumValue, 2), "0.0")
  If IsNumeric(Kousa(MinimumValue, 2)) = False Then UF1.HaHiz = Kousa(MinimumValue, 2)
  
  If UF1.HaHiz = 0 Then UF1.HaHiz = ""
  
  If IsNumeric(Kousa(MinimumValue, 3)) Then UF1.TaHiz = Format(Kousa(MinimumValue, 3), "0.0")
  If IsNumeric(Kousa(MinimumValue, 3)) = False Then UF1.TaHiz = Kousa(MinimumValue, 3)
  
  If UF1.TaHiz = 0 Then UF1.TaHiz = ""
  
  If IsNumeric(Kousa(MinimumValue, 4)) Then UF1.Naka = Format(Kousa(MinimumValue, 4), "0.0")
  If IsNumeric(Kousa(MinimumValue, 4)) = False Then UF1.Naka = "" 'Kousa(MinimumValue, 4)
  
  If UF1.Naka = 0 Then UF1.Naka = ""
  
  If IsNumeric(Kousa(MinimumValue, 5)) Then UF1.Mimi = Format(Kousa(MinimumValue, 5), "0.0")
  If IsNumeric(Kousa(MinimumValue, 5)) = False Then UF1.Mimi = "" ' Kousa(MinimumValue, 5)
  
  If UF1.Mimi = 0 Then UF1.Mimi = ""
  
  UF1.HizCome.Caption = Kousa(MinimumValue, 6)
  If Right(UF1.HizCome.Caption, 1) = "_" Then UF1.HizCome.Caption = Left(UF1.HizCome.Caption, Len(UF1.HizCome.Caption) - 1)
  Dim tMP As String:         tMP = UF1.HizCome.Caption
  Dim ModifiedTmp As String: ModifiedTmp = Replace(tMP, "。", vbCrLf)
  UF1.HizCome.Caption = ModifiedTmp
  
  UF1.DistortionFrame.Caption = "参照元:" & Kousa(MinimumValue, 7)
ElseIf UF品質基準.Visible Then 'ゲージ、幅歪、丈歪、中、耳、コメント、参照元、番号
  With UF品質基準
    If IsNumeric(Kousa(MinimumValue, 1)) Then .MG1 = Format(Kousa(MinimumValue, 1), "0.0")
    If IsNumeric(Kousa(MinimumValue, 1)) = False Then .MG1 = Kousa(MinimumValue, 1)
  
    If IsNumeric(Kousa(MinimumValue, 2)) Then .HaHiz1 = Format(Kousa(MinimumValue, 2), "0.0")
    If IsNumeric(Kousa(MinimumValue, 2)) = False Then .HaHiz1 = Kousa(MinimumValue, 2)
  
    If IsNumeric(Kousa(MinimumValue, 3)) Then .TaHiz1 = Format(Kousa(MinimumValue, 3), "0.0")
    If IsNumeric(Kousa(MinimumValue, 3)) = False Then .TaHiz1 = Kousa(MinimumValue, 3)
  
    If IsNumeric(Kousa(MinimumValue, 4)) Then .Naka1 = Format(Kousa(MinimumValue, 4), "0.0")
    If IsNumeric(Kousa(MinimumValue, 4)) = False Then .Naka1 = Kousa(MinimumValue, 4)
  
    If IsNumeric(Kousa(MinimumValue, 5)) Then .Mimi1 = Format(Kousa(MinimumValue, 5), "0.0")
    If IsNumeric(Kousa(MinimumValue, 5)) = False Then .Mimi1 = Kousa(MinimumValue, 5)
  
    .Sansyou1 = Kousa(MinimumValue, 7) & "_" & Kousa(MinimumValue, 6)
    If Right(.Sansyou1, 1) = "_" Then .Sansyou1 = Left(.Sansyou1, Len(.Sansyou1) - 1)
  End With
ElseIf UF1.Visible = False And WsCheck("全ヵ所検査") And UFMitsubishi.Visible Then
  Set ws = ThisWorkbook.Sheets("全ヵ所検査")
  ws.Unprotect 'シート保護解除
  Set RngRock = ws.Range(ws.Cells(5, 13), ws.Cells(5, 15))
  RngRock.Locked = False
  RngRock.Font.Size = 8
'  WS.Cells(1, 25) = "幅歪公差"
'  WS.Cells(1, 27) = "丈歪公差"
  ws.Cells(5, 13) = Kousa(MinimumValue, 2)
  ws.Cells(5, 15) = Kousa(MinimumValue, 3)
  RngRock.Locked = True
End If
End Sub
'配列内の数値比較
Function UpdateQpArryD(QpArryD As Variant, QpArryG As Variant) As Variant
'配列の要素を数値として取得
Dim QpArryD_1_2 As Double: QpArryD_1_2 = val(QpArryD(1, 2))
Dim QpArryD_1_3 As Double: QpArryD_1_3 = val(QpArryD(1, 3))
Dim QpArryG_1_2 As Double: QpArryG_1_2 = val(QpArryG(1, 2))
Dim QpArryG_1_3 As Double: QpArryG_1_3 = val(QpArryG(1, 3))
Dim Changed As Boolean: Changed = False

If (QpArryD_1_2 <> QpArryG_1_2) Then Changed = True
If (QpArryD_1_2 > QpArryG_1_2) Then '各要素を比較し、小さい方をQpArryDに格納
  QpArryD(1, 2) = QpArryG_1_2
  Changed = True
End If
If (QpArryD_1_3 <> QpArryG_1_3) Then Changed = True
If (QpArryD_1_3 > QpArryG_1_3) Then
  QpArryD(1, 3) = QpArryG_1_3
  Changed = True
End If

If Changed Then '変更があった場合にQpArryD(1, 6)にメッセージを追加
  QpArryD(1, 6) = QpArryD(1, 6) & vbCrLf & "納入先と取引先よる公差違い有,厳しい公差で上書き"
End If

UpdateQpArryD = QpArryD '更新されたQpArryDを返す
End Function

'設備仕様を満たしているかチェック_2024.8.19
Public Function CheckEquipmentDb(line As String, Zai As String, ATU As Single, HABA As Single, TAKE As Single, Wei As Single) As String  'DBから出力テキストを返す
'ライン,材質,厚,幅,丈,単重
Dim FPath As String:         FPath = QualityPath '"\\nlmsrvngy03\各課共有\0130_日軽稲沢\梱包課\AIM\検査\"
Dim TableName As String:     TableName = "設備仕様"
Dim target As String:        target = "伝送用ファイル.accdb"
Dim FilePath As String:      FilePath = FPath & target
Dim fieldsfilter As String:  fieldsfilter = ""
Dim PackF():                 PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)

If IsEmpty(PackF) Then
  FPath = PackingPath
  FilePath = FPath & target
  PackF = GetFieldsArr(FilePath, TableName, fieldsfilter)
  If IsEmpty(PackF) Then Exit Function
End If

Dim C_番号 As Long, C_設備名 As Long, C_製造材質 As Long, C_オーダー厚MIN As Long, C_オーダー厚MAX As Long, _
C_オーダー幅MIN As Long, C_オーダー幅MAX As Long, C_オーダー丈MIN As Long, C_オーダー丈MAX As Long, i As Long, _
C_単重MIN As Long, C_単重MAX As Long, C_作業可否 As Long, C_備考 As Long
For i = LBound(PackF, 2) To UBound(PackF, 2)
  Select Case PackF(1, i)
  Case "管理番号":      C_番号 = i
  Case "設備名":        C_設備名 = i
  Case "製造材質":      C_製造材質 = i
  Case "オーダー厚MIN": C_オーダー厚MIN = i
  Case "オーダー厚MAX": C_オーダー厚MAX = i
  Case "オーダー幅MIN": C_オーダー幅MIN = i
  Case "オーダー幅MAX": C_オーダー幅MAX = i
  Case "オーダー丈MIN": C_オーダー丈MIN = i
  Case "オーダー丈MAX": C_オーダー丈MAX = i
  Case "単重MIN":       C_単重MIN = i
  Case "単重MAX":       C_単重MAX = i
  Case "作業可否":      C_作業可否 = i
  Case "備考":          C_備考 = i
  End Select
Next i

Dim SQLfilter As String: SQLfilter = "[設備名] LIKE " & "'" & line & "'"
Dim TempArr As Variant:  TempArr = GetRecordsArr(FilePath, TableName, SQLfilter)

If IsEmpty(TempArr) Then '設備名なし
  CheckEquipmentDb = "可"
  Exit Function
End If

Dim MatchingArray() As Long ' マッチングしたインデックスを格納する配列
ReDim MatchingArray(1 To UBound(TempArr, 1)) ' 配列のサイズを設定

Dim Conditions(1 To 5) As Boolean ' 各条件の結果
Dim j As Long, a As Long, n As Long: n = 0

For i = LBound(TempArr, 1) To UBound(TempArr, 1)
'If TempArr(i, C_番号) = 2 Then 'Test
'Stop
'End If
  Conditions(1) = CheckCondition(i, C_製造材質, Zai, TempArr)
  Conditions(2) = Conditions(1) And CheckWithin(i, C_オーダー厚MIN, C_オーダー厚MAX, ATU, TempArr) '厚さのチェック
  Conditions(3) = Conditions(2) And CheckWithin(i, C_オーダー幅MIN, C_オーダー幅MAX, HABA, TempArr) '幅のチェック
  Conditions(4) = Conditions(3) And CheckWithin(i, C_オーダー丈MIN, C_オーダー丈MAX, TAKE, TempArr) '丈のチェック
  Conditions(5) = Conditions(4) And CheckWithin(i, C_単重MIN, C_単重MAX, Wei, TempArr) '単重のチェック
 ' 条件が全て満たされた場合の処理
  If Conditions(1) And Conditions(2) And Conditions(3) _
  And Conditions(4) And Conditions(5) Then
    n = n + 1
    MatchingArray(n) = TempArr(i, C_番号)
  End If
Next i

If n = 0 Then '一致なしの場合作業不可
  If Conditions(2) = False Then
    CheckEquipmentDb = "オーダー厚 範囲外"
    Exit Function
  ElseIf Conditions(3) = False Then
    CheckEquipmentDb = "C_オーダー幅 範囲外"
    Exit Function
  ElseIf Conditions(4) = False Then
    CheckEquipmentDb = "オーダー丈 範囲外"
    Exit Function
  ElseIf Conditions(5) = False Then
    CheckEquipmentDb = "単重 範囲外"
    Exit Function
  End If
End If

' MatchingArrayから0を削除する
Dim NewArray() As Long
ReDim NewArray(1 To n)
For i = 1 To n
  NewArray(i) = MatchingArray(i)
Next i
'格納
Dim HitArry()
ReDim HitArry(1 To UBound(NewArray, 1), 1 To UBound(TempArr, 2))
For a = LBound(NewArray, 1) To UBound(NewArray, 1)
  For i = LBound(TempArr, 1) To UBound(TempArr, 1)
    For j = LBound(TempArr, 2) To UBound(TempArr, 2)
      If TempArr(i, C_番号) = NewArray(a) Then
        HitArry(a, j) = TempArr(i, j)
      End If
    Next j
  Next i
Next a
For i = LBound(HitArry, 1) To UBound(HitArry, 1)
  If HitArry(i, C_作業可否) <> "" Then CheckEquipmentDb = HitArry(i, C_作業可否)
Next i

End Function

'========================================
' 外注条件 判定（品質基準.accdb）
'   外注条件 / 外注条件明細 を読み、該当すれば包装仕様書NOを返す
'   初回呼出時に GetFieldsArr / GetRecordsArr で読込み、以後はメモリ上で判定
'
'【条件の追加方法】Access でテーブルを直接開いて行を足す
'
' 手順1：外注条件（見出し）に1行追加
'   管理番号 … 既存と重複しない番号（続き番号でよい）
'   条件名   … 人が見て分かる名前
'   戻り値   … 条件がすべて成立したとき PackingNoCheck が返す包装仕様書NO
'   優先順位 … 複数の条件セットが同時に成立したとき、数字の小さいほうを採用
'   有効     … チェックを外すとその条件セットは判定されない
'
' 手順2：外注条件明細に条件を足す
'   条件管理番号 … 手順1で決めた管理番号
'   管理番号     … オートナンバーなので入力不要
'
'【判定の決まり】
'   同じ条件番号の中 … どれか1つ成立すればOK（OR）
'   条件番号どうし   … すべて成立が必要（AND）
'
'   例）材質2ST かつ 調質O かつ 厚25.5～30.5
'     条件管理番号 条件番号 対象項目 判定方法 判定値
'         3          1       材質      =      2ST
'         3          2       調質      =      O
'         3          3       厚        >=     25.5    ←下限
'         3          4       厚        <=     30.5    ←上限
'   ※厚の範囲は条件番号を分けること。同じ番号に >= と <= を入れると
'     「どちらか一方でOK」になり、厚10でも厚100でも成立してしまう
'
'   例）用途ｺｰﾄﾞが K434 または K435 → 同じ条件番号で2行
'
'【判定方法】
'   =     完全一致        例）2ST
'   Like  ワイルドカード  例）*52S（52Sで終わる） T6*（T6で始まる）
'   含む  部分一致        例）GCT
'   >=    以上（数値）    例）25.5
'   <=    以下（数値）    例）30.5
'   ※厚は数値として比較。材質・調質・ｺｰｽ・用途ｺｰﾄﾞは文字列として比較
'
'【対象項目】Lot情報履歴シートの1行目（見出し）と同じ文字を入れる
'   材質 / 調質 / 厚 / ｺｰｽ（半角ｶﾅ） / 用途ｺｰﾄﾞ（ｺｰﾄﾞが半角ｶﾅ）
'   見出しにある項目なら上記以外（巾 など）も指定可
'   1文字でも違うと一致せず、その条件は成立しない
'
'【テーブルを直した後】ReloadGaichu を実行すると反映される
'                      （フォームを開き直しても反映される）


'2026.09.18
'========================================
' 判定本体
'   Hist   : Lot情報履歴の配列（1行目が見出し）
'   LotNo  : 対象ﾛｯﾄNo
'   R_Lot  : ﾛｯﾄNo の列番号
'   戻り値 : 該当する包装仕様書NO／該当なしは ""
'========================================
Public Function GaichuPackingNo(ByRef Hist() As Variant, ByVal LotNo As String, ByVal R_Lot As Long) As String
    If Not EnsureLoaded() Then Exit Function
    If IsEmpty(m_Head) Or IsEmpty(m_Det) Then Exit Function

    Dim h As Long, bestPri As Double, hit As Boolean
    Dim headNo As Double

    For h = LBound(m_Head, 1) To UBound(m_Head, 1)
        If CBool(m_Head(h, m_HeadIdx(H_VALID))) Then
            headNo = CDbl(m_Head(h, m_HeadIdx(H_NO)))
            If MatchCondSet(Hist, LotNo, R_Lot, headNo) Then
                If Not hit Or CDbl(m_Head(h, m_HeadIdx(H_PRI))) < bestPri Then
                    hit = True
                    bestPri = CDbl(m_Head(h, m_HeadIdx(H_PRI)))
                    GaichuPackingNo = CStr(m_Head(h, m_HeadIdx(H_RET)) & "")
                End If
            End If
        End If
    Next h
End Function

'--- 条件セット1件を判定（条件番号どうしはAND、同番号内はOR）
Private Function MatchCondSet(ByRef Hist() As Variant, ByVal LotNo As String, _
                              ByVal R_Lot As Long, ByVal headNo As Double) As Boolean
    Dim condNos As Variant
    condNos = GetCondNos(headNo)
    If IsEmpty(condNos) Then Exit Function      '明細がなければ不成立（全ロット該当を防ぐ）

    Dim k As Long
    For k = LBound(condNos) To UBound(condNos)
        If Not MatchCondNo(Hist, LotNo, R_Lot, headNo, condNos(k)) Then Exit Function
    Next k
    MatchCondSet = True
End Function

'--- 条件番号1つを判定（同番号内のいずれか1行が成立すればTrue）
'    旧コード同様、ﾛｯﾄNo一致行のどれかで成立すればよい
Private Function MatchCondNo(ByRef Hist() As Variant, ByVal LotNo As String, ByVal R_Lot As Long, _
                             ByVal headNo As Double, ByVal condNo As Double) As Boolean
    Dim D As Long, i As Long, col As Long
    Dim target As String, method As String, judgeVal As String

    For D = LBound(m_Det, 1) To UBound(m_Det, 1)
        If CDbl(m_Det(D, m_DetIdx(D_HNO))) = headNo Then
            If CDbl(m_Det(D, m_DetIdx(D_CNO))) = condNo Then
                If CBool(m_Det(D, m_DetIdx(D_VALID))) Then
                    target = CStr(m_Det(D, m_DetIdx(D_TGT)) & "")
                    method = CStr(m_Det(D, m_DetIdx(D_MTD)) & "")
                    judgeVal = CStr(m_Det(D, m_DetIdx(D_VAL)) & "")
                    col = FindHistCol(Hist, target)
                    If col > 0 Then
                        For i = LBound(Hist, 1) To UBound(Hist, 1)
                            If CStr(Hist(i, R_Lot) & "") = LotNo Then
                                If JudgeOne(Hist(i, col), method, judgeVal) Then
                                    MatchCondNo = True
                                    Exit Function
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    Next D
End Function

'--- 1セルの判定
Private Function JudgeOne(ByVal cellVal As Variant, ByVal method As String, ByVal judgeVal As String) As Boolean
    Dim S As String: S = CStr(cellVal & "")

    Select Case method
        Case "="
            If IsNumeric(S) And IsNumeric(judgeVal) Then
                JudgeOne = (Round(CDbl(S), 3) = Round(CDbl(judgeVal), 3))
            Else
                JudgeOne = (S = judgeVal)
            End If
        Case "Like"
            JudgeOne = (S Like judgeVal)
        Case "含む"
            If Len(judgeVal) > 0 Then JudgeOne = (InStr(S, judgeVal) > 0)
        Case ">="
            If IsNumeric(S) And IsNumeric(judgeVal) Then JudgeOne = (CDbl(S) >= CDbl(judgeVal))
        Case "<="
            If IsNumeric(S) And IsNumeric(judgeVal) Then JudgeOne = (CDbl(S) <= CDbl(judgeVal))
        Case Else
            '未知の判定方法 → 不成立
    End Select
End Function

'--- 見出し行から対象項目の列番号を探す（見つからなければ0）
Private Function FindHistCol(ByRef Hist() As Variant, ByVal target As String) As Long
    Dim i As Long
    For i = LBound(Hist, 2) To UBound(Hist, 2)
        If CStr(Hist(1, i) & "") = target Then FindHistCol = i: Exit Function
    Next i
End Function

'--- 指定条件セットに含まれる条件番号の一覧
Private Function GetCondNos(ByVal headNo As Double) As Variant
    Dim nos() As Double, n As Long
    Dim D As Long, k As Long, exists As Boolean
    Dim v As Double

    ReDim nos(0 To 0)
    n = 0
    For D = LBound(m_Det, 1) To UBound(m_Det, 1)
        If CDbl(m_Det(D, m_DetIdx(D_HNO))) = headNo Then
            If CBool(m_Det(D, m_DetIdx(D_VALID))) Then
                v = CDbl(m_Det(D, m_DetIdx(D_CNO)))
                exists = False
                For k = 1 To n
                    If nos(k) = v Then exists = True: Exit For
                Next k
                If Not exists Then
                    n = n + 1
                    ReDim Preserve nos(0 To n)
                    nos(n) = v
                End If
            End If
        End If
    Next D

    If n = 0 Then Exit Function      'Empty を返す
    Dim ret() As Double
    ReDim ret(1 To n)
    For k = 1 To n
        ret(k) = nos(k)
    Next k
    GetCondNos = ret
End Function

'--- テーブル修正後に読み直す
Public Sub ReloadGaichu()
    m_Loaded = False
    m_LoadFailed = False
    m_Head = Empty
    m_Det = Empty
    If EnsureLoaded() Then
        MsgBox "外注条件テーブルを読み込み直しました", vbInformation
    End If
End Sub

'========================================
' 読込（初回のみ）
'========================================
Private Function EnsureLoaded() As Boolean
    If m_Loaded Then EnsureLoaded = True: Exit Function
    If m_LoadFailed Then Exit Function      '失敗済みなら再試行しない

    Dim FilePath As String
    Dim step As String
    Dim errNo As Long, errDesc As String

    On Error GoTo ErrHandler

    FilePath = QualityPath & "品質基準.accdb"

    step = TBL_H & " 読込"
    Call LoadTable(FilePath, TBL_H, _
        Array("管理番号", "条件名", "戻り値", "優先順位", "有効"), _
        m_Head, m_HeadIdx)

    step = TBL_D & " 読込"
    Call LoadTable(FilePath, TBL_D, _
        Array("条件管理番号", "条件番号", "対象項目", "判定方法", "判定値", "有効"), _
        m_Det, m_DetIdx)

    m_Loaded = True
    EnsureLoaded = True
    Exit Function

ErrHandler:
    errNo = Err.Number: errDesc = Err.Description   '★先に退避
    m_LoadFailed = True
    MsgBox "外注条件テーブルを読み込めませんでした" & vbCrLf & _
           "段階: " & step & vbCrLf & _
           "内容: " & errNo & " " & errDesc & vbCrLf & vbCrLf & _
           "外注梱包の判定は行われません。原因を解消後 ReloadGaichu を実行してください。", vbCritical
End Function

'--- 1テーブル分：列名から列番号を求め、レコードを取得
Private Sub LoadTable(ByVal FilePath As String, ByVal TableName As String, ByVal colNames As Variant, _
                      ByRef data As Variant, ByRef idx() As Long)
    Dim fieldsfilter As String:  fieldsfilter = ""
    Dim SQLfilter As String:     SQLfilter = ""
    Dim flds As Variant
    Dim hdr As Long, i As Long, k As Long

    flds = GetFieldsArr(FilePath, TableName, fieldsfilter)
    If IsEmpty(flds) Then Err.Raise vbObjectError + 1, , "「" & TableName & "」のフィールドを取得できません"

    hdr = LBound(flds, 1)
    ReDim idx(LBound(colNames) To UBound(colNames))
    For k = LBound(colNames) To UBound(colNames)
        idx(k) = -1
        For i = LBound(flds, 2) To UBound(flds, 2)
            If flds(hdr, i) = colNames(k) Then idx(k) = i: Exit For
        Next i
        If idx(k) = -1 Then
            Err.Raise vbObjectError + 2, , "「" & TableName & "」に「" & colNames(k) & "」フィールドがありません"
        End If
    Next k

    data = GetRecordsArr(FilePath, TableName, SQLfilter)
End Sub


'========================================
' ゲージ判定（品質基準.accdb / ゲージ判定条件）
'   True  … ゲージ検査（幅ｹﾞｰｼﾞ・丈ｹﾞｰｼﾞ）
'   False … 定盤検査
'
'【テーブルの読み方】1行＝1ルール
'   1行の中で、値が入っている項目はすべて満たす必要がある（AND）
'   空欄の項目は見ない。どれか1行でも成立すればゲージ判定（OR）
'   重量_超 … 重量 > 値(kg)    丈_以上 … 丈 >= 値(mm)   丈_未満 … 丈 < 値
'   丈_超   … 丈 > 値           幅_未満 … 幅 < 値
'   材質・調質 … 完全一致       コース含む … その文字を含む
'   例）重量1300超 かつ 丈2500未満 → 重量_超=1300、丈_未満=2500、他は空欄
'   ※すべて空欄の行は無視する
'【テーブルを直した後】ReloadGauge を実行
'========================================
Public Function IsGaugeRequired(ByVal 厚 As Double, ByVal 幅 As Double, ByVal 丈 As Double, _
                                ByVal 材質 As String, ByVal 調質 As String, _
                                ByVal コース As String) As Boolean
    If Not EnsureLoadedGauge() Then Exit Function
    If IsEmpty(m_Rule) Then Exit Function

    Dim 重量 As Double: 重量 = Weight(CDbl(厚), CDbl(幅), CDbl(丈))
    Dim r As Long, ok As Boolean, nCond As Long
    Dim v As Variant

    For r = LBound(m_Rule, 1) To UBound(m_Rule, 1)
        If CBool(m_Rule(r, m_GaugeIdx(G_VALID))) Then
            ok = True
            nCond = 0

            v = m_Rule(r, m_GaugeIdx(G_W_GT))
            If HasVal(v) Then nCond = nCond + 1: If Not (重量 > CDbl(v)) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_L_GE))
            If HasVal(v) Then nCond = nCond + 1: If Not (丈 >= CDbl(v)) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_L_LT))
            If HasVal(v) Then nCond = nCond + 1: If Not (丈 < CDbl(v)) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_L_GT))
            If HasVal(v) Then nCond = nCond + 1: If Not (丈 > CDbl(v)) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_H_LT))
            If HasVal(v) Then nCond = nCond + 1: If Not (幅 < CDbl(v)) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_ZAI))
            If HasVal(v) Then nCond = nCond + 1: If 材質 <> CStr(v) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_TYU))
            If HasVal(v) Then nCond = nCond + 1: If 調質 <> CStr(v) Then ok = False
            v = m_Rule(r, m_GaugeIdx(G_CRS))
            If HasVal(v) Then nCond = nCond + 1: If InStr(コース, CStr(v)) = 0 Then ok = False

            If ok And nCond > 0 Then
                IsGaugeRequired = True
                Exit Function
            End If
        End If
    Next r
End Function

'--- 値が入っているか（Null・空文字は「条件なし」）
Private Function HasVal(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    HasVal = (Len(Trim(CStr(v))) > 0)
End Function

'--- テーブル修正後に読み直す
Public Sub ReloadGauge()
    m_GaugeLoaded = False
    m_GaugeLoadFailed = False
    m_Rule = Empty
    If EnsureLoadedGauge() Then MsgBox "ゲージ判定条件を読み込み直しました", vbInformation
End Sub

'--- 読込（初回のみ）※LoadTable は外注条件側のものを共用
Private Function EnsureLoadedGauge() As Boolean
    If m_GaugeLoaded Then EnsureLoadedGauge = True: Exit Function
    If m_GaugeLoadFailed Then Exit Function

    Dim FilePath As String
    Dim errNo As Long, errDesc As String

    On Error GoTo ErrHandler
    FilePath = QualityPath & "品質基準.accdb"
    Call LoadTable(FilePath, TBL_G, _
        Array("有効", "重量_超", "丈_以上", "丈_未満", "丈_超", "幅_未満", "材質", "調質", "コース含む"), _
        m_Rule, m_GaugeIdx)

    m_GaugeLoaded = True
    EnsureLoadedGauge = True
    Exit Function

ErrHandler:
    errNo = Err.Number: errDesc = Err.Description   '★先に退避
    m_GaugeLoadFailed = True
    MsgBox "ゲージ判定条件を読み込めませんでした" & vbCrLf & _
           "内容: " & errNo & " " & errDesc & vbCrLf & vbCrLf & _
           "定盤表示になります。原因を解消後 ReloadGauge を実行してください。", vbCritical
End Function

