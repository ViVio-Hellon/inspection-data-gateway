Attribute VB_Name = "Msg"
Attribute VB_Base = "0{EC4E8E30-3776-4A4B-AA78-612B60490F4E}{7967FD4D-9211-45EB-AC91-A4303089F479}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
Public CLNo As Long 'MsgForm用戻り変数
'vbOK     : 1 「OK」ボタン
'vbCancel : 2 「キャンセル」ボタン
'vbAbort  : 3 「中止」ボタン
'vbRetry  : 4 「再試行」ボタン
'vbIgnore : 5 「無視」ボタン
'vbYes    : 6 「はい」ボタン
'vbNo     : 7 「いいえ」ボタン
Private Sub C1_Click()
CLNo = NumSet(Me.C1.Caption)  'Me.ClickNo = NumSet(Me.C3.Caption)
'Me.Hide
Unload Me 'Me.Hideが使えない場合、返り値が初期化されるため、Public変数を使用
End Sub
Private Sub C2_Click()
CLNo = NumSet(Me.C2.Caption)  'Me.ClickNo = NumSet(Me.C3.Caption)
'Me.Hide
Unload Me 'Me.Hideが使えない場合、返り値が初期化されるため、Public変数を使用
End Sub
Private Sub C3_Click()
CLNo = NumSet(Me.C3.Caption) 'Me.ClickNo = NumSet(Me.C3.Caption)
'Me.Hide
Unload Me 'Me.Hideが使えない場合、返り値が初期化されるため、Public変数を使用
End Sub
'ボタン名称を処理用番号に変換
Private Function NumSet(BtCap As String) As Long
Dim Temp As Long
Select Case BtCap
  Case "OK":          Temp = 1 '[OK]ﾎﾞﾀﾝが選択された。
  Case "キャンセル":  Temp = 2 '[キャンセル]ﾎﾞﾀﾝが選択された。
  Case "中止":        Temp = 3 '[中止]ﾎﾞﾀﾝが選択された。
  Case "再試行":      Temp = 4 '[再試行]ﾎﾞﾀﾝが選択された。
  Case "無視":        Temp = 5 '[無視]ﾎﾞﾀﾝが選択された。
  Case "はい":        Temp = 6 '[はい]ﾎﾞﾀﾝが選択された。
  Case "いいえ":      Temp = 7 '[いいえ]ﾎﾞﾀﾝが選択された。
End Select
NumSet = Temp
End Function
'×ボタンをなくす場合のみ有効化
Private Sub UserForm_Activate()
'Call WindowMenuButtonDel(Me) 'ｳｨﾝﾄﾞｳ×ﾎﾞﾀﾝを消す。
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Cancel = True
If CloseMode = 0 Then '"×"ﾎﾞﾀﾝで閉じた場合
'  Debug.Print CloseMode
  Select Case Me.typeNo
    Case 0: CLNo = 1 'vbOKOnly            : 0「OK」ボタン
    Case 1: CLNo = 2 'vbOKCancel          : 1「OK」、「キャンセル」ボタン
    Case 2: CLNo = 3 'vbAbortRetryIgnore  : 2「中止」、「再試行」、「無視」
    Case 3: CLNo = 2 'vbYesNoCancel       : 3「はい」、「いいえ」、「キャンセル」
    Case 4: CLNo = 7 'vbYesNo             : 4「はい」、「いいえ」
    Case 5: CLNo = 2 'vbRetryCancel       : 5「再試行」、「キャンセル」
    Case 6: CLNo = 2
  End Select
End If
Unload Me 'Me.Hideが使えない場合、返り値が初期化されるため、Public変数(CLNo)を使用(【06】MsgFormﾓｼﾞｭｰﾙに記載)
'Me.Hide
End Sub
Public Function Box(msg As String, Optional ModeNo As Long = 0, Optional title As String = "Microsoft Excel", Optional FS As Single = 14)

Dim typeNo As Long: typeNo = ModeNo Mod 16 'VB定数から数値を取得
Dim IndiNo As Long: IndiNo = ModeNo \ 16 'VB定数から数値を取得
Dim LB As Long

'メモ
'Windows Logon.wav’ﾛｸﾞｵﾝ
'Windows Foreground.wav’一般の警告音
'Windows Background.wav’システムエラー
'Windows Proximity Notification.wav
'Windows Unlock.wav
'Windows エラー.wav
'Windows システム エラー.wav
'Windows のシャットダウン.wav
'Windows ハードウェアの挿入.wav
'Windows ハードウェアの取り外し.wav
'Windows ベル.wav
'Windows メッセージ (警告).wav
'Windows ユーザー アカウント制御.wav’ログイン失敗とか
'recycle.wav　’ゴミ捨て時の音
'Windows リサイクル.wav’ゴミ捨て時の音

Load Me
Me.Caption = title 'タイトル

Me.L1.FontSize = FS 'フォントサイズ
Me.L1.Left = 42 'アイコンセット用位置
Me.L1.Width = 366 'アイコンセット用幅
Me.L1 = msg '表示メッセージ

'コメントの長さによってフォームやコメント表示欄のサイズを変更する。
Dim ufL As Single: ufL = 100
Dim ufMtop As Single: ufMtop = 20
Dim ufFtop As Single: ufFtop = 24
Dim COUNTLF As Long: COUNTLF = 0 '改行コードの数

COUNTLF = UBound(Split(msg, vbLf))
LB = (LenB(msg) + COUNTLF * 8) \ FS * 2
Dim n As Long: n = FS * 2 * LB

Select Case LB
  Case 0 To 1
    Me.Icon.tOP = ufMtop + 24 '24
    Me.L1.tOP = ufMtop + 24 '24
    Me.L1.Height = 24
  Case Is >= 2
    Me.Icon.tOP = ufMtop
    Me.L1.tOP = ufMtop
    Me.L1.Height = ufMtop + n
    Me.Height = ufL + n * 1.2
    Me.F1.tOP = ufFtop + n + COUNTLF * 8
End Select

'ボタンの設定を行う
Select Case typeNo
  Case 0 'vbOKOnly            : 0「OK」ボタン
    Call BottomSetting("OK")
    Me.CloseTime = 1 '自動的に閉じるモード
  Case 1 'vbOKCancel          : 1「OK」、「キャンセル」ボタン
    Call BottomSetting("OK", "キャンセル")
  Case 2 'vbAbortRetryIgnore  : 2「中止」、「再試行」、「無視」
    Call BottomSetting("中止", "再試行", "無視")
  Case 3 'vbYesNoCancel       : 3「はい」、「いいえ」、「キャンセル」
    Call BottomSetting("はい", "いいえ", "キャンセル")
  Case 4 'vbYesNo             : 4「はい」、「いいえ」
    Call BottomSetting("はい", "いいえ")
  Case 5 'vbRetryCancel       : 5「再試行」、「キャンセル」
    Call BottomSetting("再試行", "キャンセル")
  Case 6
    Call BottomSetting("OK")
End Select

'閉じるを必ずキャンセル扱いにするための前準備
Me.typeNo = typeNo 'ﾌｫｰﾑに記録（閉じる処理対応用）

'アイコンの表示に関する処理
Select Case IndiNo
  Case 0
    Me.L1.Left = 30 'アイコン無し用メッセージ位置
    Me.L1.Width = 370    'アイコン無し用メッセージ位置
    Me.Icon.Visible = False
  Case 1 'vbCritical          : 16 警告メッセージアイコンを表示
    Me.Icon.Picture = Me.Ic1.Picture
    'PlaySound "C:\Windows\Media\" & "Windows Foreground.wav", 1, 1
  Case 2 'vbQuestion          : 32 問い合わせメッセージアイコンを表示
    Me.Icon.Picture = Me.Ic2.Picture
  Case 3 'vbExclamation       : 48 注意メッセージアイコンを表示
    Me.Icon.Picture = Me.Ic3.Picture
    Beep 'PlaySound "C:\Windows\Media\" & "Windows Background.wav", 1, 1
  Case 4 'vbInformation       : 64 情報メッセージアイコンを表示
    Me.Icon.Picture = Me.Ic4.Picture
    Beep 'PlaySound "C:\Windows\Media\" & "Windows Background.wav", 1, 1
End Select

Me.Show
'Box Me.ClickNo'返り値のチェック用
'Unload Me
'Debug.Print CLNo
'Box = Me.ClickNo
Box = CLNo

End Function
Private Sub BottomSetting(Str1 As String, Optional Str2 As String = "", Optional Str3 As String = "")
Dim i As Long: i = 0

'必要なボタンの数を決定する。
If Str1 <> "" Then i = i + 1
If Str2 <> "" Then i = i + 1
If Str3 <> "" Then i = i + 1

'表示するボタンとボタンの名称を決める
Select Case i
Case 1
  Me.C1.Caption = ""
  Me.C2.Caption = Str1
  Me.C3.Caption = ""
  Me.C1.Visible = False
  Me.C2.Visible = True
  Me.C3.Visible = False
Case 2
  Me.C1.Caption = ""
  Me.C2.Caption = Str1
  Me.C3.Caption = Str2
  Me.C1.Visible = False
  Me.C2.Visible = True
  Me.C3.Visible = True
Case 2
  Me.C1.Caption = Str1
  Me.C2.Caption = Str2
  Me.C3.Caption = Str3
  Me.C1.Visible = True
  Me.C2.Visible = True
  Me.C3.Visible = True
End Select

End Sub
