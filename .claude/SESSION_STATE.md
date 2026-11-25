# SESSION_STATE

圧縮(compact)で会話が要約されても作業を続けられるように、今の状態を短く書いておく。
完了した項目は「完了した作業」に1行で残し、細部は消す(ためこまない)。

## 現在の作業目的
- VBA「UF検査判定」を Python の Web アプリ(Flask + waitress + 標準ライブラリ)へ移す。ライン PC は Python 3.12、Flask・waitress は入っている(同梱しない)。
- 本番ではまだ使わない(仕上がるまで)。ユーザーがライン PC で試し、気づいたことを直していく段階。

## 完了した作業
- PR #63〜#128 まで main に入れた(0.3.18。公差はユーザーが実機で確認済み)。0.3.19 作業中。
- 0.3.2: 最小化して放置→「バックエンドに接続できません」から戻らない件(Edge の同じ相手 6 本の通り道が詰まる)。api.js に時間の上限・abortAll、health.js で約1分で読み込み直し。
- 0.3.3: 測定:公差 80:20/外観は押すだけで複数選択(不良は1つ)/Min/Max が何もしない時に理由を出しタブへ移る/保守 測定機器ID の受信記録で画面が飛ばない/Aインプットを Min/Max 無しでは開けない(Test でも)・上 Auto(情報)/下 Manual(入力)・高さNG で測り直しを促す・「前回」列。
- 0.3.4: 保守「出力の確認」(Aインプット CSV・hostsend / 検査結果 PDF / 検査結果 CSV / 梱包検査ﾃﾞｰﾀ の最後の行。見るだけ。output_check.py)と「アイカスのライン」(伝送用ファイル 表「アイカスラインコード」。既定 = VBA の LinConv。icas_lines.py)。
- 0.3.5: 保守「日報」= all-tools の日報の紙(print_format)を写した nippou_print.py。いまの日報シートと書き出した 集計明細_*.csv を紙で出す(集計の段は出さない)。
- 0.3.6: 日報の開始 = ロット登録の時刻(workers.daily_lot_start。開始が入っている行=Aインプット送信後の次行は触らない)。検査員登録の Daily_Create は見出しだけ(start=False)。golden は Daily 3 行目だけ差(説明済み)。
- 0.3.7: キャスクに不良要因の登録(窓。不良要因印刷_<ライン> の 3 行目「ｷｬｽｸ」)・印刷で不良要因票も開く・ｻｰﾊﾞｰ保存で <Lot>_不良要因 の PDF を手元と工場共有へ(通常検査と同じ)。
- 0.3.8: デスクトップ版(Tauri。src-tauri/・bridge.py は python-web-tools から写して改名。UFKENSA_* の環境変数)。Rust は窓・起動・多重起動・Python 見張りだけ、業務は Python のまま。ブラウザ版は予備で残す。同時起動は launch_guard.hold_app_lock(両方向確認済み)。帳票 HTML に印刷/閉じる帯(report_html.PRINT_BAR)。window.open → desktop.openWindow。CI: .github/workflows/desktop-windows.yml(Windows exe)。docs/09_デスクトップ版.md。#118 で stop.bat がデスクトップ版の窓も止める(Windows CI 通過、成果物 UFKensa-windows)。実機確認待ち: 音・印刷・読み上げ・ドラッグ&ドロップ・受信機。
- 0.3.9: 画面が自動で移る(検査員登録→校正、校正→検査)前に鳴っている音を待つ(sound.waitSounds)。使用パス「Aインプットの CSV」(ainput_csv。空なら 検査履歴\<L>\<L>_Ainput。ainput.ainput_csv_root)。使用パスを 入力/出力(検査で使う)/出力(その他)/道具 と枠で分ける(paths.SECTIONS、設定画面は tbody ごと)。
- 0.3.10: VC の貼る面を大きな札で(view.vc.face = ab/a/b。VC の種類で決まり写真の有無に関係ない)。片面は貼らない面に「貼らない」の空き枠、右の列の枠の色も A=青・B=橙。golden は face の追加だけ。
- 0.3.11: Aインプット 山高さの実測を大きく・いま測った山に緑(data-just)/今回の送信結果を枠で、結果コードの一覧は畳んだ参考に/歪の面ボタンを A=青・B=橙で強く(歪タブも色)/Min/Max・Aインプット・印刷保存を「検査の流れ」の枠に出し操作タブから外す。
- 0.3.12: 使用パス「測定DB(梱包検査ﾃﾞｰﾀ)」(measure_db。空なら 工場共有\【測定DB】梱包\梱包検査ﾃﾞｰﾀ.sqlite3)、無いときは「作りますか」と聞く(measure_db_create)。Min/Max でロットの厚が空・0 のとき TypeMismatch で止まった→厚は 0.000 書式で出す。同じ返事に他の音があれば *_wait の音は鳴らさない(screen_op)。
- 0.3.13: 測定DB を「測定DB のフォルダ」(measure_db、dir、既定 = 工場共有\【測定DB】梱包)+「測定DB のファイル名」(measure_db_name、kind=name、既定 梱包検査ﾃﾞｰﾀ.sqlite3)に分けた(paths.measure_db_file。.sqlite3 で終わる旧設定はそのまま)。
- 0.3.14: 保守 出力の確認のタブ初期化エラー(init を最後に)・使用パスから下のフォルダを選んで見る。
- 0.3.15: 管理者は保存なしで新規検査。
- 0.3.16: lot.read_lot_row(第1パスは必ず第1の名前 first_lot_file。第1の値を保持し空だけ第2で補う。ロット登録・キャスク・全ヵ所で共通)。キャスク: ロット無し/ゼロ校正未は測定値を入れない・受信値はゼロ校正へ回す・大きな札。全ヵ所: ロット無しは入れない。focus_guard.js で文字を選べる。
- 0.3.17: 公差 0 の本当の原因 = 「控え(第2)から読む」(D28)・DBflag のとき第1を見ずに圧縮版だけ読んでいた(ライン PC は第1=Read\\【New】仕掛\\台帳、第2=梱包課共有\\仕掛 で別フォルダ。0.3.16 の「同じフォルダ」説は誤り)。read_lot_row はいつも第1が元。毎回ログ「ロット X の仕掛台帳: 元=…」。起動ログに版。キャスク・全ヵ所も公差 0 の理由を出す。
- 0.3.18(現場の決まり): SIKALOT の行 = オンラインは第1が正・空/無い所だけ第2で補う、オフライン・控え・DBflag は第2が正・空/無い所(公差など)だけ第1で補う。引当・受注は lot_ledger のまま(オンライン第1+第2で足りない引当を足す、オフライン第2だけ)。0.3.15〜17 の「いつも第1が元」はやめた(試験の期待も 0.3.14 に戻した)。実ファイル(SIKALOT.xlsx 由来の第1+受領の圧縮版)で 30 ロット全モード公差 OK。公差 0 は「第1にロットが無い」ときだけ再現(圧縮版にだけあるロット)→ 知らせに理由。
- 0.3.19: キャスク・全ヵ所の表を「Excel の良いところだけ」(sheet_input.js): 入力欄・Back space/この欄を消す ボタンを無くし、選んだセルに直接打つ(打っている文字をセルに重ねる)・Enter 確定・Esc・矢印で移動・Delete で消す(キャスク cask.clear_cell、判定も)・Back space はキャスク=直前の取り消し/全ヵ所=消す。日本語入力の確定 Enter で送らない(外観に文字が入らなかった原因)。キャスク外観(AB)は巾丈測定 OFF でも錠なし、錠の欄は分かる言葉で。読み上げ: ON で「読み上げ オン」・声の名前・読み上げ中を表示。ゼロ校正の許容値は保守 → 寸法のしきい値「キャスクのゼロ校正の許容値」(size_thresholds zero_micro/zero_noguisu/zero_convex。既定 0.01/0.01/0.5。VBA は 0/0/0.5)。固まる件: × の確かめを 10 分待ち→5 秒で「応答なし。終了しますか」(Rust call_within)、bridge は終了・生存確認を別の通り道、30 秒超の要求は全スレッドの呼び出しをログ。
- Claude Code の圧縮対策(この SESSION_STATE.md と .claude/hooks/session_state.py、.claude/settings.json の PreCompact / SessionStart hook)。

## 未完了の作業
- 固まる件の本当の原因はまだ(次に固まったらログの「要求が N 秒たっても終わりません」とスレッドの呼び出しで特定する)。
- ユーザーの返事待ち: Min/Max の開始・終了を選んだ時点で即計算するか(今は VBA どおりボタンを押した時だけ)。
- ライン PC は 0.3.0 のまま動いていた → 0.3.3 へ入れ替えて確かめてもらう(QF空冷時間チェックリストの使用パス欄もそれで出る)。
- 放置で切れる件は、実物の Edge で 1 時間以上の確かめが未(こちらでは応答停止→復帰の再現で確認済み)。

## 重要な仕様・設計上の決定
- 作業の流れ: ブランチ `claude/hopeful-newton-lc006l` で作業 → PR → squash で main へ(聞かずに入れてよい)→ `git fetch -q origin main && git merge -q -s ours origin/main -m "sync main" && git push -q`。
- 版は config/app.json の version。ライン PC で入れ替えを見分けるため、直すたびに上げる。
- 業務の動き・保存/印刷のセル位置は、意図して決めない限り変えない。VBA の文言(MsgBox の原文)は変えない(試験が原文を見ている)。
- golden(tests/golden)は差を説明せずに作り直さない(UFKENSA_GOLDEN_UPDATE=1)。
- 会社の DB・xlsx(アップロードされたもの)は確かめるためだけに使い、コミットしない。個人情報をコミットしない。コミットにモデル名を書かない。
- 返事は日本語・やさしい言葉。全体像はアーティファクトで(ダウンロード方式は使わない)。
- 試験: `timeout 900 python -m pytest -q tests`(約 1443 通過・7〜8 分。背景で回す)、画面の通し `UFKENSA_E2E=1 E2E_SHOTS=<dir> python -m tests.e2e`。試験の時計は env.now() 1 回で 7 秒進む(呼ぶ回数を増やさない)。

## 発生した不具合と解決状況
- 放置後に接続が戻らない → 0.3.2 で対処(実機での長時間確認は未)。
- Aインプットが Min/Max 無しで開けた(Test ラインの素通し)→ 0.3.3 で止めた。

## 変更したファイル
- 直近(0.3.3): app/static/css/kensa.css, app/static/js/views/kensa.js, app/templates/kensa.html, ufkensa/kensa/minmax.py, app/static/js/views/maint.js, app/static/js/views/master_grid.js, ufkensa/kensa/ainput_session.py, app/templates/ainput.html, app/static/js/views/ainput.js, app/static/css/ainput.css, tests/e2e/screens.spec.js, config/app.json
- 0.3.2: app/static/js/api.js, app/static/js/health.js

## 次に実行する作業
- ユーザーの次の指摘を待つ。来たら直す → 関係する試験 → 全件試験(背景)→ 版を上げる → PR・squash → 同期 → 日本語で報告。

<!-- AUTO:GIT BEGIN(圧縮の直前に hook が書き直す。手で直さない) -->
### Git の状況(自動 2026-11-23 00:44:02 / 圧縮 auto)
- ブランチ: `claude/hopeful-newton-lc006l`
- 最後のコミット: 9f5da72 sync main
- まだコミットしていない変更: なし
<!-- AUTO:GIT END -->
