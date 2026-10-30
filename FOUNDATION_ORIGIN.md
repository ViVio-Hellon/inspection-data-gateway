# 基盤の複製元

起動制御・監視・停止・画面の外枠・設定画面の骨格・共通部品は、
`ViVio-Hellon/db-viewer` から複製しました。db-viewer の基盤は `ViVio-Hellon/python-web-tools`（梱包資材総合ツール）の複製で、
その上に VBA+Access 移植用の部品（cscript ブリッジ・adoSQL・VBA 互換・設定ファイル・アクセス権限）が足されています。

| 項目 | 値 |
|---|---|
| 直接の複製元 | https://github.com/ViVio-Hellon/db-viewer `fde74b2`（0.37.0） |
| 基盤の大元 | https://github.com/ViVio-Hellon/python-web-tools `25f661f`（VER4.8.123）時点 |
| 複製日 | 2026-10-30 |

## 置換したもの

`dbviewer` → `ufkensa`、`DBVIEWER_` → `UFKENSA_`、`DBビューア` → `UF検査判定`、`nlm.db-viewer` → `nlm.uf-kensa`、`DBViewer` → `UFKensa`。
`.vbs` / `.bat` は CP932 のまま置換した。ポートは 8790 → **8800**（8800〜8803。ほかのツールと重ならない）。

## 持ってこなかったもの

- デスクトップ版（`src-tauri`・`bridge.py`・`scripts/make_dist*`・`desktop_smoke`）… ランチャーからブラウザ版で起動する（判断事項 D2）
- db-viewer の文書（`docs/` は UF検査判定のものに差し替え。基盤の規約2本だけ `docs/foundation/` に置いた）

## 基盤の改修を取り込むとき

db-viewer の `fde74b2` 以降、python-web-tools の `25f661f` 以降の差分を、基盤のファイルに当てる。
置換（上の表）を忘れずに。
