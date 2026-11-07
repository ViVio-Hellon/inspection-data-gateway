"""第3段(段7 後半・段9)の「影で走らせる」確かめ。

新しいやり方(判定を色と別に持つ・読み戻しを型に読み込む)を、いまのやり方と並べて走らせ、結果が食い違えば
ログに **[第3段7] / [第3段9]** の印を付けて出す。画面・保存はいまのやり方の結果を使うので、業務は変わらない。
ここまでの直し(段0〜6)のログにはこの印は付かないので、ログで区別できる。

- ログ: ロガー `ufkensa.stage3`(WARNING)と、処理の記録(trace)の出来事 `stage3.diff`
- 試験では `STRICT = True`(tests/_isolation)にして、食い違いがあれば失敗させる
- 止めたいときは 設定(user_settings)の `stage3_shadow` を "off" に
"""
from __future__ import annotations

from typing import Any, Iterable

from .. import trace
from ..logging_utils import get_logger
from . import model

log = get_logger("stage3")          # ロガー名 ufkensa.stage3

#: 試験用: 食い違いがあれば例外にする
STRICT = False
#: 1 回のログに出す食い違いの数
_LIMIT = 20
TAG = {"7": "[第3段7] 判定(色と別に持つ)", "9": "[第3段9] 読み戻し(型に読み込む)",
       "10": "[第3段10] 名前で読み書き(宣言の無い欄)"}


class ShadowMismatch(AssertionError):
    pass


def enabled() -> bool:
    try:
        from .. import user_settings
        return str(user_settings.get("stage3_shadow", "on") or "on") != "off"
    except Exception:                                  # noqa: BLE001 - 設定が読めなくても確かめは続ける
        return True


def report(stage: str, op: str, diffs: list[str]) -> None:
    """食い違いをログへ(印つき)。STRICT なら例外。"""
    if not diffs:
        return
    head = f"{TAG[stage]}: {op} で新旧が {len(diffs)} か所 食い違いました"
    body = "; ".join(diffs[:_LIMIT]) + (" …" if len(diffs) > _LIMIT else "")
    log.warning("%s: %s", head, body)
    try:
        trace.event("stage3.diff", head, level="WARNING", data={"stage": stage, "op": op, "diffs": diffs[:_LIMIT]})
    except Exception:                                  # noqa: BLE001
        pass
    if STRICT:
        raise ShadowMismatch(f"{head}: {body}")


# ----------------------------------------------------------------------
# 段7 後半: 判定(KensaState.verdicts)と画面の色
# ----------------------------------------------------------------------
def verdict_diffs(state: Any) -> list[str]:
    diffs = []
    for name in sorted(model.JUDGED_NAMES):
        if not state.has_ctl(name):
            continue
        c = state.ctl(name)
        old = model.verdict_of_color(c.color)
        new = state.verdicts.get(name, model.Verdict.NONE)
        if old is not new:
            diffs.append(f"{name}: 色={c.color or '(なし)'} 判定={new.value or '(なし)'}")
    return diffs


def check_verdicts(state: Any, op: str) -> None:
    if enabled():
        report("7", op, verdict_diffs(state))


# ----------------------------------------------------------------------
# 段9: 読み戻し(画面操作の再生)と、型への読み込み
# ----------------------------------------------------------------------
_SCALARS = ("lot", "atu", "hab", "tak", "ap", "am", "hp", "hm", "tp", "tm", "zai", "tyu", "youto", "strain_face")


def data_of(state: Any) -> dict[str, str]:
    """比べる業務の値(測定値・対角・直角度・ラテラル・BH・歪・不良枚数・外観・検入高さ・規格)。"""
    out: dict[str, str] = {}
    for name in model.JUDGED_NAMES:
        if state.has_ctl(name):
            out[name] = state.ctl(name).value
    for p, c in state.bad.items():
        out[model.bad_name(p)] = c.value
    for p in model.POSITIONS:
        for n in (model.visual_up_name(p), model.visual_under_name(p)):
            out[n] = state.v(n).value
    for i, c in state.ken_h.items():
        out[model.ken_h_name(i)] = c.value
    for k, c in state.strain_cb.items():
        out[k] = c.value
    for a in _SCALARS:
        out[a] = str(getattr(state, a, ""))
    return out


def restore_diffs(old: Any, new: Any, names: Iterable[str]) -> list[str]:
    a, b = data_of(old), data_of(new)
    return [f"{k}: 再生={a.get(k, '')!r} 型={b.get(k, '')!r}" for k in sorted(set(names)) if a.get(k, "") != b.get(k, "")]


def check_restore(state: Any, typed: Any, names: Iterable[str], op: str) -> None:
    if enabled():
        report("9", op, restore_diffs(state, typed, names))
