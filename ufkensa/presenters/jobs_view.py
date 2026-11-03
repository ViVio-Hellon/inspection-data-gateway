"""長時間処理(取り込みなど)の状態を画面に渡す形(db-viewer の presenters/settings から移した)"""
from __future__ import annotations

import re
from typing import Any, Optional

from .. import jobs


JOB_STATE_LABEL = {
    jobs.STATE_RUNNING: "実行中",
    jobs.STATE_DONE: "完了",
    jobs.STATE_FAILED: "失敗",
    jobs.STATE_INTERRUPTED: "中断",
}


JOB_STATE_KIND = {
    jobs.STATE_RUNNING: "run",
    jobs.STATE_DONE: "ok",
    jobs.STATE_FAILED: "ng",
    jobs.STATE_INTERRUPTED: "warn",
}


LANE_LIMIT = 6
_STEP_TAIL = re.compile(r"\s*を?[^\sを]{0,8}?(?:中|しています)[.。…]*\s*$")


def step_title(message: str) -> str:
    trimmed = _STEP_TAIL.sub("", message).strip()
    return trimmed or message.strip()


def step_dict(step: jobs.Step, *, no: int = 1, total: int = 1) -> dict[str, Any]:
    state = "run" if step.running else ("done" if step.ok else "ng")
    return {"no": no, "total": total, "label": step_title(step.label), "state": state,
            "state_label": {"run": "実行中", "done": "完了", "ng": "失敗"}[state],
            "elapsed_sec": round(step.elapsed_sec, 1)}


def lane_steps(steps: list[jobs.Step]) -> list[dict[str, Any]]:
    total = len(steps)
    if total < 2:
        return []
    order = [i for i, s in enumerate(steps) if not s.running and not s.ok]
    order += [i for i, s in enumerate(steps) if s.running]
    order += list(range(total - 1, -1, -1))
    keep: list[int] = []
    for index in order:
        if index in keep:
            continue
        keep.append(index)
        if len(keep) >= LANE_LIMIT:
            break
    keep.sort()
    return [step_dict(steps[i], no=i + 1, total=total) for i in keep]


def job_dict(job: Optional[jobs.Job]) -> Optional[dict[str, Any]]:
    if job is None:
        return None
    return {
        "id": job.id, "kind": job.kind, "label": job.label, "state": job.state,
        "state_label": JOB_STATE_LABEL.get(job.state, job.state),
        "state_kind": JOB_STATE_KIND.get(job.state, "run"),
        "pct": job.pct, "message": job.message, "ok": job.ok,
        "summary": job.summary, "error": job.error,
        "elapsed_sec": round(job.elapsed_sec, 1), "running": job.is_running,
        "steps": lane_steps(job.steps), "steps_total": len(job.steps),
    }


def jobs_dict(registry: Optional[jobs.JobRegistry] = None) -> dict[str, Any]:
    registry = registry or jobs.get_registry()
    running = registry.running()
    return {"running": [job_dict(running)] if running else [],
            "recent": [job_dict(job) for job in registry.recent()],
            "busy": running is not None}
