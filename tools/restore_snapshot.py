#!/usr/bin/env python3
"""GitHub 백업(snapshots/*.json)을 시트별 CSV로 풀어 줍니다.

구글 시트 원본이 망가졌을 때, 만들어진 CSV를 해당 시트에
'파일 → 가져오기 → 현재 시트 바꾸기'로 올리면 그 날짜 상태로 되돌아갑니다.

사용 예:
  python3 tools/restore_snapshot.py                       # 가장 최근 백업
  python3 tools/restore_snapshot.py 2026-10-15            # 특정 날짜
  python3 tools/restore_snapshot.py 2026-10-15 --only students
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SNAP_DIR = ROOT / "snapshots"


def pick_snapshot(date: str | None) -> Path:
    if date:
        path = SNAP_DIR / f"snapshot_{date}.json"
        if not path.exists():
            sys.exit(f"백업 파일이 없습니다: {path.name}")
        return path
    files = sorted(SNAP_DIR.glob("snapshot_*.json"))
    if not files:
        sys.exit("snapshots 폴더에 백업이 없습니다.")
    return files[-1]


def write_table(rows: list[dict], out_path: Path) -> None:
    # 행마다 열이 조금씩 다를 수 있어 모든 열을 처음 나온 순서대로 모읍니다.
    headers: list[str] = []
    for row in rows:
        for key in row:
            if key not in headers:
                headers.append(key)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="백업 JSON을 시트별 CSV로 변환")
    parser.add_argument("date", nargs="?", help="YYYY-MM-DD (생략 시 최신)")
    parser.add_argument("--only", help="이 시트만 변환 (예: students)")
    parser.add_argument("--out", help="저장 폴더 (기본: restore_<날짜>)")
    args = parser.parse_args()

    snap = pick_snapshot(args.date)
    data = json.loads(snap.read_text(encoding="utf-8"))
    day = snap.stem.removeprefix("snapshot_")
    out_dir = Path(args.out) if args.out else ROOT / f"restore_{day}"
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"백업: {snap.name} (저장 시각 {data.get('snapshot_at', '?')})")
    for name, rows in data.items():
        if not isinstance(rows, list) or (args.only and name != args.only):
            continue
        out_path = out_dir / f"{name}.csv"
        write_table(rows, out_path)
        print(f"  {name}: {len(rows)}행 → {out_path.relative_to(ROOT) if out_path.is_relative_to(ROOT) else out_path}")


if __name__ == "__main__":
    main()
