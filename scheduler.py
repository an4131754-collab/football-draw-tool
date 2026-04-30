from __future__ import annotations

import argparse
from pathlib import Path

from tournament_tools import (
    BASE_DIR,
    generate_artifacts,
    generate_pdf_only,
    load_config,
    load_draw_data,
    normalize_download_options,
    update_draw_runtime_options,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="依照 draw_result.json 重新產生賽程 Excel、抽籤 JSON 與 PDF。"
    )
    parser.add_argument(
        "--draw",
        required=True,
        help="抽籤結果 JSON 路徑，例如 outputs/latest/draw_result.json",
    )
    parser.add_argument(
        "--output-dir",
        help="指定輸出資料夾；未指定時會輸出到 outputs/archive/<timestamp>",
    )
    parser.add_argument(
        "--pdf-only",
        action="store_true",
        help="只重新產生 PDF，保留舊版相容用法。",
    )
    parser.add_argument(
        "--day1-latest-end",
        help="重新排程時 DAY1 最晚結束時間，例如 18:45。",
    )
    parser.add_argument(
        "--day2-latest-end",
        help="重新排程時 DAY2 最晚結束時間，例如 17:45。",
    )
    parser.add_argument(
        "--outputs",
        help="要公開產生的下載項目，逗號分隔：json,pdf,excel。未指定時沿用 JSON 內設定。",
    )
    return parser


def resolve_path(input_path: str) -> Path:
    path = Path(input_path)
    if path.is_absolute():
        return path
    return (BASE_DIR / path).resolve()


def parse_outputs(value: str | None) -> dict[str, bool] | None:
    if value is None:
        return None

    selected = {item.strip().lower() for item in value.split(",") if item.strip()}
    unknown = selected - {"json", "pdf", "excel", "schedule"}
    if unknown:
        raise ValueError(f"不支援的 outputs：{', '.join(sorted(unknown))}")
    return normalize_download_options(
        {
            "json": "json" in selected,
            "pdf": "pdf" in selected,
            "schedule": "excel" in selected or "schedule" in selected,
        }
    )


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    draw_path = resolve_path(args.draw)
    output_dir = resolve_path(args.output_dir) if args.output_dir else None
    config = load_config(BASE_DIR)

    draw_data = load_draw_data(draw_path)
    output_options = parse_outputs(args.outputs)
    draw_data = update_draw_runtime_options(
        draw_data,
        config=config,
        download_options=output_options,
        day1_latest_end=args.day1_latest_end,
        day2_latest_end=args.day2_latest_end,
    )

    if args.pdf_only:
        pdf_path, latest_pdf_path, pdf_synced = generate_pdf_only(
            draw_data,
            output_dir=output_dir,
        )
        print(f"說明 PDF：{pdf_path}")
        print(f"latest PDF：{latest_pdf_path}")
        if not pdf_synced:
            print("提醒：outputs/latest 的 PDF 可能被開啟中，無法覆蓋。")
        return 0

    artifacts = generate_artifacts(draw_data, output_dir=output_dir)

    print(f"抽籤 JSON：{artifacts.draw_json_path}")
    if artifacts.schedule_path:
        print(f"賽程 Excel：{artifacts.schedule_path}")
    if artifacts.pdf_path:
        print(f"說明 PDF：{artifacts.pdf_path}")
    print(f"latest 資料夾：{artifacts.latest_dir}")
    if not artifacts.latest_sync_complete:
        print("提醒：outputs/latest 有檔案可能被開啟中，部分檔案無法覆蓋。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
