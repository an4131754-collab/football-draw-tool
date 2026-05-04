from __future__ import annotations

import io
from pathlib import Path

import streamlit as st

from tournament_tools import (
    BASE_DIR,
    clear_latest_artifacts,
    create_draw_artifacts,
    get_artifact_filenames,
    get_latest_draw_data,
    load_config,
    load_teams,
    resolve_registration_path,
)


st.set_page_config(
    page_title="足球系際盃抽籤工具",
    page_icon="⚽",
    layout="wide",
)


def main() -> None:
    config = load_config(BASE_DIR)

    st.title("足球系際盃抽籤工具")
    st.caption("上傳同格式 Google 表單回覆 Excel，設定組數與晉級規則後抽籤。")

    with st.sidebar:
        st.header("抽籤設定")
        uploaded_file = st.file_uploader(
            "上傳報名表 Excel",
            type=["xlsx", "xlsm"],
            help="不選檔案時，會嘗試使用 repo 內的本機 fallback 報名表。",
        )
        group_count = st.number_input(
            "組數",
            min_value=2,
            max_value=int(config["max_group_count"]),
            value=int(config["default_group_count"]),
            step=1,
        )
        advance_per_group = st.number_input(
            "每組前 N 名晉級",
            min_value=1,
            value=int(config["default_advance_per_group"]),
            step=1,
        )
        wildcard_count = st.number_input(
            "最佳名次補位幾隊",
            min_value=0,
            value=int(config["default_wildcard_count"]),
            step=1,
        )
        knockout_label = st.selectbox(
            "淘汰賽階段",
            ["直接四強（4 隊晉級）", "八強賽（8 隊晉級）"],
            index=0 if str(config.get("default_knockout_format", "semifinal")) == "semifinal" else 1,
        )
        knockout_format = "quarterfinal" if "八強" in knockout_label else "semifinal"
        st.caption("直接四強需剛好 4 隊晉級；八強賽需剛好 8 隊晉級。4 組八強通常設定為每組前 2 名。")

        st.subheader("輸出")
        include_json = st.checkbox("JSON", value=True)
        include_schedule = st.checkbox("Excel", value=True)
        include_pdf = st.checkbox("PDF", value=True)

        draw_clicked = st.button("開始抽籤", type="primary", use_container_width=True)
        clear_clicked = st.button("清空目前結果", use_container_width=True)

    if clear_clicked:
        if clear_latest_artifacts(BASE_DIR):
            st.success("已清空目前結果。")
        else:
            st.warning("部分 latest 檔案可能正在被使用，無法完全清空。")

    show_fallback_teams()

    if draw_clicked:
        run_draw(
            uploaded_file=uploaded_file,
            group_count=int(group_count),
            advance_per_group=int(advance_per_group),
            wildcard_count=int(wildcard_count),
            knockout_format=knockout_format,
            download_options={
                "json": include_json,
                "schedule": include_schedule,
                "pdf": include_pdf,
            },
        )

    latest_draw = get_latest_draw_data(BASE_DIR)
    if latest_draw:
        show_latest_draw(latest_draw)
    else:
        st.info("尚未抽籤。上傳 Excel 並按下「開始抽籤」後，這裡會顯示結果。")


def show_fallback_teams() -> None:
    with st.expander("本機 fallback 名單", expanded=False):
        try:
            registration_path = resolve_registration_path(BASE_DIR)
            teams = load_teams(registration_path)
        except Exception as exc:
            st.write("尚未讀到本機報名表；線上使用時請直接上傳 Excel。")
            st.caption(str(exc))
            return

        st.write(f"目前可從 `{registration_path.name}` 讀到 {len(teams)} 隊。")
        st.write(teams)


def run_draw(
    *,
    uploaded_file,
    group_count: int,
    advance_per_group: int,
    wildcard_count: int,
    knockout_format: str,
    download_options: dict[str, bool],
) -> None:
    try:
        registration_source = None
        source_file = None
        if uploaded_file is not None:
            registration_source = io.BytesIO(uploaded_file.getvalue())
            source_file = uploaded_file.name

        draw_data, artifacts = create_draw_artifacts(
            BASE_DIR,
            registration_source=registration_source,
            source_file=source_file,
            group_count=group_count,
            advance_per_group=advance_per_group,
            wildcard_count=wildcard_count,
            knockout_format=knockout_format,
            download_options=download_options,
        )
    except Exception as exc:
        st.error(str(exc))
        return

    st.success("抽籤完成。")
    if draw_data.get("schedule_mode") == "template_schedule":
        st.caption("這次符合 12 隊 4 組，已套用 113 公開版模板產生完整賽程 Excel。")
    else:
        st.caption("這次不是 12 隊 4 組，已產生動態賽程/分組結果。")

    if not artifacts.latest_sync_complete:
        st.warning("latest 檔案同步不完整，可能有檔案被開啟中。")


def show_latest_draw(draw_data: dict) -> None:
    st.subheader("目前抽籤結果")

    col1, col2, col3 = st.columns(3)
    col1.metric("隊伍數", draw_data.get("team_count", len(draw_data.get("teams", []))))
    col2.metric("組數", draw_data.get("group_count", len(draw_data.get("groups", {}))))
    advancement = draw_data.get("advancement", {})
    col3.metric("晉級隊數", advancement.get("total_advancers", "待定"))

    st.write(f"抽籤時間：{draw_data.get('drawn_at', '')}")
    st.write(f"亂數函數：`{draw_data.get('random_function', 'secrets.SystemRandom().shuffle')}`")
    if advancement:
        st.write(f"晉級規則：{advancement.get('summary', '')}")
    st.write(f"淘汰賽階段：{advancement.get('knockout_stage', draw_data.get('knockout_format', 'semifinal'))}")

    groups = draw_data.get("groups", {})
    if groups:
        columns = st.columns(min(4, max(1, len(groups))))
        for index, (group_name, members) in enumerate(groups.items()):
            with columns[index % len(columns)]:
                st.markdown(f"### {group_name} 組")
                for seed, team in enumerate(members, start=1):
                    st.write(f"{seed}. {team}")

    show_downloads()


def show_downloads() -> None:
    artifact_filenames = get_artifact_filenames(BASE_DIR)
    latest_dir = BASE_DIR / "outputs" / "latest"

    st.subheader("下載")
    cols = st.columns(3)
    download_specs = [
        ("json", "下載 JSON", "application/json", cols[0]),
        ("schedule", "下載 Excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", cols[1]),
        ("pdf", "下載 PDF", "application/pdf", cols[2]),
    ]

    for kind, label, mime_type, column in download_specs:
        file_path = latest_dir / artifact_filenames[kind]
        with column:
            if file_path.exists():
                st.download_button(
                    label,
                    data=file_path.read_bytes(),
                    file_name=file_path.name,
                    mime=mime_type,
                    use_container_width=True,
                )
            else:
                st.button(label, disabled=True, use_container_width=True)


if __name__ == "__main__":
    main()
