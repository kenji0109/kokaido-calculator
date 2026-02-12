from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# 祝日判定
try:
    import jpholiday
except Exception:
    jpholiday = None

APP_TITLE = "公会堂料金電卓 MVP（部屋＋設備＋技術者）"
DATA_DIR = Path(__file__).parent / "data"

PRICES_CSV = DATA_DIR / "prices.csv"
CLOSED_DAYS_CSV = DATA_DIR / "closed_days.csv"
EQUIPMENT_GROUPS_CSV = DATA_DIR / "equipment_groups.csv"
EQUIPMENT_MASTER_CSV = DATA_DIR / "equipment_master.csv"

TIME_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間", "全日", "延長30分"]
EQUIPMENT_TIME_SLOTS = ["利用なし"] + TIME_SLOTS  # ★設備だけ「利用なし」


def read_csv_safely(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"CSVが見つかりません: {path}")
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="cp932")


def normalize_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def parse_list_cell(cell: str) -> List[str]:
    """
    "A,B" / "A|B" / "A、B" などを許容
    空や '*' は ['*'] を返す（全対象の意味）
    """
    s = normalize_str(cell)
    if s == "" or s == "*" or s.lower() == "all":
        return ["*"]
    for sep in ["|", "、", "，", ";"]:
        s = s.replace(sep, ",")
    xs = [x.strip() for x in s.split(",") if x.strip()]
    return xs if xs else ["*"]


def is_weekend_or_holiday(date: pd.Timestamp) -> bool:
    if date.weekday() >= 5:
        return True
    if jpholiday is not None:
        try:
            return jpholiday.is_holiday(date.date())
        except Exception:
            pass
    return False


def build_date_range(start: pd.Timestamp, end: pd.Timestamp) -> List[pd.Timestamp]:
    if end < start:
        return []
    return list(pd.date_range(start=start, end=end, freq="D"))


def load_closed_days() -> set:
    if not CLOSED_DAYS_CSV.exists():
        return set()
    df = read_csv_safely(CLOSED_DAYS_CSV)
    if "date" not in df.columns:
        raise ValueError("closed_days.csv に 'date' 列がありません")
    return set(pd.to_datetime(df["date"]).dt.date.tolist())


# ===== 設備 =====
@dataclass
class EquipmentItem:
    item_id: str
    item_name: str
    group_id: str
    unit: str
    price_per_slot: int
    price_once_yen: int
    requires: List[str]
    notes: str
    is_countable: int
    is_power_item: int


def _to_int(x) -> int:
    if pd.isna(x) or str(x).strip() == "":
        return 0
    return int(float(x))


def load_equipment_data() -> Tuple[pd.DataFrame, Dict[str, EquipmentItem]]:
    groups_df = read_csv_safely(EQUIPMENT_GROUPS_CSV)
    master_df = read_csv_safely(EQUIPMENT_MASTER_CSV)

    required_groups_cols = {"group_id", "group_name", "applies_to_rooms"}
    required_master_cols = {"item_id", "group_id", "item_name", "unit", "price_per_slot"}

    if not required_groups_cols.issubset(set(groups_df.columns)):
        missing = required_groups_cols - set(groups_df.columns)
        raise ValueError(f"equipment_groups.csv に必要な列が足りません: {missing}")

    if not required_master_cols.issubset(set(master_df.columns)):
        missing = required_master_cols - set(master_df.columns)
        raise ValueError(f"equipment_master.csv に必要な列が足りません: {missing}")

    # optional columns
    if "price_once_yen" not in master_df.columns:
        master_df["price_once_yen"] = 0
    if "requires_item_ids" not in master_df.columns:
        master_df["requires_item_ids"] = ""
    if "notes" not in master_df.columns:
        master_df["notes"] = ""
    if "is_countable" not in master_df.columns:
        master_df["is_countable"] = 1
    if "is_power_item" not in master_df.columns:
        master_df["is_power_item"] = 0

    groups_df = groups_df.copy()
    for c in ["group_id", "group_name", "applies_to_rooms"]:
        groups_df[c] = groups_df[c].map(normalize_str)

    master_df = master_df.copy()
    for c in ["item_id", "group_id", "item_name", "unit", "requires_item_ids", "notes"]:
        master_df[c] = master_df[c].map(normalize_str)

    master_df["price_per_slot"] = master_df["price_per_slot"].map(_to_int)
    master_df["price_once_yen"] = master_df["price_once_yen"].map(_to_int)
    master_df["is_countable"] = master_df["is_countable"].map(_to_int)
    master_df["is_power_item"] = master_df["is_power_item"].map(_to_int)

    items: Dict[str, EquipmentItem] = {}
    for _, r in master_df.iterrows():
        req = parse_list_cell(r["requires_item_ids"])
        req = [x for x in req if x != "*"]
        items[r["item_id"]] = EquipmentItem(
            item_id=r["item_id"],
            item_name=r["item_name"],
            group_id=r["group_id"],
            unit=r["unit"],
            price_per_slot=int(r["price_per_slot"]),
            price_once_yen=int(r["price_once_yen"]),
            requires=req,
            notes=r["notes"],
            is_countable=int(r["is_countable"]),
            is_power_item=int(r["is_power_item"]),
        )

    return groups_df, items


# ===== 技術者 =====
STAGE_TECH_FEES_PER_PERSON = {
    "午前": 22000,
    "午後": 22000,
    "夜間": 22000,
    "午前-午後": 25300,
    "午後-夜間": 25300,
    "全日": 29700,
    "延長30分": 2750,
}


# ===== 部屋代 =====
def load_prices_df() -> pd.DataFrame:
    df = read_csv_safely(PRICES_CSV)
    required = {"room", "day_type", "price_type", "slot", "amount"}
    if not required.issubset(set(df.columns)):
        missing = required - set(df.columns)
        raise ValueError(f"prices.csv に必要な列が足りません: {missing}")

    df = df.copy()
    for c in ["room", "day_type", "price_type", "slot"]:
        df[c] = df[c].map(normalize_str)
    df["amount"] = df["amount"].map(_to_int)
    return df


def calc_room_total_for_day(
    prices_df: pd.DataFrame,
    rooms: List[str],
    day_type: str,
    price_type: str,
    slot: str,
) -> Tuple[int, pd.DataFrame]:
    if not rooms:
        return 0, pd.DataFrame(columns=["種別", "部屋", "土日祝", "料金種別", "区分", "単価", "小計"])

    rows = []
    total = 0
    for room in rooms:
        m = (
            (prices_df["room"] == room)
            & (prices_df["day_type"] == day_type)
            & (prices_df["price_type"] == price_type)
            & (prices_df["slot"] == slot)
        )
        hit = prices_df[m]
        if hit.empty:
            rows.append(
                {
                    "種別": "部屋",
                    "部屋": room,
                    "土日祝": day_type,
                    "料金種別": price_type,
                    "区分": slot,
                    "単価": None,
                    "小計": None,
                    "エラー": "該当料金が prices.csv に見つかりません",
                }
            )
            continue

        amount = int(hit.iloc[0]["amount"])
        total += amount
        rows.append(
            {
                "種別": "部屋",
                "部屋": room,
                "土日祝": day_type,
                "料金種別": price_type,
                "区分": slot,
                "単価": amount,
                "小計": amount,
                "エラー": "",
            }
        )

    return total, pd.DataFrame(rows)


# ===== 設備計算 =====
def slot_to_multiplier(slot: str) -> int:
    mapping = {
        "午前": 1,
        "午後": 1,
        "夜間": 1,
        "午前-午後": 2,
        "午後-夜間": 2,
        "全日": 3,
        "延長30分": 1,
    }
    return mapping.get(slot, 1)


def collect_required_items(selected_item_ids: List[str], items: Dict[str, EquipmentItem]) -> List[str]:
    selected_set = set(selected_item_ids)
    added = True
    while added:
        added = False
        for iid in list(selected_set):
            for req in items.get(iid, EquipmentItem(iid, "", "", "", 0, 0, [], "", 1, 0)).requires:
                if req and req not in selected_set:
                    selected_set.add(req)
                    added = True
    return list(selected_set)


def calc_equipment_total_for_day(
    day_slot_default: str,
    group_overrides: Dict[str, str],
    selections: List[Dict],
    items: Dict[str, EquipmentItem],
) -> Tuple[int, pd.DataFrame]:
    if not selections:
        return 0, pd.DataFrame(
            columns=["種別", "グループ", "品目", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"]
        )

    selected_ids = [s["item_id"] for s in selections]
    full_ids = collect_required_items(selected_ids, items)

    existing = {(s["group_id"], s["item_id"]) for s in selections}
    for iid in full_ids:
        if iid not in selected_ids:
            it = items[iid]
            key = (it.group_id, it.item_id)
            if key not in existing:
                selections.append({"group_id": it.group_id, "item_id": it.item_id, "qty": 1, "auto_added": True})

    rows = []
    total = 0
    for s in selections:
        iid = s["item_id"]
        qty = int(s.get("qty", 0))
        if qty <= 0:
            continue

        it = items[iid]
        slot = group_overrides.get(it.group_id, day_slot_default)
        mult = slot_to_multiplier(slot)

        per_slot_sub = it.price_per_slot * qty * mult
        once_sub = it.price_once_yen * qty
        subtotal = per_slot_sub + once_sub
        total += subtotal

        rows.append(
            {
                "種別": "設備",
                "グループ": it.group_id,
                "品目": it.item_name,
                "区分": slot,
                "数量": qty,
                "単価(1区分)": it.price_per_slot,
                "倍率": mult,
                "区分小計": per_slot_sub,
                "一回課金": once_sub,
                "小計": subtotal,
                "備考": it.notes,
                "自動追加": bool(s.get("auto_added", False)),
            }
        )

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["グループ", "品目"]).reset_index(drop=True)
    return total, df


def calc_stage_tech_total_for_day(slot: str, people: int) -> Tuple[int, pd.DataFrame]:
    if people <= 0:
        return 0, pd.DataFrame(columns=["種別", "区分", "人数", "単価(1名)", "小計"])
    unit = STAGE_TECH_FEES_PER_PERSON.get(slot)
    if unit is None:
        return 0, pd.DataFrame(columns=["種別", "区分", "人数", "単価(1名)", "小計"])
    subtotal = unit * people
    df = pd.DataFrame([{"種別": "技術者", "区分": slot, "人数": people, "単価(1名)": unit, "小計": subtotal}])
    return subtotal, df


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    # データロード
    try:
        groups_df, items = load_equipment_data()
    except Exception as e:
        st.error(f"設備CSVの読み込みに失敗しました: {e}")
        st.stop()

    try:
        closed_days = load_closed_days()
    except Exception as e:
        st.error(f"closed_days.csv の読み込みに失敗しました: {e}")
        st.stop()

    try:
        prices_df = load_prices_df()
    except Exception as e:
        st.error(f"prices.csv の読み込みに失敗しました: {e}")
        st.stop()

    left, right = st.columns([1, 1.3], gap="large")

    # ここで宣言（UIの後半で参照するため）
    equipment_any_day = True

    with left:
        st.subheader("1) 期間・部屋・区分（部屋料金）")
        col_a, col_b = st.columns(2)
        start_date = col_a.date_input("開始日", value=pd.Timestamp.today().date())
        end_date = col_b.date_input("終了日", value=pd.Timestamp.today().date())

        room_candidates = sorted(prices_df["room"].unique().tolist())
        rooms = st.multiselect("部屋（複数OK）", room_candidates, default=[])

        default_room_slot = st.selectbox("部屋の利用区分（デフォルト）", TIME_SLOTS, index=TIME_SLOTS.index("全日"))
        is_business_default = st.checkbox("割増利用（物販/入場料徴収など）※デフォルト", value=False)

        st.divider()
        st.subheader("日別の区分設定（パターンA）")
        st.caption("期間内の各日について、部屋区分/設備区分/技術者区分/割増を個別に指定できます。")

        start_ts_preview = pd.Timestamp(start_date)
        end_ts_preview = pd.Timestamp(end_date)
        days_preview = build_date_range(start_ts_preview, end_ts_preview)

        if not days_preview:
            st.error("日付範囲が不正です（終了日が開始日より前）")
            st.stop()

        # 日別設定テーブル（編集用）
        day_rows = []
        for d in days_preview:
            day_rows.append(
                {
                    "日付": d.date().isoformat(),
                    "土日祝": "土日祝" if is_weekend_or_holiday(d) else "平日",
                    "休館日": bool(d.date() in closed_days),
                    "部屋区分": default_room_slot,
                    "割増利用": bool(is_business_default),
                    # 基本は部屋区分追随。必要なら「利用なし」に変更できる
                    "設備デフォ区分": default_room_slot,
                    "技術者区分": default_room_slot,
                }
            )

        df_days_base = pd.DataFrame(day_rows)

        key_days = f"day_settings_{start_date}_{end_date}"
        if key_days not in st.session_state:
            st.session_state[key_days] = df_days_base

        try:
            edited_days = st.data_editor(
                st.session_state[key_days],
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "日付": st.column_config.TextColumn(disabled=True),
                    "土日祝": st.column_config.TextColumn(disabled=True),
                    "休館日": st.column_config.CheckboxColumn(disabled=True),
                    "部屋区分": st.column_config.SelectboxColumn(options=TIME_SLOTS),
                    "割増利用": st.column_config.CheckboxColumn(),
                    "設備デフォ区分": st.column_config.SelectboxColumn(options=EQUIPMENT_TIME_SLOTS),
                    "技術者区分": st.column_config.SelectboxColumn(options=TIME_SLOTS),
                },
            )

            # ★空白/None を「利用なし」に補正して、表示も統一
            edited_days = edited_days.copy()
            if "設備デフォ区分" in edited_days.columns:
                def _fix_equip_cell(v):
                    s = normalize_str(v)
                    if s == "" or s.lower() == "none":
                        return "利用なし"
                    return s
                edited_days["設備デフォ区分"] = edited_days["設備デフォ区分"].apply(_fix_equip_cell)

            st.session_state[key_days] = edited_days

        except Exception:
            st.warning("この環境では日別編集UIが使えないため、日別設定は表示のみになります。")
            edited_days = df_days_base
            st.dataframe(edited_days, use_container_width=True)

        # ===== 設備入力の有効/無効（ミス防止）=====
        try:
            _df_chk = edited_days.copy()
            _df_chk = _df_chk[_df_chk["休館日"] == False]
            _equip_vals = _df_chk["設備デフォ区分"].map(normalize_str).tolist()

            def _is_equip_used(v: str) -> bool:
                vv = (v or "").strip()
                if vv == "" or vv.lower() == "none":
                    return False
                return vv != "利用なし"

            equipment_any_day = any(_is_equip_used(v) for v in _equip_vals)
        except Exception:
            equipment_any_day = True  # 万一のときは触れるようにしておく（安全側）

        if not equipment_any_day:
            st.info("✅ 日別設定がすべて「設備デフォ区分：利用なし」なので、設備入力はロックします（誤入力防止）")

            # 以前入力した数量が残って誤解を招かないよう、表示上も0に戻す
            for k in list(st.session_state.keys()):
                if str(k).startswith("qty_"):
                    st.session_state[k] = 0

        st.divider()
        st.subheader("2) 設備・技術者（追加）")
        if not equipment_any_day:
            st.caption("※この期間は設備計算を行いません（全日：利用なし）")

        default_equipment_slot = st.selectbox(
            "設備のフォールバック区分（※日別設定が優先）",
            TIME_SLOTS,
            index=TIME_SLOTS.index(default_room_slot),
            disabled=(not equipment_any_day),
        )
        st.caption("※グループごとに区分を変更できます（未変更は各日の「設備デフォ区分」を使用）")

        def group_visible(row) -> bool:
            applicable = parse_list_cell(row["applies_to_rooms"])
            if "*" in applicable:
                return True
            if not rooms:
                return True
            return any(r in applicable for r in rooms)

        visible_groups = groups_df[groups_df.apply(group_visible, axis=1)].copy()

        group_overrides: Dict[str, str] = {}
        selections: List[Dict] = []

        for _, g in visible_groups.iterrows():
            gid = g["group_id"]
            gname = g["group_name"]

            with st.expander(f"設備グループ：{gname}", expanded=False):
                col1, col2 = st.columns([1, 1])
                override = col1.selectbox(
                    f"{gname} の区分（変更する場合だけ）",
                    ["(デフォルト)"] + TIME_SLOTS,
                    index=0,
                    key=f"override_{gid}",
                    disabled=(not equipment_any_day),
                )
                if override != "(デフォルト)":
                    group_overrides[gid] = override

                group_items = sorted([it for it in items.values() if it.group_id == gid], key=lambda x: x.item_name)

                if not group_items:
                    st.info("このグループには品目がありません（equipment_master.csvを確認）")
                else:
                    for it in group_items:
                        cA, cB, cC = st.columns([3, 1, 1])
                        line = f"**{it.item_name}**（単位:{it.unit} / 1区分:{it.price_per_slot:,}円"
                        if it.price_once_yen:
                            line += f" / 1回:{it.price_once_yen:,}円"
                        line += "）"
                        cA.write(line)

                        qty = cB.number_input(
                            "数量",
                            min_value=0,
                            value=0,
                            step=1,
                            key=f"qty_{gid}_{it.item_id}",
                            disabled=(not equipment_any_day),
                        )
                        if equipment_any_day and qty > 0:
                            selections.append({"group_id": gid, "item_id": it.item_id, "qty": int(qty)})
                        if it.notes:
                            cC.caption(it.notes)

        st.divider()
        st.subheader("3) 舞台設備技術者")
        tech_people = st.number_input("技術者人数（1名あたり課金）", min_value=0, value=0, step=1)
        tech_slot_fallback = st.selectbox(
            "技術者の区分（※日別設定が優先、ここはフォールバック）",
            TIME_SLOTS,
            index=TIME_SLOTS.index(default_room_slot),
        )

        st.divider()
        run = st.button("計算する", type="primary")

    with right:
        st.subheader("結果")
        if not run:
            st.info("左で条件を入れて「計算する」を押してね 😎")
            st.stop()

        df_days = edited_days.copy()

        excluded = df_days[df_days["休館日"] == True].copy()
        df_days_calc = df_days[df_days["休館日"] == False].copy()

        if df_days_calc.empty:
            st.error("計算対象の日がありません（全日休館日など）")
            st.stop()

        if not excluded.empty:
            st.warning(f"休館日として除外した日数: {len(excluded)}（例: {excluded.iloc[0]['日付']} …）")

        st.dataframe(df_days, use_container_width=True)

        room_grand = 0
        equipment_grand = 0
        tech_grand = 0
        breakdown_rows = []

        for _, day in df_days_calc.iterrows():
            d = pd.Timestamp(day["日付"])
            day_type = "土日祝" if is_weekend_or_holiday(d) else "平日"

            room_slot = normalize_str(day["部屋区分"]) or default_room_slot

            raw_equip = normalize_str(day.get("設備デフォ区分", ""))
            # ★空白/None は「利用なし」に統一（表示も統一済みだが念のため）
            if raw_equip == "" or raw_equip.lower() == "none":
                equipment_default_slot_day = "利用なし"
            else:
                equipment_default_slot_day = raw_equip

            tech_slot_day = normalize_str(day["技術者区分"]) or tech_slot_fallback or room_slot

            is_business_day = bool(day.get("割増利用", False))
            price_type = "割増" if is_business_day else "通常"

            # 部屋
            room_total, room_df = calc_room_total_for_day(
                prices_df=prices_df,
                rooms=rooms,
                day_type=day_type,
                price_type=price_type,
                slot=room_slot,
            )
            room_grand += room_total

            if not room_df.empty:
                for _, r in room_df.iterrows():
                    breakdown_rows.append(
                        {
                            "日付": d.date(),
                            "種別": "部屋",
                            "グループ": "",
                            "品目": r["部屋"],
                            "区分": r["区分"],
                            "数量/人数": 1,
                            "単価": r["単価"],
                            "倍率": 1,
                            "小計": r["小計"],
                            "自動追加": False,
                            "備考": r.get("エラー", ""),
                        }
                    )

            # 設備（★「利用なし」ならスキップ）
            if equipment_default_slot_day == "利用なし":
                eq_total, eq_df = 0, pd.DataFrame()
            else:
                eq_total, eq_df = calc_equipment_total_for_day(
                    day_slot_default=equipment_default_slot_day,
                    group_overrides=group_overrides,
                    selections=[s.copy() for s in selections],
                    items=items,
                )
            equipment_grand += eq_total

            if not eq_df.empty:
                for _, r in eq_df.iterrows():
                    breakdown_rows.append(
                        {
                            "日付": d.date(),
                            "種別": r["種別"],
                            "グループ": r["グループ"],
                            "品目": r["品目"],
                            "区分": r["区分"],
                            "数量/人数": r["数量"],
                            "単価": r["単価(1区分)"],
                            "倍率": r["倍率"],
                            "小計": r["小計"],
                            "自動追加": r["自動追加"],
                            "備考": r["備考"],
                        }
                    )

            # 技術者
            tech_total, tech_df = calc_stage_tech_total_for_day(tech_slot_day, int(tech_people))
            tech_grand += tech_total

            if not tech_df.empty:
                r = tech_df.iloc[0].to_dict()
                breakdown_rows.append(
                    {
                        "日付": d.date(),
                        "種別": "技術者",
                        "グループ": "",
                        "品目": "舞台設備技術者",
                        "区分": r["区分"],
                        "数量/人数": r["人数"],
                        "単価": r["単価(1名)"],
                        "倍率": 1,
                        "小計": r["小計"],
                        "自動追加": False,
                        "備考": "",
                    }
                )

        st.divider()
        st.subheader("合計")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("部屋代 合計", f"{room_grand:,} 円")
        c2.metric("設備 合計", f"{equipment_grand:,} 円")
        c3.metric("技術者 合計", f"{tech_grand:,} 円")
        c4.metric("総額（部屋＋設備＋技術者）", f"{room_grand + equipment_grand + tech_grand:,} 円")

        st.divider()
        st.subheader("内訳（日別）")
        df_break = pd.DataFrame(breakdown_rows)
        st.dataframe(df_break, use_container_width=True)


if __name__ == "__main__":
    main()
