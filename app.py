from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

# 祝日判定
try:
    import jpholiday
except Exception:
    jpholiday = None

APP_TITLE = "公会堂料金電卓 MVP（部屋＋設備＋技術者＋インターネット）"
DATA_DIR = Path(__file__).parent / "data"

PRICES_CSV = DATA_DIR / "prices.csv"
CLOSED_DAYS_CSV = DATA_DIR / "closed_days.csv"
EQUIPMENT_GROUPS_CSV = DATA_DIR / "equipment_groups.csv"
EQUIPMENT_MASTER_CSV = DATA_DIR / "equipment_master.csv"

TIME_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間", "全日", "延長30分"]
EQUIPMENT_TIME_SLOTS = ["利用なし"] + TIME_SLOTS  # ★設備だけ「利用なし」


# =========================
# CSV Utility
# =========================
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


def _to_int(x) -> int:
    if pd.isna(x) or str(x).strip() == "":
        return 0
    return int(float(x))


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


# =========================
# Date / Holiday
# =========================
def holiday_name(date: pd.Timestamp) -> str:
    if jpholiday is None:
        return ""
    try:
        nm = jpholiday.is_holiday_name(date.date())
        return nm or ""
    except Exception:
        return ""


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


# =========================
# Equipment
# =========================
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


@dataclass
class GroupMeta:
    group_id: str
    group_name: str
    applies_to_rooms: str
    default_inherit_room_slot: int  # 1:日別の設備デフォ区分を継承 / 0:継承しない
    allowed_slot_override: int      # 1:グループごとの区分overrideを許可 / 0:許可しない


def load_equipment_data() -> Tuple[pd.DataFrame, Dict[str, EquipmentItem], Dict[str, GroupMeta]]:
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

    # groups optional columns
    if "default_inherit_room_slot" not in groups_df.columns:
        groups_df["default_inherit_room_slot"] = 1
    if "allowed_slot_override" not in groups_df.columns:
        groups_df["allowed_slot_override"] = 1

    # master optional columns
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
    groups_df["default_inherit_room_slot"] = groups_df["default_inherit_room_slot"].map(_to_int)
    groups_df["allowed_slot_override"] = groups_df["allowed_slot_override"].map(_to_int)

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

    group_meta: Dict[str, GroupMeta] = {}
    for _, g in groups_df.iterrows():
        gid = g["group_id"]
        group_meta[gid] = GroupMeta(
            group_id=gid,
            group_name=g["group_name"],
            applies_to_rooms=g["applies_to_rooms"],
            default_inherit_room_slot=int(g["default_inherit_room_slot"]),
            allowed_slot_override=int(g["allowed_slot_override"]),
        )

    return groups_df, items, group_meta


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
    global_fallback_slot: str,
    group_overrides: Dict[str, str],
    selections: List[Dict],
    items: Dict[str, EquipmentItem],
    group_meta: Dict[str, GroupMeta],
) -> Tuple[int, pd.DataFrame]:
    """
    事故防止の肝：
    - price_per_slot > 0 のときだけ「区分課金」扱い
    - price_per_slot == 0 かつ price_once_yen > 0 は「区分なし単価」扱い（倍率も区分も無視）
    """
    if not selections:
        return 0, pd.DataFrame(
            columns=["種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"]
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
        meta = group_meta.get(it.group_id, GroupMeta(it.group_id, it.group_id, "*", 1, 1))

        # --- 区分決定（区分課金アイテムのときのみ意味がある）---
        inherit = bool(meta.default_inherit_room_slot)
        base_slot = day_slot_default if inherit else global_fallback_slot

        # グループoverrideは許可されている場合だけ
        if meta.allowed_slot_override and it.group_id in group_overrides:
            slot = group_overrides[it.group_id]
        else:
            slot = base_slot

        # --- 課金分岐 ---
        is_slot_item = it.price_per_slot > 0
        is_once_item = (it.price_per_slot == 0) and (it.price_once_yen > 0)

        if is_slot_item:
            mult = slot_to_multiplier(slot)
            per_slot_sub = it.price_per_slot * qty * mult
            once_sub = it.price_once_yen * qty
            subtotal = per_slot_sub + once_sub
            charge_type = "区分課金"
        elif is_once_item:
            # 区分なし単価：区分・倍率は無視（事故防止）
            mult = 0
            per_slot_sub = 0
            once_sub = it.price_once_yen * qty
            subtotal = once_sub
            charge_type = "区分なし単価"
            slot = "（区分なし）"
        else:
            # 価格が両方0：データ不備
            mult = 0
            per_slot_sub = 0
            once_sub = 0
            subtotal = 0
            charge_type = "料金未設定"
            slot = "—"

        total += subtotal

        rows.append(
            {
                "種別": "設備",
                "グループ": it.group_id,
                "品目": it.item_name,
                "課金タイプ": charge_type,
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


# =========================
# Stage tech
# =========================
STAGE_TECH_FEES_PER_PERSON = {
    "午前": 22000,
    "午後": 22000,
    "夜間": 22000,
    "午前-午後": 25300,
    "午後-夜間": 25300,
    "全日": 29700,
    "延長30分": 2750,
}


def calc_stage_tech_total_for_day(slot: str, people: int) -> Tuple[int, pd.DataFrame]:
    if people <= 0:
        return 0, pd.DataFrame(columns=["種別", "区分", "人数", "単価(1名)", "小計"])
    unit = STAGE_TECH_FEES_PER_PERSON.get(slot)
    if unit is None:
        return 0, pd.DataFrame(columns=["種別", "区分", "人数", "単価(1名)", "小計"])
    subtotal = unit * people
    df = pd.DataFrame([{"種別": "技術者", "区分": slot, "人数": people, "単価(1名)": unit, "小計": subtotal}])
    return subtotal, df


# =========================
# Room prices
# =========================
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
        return 0, pd.DataFrame(columns=["種別", "部屋", "土日祝", "料金種別", "区分", "単価", "小計", "エラー"])

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


# =========================
# Internet (future-ready)
# =========================
INTERNET_POCKET_WIFI_PER_DAY = 2800
INTERNET_FIXED_FIRST_DAY = 18000
INTERNET_FIXED_AFTER_DAY = 2000
INTERNET_TEMP_LINE_BASE = 5000  # + 別途見積

FLOOR_1_ROOMS = {"大集会室"}          # 1F扱い（あなたの運用に合わせて増やしてOK）
FLOOR_3_ROOMS = {"中集会室", "小集会室"}  # 3F扱い（中・小は同フロア課金）


def infer_internet_floors(selected_rooms: List[str]) -> Dict[str, bool]:
    s = set(selected_rooms)
    return {
        "1F": bool(s & FLOOR_1_ROOMS),
        "3F": bool(s & FLOOR_3_ROOMS),
    }


def calc_internet_total(
    df_days_calc: pd.DataFrame,
    selected_rooms: List[str],
    use_pocket_wifi: bool,
    use_fixed_line: bool,
    use_temp_line: bool,
) -> Tuple[int, pd.DataFrame]:
    rows = []
    total = 0

    floors = infer_internet_floors(selected_rooms)
    active_days = [pd.Timestamp(x).date() for x in df_days_calc["日付"].tolist()]
    active_days_sorted = sorted(active_days)

    # ① ポケットWi-Fi（利用日は「部屋を使う日」に限る＝計算対象日）
    if use_pocket_wifi:
        for d in active_days_sorted:
            rows.append(
                {"日付": d, "種別": "インターネット", "品目": "ポケットWi-Fi貸出", "フロア": "全部屋", "小計": INTERNET_POCKET_WIFI_PER_DAY, "備考": "先着順/同時接続目安5台/電波不安定の可能性"}
            )
            total += INTERNET_POCKET_WIFI_PER_DAY

    # ② 常設回線（フロア毎／連続利用で段階料金）
    # 連続判定は「計算対象日」を基準にする（休館日で途切れる）
    if use_fixed_line:
        for floor_label, enabled in floors.items():
            if not enabled:
                continue

            # 連続ブロックに分割
            blocks = []
            block = []
            for d in active_days_sorted:
                if not block:
                    block = [d]
                else:
                    prev = block[-1]
                    if (pd.Timestamp(d) - pd.Timestamp(prev)).days == 1:
                        block.append(d)
                    else:
                        blocks.append(block)
                        block = [d]
            if block:
                blocks.append(block)

            for b in blocks:
                if not b:
                    continue
                # 初日
                rows.append(
                    {"日付": b[0], "種別": "インターネット", "品目": "常設回線（初日）", "フロア": floor_label, "小計": INTERNET_FIXED_FIRST_DAY, "備考": "連続利用の段階料金"}
                )
                total += INTERNET_FIXED_FIRST_DAY
                # 2日目以降
                for d in b[1:]:
                    rows.append(
                        {"日付": d, "種別": "インターネット", "品目": "常設回線（2日目以降）", "フロア": floor_label, "小計": INTERNET_FIXED_AFTER_DAY, "備考": "連続利用の段階料金"}
                    )
                    total += INTERNET_FIXED_AFTER_DAY

    # ③ 仮設回線（フロア毎に 5,000円/回 + 別途見積）
    if use_temp_line:
        for floor_label, enabled in floors.items():
            if not enabled:
                continue
            rows.append(
                {"日付": active_days_sorted[0] if active_days_sorted else None, "種別": "インターネット", "品目": "仮設回線（開通工事）", "フロア": floor_label, "小計": INTERNET_TEMP_LINE_BASE, "備考": "＋別途お見積り（NTT回線開通工事）"}
            )
            total += INTERNET_TEMP_LINE_BASE

    df = pd.DataFrame(rows, columns=["日付", "種別", "品目", "フロア", "小計", "備考"])
    return total, df


# =========================
# Day settings state sync
# =========================
def make_days_base(
    days: List[pd.Timestamp],
    closed_days: set,
    default_room_slot: str,
    is_business_default: bool,
) -> pd.DataFrame:
    rows = []
    for d in days:
        w = "土日祝" if is_weekend_or_holiday(d) else "平日"
        hn = holiday_name(d)
        rows.append(
            {
                "日付": d.date().isoformat(),
                "土日祝": w,
                "祝日名": hn,
                "休館日": bool(d.date() in closed_days),
                "部屋区分": default_room_slot,
                "割増利用": bool(is_business_default),
                "設備デフォ区分": default_room_slot,
                "技術者区分": default_room_slot,
            }
        )
    return pd.DataFrame(rows)


def _fix_equip_cell(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "利用なし"
    if s not in EQUIPMENT_TIME_SLOTS:
        return "利用なし"
    return s


def sync_defaults_into_day_df(
    df: pd.DataFrame,
    old_defaults: Dict[str, object],
    new_defaults: Dict[str, object],
) -> pd.DataFrame:
    """
    デフォルト変更を日別へ自動反映。
    ただし「手で変えたセル」を潰しにくくするため、
    「旧デフォルトと同じ値のセルだけ」を新デフォルトへ置換する。
    """
    df = df.copy()

    # 休館日・土日祝・祝日名は毎回再計算（表示ズレ防止）
    for i in range(len(df)):
        d = pd.Timestamp(df.loc[i, "日付"])
        df.loc[i, "土日祝"] = "土日祝" if is_weekend_or_holiday(d) else "平日"
        df.loc[i, "祝日名"] = holiday_name(d)

    # 主要列の差し替え
    cols = ["部屋区分", "割増利用", "設備デフォ区分", "技術者区分"]
    for c in cols:
        if c not in df.columns:
            continue
        oldv = old_defaults.get(c, None)
        newv = new_defaults.get(c, None)
        if oldv == newv:
            continue

        # "旧デフォルトと一致しているセルだけ" 新デフォルトに置換
        mask = df[c].astype(str) == str(oldv)
        df.loc[mask, c] = newv

    # 設備デフォ区分の正規化（None/空白対策）
    if "設備デフォ区分" in df.columns:
        df["設備デフォ区分"] = df["設備デフォ区分"].apply(_fix_equip_cell)

    # 技術者区分が空なら部屋区分へ（事故防止）
    if "技術者区分" in df.columns and "部屋区分" in df.columns:
        def _fix_tech(v, roomv):
            sv = normalize_str(v)
            if sv == "" or sv.lower() == "none":
                return normalize_str(roomv) or "全日"
            return sv
        df["技術者区分"] = [
            _fix_tech(df.loc[i, "技術者区分"], df.loc[i, "部屋区分"]) for i in range(len(df))
        ]

    return df


# =========================
# App
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    # データロード
    try:
        groups_df, items, group_meta = load_equipment_data()
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

    left, right = st.columns([1, 1.35], gap="large")

    # ここで宣言（UIの後半で参照するため）
    equipment_any_day = True

    with left:
        st.subheader("1) 期間・部屋・区分（部屋料金）")

        col_a, col_b = st.columns(2)
        start_date = col_a.date_input("開始日", value=pd.Timestamp.today().date())
        end_date = col_b.date_input("終了日", value=pd.Timestamp.today().date())

        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)
        days = build_date_range(start_ts, end_ts)
        if not days:
            st.error("日付範囲が不正です（終了日が開始日より前）")
            st.stop()

        room_candidates = sorted(prices_df["room"].unique().tolist())
        rooms = st.multiselect("部屋（複数OK）", room_candidates, default=[])

        default_room_slot = st.selectbox("部屋の利用区分（デフォルト）", TIME_SLOTS, index=TIME_SLOTS.index("全日"))
        is_business_default = st.checkbox("割増利用（物販/入場料徴収など）※デフォルト", value=False)

        st.divider()
        st.subheader("日別の区分設定（パターンA）")
        st.caption("デフォルト変更は日別にも自動反映（ただし手で変えた行はなるべく保持します）。")

        # --- 日別設定をsession_stateに保存（デフォルト変更時に自動同期）---
        key_days = f"day_settings_{start_date}_{end_date}"

        new_defaults = {
            "部屋区分": default_room_slot,
            "割増利用": bool(is_business_default),
            "設備デフォ区分": default_room_slot,
            "技術者区分": default_room_slot,
        }

        if key_days not in st.session_state:
            df_base = make_days_base(days, closed_days, default_room_slot, is_business_default)
            st.session_state[key_days] = df_base
            st.session_state[key_days + "_defaults"] = dict(new_defaults)
        else:
            df_existing: pd.DataFrame = st.session_state[key_days]
            # 期間が変わって行数/日付がズレたら作り直し
            if len(df_existing) != len(days) or df_existing.iloc[0]["日付"] != days[0].date().isoformat() or df_existing.iloc[-1]["日付"] != days[-1].date().isoformat():
                df_base = make_days_base(days, closed_days, default_room_slot, is_business_default)
                st.session_state[key_days] = df_base
                st.session_state[key_days + "_defaults"] = dict(new_defaults)
            else:
                old_defaults = st.session_state.get(key_days + "_defaults", dict(new_defaults))
                st.session_state[key_days] = sync_defaults_into_day_df(df_existing, old_defaults, new_defaults)
                st.session_state[key_days + "_defaults"] = dict(new_defaults)

        # 編集
        try:
            edited_days = st.data_editor(
                st.session_state[key_days],
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "日付": st.column_config.TextColumn(disabled=True),
                    "土日祝": st.column_config.TextColumn(disabled=True),
                    "祝日名": st.column_config.TextColumn(disabled=True),
                    "休館日": st.column_config.CheckboxColumn(disabled=True),
                    "部屋区分": st.column_config.SelectboxColumn(options=TIME_SLOTS),
                    "割増利用": st.column_config.CheckboxColumn(),
                    "設備デフォ区分": st.column_config.SelectboxColumn(options=EQUIPMENT_TIME_SLOTS),
                    "技術者区分": st.column_config.SelectboxColumn(options=TIME_SLOTS),
                },
            )
            edited_days = edited_days.copy()
            edited_days["設備デフォ区分"] = edited_days["設備デフォ区分"].apply(_fix_equip_cell)
            st.session_state[key_days] = edited_days
        except Exception:
            st.warning("この環境では日別編集UIが使えないため、日別設定は表示のみになります。")
            edited_days = st.session_state[key_days]
            st.dataframe(edited_days, use_container_width=True)

        # 設備入力の有効/無効（誤入力防止）
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
            equipment_any_day = True

        if not equipment_any_day:
            st.info("✅ 日別設定がすべて「設備デフォ区分：利用なし」なので、設備入力はロックします（誤入力防止）")
            for k in list(st.session_state.keys()):
                if str(k).startswith("qty_"):
                    st.session_state[k] = 0

        st.divider()
        st.subheader("2) 設備（追加）")
        if not equipment_any_day:
            st.caption("※この期間は設備計算を行いません（全日：利用なし）")

        # グローバルフォールバック（継承しないグループ用）
        global_fallback_slot = st.selectbox(
            "設備区分のフォールバック（※日別の設備デフォ区分を継承しないグループ用）",
            TIME_SLOTS,
            index=TIME_SLOTS.index(default_room_slot),
            disabled=(not equipment_any_day),
        )
        st.caption("※大道具/設営など（groupで継承しない指定）はこのフォールバックが基準になります")

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
            meta = group_meta.get(gid, GroupMeta(gid, gname, "*", 1, 1))

            with st.expander(f"設備グループ：{gname}", expanded=False):
                col1, col2 = st.columns([1, 1])

                # allowed_slot_override=0 なら UI を出さない（事故防止）
                if meta.allowed_slot_override and equipment_any_day:
                    override = col1.selectbox(
                        f"{gname} の区分（変更する場合だけ）",
                        ["(デフォルト)"] + TIME_SLOTS,
                        index=0,
                        key=f"override_{gid}",
                        disabled=(not equipment_any_day),
                    )
                    if override != "(デフォルト)":
                        group_overrides[gid] = override
                else:
                    col1.caption("🔒 このグループは区分override不可（事故防止）")

                inherit_text = "日別の設備デフォ区分を継承" if meta.default_inherit_room_slot else "継承しない（フォールバック基準）"
                col2.caption(f"設定: {inherit_text}")

                group_items = sorted([it for it in items.values() if it.group_id == gid], key=lambda x: x.item_name)
                if not group_items:
                    st.info("このグループには品目がありません（equipment_master.csvを確認）")
                else:
                    for it in group_items:
                        cA, cB, cC = st.columns([3, 1, 1])

                        # 表示：区分課金 or 区分なし単価
                        if it.price_per_slot > 0:
                            line = f"**{it.item_name}**（単位:{it.unit} / 1区分:{it.price_per_slot:,}円"
                            if it.price_once_yen:
                                line += f" / 1回:{it.price_once_yen:,}円"
                            line += "）"
                        elif it.price_once_yen > 0:
                            line = f"**{it.item_name}**（単位:{it.unit} / 1回:{it.price_once_yen:,}円 / 区分なし）"
                        else:
                            line = f"**{it.item_name}**（単位:{it.unit} / 料金未設定）"

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
            "技術者区分のフォールバック（※日別設定が空のときだけ使用）",
            TIME_SLOTS,
            index=TIME_SLOTS.index(default_room_slot),
        )
        st.caption("※日別の「技術者区分」が入っていれば、それを最優先します（ズレ防止）")

        st.divider()
        st.subheader("4) インターネット（将来込みを想定・暫定実装）")
        use_pocket_wifi = st.checkbox("① ポケットWi-Fi貸出（2,800円/日）", value=False)
        use_fixed_line = st.checkbox("② 常設回線（初日18,000円、連続2日目以降2,000円/日、フロア毎）", value=False)
        use_temp_line = st.checkbox("③ 仮設回線（5,000円/回＋別途見積、フロア毎）", value=False)

        st.divider()
        run = st.button("計算する", type="primary")

    with right:
        st.subheader("結果")
        if not run:
            st.info("左で条件を入れて「計算する」を押してね 😎")
            st.stop()

        df_days = st.session_state[key_days].copy()

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
        internet_grand = 0

        breakdown_rows = []

        for _, day in df_days_calc.iterrows():
            d = pd.Timestamp(day["日付"])

            # ★祝日も含めて day_type を決める（課題2の解消）
            day_type = "土日祝" if is_weekend_or_holiday(d) else "平日"

            room_slot = normalize_str(day["部屋区分"]) or default_room_slot
            is_business_day = bool(day.get("割増利用", False))
            price_type = "割増" if is_business_day else "通常"

            # 設備：日別の設備デフォ区分
            raw_equip = normalize_str(day.get("設備デフォ区分", ""))
            equipment_default_slot_day = _fix_equip_cell(raw_equip)

            # 技術者：日別優先 → 空ならフォールバック → それでも空なら部屋区分
            tech_slot_day = normalize_str(day.get("技術者区分", "")) or tech_slot_fallback or room_slot

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

            # 設備（利用なしならスキップ）
            if equipment_default_slot_day == "利用なし":
                eq_total, eq_df = 0, pd.DataFrame()
            else:
                eq_total, eq_df = calc_equipment_total_for_day(
                    day_slot_default=equipment_default_slot_day,
                    global_fallback_slot=global_fallback_slot,
                    group_overrides=group_overrides,
                    selections=[s.copy() for s in selections],
                    items=items,
                    group_meta=group_meta,
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
                            "単価": (r["単価(1区分)"] if pd.notna(r["単価(1区分)"]) else 0),
                            "倍率": r["倍率"],
                            "小計": r["小計"],
                            "自動追加": r["自動追加"],
                            "備考": f"{r['課金タイプ']} / {r.get('備考','')}",
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

        # インターネット（期間合算。内訳は別表に出す）
        internet_total, internet_df = calc_internet_total(
            df_days_calc=df_days_calc,
            selected_rooms=rooms,
            use_pocket_wifi=use_pocket_wifi,
            use_fixed_line=use_fixed_line,
            use_temp_line=use_temp_line,
        )
        internet_grand += internet_total

        if not internet_df.empty:
            for _, r in internet_df.iterrows():
                breakdown_rows.append(
                    {
                        "日付": r["日付"],
                        "種別": "インターネット",
                        "グループ": r.get("フロア", ""),
                        "品目": r.get("品目", ""),
                        "区分": "日",
                        "数量/人数": 1,
                        "単価": r.get("小計", 0),
                        "倍率": 1,
                        "小計": r.get("小計", 0),
                        "自動追加": False,
                        "備考": r.get("備考", ""),
                    }
                )

        st.divider()
        st.subheader("合計")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("部屋代 合計", f"{room_grand:,} 円")
        c2.metric("設備 合計", f"{equipment_grand:,} 円")
        c3.metric("技術者 合計", f"{tech_grand:,} 円")
        c4.metric("インターネット 合計", f"{internet_grand:,} 円")
        c5.metric("総額（全部）", f"{room_grand + equipment_grand + tech_grand + internet_grand:,} 円")

        st.divider()
        st.subheader("内訳（日別）")
        df_break = pd.DataFrame(breakdown_rows)
        st.dataframe(df_break, use_container_width=True)

        st.divider()
        st.subheader("インターネット内訳（参考）")
        if internet_df.empty:
            st.info("インターネット未選択（または対象部屋なし）")
        else:
            st.dataframe(internet_df, use_container_width=True)
            st.caption("※常設回線の段階料金は「連続利用（休館日で途切れる）」を基準にしています。")


if __name__ == "__main__":
    main()
