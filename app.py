from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set

import pandas as pd
import streamlit as st

# 祝日判定
try:
    import jpholiday
except Exception:
    jpholiday = None


# =========================
# App / Paths
# =========================
APP_TITLE = "公会堂料金電卓 MVP（部屋×日編集＋設備＋技術者＋インターネット）"
DATA_DIR = Path(__file__).parent / "data"

PRICES_CSV = DATA_DIR / "prices.csv"
CLOSED_DAYS_CSV = DATA_DIR / "closed_days.csv"
EQUIPMENT_GROUPS_CSV = DATA_DIR / "equipment_groups.csv"
EQUIPMENT_MASTER_CSV = DATA_DIR / "equipment_master.csv"


# =========================
# Time slots
# =========================
TIME_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間", "全日", "延長30分"]
ROOM_SLOTS_WITH_NONE = ["利用なし"] + TIME_SLOTS

# 設備は「利用なし」を許可（=倍率0で課金停止）
EQUIPMENT_TIME_SLOTS = ["利用なし"] + TIME_SLOTS

# 技術者も「利用なし」を許可（=課金停止）
TECH_TIME_SLOTS = ["利用なし"] + TIME_SLOTS


# =========================
# マイク/拡声装置：事故防止ルール（MVPはハードコードが無難）
# ★ここはCSVの item_id と部屋名に合わせて調整OK
# =========================
MIC_NEVER_ROOMS = {"第1会議室", "第2会議室", "第3会議室", "第4会議室", "第5会議室", "第9会議室", "特別室"}
MIC_C_ROOMS = {"大会議室", "小会議室"}  # 拡声装置C
MIC_D_ROOMS = {"第6会議室", "第7会議室", "第8会議室"}  # 3室すべて＆ギャラリー利用で拡声装置D

# CSVの item_id 想定（あなたのCSVに合わせる）
MIC_WIRED_ID = "mic_wired"
MIC_WIRELESS_ID = "mic_wireless"
PA_C_ID = "pa_c"
PA_D_ID = "pa_d"

# 「拡声装置にマイク1本付属」を控除する対象
PA_ITEMS_WITH_INCLUDED_MIC = {PA_C_ID, PA_D_ID, "pa_a", "pa_b"}  # A/Bが存在しても害なし
MIC_ITEMS = {MIC_WIRED_ID, MIC_WIRELESS_ID}

# マイク関連として“入力停止/課金停止”する item_id
MIC_RELATED_ITEM_IDS = {MIC_WIRED_ID, MIC_WIRELESS_ID, PA_C_ID, PA_D_ID}


def _safe_set(s: Set[str]) -> Set[str]:
    return set([x for x in s if x])


def infer_mic_allowed_for_rooms(rooms_used: Set[str], gallery_678: bool) -> Tuple[bool, str]:
    """
    その日の「使用部屋構成」からマイク可否を判定
    事故防止のため：
      - MIC_NEVER_ROOMS が混ざっていたら、その日はマイク不可
      - それ以外で、(大会議室 or 小会議室) が含まれれば OK（拡声装置C）
      - それ以外で、第6-8全て + ギャラリー利用なら OK（拡声装置D）
      - それ以外は不可
    """
    rooms_used = _safe_set(rooms_used)

    ng = sorted(list(rooms_used & MIC_NEVER_ROOMS))
    if ng:
        return False, f"マイク不可の部屋が含まれています: {', '.join(ng)}"

    if rooms_used & MIC_C_ROOMS:
        return True, "拡声装置Cの対象（大会議室/小会議室）"

    if MIC_D_ROOMS.issubset(rooms_used) and gallery_678:
        return True, "拡声装置Dの対象（第6+第7+第8 + ギャラリー利用）"

    if rooms_used & MIC_D_ROOMS:
        return False, "第6〜8会議室は「第6+第7+第8を全て」+「ギャラリー利用」の場合のみマイクOK"

    return False, "マイク対象部屋が含まれていません"


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


def parse_requires_groups(cell: str) -> List[List[str]]:
    """
    requires_item_ids 用：
    - "," は AND（複数必須）
    - "|" は OR（いずれか1つ）
    """
    s = normalize_str(cell)
    if s == "" or s == "*" or s.lower() == "all":
        return []
    for sep in ["、", "，", ";"]:
        s = s.replace(sep, ",")
    parts = [p.strip() for p in s.split(",") if p.strip()]
    groups: List[List[str]] = []
    for p in parts:
        alts = [a.strip() for a in p.split("|") if a.strip()]
        if alts:
            groups.append(alts)
    return groups


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
    requires_groups: List[List[str]]  # AND/OR混在
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
        req_groups = parse_requires_groups(r["requires_item_ids"])
        items[r["item_id"]] = EquipmentItem(
            item_id=r["item_id"],
            item_name=r["item_name"],
            group_id=r["group_id"],
            unit=r["unit"],
            price_per_slot=int(r["price_per_slot"]),
            price_once_yen=int(r["price_once_yen"]),
            requires_groups=req_groups,
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
        "利用なし": 0,
        "午前": 1,
        "午後": 1,
        "夜間": 1,
        "午前-午後": 2,
        "午後-夜間": 2,
        "全日": 3,
        "延長30分": 1,
    }
    return mapping.get(slot, 1)


def resolve_required_option(options: List[str], ctx: Dict[str, object]) -> Optional[str]:
    """
    OR候補から「どれを自動追加するか」を決める（事故防止・決め打ち）
    """
    opts = set(options)

    # pa_c|pa_d の解決
    if {PA_C_ID, PA_D_ID}.issubset(opts):
        need_c = bool(ctx.get("need_pa_c", False))
        need_d = bool(ctx.get("need_pa_d", False))

        if need_c and not need_d:
            return PA_C_ID
        if need_d and not need_c:
            return PA_D_ID

        # 判断不能 → 安全側で C
        return PA_C_ID if PA_C_ID in opts else options[0]

    return options[0] if options else None


def collect_required_items(
    selected_item_ids: List[str],
    items: Dict[str, EquipmentItem],
    requires_context: Dict[str, object],
) -> List[str]:
    """
    依存品を自動追加
    - ORは resolve_required_option で 1つだけ選ぶ
    """
    selected_set = set(selected_item_ids)
    added = True
    while added:
        added = False
        for iid in list(selected_set):
            it = items.get(iid)
            if not it:
                continue

            for group in it.requires_groups:
                if any(opt in selected_set for opt in group):
                    continue
                choice = resolve_required_option(group, requires_context)
                if choice and choice not in selected_set:
                    selected_set.add(choice)
                    added = True
    return list(selected_set)


def _fix_equip_cell(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "利用なし"
    if s not in EQUIPMENT_TIME_SLOTS:
        return "利用なし"
    return s


def _fix_room_slot(v: object, default_slot: str) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return default_slot
    if s not in ROOM_SLOTS_WITH_NONE:
        return default_slot
    return s


def _fix_tech_slot(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "利用なし"
    if s not in TECH_TIME_SLOTS:
        return "利用なし"
    return s


def calc_equipment_total_for_day(
    day_slot_default: str,
    global_fallback_slot: str,
    group_overrides: Dict[str, str],
    selections: List[Dict],
    items: Dict[str, EquipmentItem],
    group_meta: Dict[str, GroupMeta],
    requires_context: Dict[str, object],
    mic_allowed_today: bool,
) -> Tuple[int, pd.DataFrame]:
    """
    - 区分課金: price_per_slot > 0（倍率あり）
    - 区分なし単価: price_per_slot == 0 and price_once_yen > 0（倍率なし）
    - 依存品 ORは1つだけ自動追加
    - 拡声装置の付属マイク1本を控除（二重請求防止）
    - マイク不可日は MIC_RELATED_ITEM_IDS を課金0（UIもdisabledにできるが、念のため計算でも止める）
    """
    cols = ["種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"]
    if not selections:
        return 0, pd.DataFrame(columns=cols)

    selected_ids = [s["item_id"] for s in selections]
    full_ids = collect_required_items(selected_ids, items, requires_context)

    existing = {(s["group_id"], s["item_id"]) for s in selections}
    for iid in full_ids:
        if iid not in selected_ids:
            it = items[iid]
            key = (it.group_id, it.item_id)
            if key not in existing:
                selections.append({"group_id": it.group_id, "item_id": it.item_id, "qty": 1, "auto_added": True})

    # ---- 付属マイク控除：数量調整 ----
    qty_map: Dict[str, int] = {}
    for s in selections:
        iid = s.get("item_id")
        q = int(s.get("qty", 0) or 0)
        if not iid:
            continue
        qty_map[iid] = qty_map.get(iid, 0) + max(0, q)

    included_mics = sum(qty_map.get(pid, 0) for pid in PA_ITEMS_WITH_INCLUDED_MIC)
    req_wired = qty_map.get(MIC_WIRED_ID, 0)
    req_wireless = qty_map.get(MIC_WIRELESS_ID, 0)

    # 安全側：有線→ワイヤレス優先で控除
    remain = included_mics
    used_w = min(remain, req_wired)
    bill_wired = req_wired - used_w
    remain -= used_w

    used_ww = min(remain, req_wireless)
    bill_wireless = req_wireless - used_ww
    remain -= used_ww

    billed_qty_override = {
        MIC_WIRED_ID: bill_wired,
        MIC_WIRELESS_ID: bill_wireless,
    }
    deducted_note = {
        MIC_WIRED_ID: used_w,
        MIC_WIRELESS_ID: used_ww,
    }

    rows = []
    total = 0

    for s in selections:
        iid = s["item_id"]
        it = items.get(iid)
        if not it:
            continue

        meta = group_meta.get(it.group_id, GroupMeta(it.group_id, it.group_id, "*", 1, 1))

        orig_qty = int(s.get("qty", 0) or 0)
        if orig_qty <= 0:
            continue

        # マイク不可日は、マイク関連を課金0（ただし明細には残す）
        if (not mic_allowed_today) and (iid in MIC_RELATED_ITEM_IDS):
            billed_qty = 0
            forced_zero = True
        else:
            forced_zero = False
            if iid in MIC_ITEMS:
                billed_qty = int(billed_qty_override.get(iid, orig_qty))
            else:
                billed_qty = orig_qty

        inherit = bool(meta.default_inherit_room_slot)
        base_slot = day_slot_default if inherit else global_fallback_slot

        if meta.allowed_slot_override and it.group_id in group_overrides:
            slot = group_overrides[it.group_id]
        else:
            slot = base_slot

        # 設備の区分が「利用なし」の場合は倍率0
        mult = slot_to_multiplier(slot)

        is_slot_item = it.price_per_slot > 0
        is_once_item = (it.price_per_slot == 0) and (it.price_once_yen > 0)

        if is_slot_item:
            per_slot_sub = it.price_per_slot * billed_qty * mult
            once_sub = it.price_once_yen * billed_qty
            subtotal = per_slot_sub + once_sub
            charge_type = "区分課金"
        elif is_once_item:
            per_slot_sub = 0
            once_sub = it.price_once_yen * billed_qty
            subtotal = once_sub
            charge_type = "区分なし単価"
            slot = "（区分なし）"
            mult = 0
        else:
            per_slot_sub = 0
            once_sub = 0
            subtotal = 0
            charge_type = "料金未設定"
            slot = "—"
            mult = 0

        total += subtotal

        note = it.notes or ""
        if iid in PA_ITEMS_WITH_INCLUDED_MIC:
            note = (note + " / " if note else "") + "マイク1本付属"

        if iid in MIC_ITEMS:
            ded = int(deducted_note.get(iid, 0))
            if ded > 0:
                note = (note + " / " if note else "") + f"付属マイク控除:{ded}（有線→ワイヤレス優先）"
            note = (note + " / " if note else "") + f"選択:{orig_qty}→課金:{billed_qty}"

        if forced_zero:
            note = (note + " / " if note else "") + "マイク不可日のため課金0"

        # billed_qty=0 でも「選択してた」事実は明細に出す（透明性）
        if billed_qty == 0 and orig_qty > 0 and (iid in MIC_RELATED_ITEM_IDS):
            pass  # 0行でも出す
        elif billed_qty == 0:
            # それ以外の品目は、0なら出さない（見やすさ優先）
            continue

        rows.append(
            {
                "種別": "設備",
                "グループ": it.group_id,
                "品目": it.item_name,
                "課金タイプ": charge_type,
                "区分": slot,
                "数量": billed_qty,
                "単価(1区分)": it.price_per_slot,
                "倍率": mult,
                "区分小計": per_slot_sub,
                "一回課金": once_sub,
                "小計": subtotal,
                "備考": note,
                "自動追加": bool(s.get("auto_added", False)),
            }
        )

    df = pd.DataFrame(rows, columns=cols)
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
    if people <= 0 or slot == "利用なし":
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


# =========================
# Internet（MVP値：あとでCSV化OK）
# =========================
INTERNET_POCKET_WIFI_PER_DAY = 2800
INTERNET_FIXED_FIRST_DAY = 18000
INTERNET_FIXED_AFTER_DAY = 2000
INTERNET_TEMP_LINE_BASE = 5000  # + 別途見積（仮）

FLOOR_1_ROOMS = {"大集会室"}               # 1F扱い
FLOOR_3_ROOMS = {"中集会室", "小集会室"}   # 3F扱い（同フロア課金）


# =========================
# Day settings（表示用）
# =========================
def make_days_base(days: List[pd.Timestamp], closed_days: set, default_room_slot: str, is_business_default: bool) -> pd.DataFrame:
    rows = []
    for d in days:
        rows.append(
            {
                "日付": d.date().isoformat(),
                "土日祝": "土日祝" if is_weekend_or_holiday(d) else "平日",
                "祝日名": holiday_name(d),
                "休館日": bool(d.date() in closed_days),
                "割増利用": bool(is_business_default),
                "設備デフォ区分": default_room_slot,
                "技術者区分": default_room_slot,
            }
        )
    df = pd.DataFrame(rows)
    df["設備デフォ区分"] = df["設備デフォ区分"].apply(_fix_equip_cell)
    df["技術者区分"] = df["技術者区分"].apply(_fix_tech_slot)
    return df


def sync_days_df_defaults(df: pd.DataFrame, old_defaults: Dict[str, object], new_defaults: Dict[str, object]) -> pd.DataFrame:
    """
    旧デフォルトと同じ値のセルだけを新デフォルトへ置換（手編集を潰しにくい）
    """
    df = df.copy()

    # 表示系は再計算
    for i in range(len(df)):
        d = pd.Timestamp(df.loc[i, "日付"])
        df.loc[i, "土日祝"] = "土日祝" if is_weekend_or_holiday(d) else "平日"
        df.loc[i, "祝日名"] = holiday_name(d)

    cols = ["割増利用", "設備デフォ区分", "技術者区分"]
    for c in cols:
        if c not in df.columns:
            continue
        oldv = old_defaults.get(c, None)
        newv = new_defaults.get(c, None)
        if oldv == newv:
            continue
        mask = df[c].astype(str) == str(oldv)
        df.loc[mask, c] = newv

    df["設備デフォ区分"] = df["設備デフォ区分"].apply(_fix_equip_cell)
    df["技術者区分"] = df["技術者区分"].apply(_fix_tech_slot)
    return df


# =========================
# Room-Day table（計算の唯一入力）
# =========================
def _day_business_map(days_df: pd.DataFrame) -> Dict[str, bool]:
    m = {}
    for _, r in days_df.iterrows():
        m[normalize_str(r["日付"])] = bool(r.get("割増利用", False))
    return m


def build_room_day_base(days_df: pd.DataFrame, selected_rooms: List[str], default_room_slot: str) -> pd.DataFrame:
    rows = []
    day_business = _day_business_map(days_df)

    for _, drow in days_df.iterrows():
        date_str = normalize_str(drow["日付"])
        ts = pd.Timestamp(date_str)
        for room in selected_rooms:
            rows.append(
                {
                    "日付": date_str,
                    "土日祝": "土日祝" if is_weekend_or_holiday(ts) else "平日",
                    "祝日名": holiday_name(ts),
                    "休館日": bool(drow.get("休館日", False)),
                    "部屋": room,
                    "区分": default_room_slot,
                    "割増利用": bool(day_business.get(date_str, False)),
                    "手動区分": False,
                    "手動割増": False,
                }
            )
    df = pd.DataFrame(rows)
    df["区分"] = df["区分"].apply(lambda x: _fix_room_slot(x, default_room_slot))
    df["割増利用"] = df["割増利用"].astype(bool)
    return df


def merge_room_day(
    current: pd.DataFrame,
    days_df: pd.DataFrame,
    selected_rooms: List[str],
    default_room_slot: str,
) -> pd.DataFrame:
    """
    - 手動フラグFalseのセルは、日別デフォルトで上書き（=一括割増が効く）
    - 手動フラグTrueのセルは、ユーザー指定を保持
    """
    base = build_room_day_base(days_df, selected_rooms, default_room_slot)
    if current is None or current.empty:
        return base

    cur = current.copy()
    for c in ["日付", "部屋"]:
        if c in cur.columns:
            cur[c] = cur[c].map(normalize_str)

    # 旧互換
    if "手動区分" not in cur.columns:
        cur["手動区分"] = True
    if "手動割増" not in cur.columns:
        cur["手動割増"] = True

    key_cols = ["日付", "部屋"]
    keep_cols = key_cols + ["区分", "割増利用", "手動区分", "手動割増"]
    cur_small = cur[keep_cols].copy()

    merged = base.merge(cur_small, on=key_cols, how="left", suffixes=("", "_old"))

    merged["手動区分"] = merged["手動区分_old"].fillna(merged["手動区分"]).astype(bool)
    merged["手動割増"] = merged["手動割増_old"].fillna(merged["手動割増"]).astype(bool)

    merged.loc[merged["手動区分"], "区分"] = merged.loc[merged["手動区分"], "区分_old"].fillna(
        merged.loc[merged["手動区分"], "区分"]
    )
    merged.loc[merged["手動割増"], "割増利用"] = merged.loc[merged["手動割増"], "割増利用_old"].fillna(
        merged.loc[merged["手動割増"], "割増利用"]
    )

    drop_cols = [c for c in merged.columns if c.endswith("_old")]
    if drop_cols:
        merged.drop(columns=drop_cols, inplace=True)

    merged["区分"] = merged["区分"].apply(lambda x: _fix_room_slot(x, default_room_slot))
    merged["割増利用"] = merged["割増利用"].astype(bool)
    return merged


def apply_room_day_edits(full_df: pd.DataFrame, edited_subset: pd.DataFrame, default_room_slot: str) -> pd.DataFrame:
    """
    data_editorで編集された subset を full に反映し、変更があったセルは手動フラグを立てる
    """
    if full_df is None or full_df.empty or edited_subset is None or edited_subset.empty:
        return full_df

    full = full_df.copy()
    for c in ["日付", "部屋"]:
        full[c] = full[c].map(normalize_str)

    sub = edited_subset.copy()
    for c in ["日付", "部屋"]:
        sub[c] = sub[c].map(normalize_str)

    full_idx = {(r["日付"], r["部屋"]): i for i, r in full.iterrows()}

    for _, r in sub.iterrows():
        key = (r["日付"], r["部屋"])
        if key not in full_idx:
            continue
        i = full_idx[key]

        new_slot = _fix_room_slot(r.get("区分", ""), default_room_slot)
        new_bus = bool(r.get("割増利用", False))

        old_slot = normalize_str(full.loc[i, "区分"])
        old_bus = bool(full.loc[i, "割増利用"])

        if new_slot != old_slot:
            full.loc[i, "手動区分"] = True
        if new_bus != old_bus:
            full.loc[i, "手動割増"] = True

        full.loc[i, "区分"] = new_slot
        full.loc[i, "割増利用"] = new_bus

    full["割増利用"] = full["割増利用"].astype(bool)
    full["手動区分"] = full["手動区分"].astype(bool)
    full["手動割増"] = full["手動割増"].astype(bool)
    return full


# =========================
# 計算（部屋）
# =========================
def calc_rooms_from_room_day(prices_df: pd.DataFrame, room_day_df: pd.DataFrame) -> Tuple[int, pd.DataFrame]:
    if room_day_df is None or room_day_df.empty:
        return 0, pd.DataFrame(columns=["日付", "種別", "品目", "区分", "割増", "単価", "小計", "備考"])

    rows = []
    total = 0

    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)):
            continue

        date_str = normalize_str(r.get("日付", ""))
        room = normalize_str(r.get("部屋", ""))
        slot = normalize_str(r.get("区分", ""))
        is_business = bool(r.get("割増利用", False))

        dts = pd.Timestamp(date_str)

        if slot == "利用なし":
            rows.append({"日付": dts.date(), "種別": "部屋", "品目": room, "区分": "利用なし", "割増": is_business, "単価": 0, "小計": 0, "備考": ""})
            continue

        day_type = "土日祝" if is_weekend_or_holiday(dts) else "平日"
        price_type = "割増" if is_business else "通常"

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
                    "日付": dts.date(),
                    "種別": "部屋",
                    "品目": room,
                    "区分": slot,
                    "割増": is_business,
                    "単価": None,
                    "小計": None,
                    "備考": "該当料金が prices.csv に見つかりません",
                }
            )
            continue

        amount = int(hit.iloc[0]["amount"])
        total += amount
        rows.append({"日付": dts.date(), "種別": "部屋", "品目": room, "区分": slot, "割増": is_business, "単価": amount, "小計": amount, "備考": ""})

    df = pd.DataFrame(rows)
    return total, df


# =========================
# 計算用：日ごとの使用部屋を集計
# =========================
def rooms_used_by_date(room_day_df: pd.DataFrame) -> Dict[str, Set[str]]:
    """
    date_str -> set(rooms) （利用なし除外、休館日除外）
    """
    m: Dict[str, Set[str]] = {}
    if room_day_df is None or room_day_df.empty:
        return m

    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)):
            continue
        date_str = normalize_str(r.get("日付", ""))
        room = normalize_str(r.get("部屋", ""))
        slot = normalize_str(r.get("区分", ""))
        if date_str == "" or room == "":
            continue
        if slot == "利用なし":
            continue
        m.setdefault(date_str, set()).add(room)
    return m


def active_dates_from_room_day(room_day_df: pd.DataFrame) -> List[str]:
    return sorted(list(rooms_used_by_date(room_day_df).keys()))


def mic_allowed_map(room_day_df: pd.DataFrame, gallery_678: bool) -> Tuple[Dict[str, bool], Dict[str, str]]:
    """
    date_str -> mic_allowed bool / reason
    """
    ru = rooms_used_by_date(room_day_df)
    allowed: Dict[str, bool] = {}
    reason: Dict[str, str] = {}
    for d, rooms in ru.items():
        ok, why = infer_mic_allowed_for_rooms(rooms, gallery_678)
        allowed[d] = ok
        reason[d] = why
    return allowed, reason


# =========================
# 計算（設備：全日合算）
# =========================
def calc_equipment_total_all_days(
    days_df: pd.DataFrame,
    room_day_df: pd.DataFrame,
    global_default_slot: str,
    group_overrides: Dict[str, str],
    selections: List[Dict],
    items: Dict[str, EquipmentItem],
    group_meta: Dict[str, GroupMeta],
    gallery_678: bool,
) -> Tuple[int, pd.DataFrame]:
    """
    運用安全のため「実際に部屋が使われる日(active_dates)」だけ設備を課金
    """
    active_dates = active_dates_from_room_day(room_day_df)
    if not active_dates:
        return 0, pd.DataFrame(columns=["日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"])

    day_slot_map = {normalize_str(r["日付"]): _fix_equip_cell(r.get("設備デフォ区分", "")) for _, r in days_df.iterrows()}
    used_rooms = rooms_used_by_date(room_day_df)
    mic_ok_map, _ = mic_allowed_map(room_day_df, gallery_678)

    all_rows = []
    total = 0

    for d in active_dates:
        day_slot_default = day_slot_map.get(d, global_default_slot)
        rooms_today = used_rooms.get(d, set())

        # requiresの判断材料（pa_c/pa_d）
        requires_ctx = {
            "need_pa_c": bool(rooms_today & MIC_C_ROOMS),
            "need_pa_d": bool(MIC_D_ROOMS.issubset(rooms_today) and gallery_678),
        }

        mic_allowed_today = bool(mic_ok_map.get(d, False))

        day_total, day_df = calc_equipment_total_for_day(
            day_slot_default=day_slot_default,
            global_fallback_slot=global_default_slot,
            group_overrides=group_overrides,
            selections=[dict(x) for x in selections],  # 日ごとに副作用を分離
            items=items,
            group_meta=group_meta,
            requires_context=requires_ctx,
            mic_allowed_today=mic_allowed_today,
        )
        total += day_total

        if not day_df.empty:
            day_df.insert(0, "日付", pd.Timestamp(d).date())
            all_rows.append(day_df)

    if all_rows:
        out = pd.concat(all_rows, ignore_index=True)
        out = out[["日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"]]
        return total, out

    return 0, pd.DataFrame(columns=["日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"])


# =========================
# 計算（技術者：全日合算）
# =========================
def calc_stage_tech_total_all_days(days_df: pd.DataFrame, room_day_df: pd.DataFrame, people: int) -> Tuple[int, pd.DataFrame]:
    active_dates = active_dates_from_room_day(room_day_df)
    if not active_dates:
        return 0, pd.DataFrame(columns=["日付", "種別", "区分", "人数", "単価(1名)", "小計"])

    tech_slot_map = {normalize_str(r["日付"]): _fix_tech_slot(r.get("技術者区分", "")) for _, r in days_df.iterrows()}

    rows = []
    total = 0
    for d in active_dates:
        slot = tech_slot_map.get(d, "利用なし")
        sub, df = calc_stage_tech_total_for_day(slot, people)
        total += sub
        if not df.empty:
            df.insert(0, "日付", pd.Timestamp(d).date())
            rows.append(df)

    out = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["日付", "種別", "区分", "人数", "単価(1名)", "小計"])
    return total, out


# =========================
# 計算（インターネット：全日合算）
# =========================
def _split_consecutive_blocks(dates: List[pd.Timestamp]) -> List[List[pd.Timestamp]]:
    if not dates:
        return []
    dates = sorted(dates)
    blocks: List[List[pd.Timestamp]] = []
    block = [dates[0]]
    for d in dates[1:]:
        if (d - block[-1]).days == 1:
            block.append(d)
        else:
            blocks.append(block)
            block = [d]
    blocks.append(block)
    return blocks


def infer_active_days_by_floor(room_day_df: pd.DataFrame) -> Dict[str, List[pd.Timestamp]]:
    ru = rooms_used_by_date(room_day_df)
    f1 = []
    f3 = []
    for d, rooms in ru.items():
        dt = pd.Timestamp(d)
        if rooms & FLOOR_1_ROOMS:
            f1.append(dt)
        if rooms & FLOOR_3_ROOMS:
            f3.append(dt)
    return {"1F": sorted(f1), "3F": sorted(f3)}


def calc_internet_total(
    room_day_df: pd.DataFrame,
    use_pocket_wifi: bool,
    use_fixed_line: bool,
    use_temp_line: bool,
) -> Tuple[int, pd.DataFrame]:
    """
    - ポケットWi-Fi：利用日（=active_dates）ごと
    - 常設回線：フロアごとの「連続ブロック」単位で初日/2日目以降
    - 仮設回線：フロアごとの「連続ブロック開始日」に 5,000円/回（+別途見積）
    """
    active_dates = [pd.Timestamp(d) for d in active_dates_from_room_day(room_day_df)]
    if not active_dates:
        return 0, pd.DataFrame(columns=["日付", "種別", "品目", "フロア", "小計", "備考"])

    rows = []
    total = 0

    # ポケットWi-Fi
    if use_pocket_wifi:
        for d in sorted(active_dates):
            rows.append({"日付": d.date(), "種別": "インターネット", "品目": "ポケットWi-Fi貸出", "フロア": "全部屋", "小計": INTERNET_POCKET_WIFI_PER_DAY, "備考": "先着順/同時接続目安5台/電波不安定の可能性"})
            total += INTERNET_POCKET_WIFI_PER_DAY

    floors = infer_active_days_by_floor(room_day_df)

    # 常設回線（段階）
    if use_fixed_line:
        for floor_label, ds in floors.items():
            blocks = _split_consecutive_blocks(ds)
            for b in blocks:
                if not b:
                    continue
                # 初日
                rows.append({"日付": b[0].date(), "種別": "インターネット", "品目": "常設回線（初日）", "フロア": floor_label, "小計": INTERNET_FIXED_FIRST_DAY, "備考": "連続利用の段階料金"})
                total += INTERNET_FIXED_FIRST_DAY
                # 2日目以降
                for d in b[1:]:
                    rows.append({"日付": d.date(), "種別": "インターネット", "品目": "常設回線（2日目以降）", "フロア": floor_label, "小計": INTERNET_FIXED_AFTER_DAY, "備考": "連続利用の段階料金"})
                    total += INTERNET_FIXED_AFTER_DAY

    # 仮設回線（ブロック開始に1回）
    if use_temp_line:
        for floor_label, ds in floors.items():
            blocks = _split_consecutive_blocks(ds)
            for b in blocks:
                if not b:
                    continue
                rows.append({"日付": b[0].date(), "種別": "インターネット", "品目": "仮設回線（開通工事）", "フロア": floor_label, "小計": INTERNET_TEMP_LINE_BASE, "備考": "＋別途お見積り（NTT回線開通工事）"})
                total += INTERNET_TEMP_LINE_BASE

    df = pd.DataFrame(rows, columns=["日付", "種別", "品目", "フロア", "小計", "備考"])
    return total, df


# =========================
# KPI Display
# =========================
def yen(x: int) -> str:
    try:
        return f"{int(x):,} 円"
    except Exception:
        return f"{x} 円"


def render_kpis(room_total: int, equipment_total: int, tech_total: int, internet_total: int):
    st.markdown(
        """
<style>
.kpi-wrap {display:flex; gap:14px; flex-wrap:wrap;}
.kpi {
  border: 1px solid rgba(0,0,0,0.10);
  border-radius: 12px;
  padding: 12px 14px;
  min-width: 220px;
  flex: 1;
}
.kpi .label {font-size: 14px; opacity: 0.75; margin-bottom: 6px;}
.kpi .value {font-size: 22px; font-weight: 700; white-space: nowrap; overflow: hidden; text-overflow: clip;}
</style>
""",
        unsafe_allow_html=True,
    )

    grand_total = room_total + equipment_total + tech_total + internet_total

    st.markdown(
        f"""
<div class="kpi-wrap">
  <div class="kpi"><div class="label">部屋代 合計</div><div class="value">{yen(room_total)}</div></div>
  <div class="kpi"><div class="label">設備 合計</div><div class="value">{yen(equipment_total)}</div></div>
  <div class="kpi"><div class="label">技術者 合計</div><div class="value">{yen(tech_total)}</div></div>
  <div class="kpi"><div class="label">インターネット 合計</div><div class="value">{yen(internet_total)}</div></div>
  <div class="kpi"><div class="label">総額（全部）</div><div class="value">{yen(grand_total)}</div></div>
</div>
""",
        unsafe_allow_html=True,
    )


# =========================
# Main App
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

    # -------------------------
    # LEFT: 期間 / 日別 / 部屋×日
    # -------------------------
    with left:
        st.subheader("1) 期間・部屋（部屋×日テーブル編集）")

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

        default_room_slot = st.selectbox("部屋の区分（新規追加の初期値）", TIME_SLOTS, index=TIME_SLOTS.index("全日"))
        is_business_default = st.checkbox("割増利用（デフォルト）", value=False)

        st.divider()
        st.subheader("日別設定（設備・技術者・インターネットに使用）")
        st.caption("※部屋は『部屋×日テーブル』が計算の唯一入力（見た目＝計算）になります。")

        days_key = f"days_{start_date}_{end_date}"
        new_defaults = {"割増利用": bool(is_business_default), "設備デフォ区分": default_room_slot, "技術者区分": default_room_slot}

        if days_key not in st.session_state:
            df_days = make_days_base(days, closed_days, default_room_slot, is_business_default)
            st.session_state[days_key] = df_days
            st.session_state[days_key + "_defaults"] = dict(new_defaults)
        else:
            df_existing = st.session_state[days_key]
            if (
                len(df_existing) != len(days)
                or df_existing.iloc[0]["日付"] != days[0].date().isoformat()
                or df_existing.iloc[-1]["日付"] != days[-1].date().isoformat()
            ):
                df_days = make_days_base(days, closed_days, default_room_slot, is_business_default)
                st.session_state[days_key] = df_days
                st.session_state[days_key + "_defaults"] = dict(new_defaults)
            else:
                old_defaults = st.session_state.get(days_key + "_defaults", dict(new_defaults))
                st.session_state[days_key] = sync_days_df_defaults(df_existing, old_defaults, new_defaults)
                st.session_state[days_key + "_defaults"] = dict(new_defaults)

        try:
            edited_days = st.data_editor(
                st.session_state[days_key],
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "日付": st.column_config.TextColumn(disabled=True),
                    "土日祝": st.column_config.TextColumn(disabled=True),
                    "祝日名": st.column_config.TextColumn(disabled=True),
                    "休館日": st.column_config.CheckboxColumn(disabled=True),
                    "割増利用": st.column_config.CheckboxColumn(),
                    "設備デフォ区分": st.column_config.SelectboxColumn(options=EQUIPMENT_TIME_SLOTS),
                    "技術者区分": st.column_config.SelectboxColumn(options=TECH_TIME_SLOTS),
                },
            )
            edited_days = edited_days.copy()
            edited_days["設備デフォ区分"] = edited_days["設備デフォ区分"].apply(_fix_equip_cell)
            edited_days["技術者区分"] = edited_days["技術者区分"].apply(_fix_tech_slot)
            st.session_state[days_key] = edited_days
        except Exception:
            st.warning("この環境では日別編集UIが使えないため、日別設定は表示のみになります。")
            edited_days = st.session_state[days_key]
            st.dataframe(edited_days, use_container_width=True)

        st.divider()
        st.subheader("部屋×日 テーブル（ここで区分/割増を個別に調整できます）")
        st.caption("編集後は『変更を反映（確定）』を押してから計算してください（ラグ事故防止）。")

        room_day_key = f"room_day_{start_date}_{end_date}"

        if room_day_key not in st.session_state:
            st.session_state[room_day_key] = build_room_day_base(edited_days, list(rooms), default_room_slot)
        else:
            st.session_state[room_day_key] = merge_room_day(st.session_state[room_day_key], edited_days, list(rooms), default_room_slot)

        # フィルター
        st.markdown("### フィルター（編集しやすくする）")
        f1, f2 = st.columns([1, 1])
        all_dates = sorted(st.session_state[room_day_key]["日付"].unique().tolist()) if not st.session_state[room_day_key].empty else []
        date_filter = f1.multiselect("日付で絞り込み（未選択＝全日）", options=all_dates, default=[], key=f"filter_dates_{room_day_key}")
        all_rooms_in_table = sorted(st.session_state[room_day_key]["部屋"].unique().tolist()) if not st.session_state[room_day_key].empty else []
        room_filter = f2.multiselect("部屋で絞り込み（未選択＝全部屋）", options=all_rooms_in_table, default=[], key=f"filter_rooms_{room_day_key}")

        view_df = st.session_state[room_day_key].copy()
        if date_filter:
            view_df = view_df[view_df["日付"].isin(date_filter)]
        if room_filter:
            view_df = view_df[view_df["部屋"].isin(room_filter)]
        view_df = view_df.reset_index(drop=True)

        st.info("✅ 表を編集したら、下の『この表の変更を反映（確定）』を押してね（ラグ事故防止）")

        try:
            with st.form("room_day_form", clear_on_submit=False):
                edited_room_day_tmp = st.data_editor(
                    view_df,
                    use_container_width=True,
                    num_rows="fixed",
                    column_config={
                        "日付": st.column_config.TextColumn(disabled=True),
                        "土日祝": st.column_config.TextColumn(disabled=True),
                        "祝日名": st.column_config.TextColumn(disabled=True),
                        "休館日": st.column_config.CheckboxColumn(disabled=True),
                        "部屋": st.column_config.TextColumn(disabled=True),
                        "区分": st.column_config.SelectboxColumn(options=ROOM_SLOTS_WITH_NONE),
                        "割増利用": st.column_config.CheckboxColumn(),
                        "手動区分": st.column_config.CheckboxColumn(disabled=True),
                        "手動割増": st.column_config.CheckboxColumn(disabled=True),
                    },
                )

                submitted = st.form_submit_button("この表の変更を反映（確定）")

            if submitted:
                # subset を full に反映し、手動フラグを自動で立てる
                updated_full = apply_room_day_edits(st.session_state[room_day_key], edited_room_day_tmp, default_room_slot)
                st.session_state[room_day_key] = updated_full
                st.success("✅ 反映しました！このまま『計算する』を押せます。")

        except Exception:
            st.warning("この環境では部屋×日編集UIが使えないため、表示のみになります。")
            st.dataframe(view_df, use_container_width=True)

    # -------------------------
    # RIGHT: 設備 / 技術者 / ネット / 計算結果
    # -------------------------
    with right:
        st.subheader("2) 設備・技術者・インターネット（入力）")

        room_day_df = st.session_state.get(room_day_key, pd.DataFrame())
        if room_day_df is None:
            room_day_df = pd.DataFrame()

        # マイク判定：日ごと
        gallery_678 = st.checkbox("（第6〜8会議室）ギャラリー利用", value=False, help="第6+第7+第8 を全て使う場合のマイク可否に影響します。")
        mic_ok_map, mic_reason_map = mic_allowed_map(room_day_df, gallery_678)

        mic_any_allowed = any(mic_ok_map.values()) if mic_ok_map else False
        mic_any_denied = any([not v for v in mic_ok_map.values()]) if mic_ok_map else False

        if mic_any_denied:
            denied_days = [d for d, ok in mic_ok_map.items() if not ok]
            # 直近だけ見せる（長いと見づらい）
            show_days = denied_days[:7]
            st.error(
                "【マイク注意】マイク不可の日があります。不可日はマイク/拡声装置は課金0＆入力も止めます。\n"
                f"- マイク不可日（抜粋）: {', '.join(show_days)}"
                + (f" …他 {max(0, len(denied_days)-len(show_days))}日" if len(denied_days) > len(show_days) else "")
            )

        # -------------------------
        # 設備 UI（グループごと）
        # -------------------------
        st.markdown("### 設備（数量）")
        st.caption("※マイク関連（拡声装置/有線/ワイヤレス）は、マイク可の日が1日も無い場合は入力を停止します。")

        selected_rooms_now = sorted(list(set([normalize_str(x) for x in rooms])))
        selected_rooms_set = set(selected_rooms_now)

        # グループ表示のため：適用部屋が1つも一致しないグループは隠す
        def group_applies(meta: GroupMeta) -> bool:
            targets = parse_list_cell(meta.applies_to_rooms)
            if targets == ["*"]:
                return True
            return bool(set(targets) & selected_rooms_set) if selected_rooms_set else True

        # グループ override
        group_overrides: Dict[str, str] = {}
        st.markdown("#### 設備の区分（グループ単位override）")
        st.caption("「（日別デフォルト）」のままなら、日別設定の『設備デフォ区分』がそのまま使われます。")
        override_container = st.container()

        with override_container:
            for gid, meta in group_meta.items():
                if not group_applies(meta):
                    continue

                if not meta.allowed_slot_override:
                    st.text(f"・{meta.group_name}：（override不可）")
                    continue

                key = f"ov_{gid}"
                default_label = "（日別デフォルト）"
                options = [default_label] + EQUIPMENT_TIME_SLOTS
                choice = st.selectbox(f"{meta.group_name}", options=options, index=0, key=key)

                if choice != default_label:
                    group_overrides[gid] = choice

        st.divider()

        # 数量入力
        selections: List[Dict] = []
        # item_id から group_id を引く（itemsが正）
        items_by_group: Dict[str, List[EquipmentItem]] = {}
        for it in items.values():
            items_by_group.setdefault(it.group_id, []).append(it)

        # 見た目のため、グループ順を groups_df に寄せる
        group_order = [normalize_str(x) for x in groups_df["group_id"].tolist()] if "group_id" in groups_df.columns else sorted(list(group_meta.keys()))

        for gid in group_order:
            meta = group_meta.get(gid)
            if not meta:
                continue
            if not group_applies(meta):
                continue

            group_items = sorted(items_by_group.get(gid, []), key=lambda x: x.item_name)

            with st.expander(f"{meta.group_name}", expanded=False):
                st.caption(f"対象部屋: {meta.applies_to_rooms if meta.applies_to_rooms else '*'}")
                for it in group_items:
                    # マイク関連は、可の日が1日もなければ入力停止（止まる）
                    mic_disabled = (it.item_id in MIC_RELATED_ITEM_IDS) and (not mic_any_allowed)

                    qty_key = f"qty_{it.item_id}"
                    if mic_disabled:
                        # UIも状態もゼロ固定
                        st.session_state[qty_key] = 0

                    price_txt = []
                    if it.price_per_slot > 0:
                        price_txt.append(f"1区分:{it.price_per_slot:,}円")
                    if it.price_once_yen > 0:
                        price_txt.append(f"単価:{it.price_once_yen:,}円")
                    price_str = " / ".join(price_txt) if price_txt else "（料金未設定）"

                    label = f"{it.item_name}（単位:{it.unit} / {price_str}）"
                    help_txt = it.notes if it.notes else None

                    qty = st.number_input(
                        label,
                        min_value=0,
                        value=int(st.session_state.get(qty_key, 0) or 0),
                        step=1,
                        key=qty_key,
                        disabled=mic_disabled,
                        help=help_txt,
                    )

                    if int(qty) > 0:
                        selections.append({"group_id": it.group_id, "item_id": it.item_id, "qty": int(qty), "auto_added": False})

        st.divider()

        # -------------------------
        # 技術者
        # -------------------------
        st.markdown("### 舞台設備技術者")
        tech_people = st.number_input("人数", min_value=0, value=0, step=1, help="日別設定の『技術者区分』×人数で計算します。")

        st.divider()

        # -------------------------
        # インターネット
        # -------------------------
        st.markdown("### インターネット")
        use_pocket_wifi = st.checkbox("ポケットWi-Fi（2,800円/日）", value=False)
        use_fixed_line = st.checkbox("常設回線（初日18,000円、2日目以降2,000円）", value=False)
        use_temp_line = st.checkbox("仮設回線（5,000円/回 + 別途見積）", value=False)

        st.divider()

        # -------------------------
        # 計算ボタン
        # -------------------------
        do_calc = st.button("計算する", type="primary")

        if do_calc:
            # 部屋
            room_total, room_df = calc_rooms_from_room_day(prices_df, room_day_df)

            # 設備（active dates のみ）
            equipment_total, equipment_df = calc_equipment_total_all_days(
                days_df=edited_days,
                room_day_df=room_day_df,
                global_default_slot=default_room_slot,
                group_overrides=group_overrides,
                selections=selections,
                items=items,
                group_meta=group_meta,
                gallery_678=gallery_678,
            )

            # 技術者（active dates のみ）
            tech_total, tech_df = calc_stage_tech_total_all_days(edited_days, room_day_df, int(tech_people))

            # インターネット（active dates のみ）
            internet_total, internet_df = calc_internet_total(
                room_day_df=room_day_df,
                use_pocket_wifi=use_pocket_wifi,
                use_fixed_line=use_fixed_line,
                use_temp_line=use_temp_line,
            )

            st.subheader("結果")
            render_kpis(room_total, equipment_total, tech_total, internet_total)

            # マイク詳細理由（不可日だけ）
            if mic_ok_map:
                denied = [(d, mic_reason_map.get(d, "")) for d, ok in mic_ok_map.items() if not ok]
                if denied:
                    with st.expander("マイク不可日の理由（詳細）", expanded=False):
                        for d, why in denied[:20]:
                            st.write(f"- {d}: {why}")
                        if len(denied) > 20:
                            st.write(f"…他 {len(denied)-20}日")

            # 明細
            tab_all, tab_rooms, tab_eq, tab_tech, tab_net = st.tabs(["明細（全部）", "部屋", "設備", "技術者", "インターネット"])

            with tab_rooms:
                st.dataframe(room_df, use_container_width=True)

            with tab_eq:
                st.dataframe(equipment_df, use_container_width=True)

            with tab_tech:
                st.dataframe(tech_df, use_container_width=True)

            with tab_net:
                st.dataframe(internet_df, use_container_width=True)

            with tab_all:
                frames = []
                if not room_df.empty:
                    r = room_df.copy()
                    r = r.rename(columns={"品目": "名称"})
                    r["カテゴリ"] = "部屋"
                    frames.append(r[["日付", "カテゴリ", "名称", "区分", "小計", "備考"]])
                if not equipment_df.empty:
                    e = equipment_df.copy()
                    e["カテゴリ"] = "設備"
                    e = e.rename(columns={"品目": "名称"})
                    frames.append(e[["日付", "カテゴリ", "名称", "区分", "小計", "備考"]])
                if not tech_df.empty:
                    t = tech_df.copy()
                    t["カテゴリ"] = "技術者"
                    t["名称"] = "舞台設備技術者"
                    frames.append(t.rename(columns={"小計": "小計"})[["日付", "カテゴリ", "名称", "区分", "小計"]])
                if not internet_df.empty:
                    n = internet_df.copy()
                    n["カテゴリ"] = "インターネット"
                    n = n.rename(columns={"品目": "名称"})
                    frames.append(n.rename(columns={"小計": "小計"})[["日付", "カテゴリ", "名称", "フロア", "小計", "備考"]])

                if frames:
                    all_df = pd.concat(frames, ignore_index=True)
                    st.dataframe(all_df, use_container_width=True)
                else:
                    st.info("明細がありません（部屋×日が全部「利用なし」など）")


if __name__ == "__main__":
    main()