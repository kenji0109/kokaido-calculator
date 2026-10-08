from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Tuple, Optional, Set

import re
import pandas as pd
import streamlit as st

# 祝日判定（入ってなければ週末のみ判定）
try:
    import jpholiday  # type: ignore
except Exception:
    jpholiday = None

# =========================
# App / Paths
# =========================
APP_TITLE = "料金電卓"
APP_SUBTITLE = "部屋・設備・技術者・インターネットの料金を計算します"
DATA_DIR = Path(__file__).parent / "data"

PRICES_CSV = DATA_DIR / "prices.csv"
CLOSED_DAYS_CSV = DATA_DIR / "closed_days.csv"
EQUIPMENT_GROUPS_CSV = DATA_DIR / "equipment_groups.csv"
EQUIPMENT_MASTER_CSV = DATA_DIR / "equipment_master.csv"

# =========================
# 表示用 日付フォーマット
# =========================
DATE_FMT = "%Y/%m/%d"

# =========================
# Time slots
# =========================
TIME_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間", "全日", "延長30分"]

ROOM_BASE_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間", "全日"]
ROOM_SLOTS_WITH_NONE = ["利用なし"] + ROOM_BASE_SLOTS

ROOM_EXTENSION_SLOTS = ["なし", "前延長30分", "後延長30分", "前後延長30分"]

# 「全館」を選択したときは他の部屋の選択を外す
ALL_BUILDING_ROOM = "全館"

# 部屋の選択欄に出す順番（ここにない部屋は、この後ろに名前順で出す）
ROOM_DISPLAY_ORDER = [
    ALL_BUILDING_ROOM,
    "中集会室", "特別室", "小集会室",
    "第9会議室", "第6会議室", "第7会議室", "第8会議室",
    "大集会室", "控室1", "控室2", "第5会議室",
    "大会議室", "第1会議室", "第2会議室", "第3会議室", "第4会議室",
]

# 21:30終了の区分（後延長できない）
ROOM_SLOTS_ENDING_2130 = {"夜間", "午後-夜間", "全日"}
ROOM_EXTENSIONS_AFTER = {"後延長30分", "前後延長30分"}
ROOM_EXTENSION_SLOTS_2130 = [x for x in ROOM_EXTENSION_SLOTS if x not in ROOM_EXTENSIONS_AFTER]
INVALID_AFTER_EXTENSION_NOTE = "21:30終了の区分は後延長できません（延長分を計算から除外）"

EQUIPMENT_TIME_SLOTS = ["利用なし"] + TIME_SLOTS
TECH_TIME_SLOTS = ["利用なし"] + TIME_SLOTS

# =========================
# マイク/拡声装置：事故防止ルール
# =========================
MIC_NEVER_ROOMS = {"第1会議室", "第2会議室", "第3会議室", "第4会議室", "第5会議室", "第9会議室", "特別室"}
MIC_C_ROOMS = {"大会議室", "小集会室"}
MIC_D_ROOMS = {"第6会議室", "第7会議室", "第8会議室"}

MIC_WIRED_ID = "mic_wired"
MIC_WIRELESS_ID = "mic_wireless"
MIC_STAND_ID = "mic_stand"
PA_C_ID = "pa_c"
PA_D_ID = "pa_d"

# 「拡声装置にマイク1本・スタンド1本付属」を控除する対象
PA_ITEMS_WITH_INCLUDED_MIC = {PA_C_ID, PA_D_ID, "pa_a", "pa_b"}
PA_ITEMS_WITH_INCLUDED_STAND = {PA_C_ID, PA_D_ID, "pa_a", "pa_b"}

MIC_ITEMS = {MIC_WIRED_ID, MIC_WIRELESS_ID}
STAND_ITEMS = {MIC_STAND_ID}

MIC_RELATED_ITEM_IDS = {MIC_WIRED_ID, MIC_WIRELESS_ID, PA_C_ID, PA_D_ID}

EQUIPMENT_QTY_STORE_KEY = "equipment_qty_store"

# =========================
# Utility
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
    try:
        return int(float(x))
    except Exception:
        return 0

def parse_date_str(s: str) -> Optional[pd.Timestamp]:
    s = normalize_str(s)
    if not s:
        return None
    ts = pd.to_datetime(s, format=DATE_FMT, errors="coerce")
    if pd.isna(ts):
        ts = pd.to_datetime(s, errors="coerce")
    if pd.isna(ts):
        return None
    return pd.Timestamp(ts)

# =========================
# applies_to_rooms パーサ
# =========================
_RANGE_PAT = re.compile(r"第\s*(\d+)\s*[〜～\-－—]\s*(?:第\s*)?(\d+)\s*会議室")
_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")

def _normalize_digits(s: str) -> str:
    return s.translate(_FULLWIDTH_DIGITS)

_ROOM_NAME_PAT = re.compile(r"(大会議室|大集会室|中集会室|小集会室|特別室|控え室\s*1|控え室\s*2|控室\s*1|控室\s*2|第\s*\d+\s*会議室)")

def _expand_room_range(token: str) -> List[str]:
    t = normalize_str(token)
    t = _normalize_digits(t)
    m = _RANGE_PAT.search(t)
    if not m:
        return [t] if t else []
    a = int(m.group(1))
    b = int(m.group(2))
    lo, hi = (a, b) if a <= b else (b, a)
    return [f"第{i}会議室" for i in range(lo, hi + 1)]

def parse_rooms_cell(cell: str) -> List[str]:
    s = normalize_str(cell)
    s = _normalize_digits(s)
    if s == "" or s == "*" or s.lower() == "all":
        return ["*"]

    s = re.sub(r"[()\[\]{}（）【】]", " ", s)

    for sep in ["、", "，", ",", "・", "/", "／", ";", "；", "\n", "\t", "　", " "]:
        s = s.replace(sep, ",")

    tokens = [t.strip() for t in s.split(",") if t.strip()]
    out: List[str] = []
    for t in tokens:
        out.extend(_expand_room_range(t))

    return out if out else ["*"]

def infer_item_target_rooms(item_name: str, notes: str, fallback: str) -> str:
    text = _normalize_digits(f"{item_name} {notes}")
    targets: Set[str] = set()

    for m in _RANGE_PAT.finditer(text):
        a = int(m.group(1))
        b = int(m.group(2))
        lo, hi = (a, b) if a <= b else (b, a)
        for i in range(lo, hi + 1):
            targets.add(f"第{i}会議室")

    for m in _ROOM_NAME_PAT.finditer(text):
        token = m.group(0)
        token = re.sub(r"\s+", "", token)
        targets.add(token)

    if targets:
        return " / ".join(sorted(targets))
    return fallback if fallback else "*"

def parse_requires_groups(cell: str) -> List[List[str]]:
    s = normalize_str(cell)
    if s == "" or s == "*" or s.lower() == "all":
        return []
    for sep in ["、", "，", ";", "；"]:
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
            return bool(jpholiday.is_holiday(date.date()))
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

    s = pd.to_datetime(df["date"], errors="coerce").dropna()
    return set(s.dt.date.tolist())

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
    requires_groups: List[List[str]]
    notes: str
    is_countable: int
    is_power_item: int

@dataclass
class GroupMeta:
    group_id: str
    group_name: str
    applies_to_rooms: str
    default_inherit_room_slot: int
    allowed_slot_override: int

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

    if "default_inherit_room_slot" not in groups_df.columns:
        groups_df["default_inherit_room_slot"] = 1
    if "allowed_slot_override" not in groups_df.columns:
        groups_df["allowed_slot_override"] = 1

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
    opts = set(options)

    if {PA_C_ID, PA_D_ID}.issubset(opts):
        need_c = bool(ctx.get("need_pa_c", False))
        need_d = bool(ctx.get("need_pa_d", False))

        if need_d and not need_c and PA_D_ID in opts:
            return PA_D_ID
        if need_c and not need_d and PA_C_ID in opts:
            return PA_C_ID

        return PA_C_ID if PA_C_ID in opts else (options[0] if options else None)

    return options[0] if options else None

def collect_required_items(
    selected_item_ids: List[str],
    items: Dict[str, EquipmentItem],
    requires_context: Dict[str, object],
) -> List[str]:
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

def _fix_room_extension(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "なし"
    if s not in ROOM_EXTENSION_SLOTS:
        return "なし"
    return s

def is_invalid_after_extension(slot: str, ext: str) -> bool:
    return normalize_str(slot) in ROOM_SLOTS_ENDING_2130 and _fix_room_extension(ext) in ROOM_EXTENSIONS_AFTER

def _fix_tech_slot(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "利用なし"
    if s not in TECH_TIME_SLOTS:
        return "利用なし"
    return s

def _safe_set(s: Set[str]) -> Set[str]:
    return {x for x in s if x}

def infer_mic_allowed_for_rooms(rooms_used: Set[str], gallery_678: bool) -> Tuple[bool, str]:
    rooms_used = _safe_set(rooms_used)

    never = sorted(list(rooms_used & MIC_NEVER_ROOMS))
    never_note = f"（同日にマイク対象外の部屋が含まれています: {', '.join(never)}）" if never else ""

    if rooms_used & MIC_C_ROOMS:
        return True, f"拡声装置Cの対象（大会議室/小集会室）{never_note}"

    if MIC_D_ROOMS.issubset(rooms_used) and gallery_678:
        return True, f"拡声装置Dの対象（第6+第7+第8 + ギャラリー利用）{never_note}"

    if rooms_used & MIC_D_ROOMS:
        return False, f"第6〜8会議室は「第6+第7+第8を全て」かつ「ギャラリー利用」の場合のみマイク対象{never_note}"

    if rooms_used & MIC_NEVER_ROOMS:
        return False, f"マイク対象外の部屋のみです{never_note}"

    return False, "マイク対象部屋（大会議室/小集会室 または 第6〜8条件）が含まれていません"

def _is_mic_related_item_allowed_today(iid: str, ctx: Dict[str, object], mic_allowed_today: bool) -> bool:
    need_c = bool(ctx.get("need_pa_c", False))
    need_d = bool(ctx.get("need_pa_d", False))

    if iid == PA_C_ID:
        return need_c
    if iid == PA_D_ID:
        return need_d
    if iid in MIC_ITEMS:
        return mic_allowed_today
    return True

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
    cols = [
        "種別", "グループ", "品目", "課金タイプ", "区分", "数量",
        "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"
    ]
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

    # マイク控除
    included_mics = sum(qty_map.get(pid, 0) for pid in PA_ITEMS_WITH_INCLUDED_MIC)
    req_wired = qty_map.get(MIC_WIRED_ID, 0)
    req_wireless = qty_map.get(MIC_WIRELESS_ID, 0)

    remain = included_mics
    used_w = min(remain, req_wired)
    bill_wired = req_wired - used_w
    remain -= used_w

    used_ww = min(remain, req_wireless)
    bill_wireless = req_wireless - used_ww
    remain -= used_ww

    billed_qty_override = {MIC_WIRED_ID: bill_wired, MIC_WIRELESS_ID: bill_wireless}
    deducted_note = {MIC_WIRED_ID: used_w, MIC_WIRELESS_ID: used_ww}

    # スタンド控除
    included_stands = sum(qty_map.get(pid, 0) for pid in PA_ITEMS_WITH_INCLUDED_STAND)
    req_stand = qty_map.get(MIC_STAND_ID, 0)
    used_stand = min(included_stands, req_stand)
    bill_stand = req_stand - used_stand
    billed_qty_override[MIC_STAND_ID] = bill_stand
    deducted_note[MIC_STAND_ID] = used_stand

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

        # 日別可否（マイク/拡声装置）
        allowed_today = True
        forced_zero = False
        if iid in MIC_RELATED_ITEM_IDS:
            allowed_today = _is_mic_related_item_allowed_today(iid, requires_context, mic_allowed_today)
            if not allowed_today:
                forced_zero = True

        # 課金数量（付属マイク・スタンド控除）
        if forced_zero:
            billed_qty = 0
        elif iid in MIC_ITEMS or iid in STAND_ITEMS:
            billed_qty = int(billed_qty_override.get(iid, orig_qty))
        else:
            billed_qty = orig_qty

        inherit = bool(meta.default_inherit_room_slot)
        base_slot = day_slot_default if inherit else global_fallback_slot

        if meta.allowed_slot_override and it.group_id in group_overrides:
            slot = group_overrides[it.group_id]
        else:
            slot = base_slot

        mult = slot_to_multiplier(slot)

        is_slot_item = it.price_per_slot > 0
        is_once_item = (it.price_per_slot == 0) and (it.price_once_yen > 0)
        # 電源使用料は区分数にかかわらず1日単位（price_per_slot を1日あたりの単価として使う）
        is_daily_item = is_slot_item and bool(it.is_power_item)

        if is_daily_item:
            mult = 1 if mult > 0 else 0
            per_slot_sub = it.price_per_slot * billed_qty * mult
            once_sub = it.price_once_yen * billed_qty
            subtotal = per_slot_sub + once_sub
            charge_type = "日額課金"
        elif is_slot_item:
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
            note = (note + " / " if note else "") + "マイク1本・スタンド1本付属"
        if is_daily_item:
            note = (note + " / " if note else "") + "1日単位"

        if iid in MIC_ITEMS:
            ded = int(deducted_note.get(iid, 0))
            if ded > 0:
                note = (note + " / " if note else "") + f"付属マイク控除:{ded}（有線→ワイヤレス）"
            note = (note + " / " if note else "") + f"選択:{orig_qty}→課金:{billed_qty}"

        if iid == MIC_STAND_ID:
            ded = int(deducted_note.get(iid, 0))
            if ded > 0:
                note = (note + " / " if note else "") + f"付属スタンド控除:{ded}"
            note = (note + " / " if note else "") + f"選択:{orig_qty}→課金:{billed_qty}"

        if iid in MIC_RELATED_ITEM_IDS and forced_zero:
            note = (note + " / " if note else "") + "対象外日（当日は計算対象外）"

        if billed_qty == 0 and orig_qty > 0 and (iid in MIC_RELATED_ITEM_IDS or iid in STAND_ITEMS):
            pass
        elif billed_qty == 0:
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

def extension_to_pricing(ext: str, slots_in_prices: Set[str]) -> Tuple[str, int, str]:
    if ext in slots_in_prices:
        return ext, 1, ""

    if ext == "前後延長30分":
        return "延長30分", 2, "前後延長30分（延長30分×2回）"
    if ext in ("前延長30分", "後延長30分"):
        return "延長30分", 1, f"{ext}（延長30分×1回）"

    return "延長30分", 1, ""

# =========================
# Internet
# =========================
INTERNET_POCKET_WIFI_PER_DAY = 2800
INTERNET_FIXED_FIRST_DAY = 18000
INTERNET_FIXED_AFTER_DAY = 2000
INTERNET_WIFI_FIRST_DAY = 25000
INTERNET_WIFI_AFTER_DAY = 3000
INTERNET_TEMP_LINE_BASE = 5000

INTERNET_NONE = "なし"
INTERNET_WIRED = "有線LAN"
INTERNET_WIFI = "Wi-Fi"
INTERNET_WIRED_ROOMS = {"大集会室", "中集会室", "小集会室", "特別室"}
INTERNET_WIFI_ROOMS = {"大集会室", "中集会室", "小集会室", "特別室", "大会議室"}
FLOOR_1_ROOMS = {"大集会室"}
FLOOR_3_ROOMS = {"中集会室", "小集会室"}

# =========================
# Day settings
# =========================
def make_days_base(days: List[pd.Timestamp], closed_days: set, default_room_slot: str, is_business_default: bool) -> pd.DataFrame:
    rows = []
    for d in days:
        rows.append(
            {
                "日付": d.strftime(DATE_FMT),
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
    df = df.copy()

    for i in range(len(df)):
        ts = parse_date_str(df.loc[i, "日付"])
        if ts is None:
            continue
        df.loc[i, "土日祝"] = "土日祝" if is_weekend_or_holiday(ts) else "平日"
        df.loc[i, "祝日名"] = holiday_name(ts)

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
# Room-Day table
# =========================
def _day_business_map(days_df: pd.DataFrame) -> Dict[str, bool]:
    m = {}
    for _, r in days_df.iterrows():
        m[normalize_str(r["日付"])] = bool(r.get("割増利用", False))
    return m

ROOM_DAY_COLUMNS = [
    "日付", "土日祝", "祝日名", "休館日", "部屋", "区分", "延長", "割増利用", "手動区分", "手動延長", "手動割増",
]

def build_room_day_base(
    days_df: pd.DataFrame,
    selected_rooms: List[str],
    default_room_slot: str,
    default_room_extension: str = "なし",
) -> pd.DataFrame:
    rows = []
    day_business = _day_business_map(days_df)

    for _, drow in days_df.iterrows():
        date_str = normalize_str(drow["日付"])
        ts = parse_date_str(date_str)
        if ts is None:
            continue
        for room in selected_rooms:
            rows.append(
                {
                    "日付": date_str,
                    "土日祝": "土日祝" if is_weekend_or_holiday(ts) else "平日",
                    "祝日名": holiday_name(ts),
                    "休館日": bool(drow.get("休館日", False)),
                    "部屋": room,
                    "区分": default_room_slot,
                    "延長": _fix_room_extension(default_room_extension),
                    "割増利用": bool(day_business.get(date_str, False)),
                    "手動区分": False,
                    "手動延長": False,
                    "手動割増": False,
                }
            )
    df = pd.DataFrame(rows, columns=ROOM_DAY_COLUMNS)
    if not df.empty:
        df["区分"] = df["区分"].apply(lambda x: _fix_room_slot(x, default_room_slot))
        df["延長"] = df["延長"].apply(_fix_room_extension)
        df["割増利用"] = df["割増利用"].astype(bool)
    return df

def merge_room_day(
    current: pd.DataFrame,
    days_df: pd.DataFrame,
    selected_rooms: List[str],
    default_room_slot: str,
    default_room_extension: str = "なし",
) -> pd.DataFrame:
    base = build_room_day_base(days_df, selected_rooms, default_room_slot, default_room_extension)
    if current is None or current.empty:
        return base

    cur = current.copy()
    for c in ["日付", "部屋"]:
        if c in cur.columns:
            cur[c] = cur[c].map(normalize_str)

    if "手動区分" not in cur.columns:
        cur["手動区分"] = True
    if "手動割増" not in cur.columns:
        cur["手動割増"] = True
    if "延長" not in cur.columns:
        cur["延長"] = "なし"
    if "手動延長" not in cur.columns:
        cur["手動延長"] = False
    if "割増利用" not in cur.columns:
        cur["割増利用"] = False
    if "区分" not in cur.columns:
        cur["区分"] = default_room_slot

    current_map = {}
    for _, row in cur.iterrows():
        key = (normalize_str(row.get("日付", "")), normalize_str(row.get("部屋", "")))
        current_map[key] = {
            "区分": _fix_room_slot(row.get("区分", ""), default_room_slot),
            "延長": _fix_room_extension(row.get("延長", "なし")),
            "割増利用": bool(row.get("割増利用", False)),
            "手動区分": bool(row.get("手動区分", False)),
            "手動延長": bool(row.get("手動延長", False)),
            "手動割増": bool(row.get("手動割増", False)),
        }

    rows = []
    for _, row in base.iterrows():
        out = row.to_dict()
        key = (normalize_str(out.get("日付", "")), normalize_str(out.get("部屋", "")))
        old = current_map.get(key)

        if old is not None:
            out["手動区分"] = old["手動区分"]
            out["手動延長"] = old["手動延長"]
            out["手動割増"] = old["手動割増"]

            if old["手動区分"]:
                out["区分"] = old["区分"]
            if old["手動延長"]:
                out["延長"] = old["延長"]
            if old["手動割増"]:
                out["割増利用"] = old["割増利用"]

        rows.append(out)

    merged = pd.DataFrame(rows)

    if not merged.empty:
        merged["区分"] = merged["区分"].apply(lambda x: _fix_room_slot(x, default_room_slot))
        merged["延長"] = merged["延長"].apply(_fix_room_extension)
        merged["割増利用"] = merged["割増利用"].astype(bool)
        merged["手動区分"] = merged["手動区分"].astype(bool)
        merged["手動延長"] = merged["手動延長"].astype(bool)
        merged["手動割増"] = merged["手動割増"].astype(bool)
    return merged
def apply_room_day_edits(full_df: pd.DataFrame, edited_subset: pd.DataFrame, default_room_slot: str) -> pd.DataFrame:
    if full_df is None or full_df.empty or edited_subset is None or edited_subset.empty:
        return full_df

    full = full_df.copy()
    for c in ["日付", "部屋"]:
        full[c] = full[c].map(normalize_str)

    sub = edited_subset.copy()
    for c in ["日付", "部屋"]:
        sub[c] = sub[c].map(normalize_str)

    if "延長" not in full.columns:
        full["延長"] = "なし"
    if "手動延長" not in full.columns:
        full["手動延長"] = False

    full_idx = {(r["日付"], r["部屋"]): i for i, r in full.iterrows()}

    for _, r in sub.iterrows():
        key = (r["日付"], r["部屋"])
        if key not in full_idx:
            continue
        i = full_idx[key]

        new_slot = _fix_room_slot(r.get("区分", ""), default_room_slot)
        new_ext = _fix_room_extension(r.get("延長", "なし"))
        new_bus = bool(r.get("割増利用", False))

        old_slot = normalize_str(full.loc[i, "区分"])
        old_ext = _fix_room_extension(full.loc[i, "延長"])
        old_bus = bool(full.loc[i, "割増利用"])

        if new_slot != old_slot:
            full.loc[i, "手動区分"] = True
        if new_ext != old_ext:
            full.loc[i, "手動延長"] = True
        if new_bus != old_bus:
            full.loc[i, "手動割増"] = True

        full.loc[i, "区分"] = new_slot
        full.loc[i, "延長"] = new_ext
        full.loc[i, "割増利用"] = new_bus

    full["割増利用"] = full["割増利用"].astype(bool)
    full["手動区分"] = full["手動区分"].astype(bool)
    full["手動延長"] = full["手動延長"].astype(bool)
    full["手動割増"] = full["手動割増"].astype(bool)
    return full

# =========================
# 計算（部屋）
# =========================
def calc_rooms_from_room_day(prices_df: pd.DataFrame, room_day_df: pd.DataFrame) -> Tuple[int, pd.DataFrame]:
    if room_day_df is None or room_day_df.empty:
        return 0, pd.DataFrame(columns=["日付", "種別", "品目", "区分", "割増", "単価", "小計", "備考"])

    slots_in_prices = set(prices_df["slot"].unique().tolist())

    rows = []
    total = 0

    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)):
            continue

        date_str = normalize_str(r.get("日付", ""))
        room = normalize_str(r.get("部屋", ""))
        slot = normalize_str(r.get("区分", ""))
        ext = _fix_room_extension(r.get("延長", "なし"))
        is_business = bool(r.get("割増利用", False))

        dts = parse_date_str(date_str)
        if dts is None:
            continue

        day_type = "土日祝" if is_weekend_or_holiday(dts) else "平日"
        price_type = "割増" if is_business else "通常"

        if slot == "利用なし":
            note = ""
            if ext != "なし":
                note = "部屋が「利用なし」のため延長は無視されます"
            rows.append(
                {
                    "日付": dts.date(),
                    "種別": "部屋",
                    "品目": room,
                    "区分": "利用なし",
                    "割増": is_business,
                    "単価": 0,
                    "小計": 0,
                    "備考": note,
                }
            )
            continue

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
        else:
            amount = int(hit.iloc[0]["amount"])
            total += amount
            rows.append(
                {
                    "日付": dts.date(),
                    "種別": "部屋",
                    "品目": room,
                    "区分": slot,
                    "割増": is_business,
                    "単価": amount,
                    "小計": amount,
                    "備考": "",
                }
            )

        if ext != "なし" and is_invalid_after_extension(slot, ext):
            rows.append(
                {
                    "日付": dts.date(),
                    "種別": "部屋",
                    "品目": f"{room}（延長）",
                    "区分": ext,
                    "割増": is_business,
                    "単価": None,
                    "小計": None,
                    "備考": INVALID_AFTER_EXTENSION_NOTE,
                }
            )
        elif ext != "なし":
            pricing_slot, mult, note = extension_to_pricing(ext, slots_in_prices)

            m2 = (
                (prices_df["room"] == room)
                & (prices_df["day_type"] == day_type)
                & (prices_df["price_type"] == price_type)
                & (prices_df["slot"] == pricing_slot)
            )
            hit2 = prices_df[m2]
            if hit2.empty:
                rows.append(
                    {
                        "日付": dts.date(),
                        "種別": "部屋",
                        "品目": f"{room}（延長）",
                        "区分": ext,
                        "割増": is_business,
                        "単価": None,
                        "小計": None,
                        "備考": f"延長の料金が prices.csv に見つかりません（参照slot={pricing_slot}）",
                    }
                )
            else:
                unit_amount = int(hit2.iloc[0]["amount"])
                sub = unit_amount * mult
                total += sub
                rows.append(
                    {
                        "日付": dts.date(),
                        "種別": "部屋",
                        "品目": f"{room}（延長）",
                        "区分": ext,
                        "割増": is_business,
                        "単価": unit_amount,
                        "小計": sub,
                        "備考": note,
                    }
                )

    df = pd.DataFrame(rows)
    return total, df

# =========================
# 有料差額（通常料金 → 割増料金）
# =========================
PREMIUM_DIFF_COLUMNS = ["対象", "日付", "部屋", "内訳", "区分", "通常料金", "割増料金", "差額", "備考"]
PREMIUM_DIFF_ROOM = "部屋代"
PREMIUM_DIFF_EXTENSION = "延長"

def calc_premium_difference_rows(prices_df: pd.DataFrame, room_day_df: pd.DataFrame) -> pd.DataFrame:
    """利用する部屋×日ごとに、通常料金と割増料金とその差額を求める。

    部屋代と延長は別の行に分ける（延長がある日は「部屋代」行の下に「延長」行を出す）。
    部屋×日テーブルの「割増利用」の値にかかわらず、同じ区分・延長で両方の料金を計算する。
    """
    if room_day_df is None or room_day_df.empty:
        return pd.DataFrame(columns=PREMIUM_DIFF_COLUMNS)

    rows = []
    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)) or normalize_str(r.get("区分", "")) == "利用なし":
            continue

        room = normalize_str(r.get("部屋", ""))
        ext = _fix_room_extension(r.get("延長", "なし"))
        # 内訳ごとの {料金ラベル: 金額} と備考
        parts = {PREMIUM_DIFF_ROOM: ({}, []), PREMIUM_DIFF_EXTENSION: ({}, [])}

        one = pd.DataFrame([r.to_dict()])
        for label, is_business in (("通常料金", False), ("割増料金", True)):
            one["割増利用"] = is_business
            _, detail = calc_rooms_from_room_day(prices_df, one)
            for kind, (amounts, notes) in parts.items():
                item = room if kind == PREMIUM_DIFF_ROOM else f"{room}（延長）"
                d = detail[detail["品目"] == item]
                amounts[label] = int(pd.to_numeric(d["小計"], errors="coerce").fillna(0).sum())
                # 料金が見つからない等で計算できなかった行は備考に残す
                for _, x in d.iterrows():
                    note = normalize_str(x.get("備考", ""))
                    if pd.isna(x.get("小計")) and note and note not in notes:
                        notes.append(note)

        for kind, (amounts, notes) in parts.items():
            if kind == PREMIUM_DIFF_EXTENSION and ext == "なし":
                continue
            rows.append(
                {
                    "対象": True,
                    "日付": normalize_str(r.get("日付", "")),
                    "部屋": room,
                    "内訳": kind,
                    "区分": normalize_str(r.get("区分", "")) if kind == PREMIUM_DIFF_ROOM else ext,
                    "通常料金": amounts["通常料金"],
                    "割増料金": amounts["割増料金"],
                    "差額": amounts["割増料金"] - amounts["通常料金"],
                    "備考": " / ".join(notes),
                }
            )
    return pd.DataFrame(rows, columns=PREMIUM_DIFF_COLUMNS)

# =========================
# 全日差額（一部の区分 → 全日）
# =========================
ALLDAY_SLOT = "全日"
ALLDAY_FROM_SLOTS = ["午前", "午後", "夜間", "午前-午後", "午後-夜間"]
# 全日にするために追加で必要な区分（全日を使わず個別に追加した場合の参考計算用）
ALLDAY_MISSING_SLOTS = {
    "午前": ["午後-夜間"],
    "午後": ["午前", "夜間"],
    "夜間": ["午前-午後"],
    "午前-午後": ["夜間"],
    "午後-夜間": ["午前"],
}
ALLDAY_DIFF_COLUMNS = [
    "対象", "日付", "部屋", "割増", "変更前区分", "変更前料金", "全日料金", "差額", "参考：個別追加", "備考",
]

def lookup_room_price(prices_df: pd.DataFrame, room: str, date_str: str, is_business: bool, slot: str) -> Optional[int]:
    """prices.csv から部屋の基本料金を引く（見つからなければ None）。"""
    dts = parse_date_str(date_str)
    if dts is None:
        return None
    day_type = "土日祝" if is_weekend_or_holiday(dts) else "平日"
    price_type = "割増" if is_business else "通常"
    hit = prices_df[
        (prices_df["room"] == room)
        & (prices_df["day_type"] == day_type)
        & (prices_df["price_type"] == price_type)
        & (prices_df["slot"] == slot)
    ]
    if hit.empty:
        return None
    return int(hit.iloc[0]["amount"])

def allday_difference_base_rows(room_day_df: pd.DataFrame) -> pd.DataFrame:
    """全日差額の入力行（部屋×日ごと）。変更前区分は部屋×日テーブルの区分を初期値にする。

    すでに「全日」になっている行は変更前区分が分からないため、「午前-午後」を仮に入れて対象外にしておく。
    """
    rows = []
    if room_day_df is not None and not room_day_df.empty:
        for _, r in room_day_df.iterrows():
            slot = normalize_str(r.get("区分", ""))
            if bool(r.get("休館日", False)) or slot == "利用なし":
                continue
            known = slot in ALLDAY_FROM_SLOTS
            rows.append(
                {
                    "対象": known,
                    "日付": normalize_str(r.get("日付", "")),
                    "部屋": normalize_str(r.get("部屋", "")),
                    "割増": bool(r.get("割増利用", False)),
                    "変更前区分": slot if known else "午前-午後",
                }
            )
    return pd.DataFrame(rows, columns=["対象", "日付", "部屋", "割増", "変更前区分"])

def allday_separate_label(from_slot: str) -> str:
    """参考列の見出し（例：変更前が午前-午後 → 「参考：夜間1区分」）。"""
    missing = ALLDAY_MISSING_SLOTS.get(from_slot, [])
    if not missing:
        return "参考：個別追加"
    return f"参考：{'・'.join(missing)}{len(missing)}区分"

def split_separate_by_label(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """「参考：個別追加」列を、追加される区分ごとの列（例：「参考：夜間1区分」）に分ける。"""
    out = df.copy()
    labels = out["変更前区分"].map(allday_separate_label)
    order = [allday_separate_label(s) for s in ALLDAY_FROM_SLOTS]
    new_cols = [l for l in dict.fromkeys(order) if (labels == l).any()]
    pos = out.columns.get_loc("参考：個別追加")
    for i, l in enumerate(new_cols):
        out.insert(pos + i, l, out["参考：個別追加"].where(labels == l))
    out = out.drop(columns=["参考：個別追加"])
    return out, new_cols

def calc_allday_difference(prices_df: pd.DataFrame, base_rows: pd.DataFrame) -> pd.DataFrame:
    """変更前区分の料金と全日料金の差額（基本料金のみ。延長は含めない）を求める。"""
    rows = []
    for _, r in base_rows.iterrows():
        date_str = normalize_str(r["日付"])
        room = normalize_str(r["部屋"])
        is_business = bool(r["割増"])
        from_slot = normalize_str(r["変更前区分"])

        before = lookup_room_price(prices_df, room, date_str, is_business, from_slot)
        allday = lookup_room_price(prices_df, room, date_str, is_business, ALLDAY_SLOT)
        missing = [lookup_room_price(prices_df, room, date_str, is_business, s) for s in ALLDAY_MISSING_SLOTS.get(from_slot, [])]

        notes = []
        if before is None:
            notes.append(f"{from_slot}の料金が prices.csv に見つかりません")
        if allday is None:
            notes.append("全日の料金が prices.csv に見つかりません")
        separate = None if (not missing or any(m is None for m in missing)) else sum(missing)

        rows.append(
            {
                "対象": bool(r["対象"]),
                "日付": date_str,
                "部屋": room,
                "割増": is_business,
                "変更前区分": from_slot,
                "変更前料金": before,
                "全日料金": allday,
                "差額": None if (before is None or allday is None) else allday - before,
                "参考：個別追加": separate,
                "備考": " / ".join(notes),
            }
        )
    return pd.DataFrame(rows, columns=ALLDAY_DIFF_COLUMNS)

def invalid_after_extension_rows(room_day_df: pd.DataFrame) -> pd.DataFrame:
    cols = ["日付", "部屋", "区分", "延長"]
    if room_day_df is None or room_day_df.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)):
            continue
        slot = normalize_str(r.get("区分", ""))
        ext = _fix_room_extension(r.get("延長", "なし"))
        if is_invalid_after_extension(slot, ext):
            rows.append(
                {
                    "日付": normalize_str(r.get("日付", "")),
                    "部屋": normalize_str(r.get("部屋", "")),
                    "区分": slot,
                    "延長": ext,
                }
            )
    return pd.DataFrame(rows, columns=cols)

# =========================
# 計算用：日ごとの使用部屋を集計
# =========================
def rooms_used_by_date(room_day_df: pd.DataFrame) -> Dict[str, Set[str]]:
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

# =========================
# 計算（設備：全日合算）
# =========================
def calc_equipment_total_all_days(
    days_df: pd.DataFrame,
    room_day_df: pd.DataFrame,
    global_default_slot: str,
    group_overrides: Dict[str, str],
    base_selections: List[Dict],
    items: Dict[str, EquipmentItem],
    group_meta: Dict[str, GroupMeta],
    gallery_678: bool,
) -> Tuple[int, pd.DataFrame]:
    active_dates = active_dates_from_room_day(room_day_df)
    if not active_dates:
        return 0, pd.DataFrame(
            columns=[
                "日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量",
                "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"
            ]
        )

    day_slot_map = {normalize_str(r["日付"]): _fix_equip_cell(r.get("設備デフォ区分", "")) for _, r in days_df.iterrows()}
    used_rooms = rooms_used_by_date(room_day_df)

    all_rows = []
    total = 0

    for d in active_dates:
        day_slot_default = day_slot_map.get(d, global_default_slot)
        rooms_today = used_rooms.get(d, set())

        need_c = bool(rooms_today & MIC_C_ROOMS)
        need_d = bool(MIC_D_ROOMS.issubset(rooms_today) and gallery_678)
        mic_allowed_today, _reason = infer_mic_allowed_for_rooms(rooms_today, gallery_678)

        requires_ctx = {
            "need_pa_c": need_c,
            "need_pa_d": need_d,
        }

        day_selections = [dict(x) for x in base_selections]

        day_total, day_df = calc_equipment_total_for_day(
            day_slot_default=day_slot_default,
            global_fallback_slot=global_default_slot,
            group_overrides=group_overrides,
            selections=day_selections,
            items=items,
            group_meta=group_meta,
            requires_context=requires_ctx,
            mic_allowed_today=mic_allowed_today,
        )
        total += day_total

        if not day_df.empty:
            ts = parse_date_str(d)
            day_df.insert(0, "日付", ts.date() if ts is not None else d)
            all_rows.append(day_df)

    if all_rows:
        out = pd.concat(all_rows, ignore_index=True)
        out = out[
            [
                "日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量",
                "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"
            ]
        ]
        return total, out

    return 0, pd.DataFrame(
        columns=[
            "日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量",
            "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"
        ]
    )

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
            ts = parse_date_str(d)
            df.insert(0, "日付", ts.date() if ts is not None else d)
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
        dt = parse_date_str(d)
        if dt is None:
            continue
        if rooms & FLOOR_1_ROOMS:
            f1.append(dt)
        if rooms & FLOOR_3_ROOMS:
            f3.append(dt)
    return {"1F": sorted(f1), "3F": sorted(f3)}

def available_fixed_network_services(room: str) -> List[str]:
    room = normalize_str(room)
    options = [INTERNET_NONE]
    if room in INTERNET_WIRED_ROOMS:
        options.append(INTERNET_WIRED)
    if room in INTERNET_WIFI_ROOMS:
        options.append(INTERNET_WIFI)
    return options


def active_room_day_records(room_day_df: pd.DataFrame) -> List[Dict[str, object]]:
    records: List[Dict[str, object]] = []
    if room_day_df is None or room_day_df.empty:
        return records

    for _, r in room_day_df.iterrows():
        if bool(r.get("休館日", False)):
            continue
        date_str = normalize_str(r.get("日付", ""))
        room = normalize_str(r.get("部屋", ""))
        slot = normalize_str(r.get("区分", ""))
        ts = parse_date_str(date_str)
        if ts is None or room == "" or slot == "利用なし":
            continue
        records.append({"date_str": date_str, "date": ts, "room": room})
    return records


def _internet_staged_price(service: str, is_first_day: bool) -> int:
    if service == INTERNET_WIRED:
        return INTERNET_FIXED_FIRST_DAY if is_first_day else INTERNET_FIXED_AFTER_DAY
    if service == INTERNET_WIFI:
        return INTERNET_WIFI_FIRST_DAY if is_first_day else INTERNET_WIFI_AFTER_DAY
    return 0


def _internet_service_item_name(service: str, is_first_day: bool) -> str:
    suffix = "初日" if is_first_day else "2日目以降"
    if service == INTERNET_WIRED:
        return f"常設回線（有線LAN・{suffix}）"
    if service == INTERNET_WIFI:
        return f"Wi-Fi（{suffix}）"
    return service


def calc_internet_total(
    room_day_df: pd.DataFrame,
    fixed_network_selections: Dict[Tuple[str, str], str],
    use_pocket_wifi: bool,
    use_temp_line: bool,
) -> Tuple[int, pd.DataFrame]:
    active_records = active_room_day_records(room_day_df)
    active_dates = sorted({r["date"] for r in active_records})
    if not active_dates:
        return 0, pd.DataFrame(columns=["日付", "種別", "品目", "対象", "小計", "備考"])

    rows = []
    total = 0

    if use_pocket_wifi:
        for d in active_dates:
            rows.append({"日付": d.date(), "種別": "インターネット", "品目": "ポケットWi-Fi貸出", "対象": "全部屋", "小計": INTERNET_POCKET_WIFI_PER_DAY, "備考": "先着順/同時接続目安5台/電波不安定の可能性"})
            total += INTERNET_POCKET_WIFI_PER_DAY

    selected_by_room_service: Dict[Tuple[str, str], List[pd.Timestamp]] = {}
    for rec in active_records:
        date_str = str(rec["date_str"])
        room = str(rec["room"])
        service = fixed_network_selections.get((date_str, room), INTERNET_NONE)
        if service == INTERNET_NONE:
            continue
        if service not in available_fixed_network_services(room):
            continue
        selected_by_room_service.setdefault((room, service), []).append(rec["date"])

    for (room, service), ds in selected_by_room_service.items():
        blocks = _split_consecutive_blocks(ds)
        for b in blocks:
            for idx, d in enumerate(b):
                is_first_day = idx == 0
                price = _internet_staged_price(service, is_first_day)
                if price <= 0:
                    continue
                rows.append({"日付": d.date(), "種別": "インターネット", "品目": _internet_service_item_name(service, is_first_day), "対象": room, "小計": price, "備考": "連続利用の段階料金"})
                total += price

    if use_temp_line:
        floors = infer_active_days_by_floor(room_day_df)
        for floor_label, ds in floors.items():
            blocks = _split_consecutive_blocks(ds)
            for b in blocks:
                if not b:
                    continue
                rows.append({"日付": b[0].date(), "種別": "インターネット", "品目": "仮設回線（開通工事）", "対象": floor_label, "小計": INTERNET_TEMP_LINE_BASE, "備考": "＋別途お見積り（NTT回線開通工事）"})
                total += INTERNET_TEMP_LINE_BASE

    df = pd.DataFrame(rows, columns=["日付", "種別", "品目", "対象", "小計", "備考"])
    return total, df
# =========================
# KPI Display
# =========================
def build_all_details_df(
    room_df: pd.DataFrame,
    equipment_df: pd.DataFrame,
    tech_df: pd.DataFrame,
    internet_df: pd.DataFrame,
) -> pd.DataFrame:
    frames = []
    if room_df is not None and not room_df.empty:
        r = room_df.copy()
        r = r.rename(columns={"品目": "名称"})
        r["カテゴリ"] = "部屋"
        frames.append(r[["日付", "カテゴリ", "名称", "区分", "小計", "備考"]])

    if equipment_df is not None and not equipment_df.empty:
        e = equipment_df.copy()
        e["カテゴリ"] = "設備"
        e = e.rename(columns={"品目": "名称"})
        frames.append(e[["日付", "カテゴリ", "名称", "区分", "小計", "備考"]])

    if tech_df is not None and not tech_df.empty:
        t = tech_df.copy()
        t["カテゴリ"] = "技術者"
        t["名称"] = "舞台設備技術者"
        frames.append(t[["日付", "カテゴリ", "名称", "区分", "小計"]])

    if internet_df is not None and not internet_df.empty:
        n = internet_df.copy()
        n["カテゴリ"] = "インターネット"
        n = n.rename(columns={"品目": "名称"})
        frames.append(n[["日付", "カテゴリ", "名称", "対象", "小計", "備考"]])

    if not frames:
        return pd.DataFrame(columns=["日付", "カテゴリ", "名称", "区分", "小計", "備考"])
    out = pd.concat(frames, ignore_index=True)
    out = add_date_labels(out)
    out["小計"] = pd.to_numeric(out["小計"], errors="coerce").round().astype("Int64")
    return out

# 延長料金のみ明細に出す行：部屋の前・後延長と、設備の区分「延長30分」（技術者の延長は含めない）
EXTENSION_DETAIL_CATEGORIES = ("部屋", "設備")
EXTENSION_DETAIL_SLOTS = {s for s in ROOM_EXTENSION_SLOTS if s != "なし"} | {"延長30分"}
EXTENSION_DETAIL_COLUMNS = ["日付", "祝日", "カテゴリ", "名称", "区分", "小計", "備考"]

def extension_detail_rows(all_df: pd.DataFrame) -> pd.DataFrame:
    """明細（全部）から延長料金の行だけを取り出す。"""
    if all_df is None or all_df.empty or "区分" not in all_df.columns:
        return pd.DataFrame(columns=EXTENSION_DETAIL_COLUMNS)
    m = all_df["カテゴリ"].isin(EXTENSION_DETAIL_CATEGORIES) & all_df["区分"].isin(EXTENSION_DETAIL_SLOTS)
    cols = [c for c in EXTENSION_DETAIL_COLUMNS if c in all_df.columns]
    return all_df.loc[m, cols].reset_index(drop=True)

WEEKDAYS_JA = "月火水木金土日"

def format_date_label(v: object) -> str:
    """日付を「2026/11/23（月）」の形にする。日付として読めない値はそのまま返す。"""
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return "" if v is None else str(v)
    return f"{ts.strftime(DATE_FMT)}（{WEEKDAYS_JA[ts.weekday()]}）"

def holiday_label(v: object) -> str:
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return ""
    return holiday_name(ts)

def add_date_labels(df: pd.DataFrame) -> pd.DataFrame:
    """明細の「日付」に曜日を付け、その右に「祝日」列を追加する（何度呼んでも同じ結果）。"""
    if df is None or df.empty or "日付" not in df.columns or "祝日" in df.columns:
        return df
    out = df.copy()
    raw = out["日付"]
    out["日付"] = raw.map(format_date_label)
    out.insert(out.columns.get_loc("日付") + 1, "祝日", raw.map(holiday_label))
    return out

MONEY_COLUMNS = ("単価", "小計")

def show_detail_df(df: pd.DataFrame) -> None:
    """明細の表示用：日付に曜日と祝日を付け、金額を ¥ とカンマ付きで表示する。"""
    if df is None or df.empty:
        st.info("明細がありません。")
        return
    view = add_date_labels(df)
    column_config = {}
    for c in MONEY_COLUMNS:
        if c in view.columns:
            view[c] = pd.to_numeric(view[c], errors="coerce")
            column_config[c] = st.column_config.NumberColumn(format="yen")
    st.dataframe(view, width="stretch", hide_index=True, column_config=column_config)

def build_details_csv(all_df: pd.DataFrame) -> bytes:
    # Excel で文字化けしないよう BOM 付き UTF-8
    return all_df.to_csv(index=False).encode("utf-8-sig")

def _pdf_cell(v: object) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)

def build_estimate_pdf(
    title: str,
    period: str,
    rooms: List[str],
    totals: Dict[str, int],
    all_df: pd.DataFrame,
) -> bytes:
    from io import BytesIO

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    font = "HeiseiKakuGo-W5"
    if font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(font))

    h1 = ParagraphStyle("h1", fontName=font, fontSize=16, leading=22)
    body = ParagraphStyle("body", fontName=font, fontSize=9.5, leading=14)
    cell = ParagraphStyle("cell", fontName=font, fontSize=8, leading=10.5)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm
    )
    story = [
        Paragraph(f"{title}　概算見積", h1),
        Spacer(1, 4 * mm),
        Paragraph(f"作成日：{pd.Timestamp.today().strftime(DATE_FMT)}", body),
        Paragraph(f"期間：{period}", body),
        Paragraph(f"部屋：{'、'.join(rooms)}", body),
        Spacer(1, 4 * mm),
    ]

    grand = sum(int(v) for v in totals.values())
    total_rows = [["項目", "金額"]] + [[k, yen(int(v))] for k, v in totals.items()] + [["総額", yen(grand)]]
    tt = Table(total_rows, colWidths=[50 * mm, 40 * mm])
    tt.hAlign = "LEFT"
    tt.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
            ]
        )
    )
    story += [tt, Spacer(1, 6 * mm), Paragraph("明細", body), Spacer(1, 2 * mm)]

    detail_cols = ["日付", "祝日", "カテゴリ", "名称", "区分", "対象", "小計", "備考"]
    cols = [c for c in detail_cols if c in all_df.columns]
    widths = {"日付": 31, "祝日": 24, "カテゴリ": 18, "名称": 38, "区分": 18, "対象": 20, "小計": 20, "備考": 34}
    rows = [cols]
    for _, r in all_df.iterrows():
        row = []
        for c in cols:
            v = r.get(c)
            if c == "小計":
                txt = "" if _pdf_cell(v) == "" else yen(int(v))
            else:
                txt = _pdf_cell(v)
            row.append(Paragraph(txt, cell))
        rows.append(row)
    scale = (A4[0] - 28 * mm) / (sum(widths[c] for c in cols) * mm)
    dt_ = Table(rows, colWidths=[widths[c] * mm * scale for c in cols], repeatRows=1)
    dt_.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(dt_)
    doc.build(story)
    return buf.getvalue()

def build_difference_pdf(
    heading: str,
    period: str,
    rooms: List[str],
    summary: List[Tuple[str, int]],
    detail_df: pd.DataFrame,
    widths: Dict[str, float],
    money_cols: Tuple[str, ...],
    notes: Optional[List[str]] = None,
) -> bytes:
    """有料差額・全日差額の明細PDF（見積PDFと同じ体裁）。summary の最後の行を強調する。"""
    from io import BytesIO

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    font = "HeiseiKakuGo-W5"
    if font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(UnicodeCIDFont(font))

    h1 = ParagraphStyle("h1", fontName=font, fontSize=16, leading=22)
    body = ParagraphStyle("body", fontName=font, fontSize=9.5, leading=14)
    small = ParagraphStyle("small", fontName=font, fontSize=8, leading=11, textColor=colors.HexColor("#555555"))
    cell = ParagraphStyle("cell", fontName=font, fontSize=8, leading=10.5)
    cell_r = ParagraphStyle("cell_r", parent=cell, alignment=2)

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm
    )
    story = [
        Paragraph(f"料金電卓　{heading}", h1),
        Spacer(1, 4 * mm),
        Paragraph(f"作成日：{pd.Timestamp.today().strftime(DATE_FMT)}", body),
        Paragraph(f"期間：{period}", body),
        Paragraph(f"部屋：{'、'.join(rooms)}", body),
        Spacer(1, 4 * mm),
    ]

    total_rows = [["項目", "金額"]] + [[k, yen(int(v))] for k, v in summary]
    tt = Table(total_rows, colWidths=[60 * mm, 40 * mm])
    tt.hAlign = "LEFT"
    tt.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
                ("FONTSIZE", (0, -1), (-1, -1), 12),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
            ]
        )
    )
    story += [tt, Spacer(1, 3 * mm)]
    for n in notes or []:
        story.append(Paragraph(n, small))
    story += [Spacer(1, 5 * mm), Paragraph(f"明細（対象 {len(detail_df)}件）", body), Spacer(1, 2 * mm)]

    cols = [c for c in widths if c in detail_df.columns]
    head_r = ParagraphStyle("head_r", parent=cell, alignment=2)
    rows = [[Paragraph(c, head_r if c in money_cols else cell) for c in cols]]
    for _, r in detail_df.iterrows():
        row = []
        for c in cols:
            v = r.get(c)
            if c in money_cols:
                row.append(Paragraph("" if _pdf_cell(v) == "" else yen(int(v)), cell_r))
            elif pd.api.types.is_bool(v):
                row.append(Paragraph("○" if bool(v) else "", cell))
            else:
                row.append(Paragraph(_pdf_cell(v), cell))
        rows.append(row)
    # 金額列の合計行
    total_row = []
    for i, c in enumerate(cols):
        if c in money_cols:
            s = int(pd.to_numeric(detail_df[c], errors="coerce").fillna(0).sum())
            total_row.append(Paragraph(yen(s), cell_r))
        else:
            total_row.append(Paragraph("合計" if i == 0 else "", cell))
    rows.append(total_row)

    scale = (A4[0] - 28 * mm) / (sum(widths[c] for c in cols) * mm)
    dt_ = Table(rows, colWidths=[widths[c] * mm * scale for c in cols], repeatRows=1)
    dt_.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#f6f6f6")),
                ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(dt_)
    doc.build(story)
    return buf.getvalue()

def yen(x: int) -> str:
    try:
        return f"¥{int(x):,}"
    except Exception:
        return f"¥{x}"

def inject_ui_css():
    st.markdown(
        """
<style>
/* 開いているパネルの「閉じる」ボタンとパネル枠に色を付ける（全日差額＝緑、有料差額＝オレンジ、延長料金のみ明細＝青） */
div[class*="st-key-panel_open_"] button,
div[class*="st-key-panel_open_"] button:hover,
div[class*="st-key-panel_open_"] button:focus:not(:active) {
  border-color: var(--panel-color);
  background-color: var(--panel-color);
  color: #ffffff;
  font-weight: 700;
}
div[class*="st-key-panel_open_"] button:hover,
div[class*="st-key-panel_open_"] button:focus:not(:active) {
  filter: brightness(0.88);
}
div[class*="st-key-panel_open_"] button p {
  color: #ffffff;
  font-weight: 700;
}
div[class*="st-key-panel_body_"] {
  border: 2px solid var(--panel-color) !important;
  border-left-width: 6px !important;
}
div[class*="st-key-panel_open_allday"], div[class*="st-key-panel_body_allday"] {
  --panel-color: #1e9e4a;
}
div[class*="st-key-panel_open_premium"], div[class*="st-key-panel_body_premium"] {
  --panel-color: #e8750f;
}
div[class*="st-key-panel_open_detail_"], div[class*="st-key-panel_body_detail_"] {
  --panel-color: #2f6fdb;
}
/* 差額のボタン（【有料差額計算】【全日差額】）はスマホでも縦1列にせず、横2列のままにする */
.st-key-panel_toggles [data-testid="stColumn"] {
  min-width: 0 !important;
}
@media (max-width: 640px) {
  /* 「〜を閉じる」が「…」で切れないよう折り返し、2行になっても同じ段のボタンの高さをそろえる */
  .st-key-panel_toggles button {
    min-height: 3.6rem;
    padding-left: 0.5rem;
    padding-right: 0.5rem;
  }
  .st-key-panel_toggles button * {
    white-space: normal !important;
    text-overflow: clip !important;
    overflow: visible !important;
    word-break: keep-all;
    overflow-wrap: anywhere;
  }
}
:root{
  --oai-bg: var(--background-color, #ffffff);
  --oai-card-bg: var(--secondary-background-color, rgba(255,255,255,0.98));
  --oai-text: var(--text-color, #111111);
  --oai-border: rgba(49, 51, 63, 0.2);
}
.oai-sticky {
  position: sticky;
  top: 0;
  z-index: 999;
  background: var(--oai-card-bg);
  color: var(--oai-text);
  backdrop-filter: blur(6px);
  border-bottom: 1px solid var(--oai-border);
  padding: 12px 12px 10px 12px;
  margin: 0 0 12px 0;
}
.oai-sticky-inner {
  display: flex;
  gap: 14px;
  align-items: flex-end;
  justify-content: space-between;
  flex-wrap: wrap;
}
.oai-grand-label {
  font-size: 12px;
  opacity: 0.7;
  margin-bottom: 2px;
}
.oai-grand {
  font-size: 40px;
  font-weight: 800;
  line-height: 1.1;
  letter-spacing: 0.2px;
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
}
.oai-breakdown {
  font-size: 12px;
  opacity: 0.7;
  margin-top: 4px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 100%;
}
.oai-kpi-row {
  display: grid;
  /* 列の幅に合わせて折り返す（画面幅ではなく配置先の幅で決まる） */
  grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 12px;
  margin: 8px 0 8px 0;
}
.oai-kpi-row .oai-kpi-card.total { grid-column: 1 / -1; }
.oai-kpi-card {
  border: 1px solid var(--oai-border);
  border-radius: 12px;
  background: var(--oai-card-bg);
  color: var(--oai-text);
  padding: 10px 12px;
  box-shadow: 0 1px 0 rgba(0,0,0,0.03);
  overflow: hidden;
}
.oai-kpi-label {
  font-size: 12px;
  opacity: 0.75;
  margin-bottom: 4px;
  white-space: nowrap;
}
.oai-kpi-val {
  font-size: 26px;
  font-weight: 800;
  line-height: 1.15;
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
}
.oai-kpi-sub {
  font-size: 12px;
  opacity: 0.65;
  margin-top: 4px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.oai-kpi-card.total { border-width: 2px; }
@media (max-width: 640px) {
  .oai-grand { font-size: 32px; }
  .oai-kpi-row { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .oai-kpi-val { font-size: 22px; }
}
</style>
        """,
        unsafe_allow_html=True,
    )

LAST_TOTALS_KEY = "last_totals"

def render_totals_sticky(room_total: int, equipment_total: int, tech_total: int, internet_total: int):
    grand_total = room_total + equipment_total + tech_total + internet_total

    grand = yen(grand_total)
    r = yen(room_total)
    e = yen(equipment_total)
    t = yen(tech_total)
    n = yen(internet_total)

    st.markdown(
        f"""
<div class="oai-sticky">
  <div class="oai-sticky-inner">
    <div>
      <div class="oai-grand-label">総額</div>
      <div class="oai-grand">{grand}</div>
      <div class="oai-breakdown">内訳：部屋 {r} / 設備 {e} / 技術者 {t} / インターネット {n}</div>
    </div>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )

def render_kpis_cards(room_total: int, equipment_total: int, tech_total: int, internet_total: int):
    grand_total = room_total + equipment_total + tech_total + internet_total
    cards = [
        ("総額", yen(grand_total), "", "total"),
        ("部屋", yen(room_total), ""),
        ("設備", yen(equipment_total), ""),
        ("技術者", yen(tech_total), ""),
        ("インターネット", yen(internet_total), ""),
    ]

    html = ['<div class="oai-kpi-row">']
    for c in cards:
        if len(c) == 4:
            label, val, sub, cls = c
        else:
            label, val, sub = c
            cls = ""
        extra_cls = f" {cls}" if cls else ""
        html.append(
            f"""
<div class="oai-kpi-card{extra_cls}">
  <div class="oai-kpi-label">{label}</div>
  <div class="oai-kpi-val">{val}</div>
  <div class="oai-kpi-sub">{sub}</div>
</div>
            """
        )
    html.append("</div>")
    st.markdown("\n".join(html), unsafe_allow_html=True)

PREMIUM_DIFF_OPEN_KEY = "premium_diff_open"
ALLDAY_DIFF_OPEN_KEY = "allday_diff_open"
# 「detail_」で始めると inject_ui_css の .st-key-panel_*_detail_* で青色になる
EXTENSION_DETAIL_OPEN_KEY = "detail_extension_open"

def render_panel_toggle(container, label: str, open_key: str, help_text: str) -> None:
    """押すたびにパネルの開閉を切り替えるボタン。"""
    is_open = bool(st.session_state.get(open_key, False))
    if container.button(
        # \u200b（幅ゼロの空白）は、スマホで2行になるとき「を閉じる」の前で改行させるため
        f"▼ {label}\u200bを閉じる" if is_open else label,
        # 開いているときはキー名で色付きのスタイル（inject_ui_css の .st-key-panel_open_*）を当てる
        key=f"panel_open_{open_key}" if is_open else f"{open_key}_toggle",
        help=help_text,
        width="stretch",
    ):
        st.session_state[open_key] = not is_open
        st.rerun()

def render_premium_difference(
    prices_df: pd.DataFrame,
    room_day_df: pd.DataFrame,
    room_day_key: str,
    period_label: str,
    rooms: List[str],
    date_suffix: str,
) -> None:
    """【有料差額計算】ボタンで開く、通常料金→割増料金の差額計算パネル。"""
    if not st.session_state.get(PREMIUM_DIFF_OPEN_KEY, False):
        return

    with st.container(border=True, key="panel_body_premium"):
        st.markdown("#### 有料差額計算（通常料金 → 割増料金）")
        st.caption(
            "部屋×日テーブルの区分・延長で、通常料金と割増料金の両方を計算し差額を出します（部屋代と延長は別の行に分けて表示）。"
            "割増に変わる行だけ「対象」にチェックしてください（設備・技術者・インターネットは割増で変わらないため対象外）。"
        )

        diff_df = calc_premium_difference_rows(prices_df, room_day_df)
        if diff_df.empty:
            st.info("差額を計算できる部屋の利用がありません（休館日・「利用なし」は対象外です）。")
            return

        # 行の構成が変わったら（部屋・日付・区分・延長の変更）チェック状態をリセットする
        signature = "|".join(
            diff_df[["日付", "部屋", "内訳", "区分"]].astype(str).agg(",".join, axis=1).tolist()
        )
        editor_key = f"premium_diff_editor_{room_day_key}_{abs(hash(signature))}"

        b1, b2, _ = st.columns([1, 1, 2])
        if b1.button("すべて対象", key="premium_diff_all", width="stretch"):
            st.session_state[f"{editor_key}_default"] = True
            st.session_state[f"{editor_key}_ver"] = st.session_state.get(f"{editor_key}_ver", 0) + 1
        if b2.button("すべて解除", key="premium_diff_none", width="stretch"):
            st.session_state[f"{editor_key}_default"] = False
            st.session_state[f"{editor_key}_ver"] = st.session_state.get(f"{editor_key}_ver", 0) + 1

        diff_df["対象"] = bool(st.session_state.get(f"{editor_key}_default", True))
        view = diff_df.copy()
        view["日付"] = view["日付"].map(format_date_label)

        money = st.column_config.NumberColumn(format="yen", disabled=True)
        edited = st.data_editor(
            view,
            key=f"{editor_key}_{st.session_state.get(f'{editor_key}_ver', 0)}",
            width="stretch",
            hide_index=True,
            num_rows="fixed",
            column_config={
                "対象": st.column_config.CheckboxColumn(),
                "日付": st.column_config.TextColumn(disabled=True),
                "部屋": st.column_config.TextColumn(disabled=True),
                "内訳": st.column_config.TextColumn(disabled=True),
                "区分": st.column_config.TextColumn(disabled=True),
                "通常料金": money,
                "割増料金": money,
                "差額": money,
                "備考": st.column_config.TextColumn(disabled=True),
            },
        )

        target = edited[edited["対象"] == True]
        normal_total = int(target["通常料金"].sum())
        premium_total = int(target["割増料金"].sum())
        diff_total = int(target["差額"].sum())
        # 差額の内訳（部屋代・延長）
        diff_by_kind = {
            kind: int(target.loc[target["内訳"] == kind, "差額"].sum())
            for kind in (PREMIUM_DIFF_ROOM, PREMIUM_DIFF_EXTENSION)
        }
        has_extension = (target["内訳"] == PREMIUM_DIFF_EXTENSION).any()

        c1, c2, c3 = st.columns(3)
        c1.metric("通常料金（対象分）", yen(normal_total))
        c2.metric("割増料金（対象分）", yen(premium_total))
        c3.metric("差額（追加でいただく金額）", yen(diff_total))
        if has_extension:
            st.markdown(
                f"差額の内訳：部屋代 **{yen(diff_by_kind[PREMIUM_DIFF_ROOM])}** ／ "
                f"延長 **{yen(diff_by_kind[PREMIUM_DIFF_EXTENSION])}**"
            )
        st.caption(f"対象：{len(target)}件 / 全{len(edited)}件")

        if (target["備考"].astype(str) != "").any():
            st.warning("計算できない料金を含む行があります。備考を確認してください（その分は0円として計算しています）。")

        csv_df = target.drop(columns=["対象"])
        csv_df = pd.concat(
            [
                csv_df,
                pd.DataFrame(
                    [{"日付": "合計", "通常料金": normal_total, "割増料金": premium_total, "差額": diff_total}]
                ),
            ],
            ignore_index=True,
        )
        d1, d2 = st.columns(2)
        d1.download_button(
            "差額明細をCSVでダウンロード",
            data=csv_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"有料差額_{date_suffix}.csv",
            mime="text/csv",
            on_click="ignore",
            key="premium_diff_csv",
            width="stretch",
        )
        try:
            pdf_df = target.drop(columns=["対象"]).copy()
            pdf_df["日付"] = pdf_df["日付"].map(format_date_label)
            pdf_bytes = build_difference_pdf(
                heading="有料差額（通常料金 → 割増料金）",
                period=period_label,
                rooms=rooms,
                summary=[
                    ("通常料金（対象分）", normal_total),
                    ("割増料金（対象分）", premium_total),
                    *(
                        [
                            ("差額のうち部屋代", diff_by_kind[PREMIUM_DIFF_ROOM]),
                            ("差額のうち延長", diff_by_kind[PREMIUM_DIFF_EXTENSION]),
                        ]
                        if has_extension
                        else []
                    ),
                    ("差額（追加でいただく金額）", diff_total),
                ],
                detail_df=pdf_df,
                widths={"日付": 32, "部屋": 24, "内訳": 13, "区分": 26, "通常料金": 20, "割増料金": 20, "差額": 20, "備考": 28},
                money_cols=("通常料金", "割増料金", "差額"),
                notes=["※部屋代・延長料金のみ。設備・技術者・インターネットは割増で変わらないため含みません。"],
            )
            d2.download_button(
                "差額明細をPDFでダウンロード",
                data=pdf_bytes,
                file_name=f"有料差額_{date_suffix}.pdf",
                mime="application/pdf",
                on_click="ignore",
                key="premium_diff_pdf",
                width="stretch",
            )
        except Exception as e:
            d2.warning(f"PDFを作成できませんでした: {e}")

def render_allday_difference(
    prices_df: pd.DataFrame,
    room_day_df: pd.DataFrame,
    room_day_key: str,
    period_label: str,
    rooms: List[str],
    date_suffix: str,
) -> None:
    """【全日差額】ボタンで開く、一部の区分→全日に変更した場合の差額計算パネル。"""
    if not st.session_state.get(ALLDAY_DIFF_OPEN_KEY, False):
        return

    with st.container(border=True, key="panel_body_allday"):
        st.markdown("#### 全日差額（一部の区分 → 全日）")
        st.caption(
            "申し込み済みの区分（変更前区分）から全日に変更した場合に、追加でいただく金額を計算します。"
            "全日料金は午前＋午後＋夜間の合計より安いため、差額は「全日料金 − 変更前区分の料金」です。"
            "変更前区分と割増は表の中で変更できます（基本料金のみ。延長は含みません）。"
        )

        base = allday_difference_base_rows(room_day_df)
        if base.empty:
            st.info("差額を計算できる部屋の利用がありません（休館日・「利用なし」は対象外です）。")
            return

        # 行の構成（日付・部屋・区分）が変わったら入力をリセットする
        signature = "|".join(
            base[["日付", "部屋", "変更前区分"]].astype(str).agg(",".join, axis=1).tolist()
        )
        base_key = f"allday_diff_{room_day_key}_{abs(hash(signature))}"
        ver_key = f"{base_key}_ver"
        editor_key = f"{base_key}_editor_{st.session_state.get(ver_key, 0)}"
        if base_key not in st.session_state:
            st.session_state[base_key] = base

        def on_edit():
            state = st.session_state.get(editor_key, {})
            edited_rows = state.get("edited_rows", {}) if isinstance(state, dict) else {}
            cur = st.session_state[base_key].copy()
            for idx, changes in edited_rows.items():
                i = int(idx)
                if 0 <= i < len(cur):
                    for col, val in changes.items():
                        if col in ("対象", "割増", "変更前区分") and val is not None:
                            cur.at[i, col] = val
            st.session_state[base_key] = cur
            st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1

        def set_all(value: bool):
            cur = st.session_state[base_key].copy()
            cur["対象"] = value
            st.session_state[base_key] = cur
            st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1

        b1, b2, _ = st.columns([1, 1, 2])
        b1.button("すべて対象", key=f"{base_key}_all", width="stretch", on_click=set_all, args=(True,))
        b2.button("すべて解除", key=f"{base_key}_none", width="stretch", on_click=set_all, args=(False,))

        result, sep_cols = split_separate_by_label(calc_allday_difference(prices_df, st.session_state[base_key]))
        # 参考列の構成（例：参考：夜間1区分）が変わったら表を作り直す（入力は base_key 側に保持済み）
        editor_key = f"{editor_key}_{abs(hash('|'.join(sep_cols)))}"
        view = result.copy()
        view["日付"] = view["日付"].map(format_date_label)
        money = st.column_config.NumberColumn(format="yen", disabled=True)
        st.data_editor(
            view,
            key=editor_key,
            on_change=on_edit,
            width="stretch",
            hide_index=True,
            num_rows="fixed",
            column_config={
                "対象": st.column_config.CheckboxColumn(),
                "日付": st.column_config.TextColumn(disabled=True),
                "部屋": st.column_config.TextColumn(disabled=True),
                "割増": st.column_config.CheckboxColumn(),
                "変更前区分": st.column_config.SelectboxColumn(options=ALLDAY_FROM_SLOTS, required=True),
                "変更前料金": money,
                "全日料金": money,
                "差額": money,
                **{
                    c: st.column_config.NumberColumn(
                        format="yen",
                        disabled=True,
                        help=f"全日にせず、{c.replace('参考：', '')}を個別に追加した場合の料金（比較用）",
                    )
                    for c in sep_cols
                },
                "備考": st.column_config.TextColumn(disabled=True),
            },
        )

        target = result[result["対象"] == True]
        before_total = int(pd.to_numeric(target["変更前料金"], errors="coerce").fillna(0).sum())
        allday_total = int(pd.to_numeric(target["全日料金"], errors="coerce").fillna(0).sum())
        diff_total = int(pd.to_numeric(target["差額"], errors="coerce").fillna(0).sum())
        # 対象行に出てくる参考列だけを合計・CSV・PDFに使う
        target_labels = set(target["変更前区分"].map(allday_separate_label))
        target_sep_cols = [c for c in sep_cols if c in target_labels]
        separate_totals = {c: int(pd.to_numeric(target[c], errors="coerce").fillna(0).sum()) for c in target_sep_cols}
        target = target.drop(columns=[c for c in sep_cols if c not in target_sep_cols])

        c1, c2, c3 = st.columns(3)
        c1.metric("変更前の料金（対象分）", yen(before_total))
        c2.metric("全日料金（対象分）", yen(allday_total))
        c3.metric("差額（追加でいただく金額）", yen(diff_total))
        st.caption(
            f"対象：{len(target)}件 / 全{len(result)}件"
            + "".join(
                f"　（{c.replace('参考：', '参考：全日にせず')}を個別に追加した場合 {yen(v)}）"
                for c, v in separate_totals.items()
            )
        )

        if (target["備考"].astype(str) != "").any():
            st.warning("計算できない料金を含む行があります。備考を確認してください（その分は0円として計算しています）。")

        csv_df = target.drop(columns=["対象"])
        csv_df = pd.concat(
            [
                csv_df,
                pd.DataFrame(
                    [{
                        "日付": "合計",
                        "変更前料金": before_total,
                        "全日料金": allday_total,
                        "差額": diff_total,
                        **separate_totals,
                    }]
                ),
            ],
            ignore_index=True,
        )
        d1, d2 = st.columns(2)
        d1.download_button(
            "全日差額の明細をCSVでダウンロード",
            data=csv_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"全日差額_{date_suffix}.csv",
            mime="text/csv",
            on_click="ignore",
            key="allday_diff_csv",
            width="stretch",
        )
        try:
            pdf_df = target.drop(columns=["対象"]).copy()
            pdf_df["日付"] = pdf_df["日付"].map(format_date_label)
            sep_notes = [
                f"※{c.replace('参考：', '参考：全日にせず')}を個別に追加した場合 {yen(v)}"
                for c, v in separate_totals.items()
            ]
            pdf_bytes = build_difference_pdf(
                heading="全日差額（一部の区分 → 全日）",
                period=period_label,
                rooms=rooms,
                summary=[
                    ("変更前の料金（対象分）", before_total),
                    ("全日料金（対象分）", allday_total),
                    ("差額（追加でいただく金額）", diff_total),
                ],
                detail_df=pdf_df,
                widths={
                    "日付": 34, "部屋": 24, "割増": 12, "変更前区分": 22, "変更前料金": 20,
                    "全日料金": 20, "差額": 20, **{c: 22 for c in target_sep_cols}, "備考": 18,
                },
                money_cols=("変更前料金", "全日料金", "差額", *target_sep_cols),
                notes=[
                    "※差額＝全日料金 − 変更前区分の料金（基本料金のみ。延長は含みません）。",
                    *sep_notes,
                ],
            )
            d2.download_button(
                "全日差額の明細をPDFでダウンロード",
                data=pdf_bytes,
                file_name=f"全日差額_{date_suffix}.pdf",
                mime="application/pdf",
                on_click="ignore",
                key="allday_diff_pdf",
                width="stretch",
            )
        except Exception as e:
            d2.warning(f"PDFを作成できませんでした: {e}")

def render_panel_downloads(
    csv_df: pd.DataFrame,
    file_stem: str,
    key: str,
    csv_label: str,
    pdf_label: str,
    build_pdf: Callable[[], bytes],
) -> None:
    """パネルの中身（表示している明細）をそのまま CSV・PDF で出力するボタン。"""
    d1, d2 = st.columns(2)
    d1.download_button(
        csv_label,
        data=csv_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{file_stem}.csv",
        mime="text/csv",
        on_click="ignore",
        key=f"{key}_csv",
        width="stretch",
    )
    try:
        d2.download_button(
            pdf_label,
            data=build_pdf(),
            file_name=f"{file_stem}.pdf",
            mime="application/pdf",
            on_click="ignore",
            key=f"{key}_pdf",
            width="stretch",
        )
    except Exception as e:
        d2.warning(f"PDFを作成できませんでした: {e}")

def render_extension_detail(all_df: pd.DataFrame, period_label: str, rooms: List[str], date_suffix: str) -> None:
    """【延長料金のみ明細】ボタンで開く、延長料金の行だけの明細パネル。"""
    if not st.session_state.get(EXTENSION_DETAIL_OPEN_KEY, False):
        return

    with st.container(border=True, key="panel_body_detail_extension"):
        st.markdown("#### 延長料金のみ明細")
        st.caption("部屋の延長（前延長・後延長）と、設備の区分が「延長30分」の行だけを表示します（技術者は含みません）。")

        ext_df = extension_detail_rows(all_df)
        if ext_df.empty:
            st.info(
                "延長料金の明細はありません。部屋の延長は部屋×日テーブルの「延長」列、"
                "設備の延長30分は日別設定の設備区分で設定できます。"
            )
            return

        show_detail_df(ext_df)

        amounts = pd.to_numeric(ext_df["小計"], errors="coerce").fillna(0)
        room_total = int(amounts[ext_df["カテゴリ"] == "部屋"].sum())
        equipment_total = int(amounts[ext_df["カテゴリ"] == "設備"].sum())
        ext_total = room_total + equipment_total

        c1, c2, c3 = st.columns(3)
        c1.metric("部屋の延長", yen(room_total))
        c2.metric("設備の延長", yen(equipment_total))
        c3.metric("延長料金の合計", yen(ext_total))
        st.caption(f"延長の明細：{len(ext_df)}件")

        if ext_df["小計"].isna().any():
            st.warning("計算できない延長を含む行があります。備考を確認してください（その分は0円として計算しています）。")

        csv_df = pd.concat([ext_df, pd.DataFrame([{"日付": "合計", "小計": ext_total}])], ignore_index=True)
        render_panel_downloads(
            csv_df,
            file_stem=f"延長料金_{date_suffix}",
            key="extension_detail",
            csv_label="延長料金の明細をCSVでダウンロード",
            pdf_label="延長料金の明細をPDFでダウンロード",
            build_pdf=lambda: build_difference_pdf(
                heading="延長料金のみ明細",
                period=period_label,
                rooms=rooms,
                summary=[
                    ("部屋の延長", room_total),
                    ("設備の延長", equipment_total),
                    ("延長料金の合計", ext_total),
                ],
                detail_df=ext_df,
                widths={"日付": 34, "祝日": 20, "カテゴリ": 16, "名称": 36, "区分": 24, "小計": 20, "備考": 40},
                money_cols=("小計",),
                notes=["※部屋の延長と、設備の区分が「延長30分」の行のみ。技術者の延長は含みません。"],
            ),
        )

# =========================
# Main App
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    inject_ui_css()

    # 上部の合計は、入力が確定した後（計算の後）に描画する。
    # 再計算中に合計欄が消えて画面がずれないよう、まず前回の合計を表示しておき、計算後に置き換える。
    # 部屋を選ぶ前も ¥0 で表示しておく（最初に部屋を選んだときに合計欄が現れて、部屋の選択欄が下にずれないように）
    sticky_slot = st.empty()
    with sticky_slot.container():
        render_totals_sticky(*st.session_state.get(LAST_TOTALS_KEY, (0, 0, 0, 0)))

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

    with left:
        st.subheader("期間・部屋")

        today = pd.Timestamp.today().date()
        if st.session_state.get("start_date") is None:
            st.session_state["start_date"] = today
        if st.session_state.get("end_date") is None:
            st.session_state["end_date"] = st.session_state["start_date"]
        st.session_state.setdefault("start_date_prev", st.session_state["start_date"])

        def on_start_date_change():
            # 終了日が開始日と同じ（1日利用）か開始日より前なら、終了日を開始日に合わせる
            new_start = st.session_state.get("start_date")
            prev_start = st.session_state.get("start_date_prev")
            end = st.session_state.get("end_date")
            if new_start is not None and (end is None or end == prev_start or end < new_start):
                st.session_state["end_date"] = new_start
            st.session_state["start_date_prev"] = new_start

        col_a, col_b = st.columns(2)
        with col_a:
            start_date = st.date_input(
                "開始日",
                key="start_date",
                on_change=on_start_date_change,
            )
        with col_b:
            end_date = st.date_input(
                "終了日（1日のみの場合は入力不要）",
                key="end_date",
            )

        # Noneガード（未選択の場合の対策）
        if start_date is None:
            start_date = today
        if end_date is None:
            end_date = start_date

        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)
        days = build_date_range(start_ts, end_ts)
        if not days:
            st.error("日付範囲が不正です（終了日が開始日より前です）。")
            st.stop()

        if closed_days:
            closed_min, closed_max = min(closed_days), max(closed_days)
            if days[0].date() < closed_min or days[-1].date() > closed_max:
                st.warning(
                    "休館日データが登録されていない期間を含みます"
                    f"（登録済み：{closed_min.strftime(DATE_FMT)}〜{closed_max.strftime(DATE_FMT)}）。"
                    "休館日の判定ができないため、ご注意ください。"
                )

        # ROOM_DISPLAY_ORDER の順に並べる。「全館」は先頭に置き、組み込みの「Select all」の代わりに使う
        all_rooms = set(prices_df["room"].unique().tolist())
        room_candidates = [r for r in ROOM_DISPLAY_ORDER if r in all_rooms] + sorted(all_rooms - set(ROOM_DISPLAY_ORDER))

        def on_rooms_selected_change():
            # 「全館」を新たに選択したときは、他の部屋の選択を外す
            cur = list(st.session_state.get("rooms_selected", []))
            prev = set(st.session_state.get("rooms_selected_prev", []))
            if ALL_BUILDING_ROOM in cur and ALL_BUILDING_ROOM not in prev:
                st.session_state["rooms_selected"] = [ALL_BUILDING_ROOM]

        # 「全館」選択中は、各部屋を選択肢に出さない
        all_building_selected = ALL_BUILDING_ROOM in st.session_state.get("rooms_selected", [])
        room_options = [ALL_BUILDING_ROOM] if all_building_selected else room_candidates

        rooms_selected = st.multiselect(
            "部屋（複数選択可）",
            room_options,
            key="rooms_selected",
            on_change=on_rooms_selected_change,
            placeholder="選択してください",
            select_all=False,
        )
        st.session_state["rooms_selected_prev"] = list(rooms_selected)
        selected_rooms = list(rooms_selected)
        if not selected_rooms:
            st.warning("部屋を選択してください（未選択のままでは計算できません）。")

        default_room_slot = st.selectbox(
            "部屋の区分（新規追加の初期値）",
            ROOM_SLOTS_WITH_NONE,
            index=ROOM_SLOTS_WITH_NONE.index("全日") if "全日" in ROOM_SLOTS_WITH_NONE else 0,
        )

        extension_options = (
            ROOM_EXTENSION_SLOTS_2130 if default_room_slot in ROOM_SLOTS_ENDING_2130 else ROOM_EXTENSION_SLOTS
        )
        default_room_extension = st.selectbox(
            "部屋の延長（新規追加の初期値）",
            extension_options,
            index=extension_options.index("なし"),
        )

        is_business_default = st.checkbox("割増利用（デフォルト）", value=False)

        st.divider()

        days_key = f"days_{start_date}_{end_date}"
        new_defaults = {
            "割増利用": bool(is_business_default),
            "設備デフォ区分": default_room_slot,
            "技術者区分": default_room_slot,
        }

        if days_key not in st.session_state:
            df_days = make_days_base(days, closed_days, default_room_slot, is_business_default)
            st.session_state[days_key] = df_days
            st.session_state[days_key + "_defaults"] = dict(new_defaults)
        else:
            df_existing = st.session_state[days_key]
            if (
                len(df_existing) != len(days)
                or normalize_str(df_existing.iloc[0]["日付"]) != days[0].strftime(DATE_FMT)
                or normalize_str(df_existing.iloc[-1]["日付"]) != days[-1].strftime(DATE_FMT)
            ):
                df_days = make_days_base(days, closed_days, default_room_slot, is_business_default)
                st.session_state[days_key] = df_days
                st.session_state[days_key + "_defaults"] = dict(new_defaults)
            else:
                old_defaults = st.session_state.get(days_key + "_defaults", dict(new_defaults))
                st.session_state[days_key] = sync_days_df_defaults(df_existing, old_defaults, new_defaults)
                st.session_state[days_key + "_defaults"] = dict(new_defaults)

        days_expander = st.expander("日別設定（設備・技術者の区分、割増利用）", expanded=False)
        days_expander.caption(
            "設備・技術者・インターネットの計算に使います。部屋料金は下の「部屋×日テーブル」で調整してください。"
        )
        try:
            edited_days = days_expander.data_editor(
                st.session_state[days_key],
                width="stretch",
                hide_index=True,
                column_order=["日付", "割増利用", "設備デフォ区分", "技術者区分", "土日祝", "祝日名", "休館日"],
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
            st.warning("この環境では日別編集UIが利用できないため、日別設定は表示のみになります。")
            edited_days = st.session_state[days_key]
            days_expander.dataframe(edited_days, width="stretch", hide_index=True)

        if not edited_days.empty and bool(edited_days["休館日"].any()):
            closed_list = edited_days.loc[edited_days["休館日"] == True, "日付"].astype(str).tolist()
            closed_list = [d for d in closed_list if d]
            msg = " / ".join(closed_list[:20])
            suffix = "…" if len(closed_list) > 20 else ""
            st.error(f"休館日があります：{msg}{suffix}")

        st.divider()
        st.subheader("部屋×日テーブル（日ごと・部屋ごとの調整）")
        st.caption("表の編集はすぐに計算へ反映されます。")

        room_day_key = f"room_day_{start_date}_{end_date}"

        if room_day_key not in st.session_state:
            st.session_state[room_day_key] = build_room_day_base(
                edited_days, list(selected_rooms), default_room_slot, default_room_extension
            )
        else:
            st.session_state[room_day_key] = merge_room_day(
                st.session_state[room_day_key],
                edited_days,
                list(selected_rooms),
                default_room_slot,
                default_room_extension,
            )

        st.markdown("#### 絞り込み")
        f1, f2 = st.columns([1, 1])
        all_dates = sorted(st.session_state[room_day_key]["日付"].unique().tolist()) if not st.session_state[room_day_key].empty else []
        date_filter = f1.multiselect(
            "日付（未選択＝全日）", options=all_dates, default=[], key=f"filter_dates_{room_day_key}", placeholder="すべての日付"
        )
        all_rooms_in_table = sorted(st.session_state[room_day_key]["部屋"].unique().tolist()) if not st.session_state[room_day_key].empty else []
        room_filter = f2.multiselect(
            "部屋（未選択＝全部屋）", options=all_rooms_in_table, default=[], key=f"filter_rooms_{room_day_key}", placeholder="すべての部屋"
        )

        view_df = st.session_state[room_day_key].copy()
        if date_filter:
            view_df = view_df[view_df["日付"].isin(date_filter)]
        if room_filter:
            view_df = view_df[view_df["部屋"].isin(room_filter)]
        view_df = view_df.reset_index(drop=True)

        # 編集は on_change で即座に全体テーブルへ反映する。
        # 反映後はエディタの key を更新し、行番号ベースの編集差分を持ち越さない。
        editor_ver_key = f"{room_day_key}_editor_ver"
        editor_key = f"{room_day_key}_editor_{st.session_state.get(editor_ver_key, 0)}"
        st.session_state[f"{room_day_key}_view"] = view_df

        def on_room_day_edit():
            state = st.session_state.get(editor_key, {})
            edited_rows = state.get("edited_rows", {}) if isinstance(state, dict) else {}
            if not edited_rows:
                return
            view = st.session_state[f"{room_day_key}_view"].copy()
            for idx, changes in edited_rows.items():
                i = int(idx)
                if i < 0 or i >= len(view):
                    continue
                for col, val in changes.items():
                    if col in view.columns:
                        view.at[i, col] = val
            st.session_state[room_day_key] = apply_room_day_edits(
                st.session_state[room_day_key], view, default_room_slot
            )
            st.session_state[editor_ver_key] = st.session_state.get(editor_ver_key, 0) + 1

        try:
            st.data_editor(
                view_df,
                key=editor_key,
                on_change=on_room_day_edit,
                width="stretch",
                hide_index=True,
                column_order=["日付", "部屋", "区分", "延長", "割増利用", "土日祝", "祝日名", "休館日"],
                num_rows="fixed",
                column_config={
                    "日付": st.column_config.TextColumn(disabled=True),
                    "土日祝": st.column_config.TextColumn(disabled=True),
                    "祝日名": st.column_config.TextColumn(disabled=True),
                    "休館日": st.column_config.CheckboxColumn(disabled=True),
                    "部屋": st.column_config.TextColumn(disabled=True),
                    "区分": st.column_config.SelectboxColumn(options=ROOM_SLOTS_WITH_NONE),
                    "延長": st.column_config.SelectboxColumn(options=ROOM_EXTENSION_SLOTS),
                    "割増利用": st.column_config.CheckboxColumn(),
                    "手動区分": st.column_config.CheckboxColumn(disabled=True),
                    "手動延長": st.column_config.CheckboxColumn(disabled=True),
                    "手動割増": st.column_config.CheckboxColumn(disabled=True),
                },
            )
        except Exception:
            st.warning("この環境では部屋×日編集UIが利用できないため、表示のみになります。")
            st.dataframe(view_df, width="stretch", hide_index=True)

    with right:
        st.subheader("設備・技術者・インターネット")

        room_day_df = st.session_state.get(room_day_key, pd.DataFrame())
        if room_day_df is None:
            room_day_df = pd.DataFrame()

        rooms_by_day = rooms_used_by_date(room_day_df)

        has_d_days_raw = any(bool(MIC_D_ROOMS.issubset(rs)) for rs in rooms_by_day.values())
        d_rooms_explicitly_selected = MIC_D_ROOMS.issubset(set(rooms_selected))
        if d_rooms_explicitly_selected and has_d_days_raw:
            gallery_678 = st.checkbox(
                "（第6〜8会議室）ギャラリー利用",
                value=bool(st.session_state.get("gallery_678", False)),
                key="gallery_678",
                help="拡声装置D（第6+第7+第8同日利用）の計算対象判定に使用します。",
            )
        else:
            gallery_678 = False

        st.divider()

        st.markdown("### 設備（数量）")
        st.caption("各備品の「対象部屋」を確認の上、数量を入力してください。")

        selected_rooms_now = sorted(list(set([normalize_str(x) for x in selected_rooms])))
        selected_rooms_set = set(selected_rooms_now)

        def group_applies(meta: GroupMeta) -> bool:
            targets = parse_rooms_cell(meta.applies_to_rooms)
            if targets == ["*"]:
                return True
            return bool(set(targets) & selected_rooms_set) if selected_rooms_set else True

        group_overrides: Dict[str, str] = {}
        with st.expander("設備の区分（グループ単位の上書き：任意）", expanded=False):
            st.caption("未指定の場合、日別設定の「設備デフォ区分」が使用されます。")
            for gid, meta in group_meta.items():
                if not group_applies(meta):
                    continue

                if not meta.allowed_slot_override:
                    st.text(f"・{meta.group_name}：上書き不可")
                    continue

                key = f"ov_{gid}"
                default_label = "（日別デフォルト）"
                options = [default_label] + EQUIPMENT_TIME_SLOTS
                choice = st.selectbox(f"{meta.group_name}", options=options, index=0, key=key)
                if choice != default_label:
                    group_overrides[gid] = choice

        st.divider()

        q = st.text_input("備品名で検索（任意）", value="", help="部分一致で絞り込みます。例：スクリーン、マイク、プロジェクター")

        base_selections: List[Dict] = []

        items_by_group: Dict[str, List[EquipmentItem]] = {}
        for it in items.values():
            items_by_group.setdefault(it.group_id, []).append(it)

        group_order = [normalize_str(x) for x in groups_df["group_id"].tolist()] if "group_id" in groups_df.columns else sorted(list(group_meta.keys()))

        def _supplement_label(notes: str) -> str:
            n = normalize_str(notes)
            if not n:
                return ""
            keys = ["インチ", "cm", "mm", "サイズ", "幅", "高さ", "奥行"]
            if any(k in n for k in keys):
                return f" / 補足:{n}"
            return ""

        # 数量は入力欄とは別に保持する（検索で非表示・部屋の切替で対象外になっても消さない）
        qty_store: Dict[str, int] = st.session_state.setdefault(EQUIPMENT_QTY_STORE_KEY, {})
        if q:
            st.caption("検索で表示されていない備品の数量も、そのまま計算に含まれます。")

        for gid in group_order:
            meta = group_meta.get(gid)
            if not meta:
                continue
            if not group_applies(meta):
                continue

            group_items = sorted(items_by_group.get(gid, []), key=lambda x: x.item_name)

            if q:
                group_items = [it for it in group_items if q in it.item_name or q in it.notes or q in it.item_id]

            if not group_items:
                continue

            with st.expander(f"{meta.group_name}", expanded=False):
                st.caption(f"対象部屋: {meta.applies_to_rooms if meta.applies_to_rooms else '*'}")

                for it in group_items:
                    qty_key = f"qty_{it.item_id}"

                    price_txt = []
                    if it.price_per_slot > 0:
                        per = "1日" if it.is_power_item else "1区分"
                        price_txt.append(f"{per}:{it.price_per_slot:,}円")
                    if it.price_once_yen > 0:
                        price_txt.append(f"単価:{it.price_once_yen:,}円")
                    price_str = " / ".join(price_txt) if price_txt else "料金未設定"

                    target_rooms = infer_item_target_rooms(
                        it.item_name,
                        it.notes,
                        meta.applies_to_rooms if meta.applies_to_rooms else "*",
                    )
                    label = f"{it.item_name}（対象:{target_rooms} / 単位:{it.unit} / {price_str}{_supplement_label(it.notes)}）"
                    help_txt = it.notes if it.notes else None

                    if qty_key not in st.session_state:
                        st.session_state[qty_key] = int(qty_store.get(it.item_id, 0) or 0)
                    qty = st.number_input(
                        label,
                        min_value=0,
                        step=1,
                        key=qty_key,
                        help=help_txt,
                    )
                    qty_store[it.item_id] = int(qty or 0)

        excluded_items: List[str] = []
        for gid in group_order:
            meta = group_meta.get(gid)
            if not meta:
                continue
            for it in sorted(items_by_group.get(gid, []), key=lambda x: x.item_name):
                qty = int(qty_store.get(it.item_id, 0) or 0)
                if qty <= 0:
                    continue
                if not group_applies(meta):
                    excluded_items.append(f"{it.item_name}×{qty}")
                    continue
                base_selections.append(
                    {
                        "group_id": it.group_id,
                        "item_id": it.item_id,
                        "qty": qty,
                        "auto_added": False,
                    }
                )
        if excluded_items:
            st.info(
                "選択中の部屋では対象外のため、次の備品は計算に含めていません（数量は保持しています）："
                + "、".join(excluded_items)
            )

        if rooms_by_day:
            sel_map = {s["item_id"]: int(s.get("qty", 0) or 0) for s in base_selections}

            def _has_any(qty: int) -> bool:
                return int(qty or 0) > 0

            if _has_any(sel_map.get(PA_C_ID, 0)):
                eligible = [d for d, rs in rooms_by_day.items() if bool(rs & MIC_C_ROOMS)]
                if not eligible:
                    st.info("拡声装置C：対象日（大会議室/小集会室）がないため計算されません。")

            if _has_any(sel_map.get(PA_D_ID, 0)):
                eligible = [d for d, rs in rooms_by_day.items() if bool(MIC_D_ROOMS.issubset(rs) and gallery_678)]
                if not eligible:
                    st.info("拡声装置D：対象日（第6+第7+第8同日利用 かつ ギャラリー利用）がないため計算されません。")

            if _has_any(sel_map.get(MIC_WIRED_ID, 0)) or _has_any(sel_map.get(MIC_WIRELESS_ID, 0)):
                eligible = []
                for d, rs in rooms_by_day.items():
                    ok, _ = infer_mic_allowed_for_rooms(rs, gallery_678)
                    if ok:
                        eligible.append(d)
                if not eligible:
                    st.info("マイク：対象日がないため計算されません（大会議室/小集会室 または 第6〜8条件が必要です）。")

        st.divider()

        st.markdown("### 舞台設備技術者")
        tech_people = st.number_input(
            "人数",
            min_value=0,
            value=0,
            step=1,
            help="日別設定の「技術者区分」×人数で計算します。",
        )

        st.divider()

        st.markdown("### インターネット")
        st.markdown("#### 固定ネット設備（有線LAN / Wi-Fi）")
        fixed_network_selections: Dict[Tuple[str, str], str] = {}
        fixed_network_rows = [
            r for r in active_room_day_records(room_day_df)
            if len(available_fixed_network_services(str(r["room"]))) > 1
        ]
        if fixed_network_rows:
            fixed_network_dates = sorted({str(r["date_str"]) for r in fixed_network_rows})
            fixed_network_state_key = "fixed_network_saved_selections"
            active_fixed_network_keys = {
                (str(r["date_str"]), str(r["room"])) for r in fixed_network_rows
            }
            saved_fixed_network_selections = dict(
                st.session_state.get(fixed_network_state_key, {})
            )

            def save_fixed_network_choice(
                selection_key: Tuple[str, str], widget_key: str
            ) -> None:
                saved = dict(st.session_state.get(fixed_network_state_key, {}))
                saved[selection_key] = st.session_state.get(widget_key, INTERNET_NONE)
                st.session_state[fixed_network_state_key] = saved

            # 折りたたみ中や、別の日付を表示している間も、全日分の選択を
            # 永続用の session_state から復元して計算へ渡す。
            for rec in fixed_network_rows:
                date_str = str(rec["date_str"])
                room = str(rec["room"])
                options = available_fixed_network_services(room)
                selection_key = (date_str, room)
                legacy_widget_key = f"fixed_net_{date_str}_{room}"
                current = saved_fixed_network_selections.get(
                    selection_key,
                    st.session_state.get(legacy_widget_key, INTERNET_NONE),
                )
                if current not in options:
                    current = INTERNET_NONE
                saved_fixed_network_selections[selection_key] = current
                fixed_network_selections[selection_key] = current

            # 現在の部屋・日付から外れた古い選択は引き継がない。
            saved_fixed_network_selections = {
                key: value
                for key, value in saved_fixed_network_selections.items()
                if key in active_fixed_network_keys
            }
            st.session_state[fixed_network_state_key] = saved_fixed_network_selections

            configured_count = sum(
                choice != INTERNET_NONE for choice in fixed_network_selections.values()
            )
            expander_label = "固定ネット設備を設定・変更する"
            if configured_count:
                expander_label += f"（選択済み {configured_count}件）"

            with st.expander(expander_label, expanded=False):
                st.caption("部屋・日ごとに、なし / 有線LAN / Wi-Fi から1つだけ選択します。")

                if len(fixed_network_dates) > 1:
                    display_date_key = "fixed_network_display_date"
                    if st.session_state.get(display_date_key) not in fixed_network_dates:
                        st.session_state[display_date_key] = fixed_network_dates[0]
                    display_date = st.selectbox(
                        "設定する日付",
                        options=fixed_network_dates,
                        key=display_date_key,
                    )
                else:
                    display_date = fixed_network_dates[0]

                rows_for_display = [
                    rec for rec in fixed_network_rows
                    if str(rec["date_str"]) == display_date
                ]
                for rec in rows_for_display:
                    date_str = str(rec["date_str"])
                    room = str(rec["room"])
                    options = available_fixed_network_services(room)
                    selection_key = (date_str, room)
                    key = f"fixed_net_{date_str}_{room}"
                    current = fixed_network_selections[selection_key]
                    if st.session_state.get(key) not in options:
                        st.session_state[key] = current
                    choice = st.selectbox(
                        room,
                        options=options,
                        index=options.index(current),
                        key=key,
                        on_change=save_fixed_network_choice,
                        args=(selection_key, key),
                    )
                    fixed_network_selections[selection_key] = choice
        else:
            st.caption("固定ネット設備の対象室（大集会室・中集会室・小集会室・特別室・大会議室）は利用日に含まれていません。")

        st.markdown("#### その他インターネット")
        use_pocket_wifi = st.checkbox("ポケットWi-Fi（2,800円/日）", value=False)
        use_temp_line = st.checkbox("仮設回線（5,000円/回 + 別途見積）", value=False)
        st.divider()

        st.subheader("結果")
        if not selected_rooms:
            with sticky_slot.container():
                render_totals_sticky(0, 0, 0, 0)
            st.session_state.pop(LAST_TOTALS_KEY, None)
            st.info("部屋を選択すると、料金が自動で計算されます。")
        else:
            room_total, room_df = calc_rooms_from_room_day(prices_df, room_day_df)

            equipment_total, equipment_df = calc_equipment_total_all_days(
                days_df=edited_days,
                room_day_df=room_day_df,
                global_default_slot=default_room_slot,
                group_overrides=group_overrides,
                base_selections=base_selections,
                items=items,
                group_meta=group_meta,
                gallery_678=gallery_678,
            )

            tech_total, tech_df = calc_stage_tech_total_all_days(edited_days, room_day_df, int(tech_people))

            internet_total, internet_df = calc_internet_total(
                room_day_df=room_day_df,
                fixed_network_selections=fixed_network_selections,
                use_pocket_wifi=use_pocket_wifi,
                use_temp_line=use_temp_line,
            )

            with sticky_slot.container():
                render_totals_sticky(room_total, equipment_total, tech_total, internet_total)
            st.session_state[LAST_TOTALS_KEY] = (room_total, equipment_total, tech_total, internet_total)
            render_kpis_cards(room_total, equipment_total, tech_total, internet_total)

            invalid_ext_df = invalid_after_extension_rows(room_day_df)
            if not invalid_ext_df.empty:
                lines = [
                    f"- {r['日付']} / {r['部屋']} / {r['区分']} / {r['延長']}"
                    for _, r in invalid_ext_df.iterrows()
                ]
                st.error(
                    "21:30終了の区分（夜間・午後-夜間・全日）は後延長できません。"
                    "次の行の延長分は計算から除外しています（日付 / 部屋 / 区分 / 延長）：\n"
                    + "\n".join(lines)
                )

            # 上段＝差額の試算（スマホでも横2列のまま：inject_ui_css の .st-key-panel_toggles）、下段＝延長料金のみ明細
            with st.container(key="panel_toggles"):
                t1, t2 = st.columns(2)
                render_panel_toggle(
                    t1,
                    "【有料差額計算】",
                    PREMIUM_DIFF_OPEN_KEY,
                    "通常料金で申し込んだ利用が割増料金になった場合の、部屋料金の差額を計算します。",
                )
                render_panel_toggle(
                    t2,
                    "【全日差額】",
                    ALLDAY_DIFF_OPEN_KEY,
                    "午前・午後・夜間などで申し込んだ利用が、区分の追加で全日になった場合の差額を計算します。",
                )
                render_panel_toggle(
                    st,
                    "延長料金のみ明細",
                    EXTENSION_DETAIL_OPEN_KEY,
                    "部屋の延長（前延長・後延長）と、設備の「延長30分」の明細だけを表示します（技術者は含みません）。",
                )
            period_label = (
                start_ts.strftime(DATE_FMT)
                if start_ts == end_ts
                else f"{start_ts.strftime(DATE_FMT)} 〜 {end_ts.strftime(DATE_FMT)}"
            )
            date_suffix = start_ts.strftime("%Y%m%d") + (
                "" if start_ts == end_ts else f"-{end_ts.strftime('%Y%m%d')}"
            )
            file_stem = f"見積_{date_suffix}"

            diff_args = (prices_df, room_day_df, room_day_key, period_label, list(selected_rooms), date_suffix)
            render_premium_difference(*diff_args)
            render_allday_difference(*diff_args)

            all_df = build_all_details_df(room_df, equipment_df, tech_df, internet_df)
            render_extension_detail(all_df, period_label, list(selected_rooms), date_suffix)

            dl1, dl2 = st.columns(2)
            dl1.download_button(
                "明細をCSVでダウンロード",
                data=build_details_csv(all_df),
                file_name=f"{file_stem}.csv",
                mime="text/csv",
                on_click="ignore",
                width="stretch",
            )
            try:
                pdf_bytes = build_estimate_pdf(
                    title="料金電卓",
                    period=period_label,
                    rooms=list(selected_rooms),
                    totals={
                        "部屋": room_total,
                        "設備": equipment_total,
                        "技術者": tech_total,
                        "インターネット": internet_total,
                    },
                    all_df=all_df,
                )
                dl2.download_button(
                    "見積をPDFでダウンロード",
                    data=pdf_bytes,
                    file_name=f"{file_stem}.pdf",
                    mime="application/pdf",
                    on_click="ignore",
                    width="stretch",
                )
            except Exception as e:
                dl2.warning(f"PDFを作成できませんでした: {e}")

            tab_all, tab_rooms, tab_eq, tab_tech, tab_net = st.tabs(
                ["明細（全部）", "部屋", "設備", "技術者", "インターネット"]
            )

            with tab_rooms:
                show_detail_df(room_df)

            with tab_eq:
                show_detail_df(equipment_df)

            with tab_tech:
                show_detail_df(tech_df)

            with tab_net:
                show_detail_df(internet_df)

            with tab_all:
                show_detail_df(all_df)

if __name__ == "__main__":
    main()
