from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set

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
APP_TITLE = "公会堂料金電卓 （部屋×日編集＋設備＋技術者＋インターネット）"
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

EQUIPMENT_TIME_SLOTS = ["利用なし"] + TIME_SLOTS
TECH_TIME_SLOTS = ["利用なし"] + TIME_SLOTS

# =========================
# マイク/拡声装置：事故防止ルール
# =========================
MIC_NEVER_ROOMS = {"第1会議室", "第2会議室", "第3会議室", "第4会議室", "第5会議室", "第9会議室", "特別室"}
MIC_C_ROOMS = {"大会議室", "小集会室"}
MIC_D_ROOMS = {"第6会議室", "第7会議室", "第8会議室"}

PA_A_ID = "pa_a"
PA_B_ID = "pa_b"
PA_C_ID = "pa_c"
PA_D_ID = "pa_d"

MIC_WIRED_ID = "mic_wired"
MIC_WIRELESS_ID = "mic_wireless"
MIC_STAND_ID = "mic_stand"

HALLBIG_WIRELESS_ID = "wireless_a"
HALLBIG_WIRED_ID = "hallbig_wired_mic_a"
HALLBIG_STAND_ID = "hallbig_mic_stand_a"

MID_WIRELESS_ID = "mid_wireless_mic_a"
MID_WIRED_ID = "mid_wired_mic_a"
MID_STAND_ID = "mid_mic_stand_a"

# C/D判定用（会議室系のみ）
MIC_ITEMS = {MIC_WIRED_ID, MIC_WIRELESS_ID}

# 課金数量調整用（全マイク / 全スタンド）
ALL_MIC_ITEMS = {
    MIC_WIRED_ID,
    MIC_WIRELESS_ID,
    HALLBIG_WIRELESS_ID,
    HALLBIG_WIRED_ID,
    MID_WIRELESS_ID,
    MID_WIRED_ID,
}
STAND_ITEMS = {
    MIC_STAND_ID,
    HALLBIG_STAND_ID,
    MID_STAND_ID,
}

# 日別対象判定するID（C/D系のみ）
MIC_RELATED_ITEM_IDS = {MIC_WIRED_ID, MIC_WIRELESS_ID, PA_C_ID, PA_D_ID}

# 表示用
PA_ITEMS_WITH_INCLUDED_MIC = {PA_A_ID, PA_B_ID, PA_C_ID, PA_D_ID}
PA_ITEMS_WITH_INCLUDED_STAND = {PA_A_ID, PA_B_ID, PA_C_ID, PA_D_ID}

# PAごとの無料対象ルール
# mic_priority は「ワイヤレス優先 → 有線」
PA_FREE_RULES = {
    PA_A_ID: {
        "mic_priority": [HALLBIG_WIRELESS_ID, HALLBIG_WIRED_ID],
        "stand_priority": [HALLBIG_STAND_ID],
        "free_mic_count": 1,
        "free_stand_count": 1,
    },
    PA_B_ID: {
        "mic_priority": [MID_WIRELESS_ID, MID_WIRED_ID],
        "stand_priority": [MID_STAND_ID],
        "free_mic_count": 1,
        "free_stand_count": 1,
    },
    PA_C_ID: {
        "mic_priority": [MIC_WIRELESS_ID, MIC_WIRED_ID],
        "stand_priority": [MIC_STAND_ID],
        "free_mic_count": 1,
        "free_stand_count": 1,
    },
    PA_D_ID: {
        "mic_priority": [MIC_WIRELESS_ID, MIC_WIRED_ID],
        "stand_priority": [MIC_STAND_ID],
        "free_mic_count": 1,
        "free_stand_count": 1,
    },
}

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


def _apply_free_priority(
    current_qty_map: Dict[str, int],
    priority_ids: List[str],
    free_count: int,
) -> Dict[str, int]:
    deducted: Dict[str, int] = {}
    remain = max(0, int(free_count))

    for iid in priority_ids:
        current_qty = int(current_qty_map.get(iid, 0))
        if remain <= 0 or current_qty <= 0:
            deducted[iid] = 0
            continue

        use = min(remain, current_qty)
        current_qty_map[iid] = current_qty - use
        deducted[iid] = use
        remain -= use

    return deducted


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


def _fix_tech_slot(v: object) -> str:
    s = normalize_str(v)
    if s == "" or s.lower() == "none":
        return "利用なし"
    if s not in TECH_TIME_SLOTS:
        return "利用なし"
    return s


def _safe_set(s: Set[str]) -> Set[str]:
    return set([x for x in s if x])


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

    # ---- 数量マップ作成 ----
    qty_map: Dict[str, int] = {}
    for s in selections:
        iid = s.get("item_id")
        q = int(s.get("qty", 0) or 0)
        if not iid:
            continue
        qty_map[iid] = qty_map.get(iid, 0) + max(0, q)

    # ---- 課金数量初期値 ----
    billed_qty_override: Dict[str, int] = {}
    for iid, q in qty_map.items():
        billed_qty_override[iid] = int(q)

    deducted_note: Dict[str, int] = {}
    for iid in qty_map.keys():
        deducted_note[iid] = 0

    # 既知の対象IDも初期化
    for iid in ALL_MIC_ITEMS | STAND_ITEMS:
        billed_qty_override.setdefault(iid, 0)
        deducted_note.setdefault(iid, 0)

    # ---- PAごとの無料マイク・無料スタンド控除 ----
    for pa_id, rule in PA_FREE_RULES.items():
        pa_count = int(qty_map.get(pa_id, 0))
        if pa_count <= 0:
            continue

        mic_deducted = _apply_free_priority(
            billed_qty_override,
            rule["mic_priority"],
            pa_count * int(rule["free_mic_count"]),
        )
        for iid, d in mic_deducted.items():
            deducted_note[iid] = deducted_note.get(iid, 0) + int(d)

        stand_deducted = _apply_free_priority(
            billed_qty_override,
            rule["stand_priority"],
            pa_count * int(rule["free_stand_count"]),
        )
        for iid, d in stand_deducted.items():
            deducted_note[iid] = deducted_note.get(iid, 0) + int(d)

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

        # 日別可否（C/D系のみ）
        forced_zero = False
        if iid in MIC_RELATED_ITEM_IDS:
            allowed_today = _is_mic_related_item_allowed_today(iid, requires_context, mic_allowed_today)
            if not allowed_today:
                forced_zero = True

        # 課金数量
        if forced_zero:
            billed_qty = 0
        elif iid in ALL_MIC_ITEMS or iid in STAND_ITEMS:
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
        if iid in PA_ITEMS_WITH_INCLUDED_MIC or iid in PA_ITEMS_WITH_INCLUDED_STAND:
            note = (note + " / " if note else "") + "マイク1本・スタンド1本付属"

        if iid in ALL_MIC_ITEMS:
            ded = int(deducted_note.get(iid, 0))
            if ded > 0:
                note = (note + " / " if note else "") + f"付属マイク控除:{ded}（ワイヤレス優先）"
            note = (note + " / " if note else "") + f"選択:{orig_qty}→課金:{billed_qty}"

        if iid in STAND_ITEMS:
            ded = int(deducted_note.get(iid, 0))
            if ded > 0:
                note = (note + " / " if note else "") + f"付属スタンド控除:{ded}"
            note = (note + " / " if note else "") + f"選択:{orig_qty}→課金:{billed_qty}"

        if iid in MIC_RELATED_ITEM_IDS and forced_zero:
            note = (note + " / " if note else "") + "対象外日（当日は計算対象外）"

        # billed_qty=0でも、選択していた事実は残す（マイク/スタンド/PAのみ）
        if billed_qty == 0 and orig_qty > 0 and (iid in MIC_RELATED_ITEM_IDS or iid in ALL_MIC_ITEMS or iid in STAND_ITEMS):
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
INTERNET_TEMP_LINE_BASE = 5000

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


def build_room_day_base(days_df: pd.DataFrame, selected_rooms: List[str], default_room_slot: str) -> pd.DataFrame:
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
                    "延長": "なし",
                    "割増利用": bool(day_business.get(date_str, False)),
                    "手動区分": False,
                    "手動延長": False,
                    "手動割増": False,
                }
            )
    df = pd.DataFrame(rows)
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
) -> pd.DataFrame:
    base = build_room_day_base(days_df, selected_rooms, default_room_slot)
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

    key_cols = ["日付", "部屋"]
    keep_cols = key_cols + ["区分", "延長", "割増利用", "手動区分", "手動延長", "手動割増"]
    cur_small = cur[keep_cols].copy()

    merged = base.merge(cur_small, on=key_cols, how="left", suffixes=("", "_old"))

    merged["手動区分"] = merged["手動区分_old"].fillna(merged["手動区分"]).astype(bool)
    merged["手動延長"] = merged["手動延長_old"].fillna(merged["手動延長"]).astype(bool)
    merged["手動割増"] = merged["手動割増_old"].fillna(merged["手動割増"]).astype(bool)

    merged.loc[merged["手動区分"], "区分"] = merged.loc[merged["手動区分"], "区分_old"].fillna(
        merged.loc[merged["手動区分"], "区分"]
    )
    merged.loc[merged["手動延長"], "延長"] = merged.loc[merged["手動延長"], "延長_old"].fillna(
        merged.loc[merged["手動延長"], "延長"]
    )
    merged.loc[merged["手動割増"], "割増利用"] = merged.loc[merged["手動割増"], "割増利用_old"].fillna(
        merged.loc[merged["手動割増"], "割増利用"]
    )

    drop_cols = [c for c in merged.columns if c.endswith("_old")]
    if drop_cols:
        merged.drop(columns=drop_cols, inplace=True)

    if not merged.empty:
        merged["区分"] = merged["区分"].apply(lambda x: _fix_room_slot(x, default_room_slot))
        merged["延長"] = merged["延長"].apply(_fix_room_extension)
        merged["割増利用"] = merged["割増利用"].astype(bool)
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
            rows.append({"日付": dts.date(), "種別": "部屋", "品目": room, "区分": "利用なし", "割増": is_business, "単価": 0, "小計": 0, "備考": note})
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
            rows.append({"日付": dts.date(), "種別": "部屋", "品目": room, "区分": slot, "割増": is_business, "単価": amount, "小計": amount, "備考": ""})

        if ext != "なし":
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
        return 0, pd.DataFrame(columns=["日付", "種別", "グループ", "品目", "課金タイプ", "区分", "数量", "単価(1区分)", "倍率", "区分小計", "一回課金", "小計", "備考", "自動追加"])

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


def calc_internet_total(
    room_day_df: pd.DataFrame,
    use_pocket_wifi: bool,
    use_fixed_line: bool,
    use_temp_line: bool,
) -> Tuple[int, pd.DataFrame]:
    active_dates = [parse_date_str(d) for d in active_dates_from_room_day(room_day_df)]
    active_dates = [d for d in active_dates if d is not None]
    if not active_dates:
        return 0, pd.DataFrame(columns=["日付", "種別", "品目", "フロア", "小計", "備考"])

    rows = []
    total = 0

    if use_pocket_wifi:
        for d in sorted(active_dates):
            rows.append({"日付": d.date(), "種別": "インターネット", "品目": "ポケットWi-Fi貸出", "フロア": "全部屋", "小計": INTERNET_POCKET_WIFI_PER_DAY, "備考": "先着順/同時接続目安5台/電波不安定の可能性"})
            total += INTERNET_POCKET_WIFI_PER_DAY

    floors = infer_active_days_by_floor(room_day_df)

    if use_fixed_line:
        for floor_label, ds in floors.items():
            blocks = _split_consecutive_blocks(ds)
            for b in blocks:
                if not b:
                    continue
                rows.append({"日付": b[0].date(), "種別": "インターネット", "品目": "常設回線（初日）", "フロア": floor_label, "小計": INTERNET_FIXED_FIRST_DAY, "備考": "連続利用の段階料金"})
                total += INTERNET_FIXED_FIRST_DAY
                for d in b[1:]:
                    rows.append({"日付": d.date(), "種別": "インターネット", "品目": "常設回線（2日目以降）", "フロア": floor_label, "小計": INTERNET_FIXED_AFTER_DAY, "備考": "連続利用の段階料金"})
                    total += INTERNET_FIXED_AFTER_DAY

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
        return f"¥{int(x):,}"
    except Exception:
        return f"¥{x}"


def inject_ui_css():
    st.markdown(
        """
<style>
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
  grid-template-columns: repeat(5, minmax(160px, 1fr));
  gap: 12px;
  margin: 8px 0 8px 0;
}
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
@media (max-width: 1100px) {
  .oai-kpi-row { grid-template-columns: repeat(3, minmax(160px, 1fr)); }
}
@media (max-width: 640px) {
  .oai-grand { font-size: 32px; }
  .oai-kpi-row { grid-template-columns: repeat(2, minmax(150px, 1fr)); }
  .oai-kpi-val { font-size: 22px; }
}
</style>
        """,
        unsafe_allow_html=True,
    )


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
        ("部屋代 合計", yen(room_total), ""),
        ("設備 合計", yen(equipment_total), ""),
        ("技術者 合計", yen(tech_total), ""),
        ("インターネット 合計", yen(internet_total), ""),
        ("総額", yen(grand_total), "内訳は上部に表示", "total"),
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


# =========================
# 休館日情報を収集するヘルパー
# =========================
def _collect_closed_day_dates(days_df: pd.DataFrame) -> List[str]:
    """日別設定から休館日の日付文字列リストを返す"""
    if days_df is None or days_df.empty:
        return []
    mask = days_df["休館日"] == True
    return days_df.loc[mask, "日付"].astype(str).tolist()


# =========================
# Main App
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    inject_ui_css()

    # --- 修正①: 計算済みフラグがある場合のみ sticky を表示 ---
    if st.session_state.get("calc_done", False) and "last_totals" in st.session_state:
        tt = st.session_state["last_totals"]
        render_totals_sticky(tt["room"], tt["equip"], tt["tech"], tt["net"])

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
        st.subheader("1) 期間・部屋（部屋×日テーブル編集）")

        col_a, col_b = st.columns(2)
        with col_a:
            start_date = st.date_input(
                "開始日",
                value=pd.Timestamp.today().date(),
                key="start_date",
            )
        with col_b:
            end_date = st.date_input(
                "終了日",
                value=pd.Timestamp.today().date(),
                key="end_date",
            )

        if start_date is None:
            start_date = pd.Timestamp.today().date()
        if end_date is None:
            end_date = pd.Timestamp.today().date()

        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)
        days = build_date_range(start_ts, end_ts)
        if not days:
            st.error("日付範囲が不正です（終了日が開始日より前です）。")
            st.stop()

        room_candidates = sorted(prices_df["room"].unique().tolist())

        rooms_selected = st.multiselect("部屋（複数選択可）", room_candidates, default=[])
        # --- 修正②: 部屋未選択時は全部屋ではなく空リストにする ---
        selected_rooms = rooms_selected  # 空リストのまま保持

        default_room_slot = st.selectbox(
            "部屋の区分（新規追加の初期値）",
            ROOM_SLOTS_WITH_NONE,
            index=ROOM_SLOTS_WITH_NONE.index("全日") if "全日" in ROOM_SLOTS_WITH_NONE else 0,
        )

        is_business_default = st.checkbox("割増利用（デフォルト）", value=False)

        st.divider()
        st.subheader("日別設定（設備・技術者・インターネットに使用）")
        st.caption("部屋料金は「部屋×日テーブル」の内容が計算の根拠になります。")

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
            st.warning("この環境では日別編集UIが利用できないため、日別設定は表示のみになります。")
            edited_days = st.session_state[days_key]
            st.dataframe(edited_days, use_container_width=True)

        if not edited_days.empty and bool(edited_days["休館日"].any()):
            closed_list = edited_days.loc[edited_days["休館日"] == True, "日付"].astype(str).tolist()
            closed_list = [d for d in closed_list if d]
            msg = " / ".join(closed_list[:20])
            suffix = "…" if len(closed_list) > 20 else ""
            st.error(f"休館日があります：{msg}{suffix}")

        st.divider()
        st.subheader("部屋×日 テーブル（個別調整）")
        st.caption("表を編集した場合は「この表の変更を反映（確定）」を押してから計算してください。")

        room_day_key = f"room_day_{start_date}_{end_date}"

        # --- 修正②: 部屋未選択の場合はテーブルを空にする ---
        if not selected_rooms:
            st.session_state[room_day_key] = pd.DataFrame()
            st.info("部屋を選択してください。")
        else:
            if room_day_key not in st.session_state:
                st.session_state[room_day_key] = build_room_day_base(edited_days, list(selected_rooms), default_room_slot)
            else:
                st.session_state[room_day_key] = merge_room_day(st.session_state[room_day_key], edited_days, list(selected_rooms), default_room_slot)

        # 部屋が選択されている場合のみフィルターとテーブル編集を表示
        if selected_rooms:
            st.markdown("### フィルター")
            f1, f2 = st.columns([1, 1])
            all_dates = sorted(st.session_state[room_day_key]["日付"].unique().tolist()) if not st.session_state[room_day_key].empty else []
            date_filter = f1.multiselect("日付（未選択＝全日）", options=all_dates, default=[], key=f"filter_dates_{room_day_key}")
            all_rooms_in_table = sorted(st.session_state[room_day_key]["部屋"].unique().tolist()) if not st.session_state[room_day_key].empty else []
            room_filter = f2.multiselect("部屋（未選択＝全部屋）", options=all_rooms_in_table, default=[], key=f"filter_rooms_{room_day_key}")

            view_df = st.session_state[room_day_key].copy()
            if date_filter:
                view_df = view_df[view_df["日付"].isin(date_filter)]
            if room_filter:
                view_df = view_df[view_df["部屋"].isin(room_filter)]
            view_df = view_df.reset_index(drop=True)

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
                            "延長": st.column_config.SelectboxColumn(options=ROOM_EXTENSION_SLOTS),
                            "割増利用": st.column_config.CheckboxColumn(),
                            "手動区分": st.column_config.CheckboxColumn(disabled=True),
                            "手動延長": st.column_config.CheckboxColumn(disabled=True),
                            "手動割増": st.column_config.CheckboxColumn(disabled=True),
                        },
                    )
                    submitted = st.form_submit_button("この表の変更を反映（確定）")

                if submitted:
                    updated_full = apply_room_day_edits(st.session_state[room_day_key], edited_room_day_tmp, default_room_slot)
                    st.session_state[room_day_key] = updated_full
                    st.success("反映しました。続けて「計算する」を押してください。")

            except Exception:
                st.warning("この環境では部屋×日編集UIが利用できないため、表示のみになります。")
                st.dataframe(view_df, use_container_width=True)

    with right:
        st.subheader("2) 設備・技術者・インターネット（入力）")

        room_day_df = st.session_state.get(room_day_key, pd.DataFrame())
        if room_day_df is None:
            room_day_df = pd.DataFrame()

        rooms_by_day = rooms_used_by_date(room_day_df)

        has_d_days_raw = any(bool(MIC_D_ROOMS.issubset(rs)) for rs in rooms_by_day.values())
        if has_d_days_raw:
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
                        price_txt.append(f"1区分:{it.price_per_slot:,}円")
                    if it.price_once_yen > 0:
                        price_txt.append(f"単価:{it.price_once_yen:,}円")
                    price_str = " / ".join(price_txt) if price_txt else "料金未設定"

                    target_rooms = infer_item_target_rooms(it.item_name, it.notes, meta.applies_to_rooms if meta.applies_to_rooms else "*")
                    label = f"{it.item_name}（対象:{target_rooms} / 単位:{it.unit} / {price_str}{_supplement_label(it.notes)}）"
                    help_txt = it.notes if it.notes else None

                    qty = st.number_input(
                        label,
                        min_value=0,
                        value=int(st.session_state.get(qty_key, 0) or 0),
                        step=1,
                        key=qty_key,
                        help=help_txt,
                    )

                    if int(qty) > 0:
                        base_selections.append({"group_id": it.group_id, "item_id": it.item_id, "qty": int(qty), "auto_added": False})

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
        tech_people = st.number_input("人数", min_value=0, value=0, step=1, help="日別設定の「技術者区分」×人数で計算します。")

        st.divider()

        st.markdown("### インターネット")
        use_pocket_wifi = st.checkbox("ポケットWi-Fi（2,800円/日）", value=False)
        use_fixed_line = st.checkbox("常設回線（初日18,000円、2日目以降2,000円）", value=False)
        use_temp_line = st.checkbox("仮設回線（5,000円/回 + 別途見積）", value=False)

        st.divider()

        do_calc = st.button("計算する", type="primary")

        if do_calc:
            # --- 修正②: 部屋未選択時は計算を止める ---
            if not selected_rooms:
                st.warning("部屋を選択してください。左側の「部屋」欄から1つ以上の部屋を選んでから計算してください。")
                # 計算済みフラグをクリア
                st.session_state["calc_done"] = False
                st.stop()

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
                use_pocket_wifi=use_pocket_wifi,
                use_fixed_line=use_fixed_line,
                use_temp_line=use_temp_line,
            )

            st.subheader("結果")

            # --- 修正①: 計算済みフラグをセット ---
            st.session_state["calc_done"] = True
            st.session_state["last_totals"] = {
                "room": room_total,
                "equip": equipment_total,
                "tech": tech_total,
                "net": internet_total,
            }

            render_totals_sticky(room_total, equipment_total, tech_total, internet_total)

            # --- 修正③: 休館日がある場合、結果エリアにメッセージを表示 ---
            closed_day_dates = _collect_closed_day_dates(edited_days)
            if closed_day_dates:
                closed_msg = " / ".join(closed_day_dates[:20])
                suffix = "…" if len(closed_day_dates) > 20 else ""
                st.info(f"以下の休館日は計算から除外されています：{closed_msg}{suffix}")

            render_kpis_cards(room_total, equipment_total, tech_total, internet_total)

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
                    st.info("明細がありません（部屋×日が全て「利用なし」など）。")


if __name__ == "__main__":
    main()
