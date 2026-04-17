from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd


ROLE_MEAL = "meal"
ROLE_DAILY = "daily"
ROLE_SUMMARY = "summary"
VALID_ROLES = {ROLE_MEAL, ROLE_DAILY, ROLE_SUMMARY}


class MealValidationError(Exception):
    pass


def _to_plain_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    records = df.to_dict(orient="records")
    return [{k: _to_json_value(v) for k, v in row.items()} for row in records]


def _to_json_value(v: Any) -> Any:
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.isoformat()
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v


def _norm_text(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip().lower()
    return "".join(s.split())


def _coerce_num(s: pd.Series) -> pd.Series:
    cleaned = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace({"": "0", "nan": "0", "none": "0", "-": "0"})
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def _coerce_day(series: pd.Series) -> pd.Series:
    raw = series.astype(str).str.strip()

    dt = pd.to_datetime(raw, errors="coerce")
    out = dt.dt.day

    numeric = pd.to_numeric(raw, errors="coerce")
    serial_day = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce").dt.day
    out = out.fillna(serial_day)

    day_from_korean = raw.str.extract(r"(\d{1,2})\s*일", expand=False)
    day_from_korean = pd.to_numeric(day_from_korean, errors="coerce")
    out = out.fillna(day_from_korean)

    plain_day = pd.to_numeric(raw, errors="coerce")
    plain_day = plain_day.where((plain_day >= 1) & (plain_day <= 31))
    out = out.fillna(plain_day)

    return out.astype("Int64")


def _coerce_datetime(series: pd.Series) -> pd.Series:
    raw = series.astype(str).str.strip()
    dt = pd.to_datetime(raw, errors="coerce")
    numeric = pd.to_numeric(raw, errors="coerce")
    serial_dt = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    return dt.fillna(serial_dt)


def _read_workbook(path: str | Path) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(path, sheet_name=None, header=None, dtype=object)


def _sheet_name_score(sheet_name: str, role: str) -> int:
    n = _norm_text(sheet_name)
    if role == ROLE_DAILY:
        return int("식수관리상세" in n) * 6 + int("일별" in n) * 3 + int("상세" in n) * 2
    if role == ROLE_SUMMARY:
        return int("식수관리" in n and "상세" not in n) * 6 + int("합계" in n) * 3
    if role == ROLE_MEAL:
        return int("한성" in n) * 3 + int("식수" in n) * 4
    return 0


def _role_tokens(role: str) -> List[str]:
    if role == ROLE_DAILY:
        return ["인증일시", "사원번호", "사원이름", "중식", "석식", "조식", "야식2"]
    if role == ROLE_SUMMARY:
        return ["사원번호", "사원이름", "중식", "석식", "합계"]
    if role == ROLE_MEAL:
        return ["구분", "요일", "중식", "석식", "수기", "합계"]
    return []


def _sheet_content_score(df: pd.DataFrame, role: str) -> int:
    sample = df.iloc[:20, :40].fillna("").astype(str)
    text = " ".join(sample.values.ravel())
    text_norm = _norm_text(text)
    score = 0
    for t in _role_tokens(role):
        if _norm_text(t) in text_norm:
            score += 2
    if role == ROLE_DAILY and "인증일시" in text and "중식" in text and "석식" in text:
        score += 4
    if role == ROLE_SUMMARY and "합계" in text and "사원번호" in text:
        score += 4
    if role == ROLE_MEAL and "수기" in text and "요일" in text:
        score += 4
    return score


def _score_file_for_role(path: str | Path, role: str) -> int:
    wb = _read_workbook(path)
    name_hint = _norm_text(Path(path).name)
    score = 0
    if role == ROLE_DAILY and "일별" in name_hint:
        score += 3
    if role == ROLE_SUMMARY and "합계" in name_hint:
        score += 3
    if role == ROLE_MEAL and "식수" in name_hint:
        score += 3

    best_sheet_score = 0
    for sheet_name, df in wb.items():
        s = _sheet_name_score(sheet_name, role) + _sheet_content_score(df, role)
        if s > best_sheet_score:
            best_sheet_score = s
    return score + best_sheet_score


def detect_file_role(path: str | Path) -> str:
    """
    Detect file role by filename hint + sheet name + sheet column structure.
    Returns one of: "meal", "daily", "summary"
    """
    role_scores = {r: _score_file_for_role(path, r) for r in VALID_ROLES}
    best_role = max(role_scores, key=role_scores.get)
    if role_scores[best_role] <= 0:
        raise MealValidationError(f"역할 판별 실패: {path}")
    return best_role


def _find_header_row(df: pd.DataFrame, expected_tokens: Iterable[str], min_hits: int = 3) -> int:
    expected = [_norm_text(t) for t in expected_tokens]
    best_idx = -1
    best_hits = -1
    for i in range(min(40, len(df))):
        row = df.iloc[i, :80].fillna("").astype(str)
        row_text = " ".join(row.tolist())
        row_norm = _norm_text(row_text)
        hits = sum(1 for t in expected if t and t in row_norm)
        if hits > best_hits:
            best_hits = hits
            best_idx = i
    if best_hits < min_hits:
        raise MealValidationError(f"헤더 행 탐지 실패 (hits={best_hits}, min={min_hits})")
    return best_idx


def _make_unique_columns(cols: List[Any]) -> List[str]:
    out: List[str] = []
    seen: Dict[str, int] = {}
    for c in cols:
        base = str(c).strip() or "unnamed"
        count = seen.get(base, 0)
        if count == 0:
            out.append(base)
        else:
            out.append(f"{base}_{count}")
        seen[base] = count + 1
    return out


def _prepare_sheet(df_raw: pd.DataFrame, role: str) -> pd.DataFrame:
    tokens = _role_tokens(role)
    header_idx = _find_header_row(df_raw, tokens, min_hits=3)
    header = _make_unique_columns(df_raw.iloc[header_idx].tolist())
    body = df_raw.iloc[header_idx + 1 :].copy()
    body.columns = header
    body = body.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    return body


def _pick_best_sheet(path: str | Path, role: str) -> Tuple[str, pd.DataFrame]:
    wb = _read_workbook(path)
    best_name: Optional[str] = None
    best_score = -1
    best_df: Optional[pd.DataFrame] = None
    for sheet_name, df in wb.items():
        score = _sheet_name_score(sheet_name, role) + _sheet_content_score(df, role)
        if score > best_score:
            best_name = sheet_name
            best_score = score
            best_df = df
    if best_name is None or best_df is None:
        raise MealValidationError(f"시트를 찾을 수 없습니다: {path} ({role})")
    return best_name, _prepare_sheet(best_df, role)


def load_excel_file(path: str | Path) -> pd.DataFrame:
    """
    Load workbook and return a prepared DataFrame
    from the detected role's best matching sheet.
    """
    role = detect_file_role(path)
    _, df = _pick_best_sheet(path, role)
    return df


def _find_col(df: pd.DataFrame, aliases: Iterable[str], required: bool = True) -> Optional[str]:
    norm_cols = {col: _norm_text(col) for col in df.columns}
    alias_norm = [_norm_text(a) for a in aliases]
    for col, norm in norm_cols.items():
        if any(a and a in norm for a in alias_norm):
            return col
    if required:
        raise MealValidationError(f"필수 컬럼 누락: {list(aliases)}")
    return None


def normalize_daily_log(df: pd.DataFrame) -> pd.DataFrame:
    date_col = _find_col(df, ["인증일시", "인증시간", "인증일자", "일시"])
    auth_col = _find_col(df, ["인증번호"], required=False)
    emp_col = _find_col(df, ["사원번호"], required=False)
    name_col = _find_col(df, ["사원이름", "성명", "이름"], required=False)
    b_col = _find_col(df, ["조식"], required=False)
    l_col = _find_col(df, ["중식"])
    d_col = _find_col(df, ["석식"])
    n_col = _find_col(df, ["야식"], required=False)
    s_col = _find_col(df, ["간식"], required=False)
    n2_col = _find_col(df, ["야식2"], required=False)
    p_col = _find_col(df, ["건수"], required=False)

    out = pd.DataFrame()
    out["datetime"] = _coerce_datetime(df[date_col])
    out["day"] = out["datetime"].dt.day.astype("Int64")
    out["auth_id"] = df[auth_col].astype(str).str.strip() if auth_col else ""
    out["employee_id"] = df[emp_col].astype(str).str.strip() if emp_col else out["auth_id"]
    out["employee_name"] = df[name_col].astype(str).str.strip() if name_col else ""
    out["breakfast"] = _coerce_num(df[b_col]) if b_col else 0
    out["lunch"] = _coerce_num(df[l_col])
    out["dinner"] = _coerce_num(df[d_col])
    out["night"] = _coerce_num(df[n_col]) if n_col else 0
    out["snack"] = _coerce_num(df[s_col]) if s_col else 0
    out["night2"] = _coerce_num(df[n2_col]) if n2_col else 0
    out["count_flag"] = _coerce_num(df[p_col]) if p_col else 0

    meal_sum = out[["breakfast", "lunch", "dinner", "night", "snack", "night2"]].sum(axis=1)
    out["total"] = meal_sum.where(out["count_flag"] <= 0, out["count_flag"])
    out = out[out["day"].notna()].copy()
    out["day"] = out["day"].astype(int)
    return out.reset_index(drop=True)


def normalize_monthly_meal_sheet(df: pd.DataFrame) -> pd.DataFrame:
    day_col = _find_col(df, ["일자", "구분", "날짜"])
    dow_col = _find_col(df, ["요일"], required=False)
    lunch_col = _find_col(df, ["중식"])
    dinner_col = _find_col(df, ["석식"])
    manual_col = _find_col(df, ["수기"])
    total_col = _find_col(df, ["합계"])

    out = pd.DataFrame()
    out["day"] = _coerce_day(df[day_col])
    out["weekday"] = df[dow_col].astype(str).str.strip() if dow_col else ""
    out["lunch"] = _coerce_num(df[lunch_col])
    out["dinner"] = _coerce_num(df[dinner_col])
    out["manual"] = _coerce_num(df[manual_col])
    out["total"] = _coerce_num(df[total_col])
    out["calc_total"] = out["lunch"] + out["dinner"] + out["manual"]

    out = out[out["day"].notna()].copy()
    out["day"] = out["day"].astype(int)
    return out.reset_index(drop=True)


def normalize_summary_sheet(df: pd.DataFrame) -> pd.DataFrame:
    auth_col = _find_col(df, ["인증번호"], required=False)
    emp_col = _find_col(df, ["사원번호"], required=False)
    name_col = _find_col(df, ["사원이름", "성명", "이름"], required=False)
    b_col = _find_col(df, ["조식"], required=False)
    l_col = _find_col(df, ["중식"])
    d_col = _find_col(df, ["석식"])
    n_col = _find_col(df, ["야식"], required=False)
    s_col = _find_col(df, ["간식"], required=False)
    n2_col = _find_col(df, ["야식2"], required=False)
    total_col = _find_col(df, ["합계"], required=False)

    out = pd.DataFrame()
    out["auth_id"] = df[auth_col].astype(str).str.strip() if auth_col else ""
    out["employee_id"] = df[emp_col].astype(str).str.strip() if emp_col else out["auth_id"]
    out["employee_name"] = df[name_col].astype(str).str.strip() if name_col else ""
    out["breakfast"] = _coerce_num(df[b_col]) if b_col else 0
    out["lunch"] = _coerce_num(df[l_col])
    out["dinner"] = _coerce_num(df[d_col])
    out["night"] = _coerce_num(df[n_col]) if n_col else 0
    out["snack"] = _coerce_num(df[s_col]) if s_col else 0
    out["night2"] = _coerce_num(df[n2_col]) if n2_col else 0

    if total_col:
        out["total"] = _coerce_num(df[total_col])
    else:
        out["total"] = out[["breakfast", "lunch", "dinner", "night", "snack", "night2"]].sum(axis=1)

    keep = (out["employee_id"].astype(str).str.strip() != "") | (out["employee_name"].astype(str).str.strip() != "")
    out = out[keep].copy().reset_index(drop=True)
    return out


def aggregate_daily_from_log(df_daily: pd.DataFrame) -> pd.DataFrame:
    agg = (
        df_daily.groupby("day", as_index=False)[["lunch", "dinner", "total"]]
        .sum()
        .rename(columns={"lunch": "daily_lunch", "dinner": "daily_dinner", "total": "daily_total"})
    )
    return agg.sort_values("day").reset_index(drop=True)


def aggregate_employee_from_log(df_daily: pd.DataFrame) -> pd.DataFrame:
    key_cols = ["employee_id", "employee_name"]
    agg_cols = ["lunch", "dinner", "total"]
    agg = (
        df_daily.groupby(key_cols, as_index=False)[agg_cols]
        .sum()
        .rename(columns={"lunch": "daily_lunch", "dinner": "daily_dinner", "total": "daily_total"})
    )
    return agg.sort_values(["employee_name", "employee_id"]).reset_index(drop=True)


def compare_daily_vs_meal(daily_agg: pd.DataFrame, meal_df: pd.DataFrame) -> Dict[str, Any]:
    meal_agg = (
        meal_df.groupby("day", as_index=False)[["lunch", "dinner", "manual", "total", "calc_total"]]
        .sum()
        .rename(
            columns={
                "lunch": "meal_lunch",
                "dinner": "meal_dinner",
                "manual": "meal_manual",
                "total": "meal_total",
                "calc_total": "meal_calc_total",
            }
        )
    )

    merged = daily_agg.merge(meal_agg, on="day", how="outer")
    merged = merged.fillna(0)
    merged["day"] = merged["day"].astype(int)

    rows: List[Dict[str, Any]] = []
    mismatches: List[Dict[str, Any]] = []
    for _, r in merged.sort_values("day").iterrows():
        issues: List[str] = []
        if r["daily_lunch"] != r["meal_lunch"]:
            issues.append("중식 불일치")
        if r["daily_dinner"] != r["meal_dinner"]:
            issues.append("석식 불일치")
        if r["meal_total"] != r["meal_calc_total"]:
            issues.append("식수 합계식 불일치(G != D+E+F)")
        match = len(issues) == 0
        row = {
            "day": int(r["day"]),
            "daily_lunch": float(r["daily_lunch"]),
            "meal_lunch": float(r["meal_lunch"]),
            "daily_dinner": float(r["daily_dinner"]),
            "meal_dinner": float(r["meal_dinner"]),
            "meal_manual": float(r["meal_manual"]),
            "meal_total": float(r["meal_total"]),
            "meal_calc_total": float(r["meal_calc_total"]),
            "match": match,
            "issues": issues,
        }
        rows.append(row)
        if not match:
            mismatches.append(row)

    daily_total = float(daily_agg["daily_total"].sum())
    meal_lunch_total = float(meal_agg["meal_lunch"].sum())
    meal_dinner_total = float(meal_agg["meal_dinner"].sum())
    meal_manual_total = float(meal_agg["meal_manual"].sum())
    meal_final_total = float(meal_agg["meal_total"].sum())
    meal_without_manual_total = meal_lunch_total + meal_dinner_total

    manual_gap_value = meal_final_total - daily_total
    expected_gap_from_manual = meal_manual_total
    manual_gap_explained = manual_gap_value == expected_gap_from_manual
    formula_ok = all(x["meal_total"] == x["meal_calc_total"] for x in rows)

    return {
        "daily_rows": rows,
        "daily_mismatches": mismatches,
        "totals": {
            "daily_log_total": daily_total,
            "meal_lunch_total": meal_lunch_total,
            "meal_dinner_total": meal_dinner_total,
            "meal_manual_total": meal_manual_total,
            "meal_final_total": meal_final_total,
            "meal_without_manual_total": meal_without_manual_total,
        },
        "integrity_checks": {
            "meal_row_total_formula_ok": formula_ok,
            "manual_gap_value": manual_gap_value,
            "expected_gap_from_manual": expected_gap_from_manual,
            "manual_gap_explained": manual_gap_explained,
            "daily_vs_meal_match": len(mismatches) == 0 and manual_gap_explained,
        },
    }


def compare_employee_vs_summary(emp_agg: pd.DataFrame, summary_df: pd.DataFrame) -> Dict[str, Any]:
    summary_agg = (
        summary_df.groupby(["employee_id", "employee_name"], as_index=False)[["lunch", "dinner", "total"]]
        .sum()
        .rename(columns={"lunch": "summary_lunch", "dinner": "summary_dinner", "total": "summary_total"})
    )

    merged = emp_agg.merge(summary_agg, on=["employee_id", "employee_name"], how="outer").fillna(0)
    rows: List[Dict[str, Any]] = []
    mismatches: List[Dict[str, Any]] = []

    for _, r in merged.iterrows():
        issues: List[str] = []
        if r["daily_lunch"] != r["summary_lunch"]:
            issues.append("중식 불일치")
        if r["daily_dinner"] != r["summary_dinner"]:
            issues.append("석식 불일치")
        if r["daily_total"] != r["summary_total"]:
            issues.append("합계 불일치")
        match = len(issues) == 0
        row = {
            "employee_id": str(r["employee_id"]),
            "employee_name": str(r["employee_name"]),
            "daily_lunch": float(r["daily_lunch"]),
            "summary_lunch": float(r["summary_lunch"]),
            "daily_dinner": float(r["daily_dinner"]),
            "summary_dinner": float(r["summary_dinner"]),
            "daily_total": float(r["daily_total"]),
            "summary_total": float(r["summary_total"]),
            "match": match,
            "issues": issues,
        }
        rows.append(row)
        if not match:
            mismatches.append(row)

    return {
        "employee_rows": rows,
        "employee_mismatches": mismatches,
        "daily_vs_summary_match": len(mismatches) == 0,
        "summary_sheet_total": float(summary_agg["summary_total"].sum()),
    }


def _infer_period_info(daily_norm: pd.DataFrame, meal_norm: pd.DataFrame) -> Dict[str, Any]:
    period_year: Optional[int] = None
    period_month: Optional[int] = None

    if "datetime" in daily_norm.columns and daily_norm["datetime"].notna().any():
        dt = daily_norm["datetime"].dropna()
        period_year = int(dt.dt.year.mode().iloc[0])
        period_month = int(dt.dt.month.mode().iloc[0])

    days = sorted(set(meal_norm["day"].dropna().astype(int).tolist()) | set(daily_norm["day"].dropna().astype(int).tolist()))
    return {"year": period_year, "month": period_month, "days_in_file": days}


def _basename(path: str | Path) -> str:
    return Path(path).name


def validate_meal_files(
    meal_path: str | Path,
    daily_path: str | Path,
    summary_path: str | Path,
) -> Dict[str, Any]:
    errors: List[str] = []
    try:
        _, daily_raw = _pick_best_sheet(daily_path, ROLE_DAILY)
        _, summary_raw = _pick_best_sheet(summary_path, ROLE_SUMMARY)
        _, meal_raw = _pick_best_sheet(meal_path, ROLE_MEAL)

        daily_norm = normalize_daily_log(daily_raw)
        summary_norm = normalize_summary_sheet(summary_raw)
        meal_norm = normalize_monthly_meal_sheet(meal_raw)

        daily_agg = aggregate_daily_from_log(daily_norm)
        emp_agg = aggregate_employee_from_log(daily_norm)

        cmp_daily = compare_daily_vs_meal(daily_agg, meal_norm)
        cmp_emp = compare_employee_vs_summary(emp_agg, summary_norm)
        period_info = _infer_period_info(daily_norm, meal_norm)

        result: Dict[str, Any] = {
            "success": True,
            "detected_files": {
                "meal": _basename(meal_path),
                "daily": _basename(daily_path),
                "summary": _basename(summary_path),
            },
            "period_info": period_info,
            "summary": {
                "daily_log_total": cmp_daily["totals"]["daily_log_total"],
                "summary_sheet_total": cmp_emp["summary_sheet_total"],
                "meal_lunch_total": cmp_daily["totals"]["meal_lunch_total"],
                "meal_dinner_total": cmp_daily["totals"]["meal_dinner_total"],
                "meal_manual_total": cmp_daily["totals"]["meal_manual_total"],
                "meal_final_total": cmp_daily["totals"]["meal_final_total"],
                "meal_without_manual_total": cmp_daily["totals"]["meal_without_manual_total"],
                "daily_vs_summary_match": cmp_emp["daily_vs_summary_match"],
                "daily_vs_meal_match": cmp_daily["integrity_checks"]["daily_vs_meal_match"],
                "manual_gap_explained": cmp_daily["integrity_checks"]["manual_gap_explained"],
            },
            "daily_comparison_rows": cmp_daily["daily_rows"],
            "employee_comparison_rows": cmp_emp["employee_rows"],
            "daily_mismatches": cmp_daily["daily_mismatches"],
            "employee_mismatches": cmp_emp["employee_mismatches"],
            "integrity_checks": {
                "meal_row_total_formula_ok": cmp_daily["integrity_checks"]["meal_row_total_formula_ok"],
                "manual_gap_value": cmp_daily["integrity_checks"]["manual_gap_value"],
                "expected_gap_from_manual": cmp_daily["integrity_checks"]["expected_gap_from_manual"],
            },
            "errors": [],
        }
        return result
    except Exception as e:
        errors.append(str(e))
        return {
            "success": False,
            "detected_files": {
                "meal": _basename(meal_path),
                "daily": _basename(daily_path),
                "summary": _basename(summary_path),
            },
            "period_info": {"year": None, "month": None, "days_in_file": []},
            "summary": {},
            "daily_comparison_rows": [],
            "employee_comparison_rows": [],
            "daily_mismatches": [],
            "employee_mismatches": [],
            "integrity_checks": {},
            "errors": errors,
        }


def validate_meal_bundle(file_paths: List[str | Path]) -> Dict[str, Any]:
    role_map: Dict[str, str | Path] = {}
    errors: List[str] = []

    if len(file_paths) < 3:
        return {
            "success": False,
            "detected_files": {},
            "period_info": {"year": None, "month": None, "days_in_file": []},
            "summary": {},
            "daily_comparison_rows": [],
            "employee_comparison_rows": [],
            "daily_mismatches": [],
            "employee_mismatches": [],
            "integrity_checks": {},
            "errors": ["파일 수가 부족합니다. 식수/일별/합계 3개 파일이 필요합니다."],
        }

    for p in file_paths:
        try:
            role = detect_file_role(p)
            if role in role_map:
                existing = role_map[role]
                old_score = _score_file_for_role(existing, role)
                new_score = _score_file_for_role(p, role)
                if new_score > old_score:
                    role_map[role] = p
            else:
                role_map[role] = p
        except Exception as e:
            errors.append(f"{Path(p).name}: {e}")

    missing = [r for r in (ROLE_MEAL, ROLE_DAILY, ROLE_SUMMARY) if r not in role_map]
    if missing:
        errors.append(f"역할 판별 실패(누락): {missing}")
        return {
            "success": False,
            "detected_files": {k: _basename(v) for k, v in role_map.items()},
            "period_info": {"year": None, "month": None, "days_in_file": []},
            "summary": {},
            "daily_comparison_rows": [],
            "employee_comparison_rows": [],
            "daily_mismatches": [],
            "employee_mismatches": [],
            "integrity_checks": {},
            "errors": errors,
        }

    result = validate_meal_files(
        meal_path=role_map[ROLE_MEAL],
        daily_path=role_map[ROLE_DAILY],
        summary_path=role_map[ROLE_SUMMARY],
    )

    if errors:
        result["errors"] = list(result.get("errors", [])) + errors
    return result

