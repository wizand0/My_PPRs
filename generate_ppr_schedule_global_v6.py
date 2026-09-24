#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_ppr_schedule_global_v6.py

Глобальный планировщик ППР для листа «ППР».

Идея:
  1. Сначала читаются ВСЕ разделы и их технические ограничения.
  2. Для каждого раздела строятся варианты размещения ТО-2/ТО-3/ТО-4.
  3. Затем выполняется несколько глобальных проходов оптимизации: раздел
     может менять свои специальные ТО, если это уменьшает общую плотность,
     конфликты и отклонения от технических окон.
  4. Остальные месяцы сезона получают ТО-1.

Правила по умолчанию:
  ТО-2 = 2 раза в год
  ТО-3 = 1 раз в год
  ТО-4 = 1 раз в год

Для стандартного набора 2/1/1 целевые точки сезона примерно такие:
  ТО-4 ~ 12% сезона (обычно I квартал)
  ТО-2 ~ 34% (обычно II квартал)
  ТО-3 ~ 60% (обычно III квартал)
  ТО-2 ~ 84% (обычно IV квартал)

Главное правило: специальные ТО (ТО-2/ТО-3/ТО-4) взаимоисключающие
внутри раздела по кварталу — не более одного специального ТО в каждом
календарном квартале. Если технические окна делают это невозможным,
алгоритм допускает вынужденный fallback и выдаёт предупреждение.

Это даёт интервал ТО-4 -> ТО-3 около половины сезона и не создаёт
искусственных цепочек вроде ТО-3 на неделях 18,19,20.

A:C и технические столбцы BM:BX не изменяются.
"""

import math
import random
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

TO_TEXT = {1: "то1", 2: "то2", 3: "то3", 4: "то4"}
TO_FROM_TEXT = {v: k for k, v in TO_TEXT.items()}
HEADER_NUM_RE = re.compile(r"^\d+$")

DEFAULT_N2 = 2
DEFAULT_N3 = 1
DEFAULT_N4 = 1
DEFAULT_DELAY = 0.10
DEFAULT_ROW_DELAY = 0.01
OPT_PASSES = 7
RANDOM_SEED = 20270923
SPECIAL_LOAD_WEIGHT = 22.0
SPECIAL_SPREAD_WEIGHT = 2.5
SPECIAL_MONTH_WEIGHT = 55.0
SPECIAL_MONTH_ZERO_BONUS = 18.0
TO1_LOAD_WEIGHT = 100.0
MIN_SPECIAL_GAP_FLOOR = 7.0
SPECIAL_GAP_WEIGHT = 180.0
MIN_SPECIAL_GAP_FLOOR = 7.0
SPECIAL_GAP_WEIGHT = 180.0
SPECIAL_SEQUENCE_WEIGHT = 3.0


def normalize(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip().lower()


def to_int(v):
    if v in (None, ""):
        return None
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def week_in_window(week, start, end):
    if start is None or end is None:
        return True
    if start <= end:
        return start <= week <= end
    return week >= start or week <= end


def cyclic_week_distance(a, b):
    d = abs(a - b)
    return min(d, 53 - d)


def quarter_of_month(month):
    """Квартал календарного года: 1=янв-мар, 2=апр-июн, 3=июл-сен, 4=окт-дек."""
    return (month - 1) // 3 + 1


def find_week_row_and_data_col(ws):
    best = None
    for r in range(1, min(10, ws.max_row) + 1):
        run = []
        start = None
        for c in range(1, min(ws.max_column, 100) + 1):
            w = to_int(ws.cell(r, c).value)
            if w is not None and 1 <= w <= 53:
                if start is None:
                    start = c
                    run = []
                run.append((c, w))
            else:
                if len(run) >= 10:
                    cand = (len(run), r, start, run[-1][0])
                    if best is None or cand[0] > best[0]:
                        best = cand
                start = None
                run = []
        if len(run) >= 10:
            cand = (len(run), r, start, run[-1][0])
            if best is None or cand[0] > best[0]:
                best = cand
    if best is None:
        raise ValueError('Не удалось найти непрерывную строку недель 1..53 на листе «ППР».')
    return best[1], best[2]


def find_last_week_col(ws, week_row):
    last = None
    for c in range(1, ws.max_column + 1):
        w = to_int(ws.cell(week_row, c).value)
        if w is not None and 1 <= w <= 53:
            last = c
    if last is None:
        raise ValueError("Не найден последний столбец недель.")
    return last


def build_month_map(ws, week_row, first_col, last_col):
    header_row = week_row - 2
    month_names = [
        "январь", "февраль", "март", "апрель", "май", "июнь",
        "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
    ]
    result = {}
    merges = [
        m for m in ws.merged_cells.ranges
        if m.min_row == header_row and m.max_row == header_row
        and m.min_col >= first_col and m.max_col <= last_col
    ]
    merges.sort(key=lambda m: m.min_col)
    if len(merges) == 12:
        for month, m in enumerate(merges, 1):
            for c in range(m.min_col, m.max_col + 1):
                result[c] = month
    else:
        current = None
        for c in range(first_col, last_col + 1):
            v = normalize(ws.cell(header_row, c).value)
            for i, name in enumerate(month_names, 1):
                if v.startswith(name):
                    current = i
                    break
            result[c] = current
    return result


def build_groups(ws, first_col, last_col, config_start):
    headers = []
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a is not None and HEADER_NUM_RE.match(str(a).strip()) and b not in (None, ""):
            headers.append((r, str(b).strip()))

    groups = []
    for i, (header_row, name) in enumerate(headers):
        end = headers[i + 1][0] if i + 1 < len(headers) else ws.max_row + 1
        leaf_rows = list(range(header_row + 1, end))
        def get(offset):
            return ws.cell(header_row, config_start + offset).value
        cfg = {
            "season_start": to_int(get(0)), "season_end": to_int(get(1)),
            "n2": to_int(get(2)), "n3": to_int(get(3)), "n4": to_int(get(4)),
            "to2_pref_start": to_int(get(5)), "to2_pref_end": to_int(get(6)),
            "to3_pref_start": to_int(get(7)), "to3_pref_end": to_int(get(8)),
            "to4_pref_start": to_int(get(9)), "to4_pref_end": to_int(get(10)),
            "locked": get(11) not in (None, ""),
        }
        if cfg["n2"] is None: cfg["n2"] = DEFAULT_N2
        if cfg["n3"] is None: cfg["n3"] = DEFAULT_N3
        if cfg["n4"] is None: cfg["n4"] = DEFAULT_N4
        groups.append({
            "name": name, "header_row": header_row, "leaf_rows": leaf_rows,
            "cfg": cfg, "id": f"{name} [строка {header_row}]",
        })
    return groups


def validate_cfg(g, warnings):
    c = g["cfg"]
    for k in ("n2", "n3", "n4"):
        if c[k] < 0:
            warnings.append(f'{g["id"]}: {k}={c[k]} недопустимо; использовано 0.')
            c[k] = 0
    for p in ("to2", "to3", "to4"):
        a, b = c[f"{p}_pref_start"], c[f"{p}_pref_end"]
        if (a is None) != (b is None):
            warnings.append(f'{g["id"]}: окно {p.upper()} задано неполностью и проигнорировано.')
            c[f"{p}_pref_start"] = c[f"{p}_pref_end"] = None


def prepare_group(g, cols_by_month, warnings):
    c = g["cfg"]
    ss, se = c["season_start"], c["season_end"]
    months = []
    for m in range(1, 13):
        cols = cols_by_month.get(m, [])
        if not cols:
            continue
        if ss is None or se is None or any(week_in_window(w, ss, se) for _, w in cols):
            months.append(m)
    g["months"] = months
    if not months:
        warnings.append(f'{g["id"]}: сезон не содержит ни одного месяца.')
        g["events"] = []
        return

    requested = [4] * c["n4"] + [3] * c["n3"] + [2] * c["n2"]
    if len(requested) > len(months):
        warnings.append(
            f'{g["id"]}: специальных ТО {len(requested)} больше доступных месяцев {len(months)}; '
            'лишние ТО будут отброшены.'
        )
        # Сохраняем более сложные ТО, режем сначала ТО-2.
        while len(requested) > len(months) and 2 in requested:
            requested.remove(2)
        while len(requested) > len(months) and 3 in requested:
            requested.remove(3)
        while len(requested) > len(months) and 4 in requested:
            requested.remove(4)
    g["events"] = requested


def month_candidates(g, month, cols_by_month):
    c = g["cfg"]
    ss, se = c["season_start"], c["season_end"]
    cols = cols_by_month.get(month, [])
    if ss is not None and se is not None:
        filtered = [(col, w) for col, w in cols if week_in_window(w, ss, se)]
        if filtered:
            return filtered
    return list(cols)


def special_phase(g, to_num, occurrence):
    """Целевая доля сезона для специального ТО внутри раздела.

    Для стандартного набора получается цепочка:
    ТО-4 -> ТО-2 -> ТО-3 -> ТО-2.
    Это помогает избежать ТО-4 и ТО-2 в соседних месяцах и длинного
    периода без специальных работ.
    """
    if to_num == 4:
        return 0.12
    if to_num == 3:
        return 0.60
    if to_num == 2:
        return 0.34 if occurrence == 0 else 0.84
    return 0.50


def target_week(g, to_num, occurrence, cols_by_month):
    """Мягкая целевая точка специального ТО внутри сезона."""
    all_weeks = [w for m in g["months"] for _, w in month_candidates(g, m, cols_by_month)]
    if not all_weeks:
        return 1
    lo, hi = min(all_weeks), max(all_weeks)
    return lo + (hi - lo) * special_phase(g, to_num, occurrence)


def min_special_gap(g, cols_by_month):
    """Желательный минимальный интервал между специальными ТО раздела."""
    all_weeks = [w for m in g["months"] for _, w in month_candidates(g, m, cols_by_month)]
    if len(all_weeks) < 2 or len(g["events"]) <= 1:
        return MIN_SPECIAL_GAP_FLOOR
    span = max(all_weeks) - min(all_weeks)
    # Для 4 специальных ТО в сезоне 1-51 получается примерно 8 недель.
    return max(MIN_SPECIAL_GAP_FLOOR, span / (len(g["events"]) + 1) * 0.65)


def preferred_window(g, to_num):
    c = g["cfg"]
    return c[f"to{to_num}_pref_start"], c[f"to{to_num}_pref_end"]

def event_label(to_num, occurrence):
    return f"ТО-{to_num}" + (f" #{occurrence + 1}" if to_num == 2 else "")


def build_initial_assignment(g, cols_by_month, used3, used4, to2_usage):
    """Жадный старт с учётом уже занятых недель."""
    events = list(g["events"])
    # Размещаем сначала ТО-4, затем ТО-3, затем ТО-2.
    order = sorted(enumerate(events), key=lambda x: {4: 0, 3: 1, 2: 2}[x[1]])
    chosen = {}
    used_months = set()
    occ_counter = Counter()
    for idx, to_num in order:
        occ = occ_counter[to_num]
        occ_counter[to_num] += 1
        candidates = candidate_weeks_for_event(g, to_num, occ, cols_by_month, 12)
        best = None
        for _, m, col, week, in_pref in candidates:
            if m in used_months:
                continue
            hard_conflict = week in (used3 if to_num == 3 else used4 if to_num == 4 else set())
            if hard_conflict:
                continue
            score = abs(week - target_week(g, to_num, occ, cols_by_month))
            if to_num in (3, 4):
                score += (0 if in_pref else 25)
            else:
                score += to2_usage.get(week, 0) * 4
            # Не ставим ТО-3/4 рядом с другим сложным ТО этого же раздела.
            for other_idx, (om, ocol, ow) in chosen.items():
                if events[other_idx] in (3, 4) and to_num in (3, 4):
                    score += max(0, 8 - abs(week - ow)) * 8
            if best is None or score < best[0]:
                best = (score, m, col, week)
        if best is None:
            # Выбираем любой допустимый месяц, даже если глобальная неделя занята;
            # это будет отмечено финальной проверкой/предупреждением.
            for _, m, col, week, _ in candidates:
                if m not in used_months:
                    best = (9999, m, col, week)
                    break
        if best:
            _, m, col, week = best
            chosen[idx] = (m, col, week)
            used_months.add(m)
    return chosen


def assignment_score(g, assignment, all_assignments, cols_by_month, week_load3, week_load4, week_load2):
    """Чем меньше, тем лучше."""
    events = g["events"]
    score = 0.0
    specials = []
    for idx, (m, col, week) in assignment.items():
        to_num = events[idx]
        occ = sum(1 for j in range(idx) if events[j] == to_num)
        target = target_week(g, to_num, occ, cols_by_month)
        score += abs(week - target) * (1.0 if to_num == 2 else 1.8)
        ps, pe = preferred_window(g, to_num)
        if ps is not None and not week_in_window(week, ps, pe):
            # Вне технического окна — сильный штраф, но не бесконечный:
            # если окно уже перегружено, планировщик сможет найти компромисс.
            dist = min(abs(week - ps), abs(week - pe))
            score += 35 + dist * 5
        if to_num == 2:
            score += week_load2.get(week, 0) * 8
        elif to_num == 3:
            score += week_load3.get(week, 0) * 10000
            specials.append((3, week))
        elif to_num == 4:
            score += week_load4.get(week, 0) * 10000
            specials.append((4, week))
    # Внутри раздела: сложные ТО должны быть разнесены.
    hard_weeks = [w for t, w in specials]
    if len(hard_weeks) >= 2:
        for i in range(len(hard_weeks)):
            for j in range(i + 1, len(hard_weeks)):
                d = abs(hard_weeks[i] - hard_weeks[j])
                score += max(0, 20 - d) * 25
    # ТО-3/ТО-4 желательно около половины сезона друг от друга.
    w3 = [w for t, w in specials if t == 3]
    w4 = [w for t, w in specials if t == 4]
    if w3 and w4:
        season_weeks = [w for m in g["months"] for _, w in month_candidates(g, m, cols_by_month)]
        span = max(season_weeks) - min(season_weeks) if season_weeks else 48
        desired = max(10, span * 0.50)
        d = abs(w3[0] - w4[0])
        score += abs(d - desired) * 3
    # Нельзя два специальных ТО в одном месяце.
    months = [x[0] for x in assignment.values()]
    weeks = [x[2] for x in assignment.values()]
    score += (len(months) - len(set(months))) * 100000
    score += (len(weeks) - len(set(weeks))) * 100000
    return score


def _special_week_load(used3, used4):
    """Общая текущая загрузка недели сложными ТО."""
    load = Counter()
    for w in used3:
        load[w] += 1
    for w in used4:
        load[w] += 1
    return load


def global_hard_assignment(groups, cols_by_month, used3_locked, used4_locked):
    """Глобально размещает ТО-3/ТО-4 с балансировкой по месяцам и неделям.

    Важное отличие v4: алгоритм сначала стремится получить равномерную
    месячную загрузку крупных ТО по всему году. Технические окна остаются
    жесткими. Для разделов без узких окон свободные месяцы (в том числе
    январь/февраль/декабрь) теперь получают приоритет, если они недогружены.
    """
    events = []
    for g in groups:
        if g["cfg"]["locked"]:
            continue
        occ = Counter()
        for idx, t in enumerate(g["events"]):
            if t not in (3, 4):
                continue
            o = occ[t]
            occ[t] += 1
            ps, pe = preferred_window(g, t)
            candidates = []
            for m in g["months"]:
                for col, week in month_candidates(g, m, cols_by_month):
                    # Техническое окно является жестким, если оно задано.
                    if ps is not None and not week_in_window(week, ps, pe):
                        continue
                    candidates.append((m, col, week))
            if not candidates:
                # Если техническое окно невозможно реализовать, допускаем
                # сезонный fallback и обязательно сообщаем об этом в main().
                for m in g["months"]:
                    for col, week in month_candidates(g, m, cols_by_month):
                        candidates.append((m, col, week))
            target = target_week(g, t, o, cols_by_month)
            events.append({
                "g": g, "idx": idx, "t": t, "occ": o,
                "candidates": candidates, "target": target,
                "pref": (ps, pe),
            })

    # Сначала наиболее ограниченные события. ТО-4 раньше ТО-3 при равном
    # числе кандидатов, потому что его окна в исходной таблице обычно уже.
    events.sort(key=lambda e: (len(e["candidates"]), -e["t"]))

    total_events = len(events) + len(used3_locked) + len(used4_locked)
    # Средняя желаемая загрузка месяца. Это не жесткая квота: технические
    # окна имеют приоритет. При 27 крупных ТО получается около 2.25/месяц.
    desired_month_load = total_events / 12.0 if total_events else 0.0

    # Состояние: стоимость, назначения, занятые месяцы и кварталы разделов,
    # недели ТО-3/ТО-4 и месячная загрузка крупных ТО.
    states = [(0.0, {}, defaultdict(set), defaultdict(set), set(used3_locked),
               set(used4_locked), Counter())]
    BEAM = 250

    for e in events:
        new_states = []
        gid = e["g"]["id"]
        for cost, amap, group_months, group_quarters, u3, u4, month_load in states:
            gm = group_months[gid]
            gq = group_quarters[gid]
            week_load = _special_week_load(u3, u4)
            for m, col, week in e["candidates"]:
                if m in gm:
                    continue
                # Жёсткое правило: не более одного специального ТО в квартале
                # внутри одного раздела. Если это ТО-2, ТО-3 или ТО-4 —
                # правило одинаковое: они взаимоисключающие.
                q = quarter_of_month(m)
                if q in gq:
                    continue
                if e["t"] == 3 and week in u3:
                    continue
                if e["t"] == 4 and week in u4:
                    continue

                # v5: если для этого события есть варианты с нормальным
                # интервалом, не рассматриваем кандидатов, которые ломают
                # минимальный интервал внутри данного раздела.
                prior_weeks = [v[2] for (kgid, _), v in amap.items() if kgid == gid]
                gap_min = min_special_gap(e["g"], cols_by_month)
                if prior_weeks and any(abs(week - pw) < gap_min for pw in prior_weeks):
                    feasible_gap = any(
                        m2 not in gm and
                        quarter_of_month(m2) not in gq and
                        all(abs(w2 - pw) >= gap_min for pw in prior_weeks)
                        for m2, _, w2 in e["candidates"]
                    )
                    if feasible_gap:
                        continue

                # Недельная загрузка.
                new_week_load = week_load[week] + 1
                load_penalty = (new_week_load ** 2) * SPECIAL_LOAD_WEIGHT

                # Месячная загрузка — теперь основной глобальный фактор.
                # Штрафуем отклонение от 2.25 (для 27 событий), но не
                # запрещаем 3-ю работу в месяце.
                new_month_load = month_load[m] + 1
                month_penalty = ((new_month_load - desired_month_load) ** 2
                                  - ((month_load[m] - desired_month_load) ** 2)) * SPECIAL_MONTH_WEIGHT

                # Пока месяц совсем пуст, небольшой бонус помогает заполнить
                # периферийные месяцы, а не ждать, пока центр уже перегружен.
                coverage_bonus = -SPECIAL_MONTH_ZERO_BONUS if month_load[m] == 0 else 0

                target_penalty = abs(week - e["target"]) * SPECIAL_SPREAD_WEIGHT

                # v5: равномерность специальных ТО внутри одного раздела.
                # Не допускаем искусственной цепочки ТО-4 -> ТО-2 и затем
                # длительного периода только с ТО-1.
                prior_weeks = [v[2] for (kgid, _), v in amap.items() if kgid == gid]
                if prior_weeks:
                    gap_min = min_special_gap(e["g"], cols_by_month)
                    for pw in prior_weeks:
                        gap = abs(week - pw)
                        if gap < gap_min:
                            target_penalty += (gap_min - gap) ** 2 * SPECIAL_GAP_WEIGHT

                # Выход за техническое окно возможен только в fallback-случае.
                ps, pe = e["pref"]
                if ps is not None and not week_in_window(week, ps, pe):
                    dist = min(abs(week - ps), abs(week - pe))
                    target_penalty += 1000 + dist * 100

                cross_penalty = (
                    5 if (e["t"] == 3 and week in u4) or
                         (e["t"] == 4 and week in u3) else 0
                )

                c = cost + load_penalty + month_penalty + coverage_bonus \
                    + target_penalty + cross_penalty

                ngm = defaultdict(set, {k: set(v) for k, v in group_months.items()})
                ngm[gid].add(m)
                ngq = defaultdict(set, {k: set(v) for k, v in group_quarters.items()})
                ngq[gid].add(q)
                nu3 = set(u3)
                nu4 = set(u4)
                if e["t"] == 3:
                    nu3.add(week)
                else:
                    nu4.add(week)
                nml = Counter(month_load)
                nml[m] += 1
                nam = dict(amap)
                nam[(gid, e["idx"])] = (m, col, week)
                new_states.append((c, nam, ngm, ngq, nu3, nu4, nml))

        if not new_states:
            # Квартальное правило имеет приоритет, но если технические окна
            # физически не позволяют разместить событие в свободном квартале,
            # не теряем ТО: разрешаем технический fallback и пометим его.
            fallback_states = []
            for cost, amap, group_months, group_quarters, u3, u4, month_load in states:
                gm = group_months[gid]
                gq = group_quarters[gid]
                week_load = _special_week_load(u3, u4)
                for m, col, week in e["candidates"]:
                    if m in gm:
                        continue
                    if e["t"] == 3 and week in u3:
                        continue
                    if e["t"] == 4 and week in u4:
                        continue
                    q = quarter_of_month(m)
                    # Здесь уже нет состояния, где событие удаётся разместить
                    # в свободном квартале. Разрешаем второй спец-ТО в квартале
                    # только как вынужденный fallback.
                    if q in gq:
                        pass
                    # Если дошли сюда, причина невозможности — другие ограничения.
                    new_week_load = week_load[week] + 1
                    new_month_load = month_load[m] + 1
                    c = cost + (new_week_load ** 2) * SPECIAL_LOAD_WEIGHT
                    c += ((new_month_load - desired_month_load) ** 2 -
                          ((month_load[m] - desired_month_load) ** 2)) * SPECIAL_MONTH_WEIGHT
                    c += abs(week - e["target"]) * SPECIAL_SPREAD_WEIGHT
                    ps, pe = e["pref"]
                    if ps is not None and not week_in_window(week, ps, pe):
                        dist = min(abs(week - ps), abs(week - pe))
                        c += 1000 + dist * 100
                    ngm = defaultdict(set, {k: set(v) for k, v in group_months.items()})
                    ngm[gid].add(m)
                    ngq = defaultdict(set, {k: set(v) for k, v in group_quarters.items()})
                    ngq[gid].add(q)
                    nu3, nu4 = set(u3), set(u4)
                    if e["t"] == 3: nu3.add(week)
                    else: nu4.add(week)
                    nml = Counter(month_load); nml[m] += 1
                    nam = dict(amap); nam[(gid, e["idx"])] = (m, col, week)
                    fallback_states.append((c, nam, ngm, ngq, nu3, nu4, nml))
            if not fallback_states:
                # Совсем нет допустимого места — оставляем событие без назначения.
                e["unassigned"] = True
                continue
            e["quarter_fallback"] = True
            new_states = fallback_states
        new_states.sort(key=lambda x: x[0])
        unique = {}
        for st in new_states:
            # Сохраняем разные месячные профили и недельные профили.
            key = (
                tuple(sorted(st[4])), tuple(sorted(st[5])),
                tuple(sorted(st[6].items())),
                tuple(sorted((k, v[0]) for k, v in st[1].items())),
            )
            if key not in unique:
                unique[key] = st
        states = list(unique.values())[:BEAM]

    if not states:
        return {}
    best = min(states, key=lambda x: x[0])
    result = defaultdict(dict)
    for (gid, idx), value in best[1].items():
        result[gid][idx] = value
    return dict(result)

def assign_soft_to2(groups, cols_by_month, hard_assignments, used2_locked):
    """Глобально распределяет ТО-2 с учетом уже размещенных ТО-3/ТО-4.

    В v4 месячная нагрузка считается общей для всех специальных ТО. Поэтому
    ТО-2 не заполняет середину года только потому, что там находятся удобные
    целевые недели: оно старается дополнить недогруженные месяцы.
    """
    events = []
    for g in groups:
        if g["cfg"]["locked"]:
            continue
        hard = hard_assignments.get(g["id"], {})
        used_months = {v[0] for v in hard.values()}
        occ = 0
        for idx, t in enumerate(g["events"]):
            if t != 2:
                continue
            candidates = []
            target = target_week(g, 2, occ, cols_by_month)
            ps, pe = preferred_window(g, 2)
            for m in g["months"]:
                if m in used_months:
                    continue
                for col, week in month_candidates(g, m, cols_by_month):
                    in_pref = ps is None or week_in_window(week, ps, pe)
                    candidates.append((m, col, week, in_pref))
            events.append((g, idx, occ, candidates, target))
            occ += 1

    # Чем меньше доступных месяцев, тем раньше размещаем ТО-2.
    events.sort(key=lambda x: len(x[3]))

    load = Counter(used2_locked)
    month_load = Counter()
    for amap in hard_assignments.values():
        for m, _, _ in amap.values():
            month_load[m] += 1
    for w in used2_locked:
        # locked_to2 содержит только недельную информацию; месяц будет
        # восстановлен ниже из cols_by_month.
        for m, cols in cols_by_month.items():
            if any(ww == w for _, ww in cols):
                month_load[m] += 1
                break

    total_special = sum(len(a) for a in hard_assignments.values()) + len(events) + sum(used2_locked.values())
    desired_month = total_special / 12.0 if total_special else 0.0

    result = defaultdict(dict)
    used_months_by_group = {
        g["id"]: {v[0] for v in hard_assignments.get(g["id"], {}).values()}
        for g in groups
    }
    used_quarters_by_group = {
        g["id"]: {quarter_of_month(v[0]) for v in hard_assignments.get(g["id"], {}).values()}
        for g in groups
    }

    for g, idx, occ, candidates, target in events:
        if not candidates:
            continue
        available = [x for x in candidates
                     if x[0] not in used_months_by_group[g["id"]]
                     and quarter_of_month(x[0]) not in used_quarters_by_group[g["id"]]]
        quarter_fallback = False
        if not available:
            # Вынужденный fallback: если технические окна/сезон не дают
            # свободного квартала, лучше сохранить требуемое ТО и сообщить об этом.
            available = [x for x in candidates if x[0] not in used_months_by_group[g["id"]]]
            quarter_fallback = True
        if not available:
            available = candidates
        ps, pe = preferred_window(g, 2)

        # v5: сначала пытаемся оставить только кандидатов с нормальным
        # интервалом до уже поставленных специальных ТО этого раздела.
        prior_weeks = [v[2] for v in hard_assignments.get(g["id"], {}).values()]
        prior_weeks += [v[2] for v in result.get(g["id"], {}).values()]
        gap_min = min_special_gap(g, cols_by_month)
        if prior_weeks:
            # ВАЖНО: квартальное правило жёстче, чем правило интервала.
            # Сначала оставляем только свободные кварталы, затем среди них
            # стараемся найти хороший интервал. Если хороший интервал не
            # получается, выбираем любой свободный квартал, но НЕ нарушаем
            # квартальную взаимоисключаемость.
            gap_ok = [x for x in available if all(abs(x[2] - pw) >= gap_min for pw in prior_weeks)]
            if gap_ok:
                available = gap_ok

        # Защитная проверка: после любых фильтров нельзя вернуть занятый квартал.
        occupied_quarters = used_quarters_by_group[g["id"]]
        quarter_safe = [x for x in available if quarter_of_month(x[0]) not in occupied_quarters]
        if quarter_safe:
            available = quarter_safe
        elif not quarter_fallback:
            # Теоретически сюда не должны попадать, но если попали — это
            # означает, что текущий набор кандидатов исчерпал свободные кварталы.
            # Восстанавливаем кандидатов и фиксируем вынужденный fallback.
            available = [x for x in candidates if x[0] not in used_months_by_group[g["id"]]]
            quarter_fallback = True

        def score(x):
            m, col, week, in_pref = x
            old = month_load[m]
            new = old + 1
            # Прирост квадратичного штрафа относительно средней месячной цели.
            month_penalty = ((new - desired_month) ** 2 -
                             (old - desired_month) ** 2) * 42.0
            # Небольшой бонус полностью пустому месяцу: это помогает
            # заполнить январь/февраль/декабрь, когда они технически доступны.
            coverage = -14.0 if old == 0 else 0.0
            week_penalty = load[week] * 24.0
            target_penalty = abs(week - target) * 0.8

            # v5: ТО-2 тоже входит в общую последовательность специальных ТО.
            prior = [v[2] for v in hard_assignments.get(g["id"], {}).values()]
            prior += [v[2] for v in result.get(g["id"], {}).values()]
            if prior:
                gap_min = min_special_gap(g, cols_by_month)
                for pw in prior:
                    gap = abs(week - pw)
                    if gap < gap_min:
                        target_penalty += (gap_min - gap) ** 2 * SPECIAL_GAP_WEIGHT

            pref_penalty = 0.0 if (ps is None or in_pref) else 60.0
            return month_penalty + coverage + week_penalty + target_penalty + pref_penalty

        best = min(available, key=lambda x: (score(x), x[2]))
        m, col, week, _ = best
        used_months_by_group[g["id"]].add(m)
        used_quarters_by_group[g["id"]].add(quarter_of_month(m))
        result[g["id"]][idx] = (m, col, week)
        load[week] += 1
        month_load[m] += 1

    return dict(result)

def global_optimize(groups, cols_by_month, used3_locked, used4_locked, to2_locked, passes=OPT_PASSES):
    hard = global_hard_assignment(groups, cols_by_month, used3_locked, used4_locked)
    soft = assign_soft_to2(groups, cols_by_month, hard, to2_locked)
    result = {}
    for g in groups:
        if g["cfg"]["locked"]:
            continue
        result[g["id"]] = dict(hard.get(g["id"], {}))
        result[g["id"]].update(soft.get(g["id"], {}))
    return result

def choose_normal_week(g, month, cols_by_month, special_months, week_load1):
    """Выбирает неделю ТО-1 с глобальным балансом по неделям.

    ТО-1 получает одну работу в каждом месяце сезона, где нет специального
    ТО. Внутри месяца выбирается наименее загруженная неделя. Поэтому разные
    разделы больше не образуют вертикальные столбцы в одной неделе.
    """
    candidates = month_candidates(g, month, cols_by_month)
    if not candidates:
        return None
    center = sum(w for _, w in candidates) / len(candidates)
    return min(candidates, key=lambda cw: (
        week_load1.get(cw[1], 0) * TO1_LOAD_WEIGHT,
        abs(cw[1] - center),
        cw[1],
    ))

def main():
    if len(sys.argv) < 2:
        print("Использование: python generate_ppr_schedule_global_v6.py вход.xlsx [выход.xlsx] [задержка_сек]")
        sys.exit(1)

    in_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else in_path
    delay = float(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_DELAY
    random.seed(RANDOM_SEED)

    print("=" * 78)
    print("ГЛОБАЛЬНАЯ ГЕНЕРАЦИЯ ГОДОВОГО ГРАФИКА ППР")
    print("=" * 78)

    wb = openpyxl.load_workbook(in_path, data_only=False)
    if "ППР" not in wb.sheetnames:
        raise SystemExit('Ошибка: в книге нет листа «ППР».')
    ws = wb["ППР"]

    week_row, first_col = find_week_row_and_data_col(ws)
    last_col = find_last_week_col(ws, week_row)
    month_map = build_month_map(ws, week_row, first_col, last_col)
    cols_by_month = defaultdict(list)
    for c in range(first_col, last_col + 1):
        m = month_map.get(c)
        w = to_int(ws.cell(week_row, c).value)
        if m and w is not None:
            cols_by_month[m].append((c, w))
    config_start = last_col + 1
    groups = build_groups(ws, first_col, last_col, config_start)
    warnings = []
    for g in groups:
        validate_cfg(g, warnings)
        prepare_group(g, cols_by_month, warnings)

    # Защитный снимок A:C и технических столбцов.
    protected = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, config_start + 12):
            if c <= 3 or c >= config_start:
                cell = ws.cell(r, c)
                if not isinstance(cell, MergedCell):
                    protected[(r, c)] = cell.value

    used3, used4 = set(), set()
    locked_to2 = Counter()
    locked_groups = [g for g in groups if g["cfg"]["locked"]]
    unlocked = [g for g in groups if not g["cfg"]["locked"]]
    for g in locked_groups:
        for r in g["leaf_rows"]:
            for c in range(first_col, last_col + 1):
                v = normalize(ws.cell(r, c).value)
                w = to_int(ws.cell(week_row, c).value)
                if v in TO_FROM_TEXT and w is not None:
                    if v == "то3": used3.add(w)
                    elif v == "то4": used4.add(w)
                    elif v == "то2": locked_to2[w] += 1

    print(f"Строка недель: {week_row}")
    print(f"Недели: {get_column_letter(first_col)}:{get_column_letter(last_col)}")
    print(f"Технические столбцы: {get_column_letter(config_start)}:{get_column_letter(config_start + 11)}")
    print(f"Разделов: {len(groups)}; зафиксировано: {len(locked_groups)}; генерируется: {len(unlocked)}")
    print(f"Правила по умолчанию: ТО-2×{DEFAULT_N2}, ТО-3×{DEFAULT_N3}, ТО-4×{DEFAULT_N4}")
    print("Глобальная оптимизация v6: квартальная взаимоисключаемость специальных ТО + равномерность + балансировка ТО-1")
    print()

    print("Построение общего календаря...")
    assignments = global_optimize(groups, cols_by_month, used3, used4, locked_to2, OPT_PASSES)
    print()

    # Теперь очищаем и записываем только недельную сетку.
    totals = Counter()
    week_to3 = defaultdict(list)
    week_to4 = defaultdict(list)
    week_to2 = defaultdict(list)
    week_to1 = Counter()

    for idx, g in enumerate(groups, 1):
        name = g["name"]
        cfg = g["cfg"]
        print(f"[{idx:02d}/{len(groups):02d}] Приступаю к разделу «{name}» [строка {g['header_row']}]")
        if cfg["locked"]:
            print("       Зафиксирован — существующий график не изменяется.")
            time.sleep(delay)
            continue

        amap = assignments[g["id"]]
        assignment = {}
        for event_idx, (month, col, week) in amap.items():
            assignment[col] = g["events"][event_idx]

        special_months = {month for month, _, _ in amap.values()}

        # Контроль технических окон. Если глобальная оптимизация была вынуждена
        # выйти за заданное окно, это явно фиксируем в логе.
        for event_idx, (month, col, week) in amap.items():
            t = g["events"][event_idx]
            ps, pe = preferred_window(g, t)
            if ps is not None and not week_in_window(week, ps, pe):
                warnings.append(
                    f'{g["id"]}: ТО-{t} поставлено на неделю {week}, вне технического окна {ps}-{pe}.'
                )

        # ТО-1 во все остальные месяцы сезона.
        for m in g["months"]:
            if m in special_months:
                continue
            picked = choose_normal_week(g, m, cols_by_month, special_months, week_to1)
            if picked:
                assignment[picked[0]] = 1
                week_to1[picked[1]] += 1

        counts = Counter(assignment.values())
        print(f"       Сезон: {cfg['season_start'] or 1}-{cfg['season_end'] or 53}; "
              f"ТО-2×{cfg['n2']}, ТО-3×{cfg['n3']}, ТО-4×{cfg['n4']}")
        print(f"       Добавлено: ТО-1×{counts[1]}, ТО-2×{counts[2]}, ТО-3×{counts[3]}, ТО-4×{counts[4]}")
        special_week_list = []
        for t in (2, 3, 4):
            weeks = sorted(to_int(ws.cell(week_row, c).value) for c,v in assignment.items() if v == t)
            if weeks:
                print(f"       ТО-{t}: недели {', '.join(map(str, weeks))}")
                special_week_list.extend((w, t) for w in weeks)

        special_week_list.sort()
        if len(special_week_list) >= 2:
            gaps = [special_week_list[i+1][0] - special_week_list[i][0] for i in range(len(special_week_list)-1)]
            print(f"       Интервалы между специальными ТО: {', '.join(map(str, gaps))} нед.")

        # Очистка ТОЛЬКО недель D:BL.
        for r in g["leaf_rows"]:
            for c in range(first_col, last_col + 1):
                ws.cell(r, c).value = None
            time.sleep(DEFAULT_ROW_DELAY)
        for col, t in assignment.items():
            for r in g["leaf_rows"]:
                ws.cell(r, col).value = TO_TEXT[t]

        for t in assignment.values():
            totals[t] += 1
        r = g["leaf_rows"][0] if g["leaf_rows"] else None
        if r:
            for col, t in assignment.items():
                w = to_int(ws.cell(week_row, col).value)
                if t == 2: week_to2[w].append(g["id"])
                elif t == 3: week_to3[w].append(g["id"])
                elif t == 4: week_to4[w].append(g["id"])
        print("       Готово.")
        print()
        time.sleep(delay)

    # Статистика за locked.
    for g in locked_groups:
        if not g["leaf_rows"]: continue
        r = g["leaf_rows"][0]
        for c in range(first_col, last_col + 1):
            v = normalize(ws.cell(r,c).value)
            if v in TO_FROM_TEXT:
                totals[TO_FROM_TEXT[v]] += 1

    # Проверки A:C и BM:BX.
    changed_protected = []
    for (r,c), old in protected.items():
        cell = ws.cell(r,c)
        if not isinstance(cell, MergedCell) and cell.value != old:
            changed_protected.append((r,c,old,cell.value))
    if changed_protected:
        warnings.append(f"Защитная проверка: изменено защищённых ячеек A:C/BM:BX: {len(changed_protected)}.")
        for r,c,old,_ in changed_protected:
            ws.cell(r,c).value = old

    # Финальные коллизии.
    for week,names in sorted(week_to3.items()):
        if len(set(names)) > 1:
            warnings.append(f"Неделя {week}: конфликт ТО-3: {', '.join(sorted(set(names)))}")
    for week,names in sorted(week_to4.items()):
        if len(set(names)) > 1:
            warnings.append(f"Неделя {week}: конфликт ТО-4: {', '.join(sorted(set(names)))}")

    # Проверка: не более одного специального ТО на квартал для каждого раздела.
    for g in unlocked:
        amap = assignments[g["id"]]
        months = [m for m,_,_ in amap.values()]
        weeks = [w for _,_,w in amap.values()]
        if len(months) != len(set(months)):
            warnings.append(f"{g['id']}: более одного специального ТО попало в один месяц.")
        quarters = [quarter_of_month(m) for m in months]
        if len(quarters) != len(set(quarters)):
            warnings.append(
                f"{g['id']}: в одном квартале размещено более одного специального ТО "
                f"(кварталы: {', '.join(map(str, sorted(quarters)))}) — проверьте технические окна."
            )
        if len(weeks) != len(set(weeks)):
            warnings.append(f"{g['id']}: два специальных ТО попали на одну неделю.")
        if len(weeks) >= 2:
            ordered = sorted(weeks)
            gap_min = min_special_gap(g, cols_by_month)
            for a, b in zip(ordered, ordered[1:]):
                if b - a < gap_min - 1e-9:
                    warnings.append(
                        f"{g['id']}: малый интервал между специальными ТО: {b-a} нед. "
                        f"(желательный минимум {gap_min:.1f} нед.)."
                    )

    wb.save(out_path)

    # Сводка распределения специальных недель.
    print("=" * 78)
    print("ИТОГ ГЛОБАЛЬНОЙ ОПТИМИЗАЦИИ")
    print("=" * 78)
    print(f"Обработано разделов: {len(groups)}")
    print(f"ТО-1: {totals[1]}")
    print(f"ТО-2: {totals[2]}")
    print(f"ТО-3: {totals[3]}")
    print(f"ТО-4: {totals[4]}")
    print(f"Уникальных недель ТО-3: {len(week_to3)}")
    print(f"Уникальных недель ТО-4: {len(week_to4)}")
    print(f"ТО-2: {totals[2]} работ, распределены по {len(week_to2)} неделям")

    def compact_load(d):
        return ", ".join(f"{w}:{len(set(ns))}" for w,ns in sorted(d.items()))
    print(f"Распределение ТО-3: {compact_load(week_to3)}")
    print(f"Распределение ТО-4: {compact_load(week_to4)}")
    print(f"Распределение ТО-2: {compact_load(week_to2)}")

    # Месячная сводка крупных ТО — главный контроль v4.
    month_special = Counter()
    month_to1 = Counter()
    month_names = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
    for g in groups:
        if g["cfg"]["locked"]:
            continue
        amap = assignments.get(g["id"], {})
        special_cols = set()
        for _, (m, col, _) in amap.items():
            special_cols.add(col)
            month_special[m] += 1
        for m in g["months"]:
            if m not in {v[0] for v in amap.values()}:
                month_to1[m] += 1
    print("Месячное распределение крупных ТО (ТО-2/3/4):")
    print("  " + ", ".join(f"{month_names[m-1]}:{month_special[m]}" for m in range(1, 13)))
    print("Месячное распределение ТО-1:")
    print("  " + ", ".join(f"{month_names[m-1]}:{month_to1[m]}" for m in range(1, 13)))

    if warnings:
        print(f"\nПРЕДУПРЕЖДЕНИЯ ({len(warnings)}):")
        for w in warnings:
            print(" - " + w)
    else:
        print("\nПроверки завершены: предупреждений нет.")
    print(f"\nСохранено: {out_path}")


if __name__ == "__main__":
    main()
