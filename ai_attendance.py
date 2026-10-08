"""
AI Attendance Intelligence Module
===================================
STEP 1 — Create this as a NEW FILE called: ai_attendance.py
Place it in the same folder as APP.PY

Logic:
- Calculates attendance % per student per course
- Probability of reaching 80% by semester end using logistic model
- Suppresses alerts if < 5% of total semester classes completed (too early)
- Alert triggers only when probability of hitting 80% drops below 55%
- Calculates minimum consecutive days needed to recover
- Risk badges: SAFE / WARNING / CRITICAL / TOO_EARLY
"""

import math


# ── CORE CALCULATIONS ──────────────────────────────────────────────────────

def calculate_attendance_percentage(present, total):
    """Returns current attendance % for a student in a course."""
    if total == 0:
        return 0.0
    return round((present / total) * 100, 2)


def calculate_days_to_recover(present, total, remaining_classes, target=80):
    """
    Returns minimum consecutive days a student must attend
    to reach the target attendance percentage.

    Formula derived from:
        (present + x) / (total + x) >= target / 100
        => x >= (target * total - 100 * present) / (100 - target)
    """
    if target >= 100:
        return remaining_classes  # impossible unless all remaining attended

    numerator = (target * total) - (100 * present)
    denominator = 100 - target

    if numerator <= 0:
        return 0  # Already at or above target

    days_needed = math.ceil(numerator / denominator)
    return min(days_needed, remaining_classes)  # Can't exceed what's remaining


def estimate_probability(present, total, remaining_classes, target=80):
    """
    Estimates the probability (0–100) that a student will reach
    the target attendance by semester end.

    Model:
    - If < 5% of semester done → return None (too early, suppress alerts)
    - If already above target  → probability based on buffer days available
    - If below target          → probability based on effort ratio
                                 (days needed vs days remaining)
    - If mathematically impossible → 0%
    """
    if total == 0:
        return None  # No classes held yet

    future_total = total + remaining_classes
    best_case_pct = ((present + remaining_classes) / future_total) * 100

    # Mathematical impossibility
    if best_case_pct < target:
        return 0.0

    current_pct = (present / total) * 100

    if current_pct >= target:
        # Already safe — calculate how many classes they can still afford to miss
        # max_skippable: skip this many and still stay at target
        max_skippable = math.floor(
            (current_pct / 100 * future_total) - (target / 100 * future_total)
        )
        if max_skippable <= 0:
            return 72.0
        # More buffer → higher confidence
        prob = min(99.0, 75.0 + (max_skippable / max(remaining_classes, 1)) * 24)
        return round(prob, 1)

    # Below target — how hard is recovery?
    days_needed = calculate_days_to_recover(present, total, remaining_classes, target)

    if days_needed > remaining_classes:
        return 0.0

    effort_ratio = days_needed / max(remaining_classes, 1)

    # Logistic-style decay: low effort → high prob, high effort → low prob
    # effort_ratio=0.0 → ~95%, effort_ratio=1.0 → ~5%
    prob = 95 * (1 - effort_ratio) + 5
    prob = max(0.0, min(99.0, prob))
    return round(prob, 1)


def get_risk_badge(probability, current_pct, classes_completed_pct):
    """
    Returns (badge_label, badge_color, message) based on probability and context.

    Thresholds:
      - classes_completed_pct < 5%  → TOO_EARLY  (grey)
      - probability >= 75            → SAFE        (green)
      - 55 <= probability < 75       → WARNING     (orange)  ← alert sent
      - probability < 55             → CRITICAL    (red)     ← alert sent
      - probability is None          → TOO_EARLY
    """
    if probability is None or classes_completed_pct < 5.0:
        return "TOO_EARLY", "#6c757d", "Not enough data yet"

    if current_pct >= 80 and probability >= 75:
        return "SAFE", "#198754", "You are on track ✓"

    if probability >= 75:
        return "SAFE", "#198754", "You are on track ✓"

    if probability >= 55:
        return "WARNING", "#fd7e14", f"Attend consistently — {100 - round(probability)}% chance of falling short"

    return "CRITICAL", "#dc3545", "High risk of not meeting 80% — immediate action needed"


def should_send_alert(probability, classes_completed_pct):
    """
    Returns True only if:
      1. At least 5% of semester classes are completed (not too early)
      2. Probability of reaching 80% is below 55%
    """
    if classes_completed_pct < 5.0:
        return False
    if probability is None:
        return False
    return probability < 55.0


# ── FULL ANALYSIS PER COURSE ───────────────────────────────────────────────

def analyse_student_course(present, total, total_semester_classes):
    """
    Full analysis for one student in one course.

    Parameters:
        present              : classes attended so far
        total                : total classes held so far
        total_semester_classes: total classes planned for full semester

    Returns a dict with all computed values.
    """
    remaining = max(0, total_semester_classes - total)
    current_pct = calculate_attendance_percentage(present, total)
    classes_completed_pct = (total / total_semester_classes * 100) if total_semester_classes > 0 else 0
    probability = estimate_probability(present, total, remaining)
    days_needed = calculate_days_to_recover(present, total, remaining) if probability is not None else None
    badge, color, message = get_risk_badge(probability, current_pct, classes_completed_pct)
    alert = should_send_alert(probability, classes_completed_pct)

    return {
        "present": present,
        "total_held": total,
        "remaining_classes": remaining,
        "attendance_pct": current_pct,
        "classes_completed_pct": round(classes_completed_pct, 1),
        "probability_of_80": probability,
        "days_needed_to_recover": days_needed,
        "risk_badge": badge,
        "badge_color": color,
        "message": message,
        "send_alert": alert,
    }