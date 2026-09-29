"""
Generate synthetic multi-department sample data for the FPY dashboard.

Creates SAMPLE_KPI_DATA.xlsx with the same general schema as a real
manufacturing test-data export, but with 100% fictional departments, part
numbers, serials, technicians, and comments. Safe to commit publicly.

The generator deliberately plants scenarios the dashboard exercises:
  * Clean units (PT+FT PASS)              -> counted in True FPY
  * Units that FAIL then pass on retest   -> duplicate rows removed by dedup;
                                             still failed for serial-level FPY
  * Repeat-offender combos (same Part+Station+Failure Code failing often)
  * Multiple departments (FF) with different yield profiles
  * Blank Pass/Fail rows (data-quality edge case)
  * TBD / noise values (exercise the noise-filtering logic)
"""

import random
from datetime import datetime, timedelta

import pandas as pd

random.seed(42)  # reproducible

OUTPUT_FILE = "SAMPLE_KPI_DATA.xlsx"

# ---------------- Fictional vocabularies ----------------
DEPARTMENTS = ["DEPT-A", "DEPT-B", "DEPT-C", "DEPT-D"]
STAGES = ["PT", "FT", "BI"]  # Pre-Test, Final Test, Burn-In

# Per-department synthetic vocabularies (stations, parts, failure modes)
DEPT_CONFIG = {
    "DEPT-A": {
        "stations": ["A-STATION-1", "A-STATION-2", "A-STATION-3"],
        "parts": ["PA-1000-01R", "PA-1000-02R", "PA-2000-01R"],
        "failure_codes": ["Component", "Wiring", "PCB", "Firmware"],
        "defect_codes": ["E01-Open", "E02-Shorted Internally", "C03-Physical Damage",
                         "W01-Miswire/Swapped wires", "E10-Undetermined Comp. Defect"],
        "ate_steps": ["8.2.9", "8.2.13", "8.4.1.5", "8.5.8", "9.2.3.10"],
        "family": "PRODUCT-A",
    },
    "DEPT-B": {
        "stations": ["B-STATION-1", "B-STATION-2"],
        "parts": ["PB-3000-01R", "PB-3000-04R"],
        "failure_codes": ["HARDWARE", "Calibration", "Component"],
        "defect_codes": ["C06-Improperly Seated/Installed", "E01-Open",
                         "OTHERS - See Tech Comments"],
        "ate_steps": ["5.2.21", "5.3.8", "8.6.19"],
        "family": "PRODUCT-B",
    },
    "DEPT-C": {
        "stations": ["C-STATION-1", "C-STATION-2", "C-STATION-3", "C-STATION-4"],
        "parts": ["PC-500-01R", "PC-600-02R", "PC-700-01R", "PC-800-01R"],
        "failure_codes": ["Component", "MODULE DEFECT", "PCA DEFECT - VENDOR", "OTHER"],
        "defect_codes": ["E02-Shorted Internally", "C03-Physical Damage",
                         "TBD - TO BE DETERMINED", "E10-Undetermined Comp. Defect"],
        "ate_steps": ["8.2.19", "8.4.2.9", "8.5.11", "8.5.16"],
        "family": "PRODUCT-C",
    },
    "DEPT-D": {
        "stations": ["D-STATION-1"],
        "parts": ["PD-900-01R"],
        "failure_codes": ["WIRING", "Component"],
        "defect_codes": ["W01-Miswire/Swapped wires", "E01-Open"],
        "ate_steps": ["TBD - TO BE DETERMINED"],
        "family": "PRODUCT-D",
    },
}

TECHS = ["Tech A", "Tech B", "Tech C", "Tech D", "Tech E", "Tech F"]
TROUBLESHOOTERS = ["Tech G", "Tech H"]

FAIL_COMMENT_TEMPLATES = [
    "Failed Test Step {step}: measured value out of tolerance.",
    "Test Step {step} FAIL - no output detected.",
    "Unit failed at step {step}, see rework instructions.",
    "REPLACED defective component, retest required.",
    "Found shorted part, REPLACED and FIXED.",
    "Intermittent failure observed at step {step}.",
]
REWORK_TEMPLATES = [
    "Replace failed component and retest.",
    "Verify wiring and re-run test.",
    "Send board to repair, then retest.",
    "Re-seat connector and retest.",
]
PASS_COMMENTS = [
    "All tests passed within limits.",
    "No issues found.",
    "Visual and functional test OK.",
    "",
]

COLUMNS = ["FF", "Product Family", "Part No", "Revision", "Serial Number",
           "Date Tested", "Pass or Fail", "Process Stage", "Station",
           "Failed ATE Step", "Failed Assy", "Reference Designator",
           "Defect Code", "Failure Code", "Tested By",
           "Troubleshoot Completed By", "# Failures", "Tech Comments",
           "Rework Instructions", "Work Order"]

START_DATE = datetime(2026, 1, 5)
END_DATE = datetime(2026, 8, 28)


def rand_dt(start, end):
    delta = end - start
    return start + timedelta(seconds=random.randint(0, int(delta.total_seconds())))


def make_row(dept, serial, dt, stage, result, step="", comment="",
             rework="", station=None, part=None):
    cfg = DEPT_CONFIG[dept]
    station = station or random.choice(cfg["stations"])
    part = part or random.choice(cfg["parts"])
    fail = result == "FAIL"
    return {
        "FF": dept,
        "Product Family": cfg["family"],
        "Part No": part,
        "Revision": random.choice(["RA", "RB", "RC"]),
        "Serial Number": serial,
        "Date Tested": dt,
        "Pass or Fail": result,
        "Process Stage": stage,
        "Station": station,
        "Failed ATE Step": step if fail else "",
        "Failed Assy": part if fail else "",
        "Reference Designator": (random.choice(["Q1", "T2", "A4", "D9", "F1"])
                                 if fail else ""),
        "Defect Code": random.choice(cfg["defect_codes"]) if fail else "",
        "Failure Code": random.choice(cfg["failure_codes"]) if fail else "",
        "Tested By": random.choice(TECHS),
        "Troubleshoot Completed By": random.choice(TROUBLESHOOTERS) if fail else "",
        "# Failures": 1 if fail else 0,
        "Tech Comments": comment,
        "Rework Instructions": rework,
        "Work Order": f"WO{random.randint(700000, 799999)}",
    }


def main():
    rows = []
    counters = {d: 0 for d in DEPARTMENTS}

    def next_serial(dept):
        counters[dept] += 1
        return f"{dept}-SN-{counters[dept]:05d}"

    # Per-department volumes and first-pass-fail rates
    dept_plan = {
        "DEPT-A": {"units": 120, "fp_fail_rate": 0.10},   # ~90% FPY
        "DEPT-B": {"units": 80,  "fp_fail_rate": 0.20},   # ~80% FPY
        "DEPT-C": {"units": 150, "fp_fail_rate": 0.15},   # ~85% FPY
        "DEPT-D": {"units": 50,  "fp_fail_rate": 0.06},   # ~94% FPY
    }

    for dept, plan in dept_plan.items():
        cfg = DEPT_CONFIG[dept]
        n_units = plan["units"]

        for _ in range(n_units):
            sn = next_serial(dept)
            t_pt = rand_dt(START_DATE, END_DATE - timedelta(days=4))

            # PT record (usually pass)
            rows.append(make_row(
                dept, sn, t_pt, "PT",
                "PASS" if random.random() > 0.05 else "FAIL",
                comment=random.choice(PASS_COMMENTS),
            ))

            # Decide if this unit fails its first FT
            fails_first = random.random() < plan["fp_fail_rate"]
            t_ft = t_pt + timedelta(days=random.randint(1, 3))

            if fails_first:
                step = random.choice(cfg["ate_steps"])
                comment = random.choice(FAIL_COMMENT_TEMPLATES).format(step=step)
                rows.append(make_row(
                    dept, sn, t_ft, "FT", "FAIL", step=step,
                    comment=comment, rework=random.choice(REWORK_TEMPLATES),
                ))
                # ~75% get retested and pass; rest fail again
                if random.random() < 0.75:
                    rows.append(make_row(
                        dept, sn, t_ft + timedelta(days=random.randint(1, 3)),
                        "FT", "PASS",
                        comment="Retest after rework - PASS",
                    ))
                else:
                    rows.append(make_row(
                        dept, sn, t_ft + timedelta(days=2), "FT", "FAIL",
                        step=random.choice(cfg["ate_steps"]),
                        comment="Failed again after rework.",
                        rework="Escalate to engineering.",
                    ))
            else:
                rows.append(make_row(dept, sn, t_ft, "FT", "PASS",
                                     comment=random.choice(PASS_COMMENTS)))

            # A few units also get a Burn-In record
            if random.random() < 0.15:
                rows.append(make_row(
                    dept, sn, t_ft + timedelta(days=1), "BI",
                    "PASS" if random.random() > 0.08 else "FAIL",
                    comment="" if random.random() > 0.08 else "Burn-in failure.",
                ))

        # Plant a chronic repeat-offender combo in each dept
        chronic_part = cfg["parts"][0]
        chronic_station = cfg["stations"][0]
        chronic_code = cfg["failure_codes"][0]
        for _ in range(random.randint(4, 7)):
            sn = next_serial(dept)
            t = rand_dt(END_DATE - timedelta(days=60), END_DATE)
            rows.append(make_row(
                dept, sn, t, "FT", "FAIL",
                step=cfg["ate_steps"][0] if cfg["ate_steps"] else "",
                comment=random.choice(FAIL_COMMENT_TEMPLATES).format(
                    step=cfg["ate_steps"][0]),
                rework=random.choice(REWORK_TEMPLATES),
                station=chronic_station, part=chronic_part,
            ))
            # Force the chronic failure code on this combo
            rows[-1]["Failure Code"] = chronic_code

        # A couple of blank Pass/Fail rows (data-quality edge case)
        for _ in range(2):
            sn = next_serial(dept)
            rows.append(make_row(dept, sn, rand_dt(START_DATE, END_DATE),
                                 "FT", ""))

    df = pd.DataFrame(rows, columns=COLUMNS)
    df = df.sort_values("Date Tested").reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, sheet_name="Export", index=False)

    # Report the expected headline metrics so users can sanity-check
    serials = df[df["Serial Number"] != ""]["Serial Number"].unique()
    failed_serials = df[(df["Pass or Fail"] == "FAIL")]["Serial Number"].unique()
    fpy = (len(serials) - len(set(serials) & set(failed_serials))) / len(serials) * 100
    print(f"✅ Wrote {OUTPUT_FILE}: {len(df)} rows, {len(serials)} unique serials "
          f"across {len(DEPARTMENTS)} departments")
    print(f"   Expected overall serial-level FPY ≈ {fpy:.1f}%")
    for d in DEPARTMENTS:
        dd = df[df["FF"] == d]
        ds = dd["Serial Number"].nunique()
        dfail = dd[dd["Pass or Fail"] == "FAIL"]["Serial Number"].nunique()
        print(f"   {d}: {ds} serials, {dfail} failed, "
              f"FPY ≈ {(ds - dfail) / ds * 100:.1f}%")


if __name__ == "__main__":
    main()
