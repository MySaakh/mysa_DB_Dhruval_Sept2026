#!/usr/bin/env python3
"""Generate synthetic dummy files that mimic the messy patterns found in a
real leads corpus. All values are fake. Used by smoke_test.py.

Usage: python tests/make_fixtures.py [out_dir]   (default: tests/fixtures)
"""

import os
import sys

import openpyxl


def build(base):
    os.makedirs(base, exist_ok=True)

    # fx1: clean header + 40 fake data rows (tests sample cap + masking)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(["Company Name", "Owner", "Mobile", "Email", "City",
               "Pin Code", "GSTIN", "Products"])
    for i in range(1, 41):
        ws.append([
            f"DummyCo {i}", f"Owner {i}", f"90000000{i:02d}",
            f"owner{i}@example.com", "TestCity", f"4000{i:02d}",
            "27ABCDE1234F1Z5", "Widgets",
        ])
    wb.save(os.path.join(base, "fx1_clean.xlsx"))

    # fx2: sparse banner row, blank row, then real header at row 3
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["SAMPLE VENDOR LIST 2024", None, None, None])
    ws.append([None, None, None, None])
    ws.append(["Firm Name", "Contact Person", "Phone", "Area"])
    for i in range(1, 6):
        ws.append([f"FakeFirm {i}", f"Person {i}", f"80000000{i:02d}", "TestArea"])
    wb.save(os.path.join(base, "fx2_banner.xlsx"))

    # fx3: headerless CSV — data starts at line 1 (tests low-confidence flag)
    with open(os.path.join(base, "fx3_headerless.csv"), "w", encoding="utf-8") as f:
        f.write("Dummy Traders,Ramesh Test,9111100001,TestTown,400001\n")
        f.write("Sample Exports,Suresh Test,9111100002,TestTown,400002\n")
        f.write("Mock Industries,Mahesh Test,9111100003,TestTown,400003\n")

    # fx4: cp1252-encoded CSV (invalid as UTF-8; must not be skipped)
    with open(os.path.join(base, "fx4_cp1252.csv"), "wb") as f:
        f.write("name,city\nCafé Test,S\xe3o Paulo\nPlain Test,Delhi\n".encode("cp1252"))

    # fx5: CSV whose data contains a formula-looking string (injection test)
    with open(os.path.join(base, "fx5_inject.csv"), "w", encoding="utf-8") as f:
        f.write("company,note\n")
        f.write('Dummy Ltd,"=HYPERLINK(""http://attacker.example/x"",""open me"")"\n')
        f.write("Safe Ltd,normal note\n")

    # fx6: CSV with a control character (crashed the previous version)
    with open(os.path.join(base, "fx6_ctrl.csv"), "wb") as f:
        f.write(b"company,note\nBad Ltd,has\x01control char\nOk Ltd,fine\n")

    # fx7: an ordinary text note (must be excluded by default)
    with open(os.path.join(base, "fx7_notes.txt"), "w", encoding="utf-8") as f:
        f.write("These are my private notes about vendors.\n")
        f.write("Remember to call the Pune supplier.\n")

    # fx8: Excel lock file (must be skipped silently)
    with open(os.path.join(base, "~$fx1_clean.xlsx"), "wb") as f:
        f.write(b"lockfile junk")

    return sorted(os.listdir(base))


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(__file__), "fixtures"
    )
    print("fixtures written:", build(out))
