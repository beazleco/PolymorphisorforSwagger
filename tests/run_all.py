#!/usr/bin/env python3
"""Run every suite. Exit non-zero if anything failed."""
import subprocess, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUITES = ["tests.test_workbook_contract", "tests.test_sor_verification",
          "tests.test_regressions"]
totals = []
rc = 0
for mod in SUITES:
    print("=" * 72)
    print(mod)
    print("=" * 72)
    p = subprocess.run([sys.executable, "-m", mod], cwd=ROOT,
                       capture_output=True, text=True)
    tail = [l for l in p.stdout.strip().split("\n") if "passed," in l]
    print(tail[-1] if tail else p.stdout[-400:])
    if p.returncode:
        rc = 1
        print(p.stdout[-3000:])
        print(p.stderr[-2000:])
    totals.append((mod, tail[-1] if tail else "?"))
print("")
print("=" * 72)
for mod, line in totals:
    print("  %-40s %s" % (mod, line))
print("=" * 72)
raise SystemExit(rc)
