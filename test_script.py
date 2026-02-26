#!/usr/bin/env python3
"""
Cross-platform smoke test for the bills_monthly project.

Tests that core dependencies import, template files exist, and shared
utility functions work correctly. Safe to run on any OS.
"""

import os
import sys
import platform


def test_imports():
    """Test that core dependencies are importable."""
    print("1. Testing imports...")
    import openpyxl
    print(f"   OK — openpyxl {openpyxl.__version__}")
    from dateutil.relativedelta import relativedelta
    print("   OK — python-dateutil")
    from bill_utils import generate_random_bill_no, compute_billing_dates, update_excel_file
    print("   OK — bill_utils")
    return True


def test_template_files():
    """Test that required template files exist."""
    print("\n2. Checking required files...")
    templates = ['Mobile_Bill_Template.xlsx', 'Landline_Bill_Template.xlsx']
    all_ok = True
    for f in templates:
        if os.path.exists(f):
            print(f"   OK — {f}")
        else:
            print(f"   FAIL — {f} NOT FOUND")
            all_ok = False
    return all_ok


def test_bill_utils():
    """Test shared utility functions."""
    print("\n3. Testing bill_utils functions...")
    from bill_utils import generate_random_bill_no, compute_billing_dates

    # Random bill number format: 2 letters + 14 digits = 16 chars
    bill_no = generate_random_bill_no()
    assert len(bill_no) == 16, f"Bill number length should be 16, got {len(bill_no)}"
    assert bill_no[:2].isalpha(), "First 2 chars should be letters"
    assert bill_no[2:].isdigit(), "Last 14 chars should be digits"
    print(f"   OK — generate_random_bill_no() → {bill_no}")

    # Billing dates should return all expected keys
    dates = compute_billing_dates()
    expected_keys = [
        'statement_date_str', 'statement_period_str',
        'due_date_q7_str', 'due_date_s12_str', 'bill_no_str'
    ]
    for key in expected_keys:
        assert key in dates, f"Missing key: {key}"
    print(f"   OK — compute_billing_dates() returned {len(dates)} fields")

    return True


def test_windows_deps():
    """Test Windows-specific dependencies (skipped on non-Windows)."""
    print("\n4. Testing Windows dependencies...")
    if platform.system() != 'Windows':
        print("   SKIP — not on Windows")
        return True

    try:
        import win32com.client
        print("   OK — win32com.client")
    except ImportError:
        print("   FAIL — pywin32 not installed (pip install pywin32)")
        return False

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        print(f"   OK — Excel version: {excel.Version}")
        excel.Quit()
        print("   OK — Excel COM working")
    except Exception as e:
        print(f"   FAIL — Excel COM: {e}")
        return False

    return True


def main():
    print("=" * 50)
    print(f"BILLS_MONTHLY SMOKE TEST ({platform.system()})")
    print("=" * 50)

    all_passed = True
    try:
        all_passed &= test_imports()
        all_passed &= test_template_files()
        all_passed &= test_bill_utils()
        all_passed &= test_windows_deps()
    except Exception as e:
        print(f"\n   FAIL — Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        all_passed = False

    print("\n" + "=" * 50)
    if all_passed:
        print("ALL TESTS PASSED ✓")
    else:
        print("SOME TESTS FAILED ✗")
        sys.exit(1)
    print("=" * 50)


if __name__ == "__main__":
    main()
