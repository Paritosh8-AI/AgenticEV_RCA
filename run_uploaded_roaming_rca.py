"""
ElectreeFi Studio - Automated Launcher for Uploaded Roaming RCA
Intelligently detects whether dual files (completed + cancelled), a single file,
or a custom CLI input file is passed, invokes RoamingUploadAnalyzer with 100% production safety,
and outputs both Word (.docx) and Excel (.xlsx) reports.
"""
import os
import sys

# Ensure project root in sys.path
_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "."))
if _root not in sys.path:
    sys.path.insert(0, _root)

from src.rca.roaming_upload_analyzer import RoamingUploadAnalyzer

def main():
    completed_file = "uploads/latest_completed.xlsx"
    cancelled_file = "uploads/latest_cancelled.xlsx"
    single_file = "uploads/latest_roaming_upload.xlsx"
    sample_file = "uploads/sample_roaming_dual_sheet.xlsx"
    output_docx = "ElectreeFi_Uploaded_Roaming_RCA_Report.docx"
    output_xlsx = "ElectreeFi_Uploaded_Roaming_RCA_Report.xlsx"

    analyzer = RoamingUploadAnalyzer()
    try:
        # Check CLI argument first
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            input_path = sys.argv[1]
            print(f"[LAUNCHER] Using user-specified input file: {input_path}")
            # Check if excel has multiple sheets
            import openpyxl
            wb = openpyxl.load_workbook(input_path, read_only=True)
            sheet_names = [s.lower() for s in wb.sheetnames]
            if "completed" in sheet_names and "cancelled" in sheet_names:
                print("  • Detected multi-sheet workbook with 'Completed' and 'Cancelled' tabs.")
                instances = analyzer.analyze_dual_files(completed_file=input_path, cancelled_file=input_path)
            else:
                rows = analyzer.parse_uploaded_excel(input_path)
                instances = analyzer.analyze_instances(rows)
            sources = [os.path.basename(input_path)]
        elif os.path.exists(completed_file) and os.path.exists(cancelled_file):
            print(f"[LAUNCHER] Found dual upload files:\n  - Completed: {completed_file}\n  - Cancelled: {cancelled_file}")
            instances = analyzer.analyze_dual_files(completed_file=completed_file, cancelled_file=cancelled_file)
            sources = [os.path.basename(completed_file), os.path.basename(cancelled_file)]
        elif os.path.exists(cancelled_file):
            print(f"[LAUNCHER] Found cancelled upload file: {cancelled_file}")
            rows = analyzer.parse_uploaded_excel(cancelled_file, default_category="Cancelled")
            instances = analyzer.analyze_instances(rows)
            sources = [os.path.basename(cancelled_file)]
        elif os.path.exists(single_file):
            print(f"[LAUNCHER] Found single upload file: {single_file}")
            rows = analyzer.parse_uploaded_excel(single_file)
            instances = analyzer.analyze_instances(rows)
            sources = [os.path.basename(single_file)]
        elif os.path.exists(sample_file):
            print(f"[LAUNCHER] Using default sample file: {sample_file}")
            instances = analyzer.analyze_dual_files(completed_file=sample_file, cancelled_file=sample_file)
            sources = [os.path.basename(sample_file)]
        else:
            print("[ERROR] No uploaded files found in uploads/. Please upload an Excel file via the GUI or pass a path.")
            return 1

        if not instances:
            print("[WARN] No valid instances found across target parties.")
            return 0

        analyzer.generate_word_report(instances, output_docx=output_docx, source_filenames=sources)
        analyzer.generate_excel_report(instances, output_xlsx=output_xlsx)
        print("\n" + "=" * 70)
        print("RCA COMPLETE (100% PRODUCTION SAFE):")
        print(f"  • Word Report:  {os.path.abspath(output_docx)}")
        print(f"  • Excel Report: {os.path.abspath(output_xlsx)}")
        print("=" * 70)
        return 0
    finally:
        analyzer.close()

if __name__ == "__main__":
    sys.exit(main() or 0)
