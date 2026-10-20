import os
import glob
import csv
import json

def update_programs():
    csv_files = glob.glob('Program Summary*.csv') + glob.glob('program summary*.csv') + glob.glob('program-summary*.csv')
    if not csv_files:
        print("No Program Summary CSV files found in workspace.")
        return

    # Sort to pick the newest by filename/modification
    csv_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    target_csv = csv_files[0]
    print(f"Reading latest Program Summary file: {target_csv}")

    programs = {}
    with open(target_csv, 'r', encoding='utf-8', errors='replace') as f:
        reader = csv.DictReader(f)
        for row in reader:
            status = (row.get('Program Status') or '').strip()
            if status.lower() != 'active':
                continue

            code = (row.get('Local Code') or '').strip()
            title = (row.get('Program Entity Title') or '').strip()
            if not code or not title:
                continue

            prog_id = f"{code} - {title}"
            code_upper = code.upper()
            is_associate = code_upper.startswith('A') or code_upper.startswith('S')

            if prog_id not in programs:
                programs[prog_id] = {
                    "id": prog_id,
                    "code": code,
                    "title": title,
                    "isAssociate": is_associate,
                    "courses": []
                }

            subj = (row.get('Program_Course Block_Course Block Item Course Subject Code') or '').strip().upper()
            num = (row.get('Program_Course Block_Course Block Item Course Number') or '').strip().upper()
            if subj and num:
                crse = f"{subj}-{num}"
                if crse not in programs[prog_id]["courses"]:
                    programs[prog_id]["courses"].append(crse)

    print(f"Extracted {len(programs)} active programs from {target_csv}.")

    # Generate program-data.js
    js_content = "/**\n * Chaffey College Academic Program & Course Mapping\n * Auto-generated from " + target_csv + "\n */\n"
    js_content += "window.CHAFFEY_PROGRAM_DATA = " + json.dumps(programs, indent=2) + ";\n"

    with open('program-data.js', 'w', encoding='utf-8') as f:
        f.write(js_content)
    print("Successfully updated program-data.js")

    # If target is not 'Program Summary.csv', create/update standardized copy
    if target_csv != 'Program Summary.csv':
        try:
            with open(target_csv, 'rb') as src, open('Program Summary.csv', 'wb') as dst:
                dst.write(src.read())
            print("Successfully updated standardized 'Program Summary.csv'")
        except Exception as e:
            print(f"Could not copy to Program Summary.csv: {e}")

if __name__ == '__main__':
    update_programs()
