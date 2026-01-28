from flask import Flask, request, jsonify
import os
import pandas as pd

app = Flask(__name__)

def get_excel_file(folder_path):
    """Find single Excel file in folder"""
    if not os.path.exists(folder_path):
        raise ValueError(f"Folder does not exist: {folder_path}")
    
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    if len(files) != 1:
        raise ValueError(f"Expected exactly 1 Excel file in {folder_path}, found {len(files)}")
    return os.path.join(folder_path, files[0])

@app.route('/process-files', methods=['POST'])
def process_files():
    try:
        data = request.json or {}
        print("✅ Received request from Zapier")

        # Use environment variables or default paths
        base_path = os.path.expanduser("~")  # Works on Windows and Linux
        
        # For local testing (Windows)
        if os.name == 'nt':  # Windows
            folder1 = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\input\file1(latest date)"
            folder2 = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\input\file2(older date)"
            output_path = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\output-drive"
        else:  # Linux (Render)
            # On Render, use /tmp or environment variable
            folder1 = os.getenv('INPUT_FOLDER_1', '/tmp/input/file1')
            folder2 = os.getenv('INPUT_FOLDER_2', '/tmp/input/file2')
            output_path = os.getenv('OUTPUT_PATH', '/tmp/output')

        # Create folders if they don't exist
        os.makedirs(folder1, exist_ok=True)
        os.makedirs(folder2, exist_ok=True)
        os.makedirs(output_path, exist_ok=True)

        # Get Excel files
        file1 = get_excel_file(folder1)
        file2 = get_excel_file(folder2)

        # Read Excel files
        df_latest = pd.read_excel(file1)
        df_older = pd.read_excel(file2)

        # Extract date labels from filenames
        date1 = os.path.splitext(os.path.basename(file1))[0]
        date2 = os.path.splitext(os.path.basename(file2))[0]

        df_latest["Date"] = date1
        df_older["Date"] = date2

        # Merge data
        df_merge = pd.concat([df_latest, df_older], ignore_index=True)
        df_merge["Remarks"] = ""

        # Find common items
        common = pd.merge(df_latest, df_older, on="Ref.No", how="inner")

        # Add remarks
        df_merge.loc[
            df_merge["Ref.No"].isin(common["Ref.No"]),
            "Remarks"
        ] = "In Stock"

        df_merge.loc[
            (df_merge["Ref.No"].isin(df_older["Ref.No"])) &
            (~df_merge["Ref.No"].isin(df_latest["Ref.No"])),
            "Remarks"
        ] = "Sold"

        df_merge.loc[
            (df_merge["Ref.No"].isin(df_latest["Ref.No"])) &
            (~df_merge["Ref.No"].isin(df_older["Ref.No"])),
            "Remarks"
        ] = "New Arrival"

        # Save output files
        data_file = os.path.join(output_path, "Data.xlsx")
        common_file = os.path.join(output_path, "common.xlsx")

        df_merge.to_excel(data_file, index=False)
        common.to_excel(common_file, index=False)

        print(f"✅ Files processed successfully")
        print(f"   - Data.xlsx saved")
        print(f"   - common.xlsx saved")

        return jsonify({
            "status": "success",
            "message": "Files processed successfully",
            "output_files": ["Data.xlsx", "common.xlsx"],
            "data_file": data_file,
            "common_file": common_file
        })

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/', methods=['GET'])
def health():
    return jsonify({"status": "alive", "message": "Webhook is running"}), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)