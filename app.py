from flask import Flask, request, jsonify
import os
import pandas as pd

app = Flask(__name__)

def get_excel_file(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    if len(files) != 1:
        raise ValueError(f"Expected exactly 1 Excel file in {folder_path}")
    return os.path.join(folder_path, files[0])

@app.route('/process-files', methods=['POST'])
def process_files():
    try:
        data = request.json or {}
        print("Received request from Zapier")

        folder1 = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\input\file1(latest date)"
        folder2 = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\input\file2(older date)"

        file1 = get_excel_file(folder1)
        file2 = get_excel_file(folder2)

        df_latest = pd.read_excel(file1)
        df_older = pd.read_excel(file2)

        date1 = os.path.splitext(os.path.basename(file1))[0]
        date2 = os.path.splitext(os.path.basename(file2))[0]

        df_latest["Date"] = date1
        df_older["Date"] = date2

        df_merge = pd.concat([df_latest, df_older], ignore_index=True)
        df_merge["Remarks"] = ""

        common = pd.merge(df_latest, df_older, on="Ref.No", how="inner")

        df_merge.loc[df_merge["Ref.No"].isin(common["Ref.No"]), "Remarks"] = "In Stock"
        df_merge.loc[(df_merge["Ref.No"].isin(df_older["Ref.No"])) & (~df_merge["Ref.No"].isin(df_latest["Ref.No"])), "Remarks"] = "Sold"
        df_merge.loc[(df_merge["Ref.No"].isin(df_latest["Ref.No"])) & (~df_merge["Ref.No"].isin(df_older["Ref.No"])), "Remarks"] = "New Arrival"

        OUTPUT = r"C:\Users\Vasu\PY\.venv\PROJECTS\project1\output-drive"
        os.makedirs(OUTPUT, exist_ok=True)

        df_merge.to_excel(os.path.join(OUTPUT, "Data.xlsx"), index=False)
        common.to_excel(os.path.join(OUTPUT, "common.xlsx"), index=False)

        return jsonify({
            "status": "success",
            "message": "Files processed successfully",
            "output_files": ["Data.xlsx", "common.xlsx"]
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/', methods=['GET'])
def health():
    return jsonify({"status": "alive"}), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)