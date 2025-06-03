from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    pdf_file = request.files['pdf']
    excel_file = request.files['excel']

    if not pdf_file or not excel_file:
        return "PDFとExcelファイルの両方をアップロードしてください。"

    # 保存
    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)

    pdf_file.save(pdf_path)
    excel_file.save(excel_path)

    # 1. PDF → テキスト変換
    output_txt_path = os.path.join(UPLOAD_FOLDER, "output.txt")
    with pdfplumber.open(pdf_path) as pdf, open(output_txt_path, "w", encoding="utf-8") as f_out:
        for page in pdf.pages:
            lines = page.extract_text(layout=True).split("\n")
            for line in lines:
                f_out.write(line + "\n")

    # 2. テキスト処理とデータ抽出
    results = []
    current_race_info = []
    joyo = r"[A-Za-z０-９Ａ-Ｚａ-ｚ0-9\u4E00-\u9FFF\u3040-\u309F\u30A0-\u30FF]+"
    name_sep = r"(?:\s|\u3000)"
    athlete_pattern = fr"(?:\(?\d*\)?)?\s*({joyo}{name_sep}{joyo})\(?\d*\)?\s+({joyo}(?:\s+{joyo})?)\s+(DNS|DNF|\d+(?::\d+)?\.\d+)"
    event_pattern = r"([Ａ-Ｚａ-ｚ0-９０-９一-龯々〆〤ぁ-んァ-ヶー]+(?:\s?[Ａ-Ｚａ-ｚ0-９０-９一-龯々〆〤ぁ-んァ-ヶー]+)*\s?\d+m(?:H|SC)?)"
    results_pattern_1 = fr"(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}"
    results_pattern_2 = fr"(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}\s+(?:(\d+)\s+)?(\d+)\s+(\d+)\s+{athlete_pattern}"

    with open(output_txt_path, "r", encoding="utf-8") as f:
        data = f.read()
    lines = data.strip().split('\n')

    current_date_str = None
    def convert_to_full_date(date_str):
        try:
            m, d = map(int, re.findall(r'(\d+)月(\d+)', date_str)[0])
            return f"2025/{m:02}/{d:02}"
        except:
            return None

    def infer_gender(event_name):
        if '男子' in event_name:
            return '男子'
        elif '女子' in event_name:
            return '女子'
        return None

    grade_df = pd.read_excel(excel_path)
    name_to_grade = dict(zip(grade_df["氏名"], grade_df["学年"]))
    def get_grade(name):
        return name_to_grade.get(name, None)

    for line in lines:
        line = line.strip()
        if re.match(r"\d+月\d+日", line):
            current_date_str = convert_to_full_date(line)
            continue

        match_event = re.match(event_pattern, line)
        if match_event:
            current_event = match_event.group(1).strip()
            gender = infer_gender(current_event)
            current_race_info = [(1, None)]
            continue

        match_race = re.findall(r"(\d+)組\s*(?:\(風:([+-]?[0-9.]+)\))?", line)
        if match_race:
            current_race_info = [(int(g), float(w) if w else None) for g, w in match_race]

        match2 = re.match(results_pattern_2, line)
        if match2:
            data_list = match2.groups()
            順位1, レーン1, ナンバー1, 氏名1, 所属1, 記録1 = data_list[0:6]
            組1, 風1 = current_race_info[0]
            results.append({
                "種目": current_event, "組": 組1, "風": 風1,
                "順位": int(順位1) if 順位1 else None, "レーン": int(レーン1), "ナンバー": int(ナンバー1),
                "氏名": 氏名1.strip(), "所属": 所属1.strip(), "記録": 記録1,
                "性別": gender, "日付": current_date_str, "学年": get_grade(氏名1.strip())
            })
            if len(current_race_info) > 1:
                順位2, レーン2, ナンバー2, 氏名2, 所属2, 記録2 = data_list[6:12]
                組2, 風2 = current_race_info[1]
                results.append({
                    "種目": current_event, "組": 組2, "風": 風2,
                    "順位": int(順位2) if 順位2 else None, "レーン": int(レーン2), "ナンバー": int(ナンバー2),
                    "氏名": 氏名2.strip(), "所属": 所属2.strip(), "記録": 記録2,
                    "性別": gender, "日付": current_date_str, "学年": get_grade(氏名2.strip())
                })
            continue

        match1 = re.match(results_pattern_1, line)
        if match1 and len(current_race_info) >= 1:
            順位, レーン, ナンバー, 氏名, 所属, 記録 = match1.groups()
            組, 風 = current_race_info[0]
            results.append({
                "種目": current_event, "組": 組, "風": 風,
                "順位": int(順位) if 順位 else None, "レーン": int(レーン), "ナンバー": int(ナンバー),
                "氏名": 氏名.strip(), "所属": 所属.strip(), "記録": 記録,
                "性別": gender, "日付": current_date_str, "学年": get_grade(氏名.strip())
            })

    df = pd.DataFrame(results)
    df = df[df["記録"].str.match(r"^\d+(?::\d+)?\.\d+$")]
    df = (
        df.sort_values(by=["種目", "記録"])
          .groupby("種目", group_keys=False)
          .apply(lambda d: d.assign(全体順位=range(1, len(d)+1)))
          .reset_index(drop=True)
    )
    df["種目"] = df["種目"].str.replace(r"^(一般)?(男|女)子", "", regex=True)
    newdf = df[df['所属'] == '東北大'].sort_values(by=["種目", "記録"]).reset_index(drop=True)
    newdf.index = newdf.index + 1
    newdf["種目"] = newdf["種目"].str.replace(r"^(一般)?(男|女)子", "", regex=True)

    newdf = newdf[[ "性別", "種目", "氏名", "学年", "記録", "風", "日付", "組", "レーン", "順位", "全体順位" ]]

    output_excel_path = os.path.join(UPLOAD_FOLDER, "結果.xlsx")
    newdf.to_excel(output_excel_path, index=False)

    return send_file(output_excel_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
