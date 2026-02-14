"""
人材開発支援助成金（事業展開等リスキリング支援コース）書類作成ツール
質問に答えるだけで全書類が完成するWebアプリケーション
"""

import os
import json
from flask import Flask, render_template, request, jsonify, send_file
from generator import generate_all_documents

app = Flask(__name__)
app.config['SECRET_KEY'] = 'jinzai-kaihatsu-joseikin-tool-2026'

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "データが送信されていません"}), 400

        # データの前処理
        processed = preprocess_data(data)

        # 全書類を生成
        zip_path, files = generate_all_documents(processed)

        return jsonify({
            "success": True,
            "zip_path": zip_path,
            "files": [os.path.basename(f) for f in files],
            "message": f"{len(files)}件の書類を生成しました"
        })
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route('/download')
def download():
    output_dir = os.path.join(BASE_DIR, "tool", "output")
    zip_path = os.path.join(output_dir, "人材開発支援助成金_申請書類一式.zip")
    if os.path.exists(zip_path):
        return send_file(zip_path, as_attachment=True,
                        download_name="人材開発支援助成金_申請書類一式.zip")
    return "ファイルが見つかりません", 404


def preprocess_data(data):
    """フォームデータを各書類生成関数が使いやすい形に変換"""
    processed = dict(data)

    # 労働者データの整形
    workers = []
    i = 1
    while f"worker_{i}_name" in data:
        worker = {
            "name": data.get(f"worker_{i}_name", ""),
            "name_kana": data.get(f"worker_{i}_name_kana", ""),
            "insurance_1": data.get(f"worker_{i}_insurance_1", ""),
            "insurance_2": data.get(f"worker_{i}_insurance_2", ""),
            "insurance_3": data.get(f"worker_{i}_insurance_3", ""),
            "insurance_number": f"{data.get(f'worker_{i}_insurance_1', '')}-{data.get(f'worker_{i}_insurance_2', '')}-{data.get(f'worker_{i}_insurance_3', '')}",
            "employment_type": data.get(f"worker_{i}_type", "regular"),
        }
        if worker["name"]:
            workers.append(worker)
        i += 1
    processed["workers"] = workers

    # bool変換
    processed["has_agent"] = data.get("has_agent") == "yes"
    processed["is_subscription"] = data.get("is_subscription") == "yes"
    processed["has_exam"] = data.get("has_exam") == "yes"
    processed["auto_renewal"] = data.get("auto_renewal") == "yes"
    processed["is_sme"] = data.get("is_sme", "yes") == "yes"
    processed["is_voluntary"] = data.get("is_voluntary") == "yes"
    processed["is_batch_application"] = data.get("is_batch_application") == "yes"

    # 支給申請の日付（計画届と同じにデフォルト設定）
    if not processed.get("app_year"):
        processed["app_year"] = processed.get("submit_year")
        processed["app_month"] = processed.get("submit_month")
        processed["app_day"] = processed.get("submit_day")

    # 証明日付（提出日と同じにデフォルト設定）
    if not processed.get("cert_year"):
        processed["cert_year"] = processed.get("submit_year")
        processed["cert_month"] = processed.get("submit_month")
        processed["cert_day"] = processed.get("submit_day")

    return processed


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("人材開発支援助成金 書類作成ツール")
    print("（事業展開等リスキリング支援コース）")
    print("=" * 60)
    print("\nブラウザで以下のURLを開いてください：")
    print("  http://localhost:5000")
    print("\n終了するには Ctrl+C を押してください\n")
    app.run(debug=True, port=5000)
