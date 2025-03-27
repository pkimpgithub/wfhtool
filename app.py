from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files['file']
    emails = request.form['emails']
    dates = request.form['dates']

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    try:
        df = pd.read_csv(file_path, encoding='latin1')
        df['date'] = pd.to_datetime(df['date_time_utc'], errors='coerce').dt.date

        email_list = [e.strip() for e in emails.replace("\n", ",").split(",") if e.strip()]
        date_list = [datetime.strptime(d.strip(), "%Y-%m-%d").date() for d in dates.replace("\n", ",").split(",") if d.strip()]

        filtered_df = df[
            (df['date'].isin(date_list)) &
            (df['sender_address'].isin(email_list)) &
            (~df['message_subject'].str.contains("Automatic reply", case=False, na=False))
        ]

        result = filtered_df.groupby(['sender_address', 'date']).size().reset_index(name='email_count')
        export_path = os.path.join(UPLOAD_FOLDER, 'email_report.csv')
        result.to_csv(export_path, index=False)

        return jsonify(result.to_dict(orient='records'))

    except Exception as e:
        return str(e), 400

@app.route('/download')
def download():
    return send_file(os.path.join(UPLOAD_FOLDER, 'email_report.csv'), as_attachment=True)

if __name__ == '__main__':   
    app.run(host='0.0.0.0', port=5000)
