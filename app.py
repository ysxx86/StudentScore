from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max-limit

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def get_excellent_threshold(grade):
    if grade in [1, 2]:
        return 90
    elif grade in [3, 4]:
        return 85
    elif grade in [5, 6]:
        return 80
    return 90  # 默认值

def analyze_scores(file_path, grade):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 假设成绩列名为'成绩'或'分数'
    score_column = None
    for col in ['成绩', '分数', 'score', 'Score']:
        if col in df.columns:
            score_column = col
            break
    
    if score_column is None:
        return {'error': '未找到成绩列'}
    
    scores = df[score_column]
    excellent_threshold = get_excellent_threshold(grade)
    
    # 统计各分数段人数
    ranges = {
        '90-100': len(scores[(scores >= 90) & (scores <= 100)]),
        '80-89': len(scores[(scores >= 80) & (scores < 90)]),
        '70-79': len(scores[(scores >= 70) & (scores < 80)]),
        '60-69': len(scores[(scores >= 60) & (scores < 70)]),
        '60以下': len(scores[scores < 60])
    }
    
    # 计算其他统计数据
    total_students = len(scores)
    total_score = scores.sum()
    average_score = scores.mean()
    excellent_count = len(scores[scores >= excellent_threshold])  # 根据年级判断优秀线
    pass_count = len(scores[scores >= 60])      # 60分及格线
    
    return {
        'ranges': ranges,
        'total_score': round(float(total_score), 2),
        'average_score': round(float(average_score), 2),
        'excellent_rate': round(excellent_count / total_students * 100, 2),
        'pass_rate': round(pass_count / total_students * 100, 2),
        'excellent_count': excellent_count,
        'total_students': total_students
    }

@app.route('/')
def index():
    return send_file('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件被上传'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'})
    
    if not file.filename.endswith(('.xls', '.xlsx')):
        return jsonify({'error': '请上传Excel文件'})
    
    grade = request.form.get('grade', type=int)
    if not grade or grade not in range(1, 7):
        return jsonify({'error': '请选择正确的年级（1-6年级）'})
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    try:
        result = analyze_scores(file_path, grade)
        os.remove(file_path)  # 分析完成后删除文件
        return jsonify(result)
    except Exception as e:
        os.remove(file_path)  # 发生错误时也删除文件
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)