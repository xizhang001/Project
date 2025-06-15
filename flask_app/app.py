import os
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash
from werkzeug.utils import secure_filename
from utils.extract_logic import process_student_files, init_ranking_data, lookup_institution_ranking
from datetime import datetime

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'jpg', 'jpeg', 'png', 'xlsx'}

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create upload folder if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Load default ranking Excel
default_ranking_path = os.path.join("data", "Indianranking2025.xlsx")
default_sheet = "TBS India 25"
ranking_df, institution_list = init_ranking_data(default_ranking_path, default_sheet)

@app.before_request
def before_request():
    """Initialize session history if it doesn't exist"""
    if 'history' not in session:
        session['history'] = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # Save uploaded files
        transcript = request.files.get('transcript')
        cv = request.files.get('cv')
        reference = request.files.get('reference')
        custom_excel = request.files.get('ranking_excel')

        file_paths = {}
        file_names = {}  # Track original filenames for history

        for label, file in [('transcript', transcript), ('cv', cv), ('reference', reference), ('ranking_excel', custom_excel)]:
            if file and file.filename:
                filename = secure_filename(file.filename)
                path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(path)
                file_paths[label] = path
                file_names[label] = filename

        # Use custom ranking file if uploaded
        ranking_path = file_paths.get('ranking_excel', default_ranking_path)
        ranking_df_local, institution_list_local = init_ranking_data(ranking_path, default_sheet)

        # Run the extraction and matching logic
        result = process_student_files(
            transcript_path=file_paths.get('transcript'),
            cv_path=file_paths.get('cv'),
            reference_paths=[file_paths['reference']] if 'reference' in file_paths else None,
            ranking_df=ranking_df_local,
            institution_list=institution_list_local
        )

        # ✅ Always store these for manual check
        session['source_file'] = result.get("_source")
        session['raw_text'] = result.get("_raw_text")[:2000] if result.get("_raw_text") else None
        session['match_score'] = result.get("_match_score")

        # ✅ Store main result only if a match was found
        if result.get("Name of Institution"):
            session['result'] = {
                "Name of Institution": result.get("Name of Institution"),
                "City": result.get("City"),
                "State": result.get("State"),
                "Tier 1": result.get("Tier 1"),
                "Tier 2": result.get("Tier 2")
            }
        else:
            session['result'] = None

        # Create history entry
        history_entry = {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'files': {
                'transcript': file_names.get('transcript'),
                'cv': file_names.get('cv'),
                'reference': file_names.get('reference')
            },
            'result': {
                'institution': result.get("Name of Institution"),
                'city': result.get("City"),
                'state': result.get("State"),
                'match_score': result.get("_match_score")
            }
        }

        # Add to session history (limit to last 10 entries)
        session['history'].insert(0, history_entry)
        session['history'] = session['history'][:10]
        session.modified = True

        return redirect(url_for('results'))

    # GET request fallback
    return render_template('upload.html')

@app.route('/results')
def results():
    result = session.get('result')
    return render_template('results.html', result=result)

@app.route('/manual-check')
def manual_check():
    return render_template('manual_check.html',
                           source=session.get('source_file'),
                           raw_text=session.get('raw_text'),
                           match_score=session.get('match_score'))

@app.route("/institution-names")
def institution_names():
    unique_names = sorted(set(name.strip() for name in institution_list))
    return jsonify(unique_names)

@app.route('/search', methods=['GET', 'POST'])
def search():
    result = None
    query = None

    if request.method == 'POST':
        query = request.form.get('institution_name', '').strip().lower()
        if query:
            matches = ranking_df[ranking_df['Name of Institution'].str.lower() == query]
            if not matches.empty:
                row = matches.iloc[0]
                result = lookup_institution_ranking(row['Name of Institution'], ranking_df)

    return render_template('search.html', result=result, query=query)

@app.route('/history')
def history():
    return render_template('history.html', history=session.get('history', []))

@app.route('/clear-history', methods=['POST'])
def clear_history():
    session['history'] = []
    session.modified = True
    flash('History cleared successfully', 'success')
    return redirect(url_for('history'))

@app.route("/privacy")
def privacy():
    return render_template("privacy.html")


if __name__ == '__main__':
    app.run(debug=True)