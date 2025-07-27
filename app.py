#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import os
import uuid
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import generation_cv  # Importa il modulo generation_cv esistente per la generazione del CV

app = Flask(__name__, static_folder='static')
app.secret_key = "cv-update-secret-key"  # Chiave segreta per i messaggi flash

# Percorso del file JSON del CV
CV_JSON_PATH = 'cv.json'
# Cartella per le immagini del profilo
PROFILE_IMAGES_DIR = os.path.join('static', 'img')
# Estensioni consentite per le immagini
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

# Verifica delle estensioni consentite
def allowed_file(filename):
    """Verifica se l'estensione del file è consentita."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Assicurati che la cartella per le immagini esista
os.makedirs(PROFILE_IMAGES_DIR, exist_ok=True)

def load_cv_data(lang='it'):
    """Carica i dati dal file JSON del CV nella lingua selezionata."""
    file_path = CV_JSON_PATH if lang == 'it' else 'cv_en.json'
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

def save_cv_data(cv_data):
    """Salva i dati nel file JSON del CV."""
    with open(CV_JSON_PATH, 'w', encoding='utf-8') as file:
        json.dump(cv_data, file, indent=2, ensure_ascii=False)

@app.route('/')
def index():
    """Route principale che mostra il form con i dati del CV."""
    lang = request.args.get('lang', 'it')
    cv_data = load_cv_data(lang)
    return render_template('index.html', cv=cv_data, lang=lang)

@app.route('/update/basics', methods=['POST'])
def update_basics():
    """Aggiorna le informazioni di base del CV."""
    cv_data = load_cv_data()
    
    # Aggiorna i dati di base
    cv_data['basics']['name'] = request.form['name']
    cv_data['basics']['tagline'] = request.form['tagline']
    cv_data['basics']['email'] = request.form['email']
    cv_data['basics']['phone']['mobile'] = request.form['mobile']
    cv_data['basics']['phone']['fixed'] = request.form['fixed']
    cv_data['basics']['profiles']['github'] = request.form['github']
    cv_data['basics']['profiles']['telegram'] = request.form['telegram']
    cv_data['basics']['location'] = request.form['location']
    cv_data['basics']['birth']['date'] = request.form['birth_date']
    cv_data['basics']['birth']['place'] = request.form['birth_place']
    cv_data['basics']['nationality'] = request.form['nationality']
    
    save_cv_data(cv_data)
    flash('Informazioni di base aggiornate con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/update/work', methods=['POST'])
def update_work():
    """Aggiorna le esperienze lavorative del CV."""
    cv_data = load_cv_data()
    work_index = int(request.form['work_index'])
    
    if work_index < len(cv_data['work']):
        # Aggiorna un'esperienza lavorativa esistente
        cv_data['work'][work_index]['company'] = request.form['company']
        cv_data['work'][work_index]['position'] = request.form['position']
        cv_data['work'][work_index]['duration'] = request.form['duration']
        
        # Gestisci gli achievement come array separati da nuova riga
        achievements = request.form['achievements'].strip().split('\n')
        cv_data['work'][work_index]['achievements'] = [a.strip() for a in achievements if a.strip()]
    
    save_cv_data(cv_data)
    flash('Esperienza lavorativa aggiornata con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/add/work', methods=['POST'])
def add_work():
    """Aggiunge una nuova esperienza lavorativa al CV."""
    cv_data = load_cv_data()
    
    # Crea una nuova esperienza lavorativa
    new_work = {
        'company': request.form['company'],
        'position': request.form['position'],
        'duration': request.form['duration'],
        'achievements': [a.strip() for a in request.form['achievements'].strip().split('\n') if a.strip()]
    }
    
    cv_data['work'].append(new_work)
    save_cv_data(cv_data)
    flash('Nuova esperienza lavorativa aggiunta con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/delete/work/<int:index>')
def delete_work(index):
    """Elimina un'esperienza lavorativa dal CV."""
    cv_data = load_cv_data()
    
    if 0 <= index < len(cv_data['work']):
        deleted = cv_data['work'].pop(index)
        save_cv_data(cv_data)
        flash(f'Esperienza lavorativa "{deleted["position"]} presso {deleted["company"]}" eliminata!', 'success')
    
    return redirect(url_for('index'))

@app.route('/update/education', methods=['POST'])
def update_education():
    """Aggiorna le informazioni sull'istruzione."""
    cv_data = load_cv_data()
    edu_index = int(request.form['edu_index'])
    
    if edu_index < len(cv_data['education']):
        cv_data['education'][edu_index]['degree'] = request.form['degree']
        cv_data['education'][edu_index]['institution'] = request.form['institution']
    
    save_cv_data(cv_data)
    flash('Informazioni sull\'istruzione aggiornate con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/add/education', methods=['POST'])
def add_education():
    """Aggiunge una nuova istruzione al CV."""
    cv_data = load_cv_data()
    
    new_education = {
        'degree': request.form['degree'],
        'institution': request.form['institution']
    }
    
    cv_data['education'].append(new_education)
    save_cv_data(cv_data)
    flash('Nuova istruzione aggiunta con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/delete/education/<int:index>')
def delete_education(index):
    """Elimina un'istruzione dal CV."""
    cv_data = load_cv_data()
    
    if 0 <= index < len(cv_data['education']):
        deleted = cv_data['education'].pop(index)
        save_cv_data(cv_data)
        flash(f'Istruzione "{deleted["degree"]}" eliminata!', 'success')
    
    return redirect(url_for('index'))

@app.route('/update/skills', methods=['POST'])
def update_skills():
    """Aggiorna le competenze del CV."""
    cv_data = load_cv_data()
    
    # Aggiorna le competenze AI
    cv_data['skills']['ai'] = [s.strip() for s in request.form['ai_skills'].strip().split(',') if s.strip()]
    
    # Aggiorna le competenze di programmazione
    cv_data['skills']['programming']['advanced'] = [s.strip() for s in request.form['prog_advanced'].strip().split(',') if s.strip()]
    cv_data['skills']['programming']['intermediate'] = [s.strip() for s in request.form['prog_intermediate'].strip().split(',') if s.strip()]
    cv_data['skills']['programming']['basic'] = [s.strip() for s in request.form['prog_basic'].strip().split(',') if s.strip()]
    
    # Altre competenze
    cv_data['skills']['industrialAutomation'] = [s.strip() for s in request.form['industrial_automation'].strip().split(',') if s.strip()]
    cv_data['skills']['systems']['windows'] = request.form['systems_windows']
    cv_data['skills']['systems']['linux'] = request.form['systems_linux']
    cv_data['skills']['software'] = [s.strip() for s in request.form['software_skills'].strip().split(',') if s.strip()]
    cv_data['skills']['devOps'] = [s.strip() for s in request.form['devops_skills'].strip().split(',') if s.strip()]
    
    save_cv_data(cv_data)
    flash('Competenze aggiornate con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/update/languages', methods=['POST'])
def update_languages():
    """Aggiorna le lingue conosciute."""
    cv_data = load_cv_data()
    lang_index = int(request.form['lang_index'])
    
    if lang_index < len(cv_data['languages']):
        cv_data['languages'][lang_index]['language'] = request.form['language']
        cv_data['languages'][lang_index]['level'] = request.form['level']
    
    save_cv_data(cv_data)
    flash('Lingua aggiornata con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/add/language', methods=['POST'])
def add_language():
    """Aggiunge una nuova lingua al CV."""
    cv_data = load_cv_data()
    
    new_language = {
        'language': request.form['language'],
        'level': request.form['level']
    }
    
    cv_data['languages'].append(new_language)
    save_cv_data(cv_data)
    flash('Nuova lingua aggiunta con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/delete/language/<int:index>')
def delete_language(index):
    """Elimina una lingua dal CV."""
    cv_data = load_cv_data()
    
    if 0 <= index < len(cv_data['languages']):
        deleted = cv_data['languages'].pop(index)
        save_cv_data(cv_data)
        flash(f'Lingua "{deleted["language"]}" eliminata!', 'success')
    
    return redirect(url_for('index'))

@app.route('/update/digital-skills', methods=['POST'])
def update_digital_skills():
    """Aggiorna le competenze digitali."""
    cv_data = load_cv_data()
    skill_index = int(request.form['skill_index'])
    
    if skill_index < len(cv_data['digitalSkills']):
        cv_data['digitalSkills'][skill_index]['skill'] = request.form['skill']
        cv_data['digitalSkills'][skill_index]['level'] = request.form['level']
    
    save_cv_data(cv_data)
    flash('Competenza digitale aggiornata con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/add/digital-skill', methods=['POST'])
def add_digital_skill():
    """Aggiunge una nuova competenza digitale al CV."""
    cv_data = load_cv_data()
    
    new_skill = {
        'skill': request.form['skill'],
        'level': request.form['level']
    }
    
    cv_data['digitalSkills'].append(new_skill)
    save_cv_data(cv_data)
    flash('Nuova competenza digitale aggiunta con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/delete/digital-skill/<int:index>')
def delete_digital_skill(index):
    """Elimina una competenza digitale dal CV."""
    cv_data = load_cv_data()
    
    if 0 <= index < len(cv_data['digitalSkills']):
        deleted = cv_data['digitalSkills'].pop(index)
        save_cv_data(cv_data)
        flash(f'Competenza digitale "{deleted["skill"]}" eliminata!', 'success')
    
    return redirect(url_for('index'))

@app.route('/update/other', methods=['POST'])
def update_other():
    """Aggiorna le altre informazioni del CV."""
    cv_data = load_cv_data()
    
    cv_data['other']['drivingLicense'] = request.form['driving_license']
    cv_data['other']['hobbies'] = [h.strip() for h in request.form['hobbies'].strip().split(',') if h.strip()]
    cv_data['other']['qualities'] = [q.strip() for q in request.form['qualities'].strip().split(',') if q.strip()]
    
    save_cv_data(cv_data)
    flash('Altre informazioni aggiornate con successo!', 'success')
    return redirect(url_for('index'))

@app.route('/generate-cv')
def generate_cv():
    """Genera il CV in formato DOCX nella lingua selezionata."""
    try:
        import importlib
        importlib.reload(generation_cv)
        lang = request.args.get('lang', 'it')
        json_path = CV_JSON_PATH if lang == 'it' else 'cv_en.json'
        cv_data = load_cv_data(lang)
        output_filename = f"CV_{cv_data['basics']['name'].replace(' ', '_')}_{lang}.docx"
        generation_cv.generate_cv(json_path, output_filename)
        flash(f'CV generated successfully as "{output_filename}"!', 'success')
        return redirect(url_for('index', lang=lang))
    except Exception as e:
        flash(f'Error generating CV: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.route('/download-cv')
def download_cv():
    """Scarica il file CV generato nella lingua selezionata."""
    lang = request.args.get('lang', 'it')
    cv_data = load_cv_data(lang)
    output_filename = f"CV_{cv_data['basics']['name'].replace(' ', '_')}_{lang}.docx"
    if os.path.exists(output_filename):
        return send_file(output_filename, as_attachment=True)
    else:
        flash('CV file not generated yet. Please generate the CV first.', 'warning')
        return redirect(url_for('index', lang=lang))

@app.route('/upload-photo', methods=['POST'])
def upload_photo():
    """Gestisce l'upload della foto profilo."""
    # Controlla se è stato inviato un file
    if 'profile_photo' not in request.files:
        flash('Nessun file inviato', 'danger')
        return redirect(url_for('index'))
        
    file = request.files['profile_photo']
    
    # Se l'utente non seleziona un file, il browser invia un file vuoto senza nome
    if file.filename == '':
        flash('Nessun file selezionato', 'danger')
        return redirect(url_for('index'))
        
    # Se il file esiste ed è un'estensione consentita
    if file and allowed_file(file.filename):
        # Genera un nome di file sicuro e unico
        filename = str(uuid.uuid4()) + os.path.splitext(secure_filename(file.filename))[1]
        file_path = os.path.join(PROFILE_IMAGES_DIR, filename)
        
        # Salva il file
        file.save(file_path)
        
        # Aggiorna il JSON con il percorso della nuova immagine
        cv_data = load_cv_data()
        
        # Salva il percorso relativo dell'immagine nel JSON
        cv_data['basics']['photo'] = os.path.join('img', filename).replace('\\', '/')
        save_cv_data(cv_data)
        
        flash('Foto profilo caricata con successo!', 'success')
    else:
        flash('Formato file non supportato. Utilizzare PNG, JPG o JPEG.', 'danger')
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)