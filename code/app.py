from flask import Flask, request, jsonify, render_template, url_for, send_from_directory
import os
import oletools.olevba
import google.generativeai as genai
import re
import matplotlib.pyplot as plt
import networkx as nx

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
IMAGE_FOLDER = 'static/images'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGE_FOLDER, exist_ok=True)
api_key = "AIzaSyBIj_b-axuFTRVrc4ya65CCfazNejT6u-Y"
genai.configure(api_key=api_key)

def generate_diagram(functions):
    G = nx.DiGraph()
    for function in functions:
        G.add_node(function)
    for i in range(len(functions) - 1):
        G.add_edge(functions[i], functions[i + 1])
    
    pos = nx.spring_layout(G)
    plt.figure()
    nx.draw(G, pos, with_labels=True, node_size=3000, node_color='lightblue', font_size=10, font_weight='bold', arrows=True)
    image_path = os.path.join(IMAGE_FOLDER, 'vba_process_flow.png')
    plt.savefig(image_path)
    plt.close()
    return image_path

def parse_vba_code(vba_code):
    functions = re.findall(r'(Sub\s+\w+|Function\s+\w+)', vba_code)
    return functions

def extract_vba_macros(file_path):
    vbaparser = oletools.olevba.VBA_Parser(file_path)
    vba_macros = []
    vba_functions = []
    if vbaparser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
            vba_functions.extend(parse_vba_code(vba_code))  # Use extend instead of append
            vba_macros.append({
                'filename': filename,
                'stream_path': stream_path,
                'vba_filename': vba_filename,
                'vba_code': vba_code
            })
    vbaparser.close()
    return [vba_macros, vba_functions]

def generate_content(vba_macros):
    if not vba_macros:
        return 'No VBA Macros found'
    
    vba_code_combined = "\n".join([macro['vba_code'] for macro in vba_macros])
    model = genai.GenerativeModel(model_name="gemini-pro")
    response = model.generate_content([vba_code_combined])
    
    return response.text if response else 'No content generated'

@app.route('/home')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and file.filename.endswith('.xlsm'):
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        t = extract_vba_macros(file_path)
        vba_macros = t[0]
        vba_functions = t[1]
        generated_content = generate_content(vba_macros)
        image_url = ''
        if vba_functions:
            image_path = generate_diagram(vba_functions)
            image_url = url_for('static', filename='images/vba_process_flow.png')
        
        return jsonify({
            'message': 'File successfully uploaded',
            'file_path': file_path,
            'vba_macros': vba_macros,
            'generated_content': generated_content,
            'image_url': image_url
        }), 200
    else:
        return jsonify({'error': 'Invalid file type'}), 400

if __name__ == '__main__':
    app.run(debug=True)
