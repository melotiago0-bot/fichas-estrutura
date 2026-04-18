from flask import Flask, request, jsonify, send_file, send_from_directory
import json, re, shutil, os, subprocess, tempfile

app = Flask(__name__, static_folder='static')

BASE_DIR  = os.path.dirname(__file__)
TEMPLATE  = os.path.join(BASE_DIR, 'template_clean.docx')
UNPACKED  = os.path.join(BASE_DIR, 'clean_unpacked')
PACK_SCRIPT = os.path.join(BASE_DIR, 'scripts', 'pack.py')

HEADER = {'nome_uc':10,'autor':47,'apresentacao':175,'keywords':437}
UE_MAP = [None,
    {'titulo':812, 'desc':884, 'ea':[(945,1007),(1059,1121),(1173,1235),(1289,1353)]},
    {'titulo':1679,'desc':1755,'ea':[(1819,1881),(1937,1999),(2055,2117),(2172,2236)]},
    {'titulo':2518,'desc':2594,'ea':[(2658,2720),(2776,2838),(2894,2956),(3011,3075)]},
    {'titulo':3357,'desc':3433,'ea':[(3497,3559),(3615,3677),(3733,3795),(3850,3914)]},
    {'titulo':4196,'desc':4272,'ea':[(4336,4398),(4454,4516),(4572,4634),(4689,4753)]},
    {'titulo':5035,'desc':5111,'ea':[(5175,5237),(5293,5355),(5411,5473),(5528,5592)]},
    {'titulo':5874,'desc':5950,'ea':[(6014,6076),(6132,6194),(6250,6312),(6367,6431)]},
]

def xml_esc(s):
    return str(s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def replace_line(lines, line_no, new_text):
    idx = line_no - 1
    lines[idx] = re.sub(
        r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)',
        lambda m: m.group(1) + xml_esc(new_text) + m.group(3),
        lines[idx]
    )

def fill_and_pack(data):
    work_dir = tempfile.mkdtemp(prefix='fe_')
    try:
        shutil.copytree(UNPACKED, work_dir + '/doc')
        doc_xml = work_dir + '/doc/word/document.xml'
        with open(doc_xml, encoding='utf-8') as f:
            lines = f.readlines()

        replace_line(lines, HEADER['nome_uc'],      data.get('nomeUC', ''))
        replace_line(lines, HEADER['autor'],        data.get('autor', ''))
        replace_line(lines, HEADER['apresentacao'], data.get('apresentacao', ''))
        replace_line(lines, HEADER['keywords'],     '; '.join(data.get('palavrasChave', [])))

        for ue in data.get('ues', []):
            n = ue.get('numero')
            m = UE_MAP[n] if n and 1 <= n <= 7 else None
            if not m:
                continue
            replace_line(lines, m['titulo'], ue.get('titulo', ''))
            replace_line(lines, m['desc'],   ue.get('descricao', ''))
            eacts = ue.get('eatividades', [])
            for i, (tl, typl) in enumerate(m['ea']):
                if i < len(eacts):
                    replace_line(lines, tl,   eacts[i].get('titulo', ''))
                    replace_line(lines, typl, eacts[i].get('tipo', ''))
                else:
                    replace_line(lines, tl,   '')
                    replace_line(lines, typl, '')

        with open(doc_xml, 'w', encoding='utf-8') as f:
            f.writelines(lines)

        out = work_dir + '/output.docx'
        r = subprocess.run(
            ['python3', PACK_SCRIPT, work_dir + '/doc', out,
             '--original', TEMPLATE, '--validate', 'false'],
            capture_output=True, text=True,
            cwd=os.path.join(BASE_DIR, 'scripts')
        )
        if r.returncode != 0:
            raise RuntimeError(r.stderr or r.stdout)
        with open(out, 'rb') as f:
            return f.read()
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/sugestao', methods=['POST', 'OPTIONS'])
def sugestao():
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        return resp
    try:
        import urllib.request, urllib.parse
        data = request.get_json(force=True)
        SHEETS_URL = 'https://script.google.com/macros/s/AKfycbzx0Ky5OzwOkvJ7JmoVQUP162TJIYWW6BrD5-w6R7EKB_uxt2omexmCiMQ-sYSg0w8s/exec'
        params = urllib.parse.urlencode({
            'texto': data.get('texto', ''),
            'autor': data.get('autor', 'Equipa')
        })
        url = SHEETS_URL + '?' + params
        response = urllib.request.urlopen(url, timeout=15)
        return jsonify({'status': 'ok', 'response': response.read().decode('utf-8')})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        return resp
    try:
        data = request.get_json(force=True)
        docx_bytes = fill_and_pack(data)
        slug = re.sub(r'[^a-zA-Z0-9À-ÿ ]', '', data.get('nomeUC', 'ficha')).strip().replace(' ', '_')
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        tmp.write(docx_bytes)
        tmp.close()
        return send_file(
            tmp.name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{slug}_ficha_estrutura.docx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5050)))
