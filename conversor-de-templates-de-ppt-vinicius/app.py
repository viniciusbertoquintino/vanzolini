
import os
import zipfile
import shutil
import json
import datetime
import unicodedata
from pathlib import Path
from threading import Thread

from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

# COM / PowerPoint
import pythoncom
import win32com.client
import pywintypes

# --------------------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / 'static' / 'uploads'
DOWNLOAD_FOLDER = BASE_DIR / 'static' / 'downloads'
PROGRESS_DIR = BASE_DIR / 'data' / 'progress'
LOG_FILE = BASE_DIR / 'data' / 'logs' / 'conversions.jsonl'

for p in [UPLOAD_FOLDER, DOWNLOAD_FOLDER, PROGRESS_DIR, LOG_FILE.parent]:
    p.mkdir(parents=True, exist_ok=True)

ALLOWED_TEMPLATE_EXT = {'.ppt', '.pptx'}

app = Flask(
    __name__,
    static_folder=str(BASE_DIR / 'static'),
    static_url_path='/static',
    template_folder=str(BASE_DIR / 'templates')
)

# --------------------------------------------------------------------------------------
# Util: logging JSONL
# --------------------------------------------------------------------------------------
def log_conversion(event: str, conversion_id: str, **kwargs):
    try:
        entry = {
            'ts': datetime.datetime.now().isoformat(),
            'event': event,
            'conversion_id': conversion_id
        }
        entry.update(kwargs)
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry, ensure_ascii=False) + '\n')
    except Exception as e:
        print(f'[LOG][WARN] Falha ao registrar log: {e}')

# --------------------------------------------------------------------------------------
# Util: validações
# --------------------------------------------------------------------------------------
def is_template_file(filename: str) -> bool:
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_TEMPLATE_EXT

def is_zip_file(filename: str) -> bool:
    return filename.lower().endswith('.zip')

def zip_contains_only_ppt(zip_path: Path):
    """Retorna (ok, arquivos_ppt, outros). ok=True se tem pelo menos 1 ppt/pptx;
    se não houver ppt/pptx, ok=False. Outros arquivos são apenas informados.
    """
    arquivos_ppt = []
    outros = []
    with zipfile.ZipFile(str(zip_path), 'r') as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            if name.lower().endswith(('.ppt', '.pptx')):
                arquivos_ppt.append(name)
            else:
                outros.append(name)
    return (len(arquivos_ppt) > 0, arquivos_ppt, outros)

# --------------------------------------------------------------------------------------
# Progresso: arquivo JSON por conversão
# --------------------------------------------------------------------------------------
def progress_path(conversion_id: str) -> Path:
    return PROGRESS_DIR / f"{conversion_id}.json"

def write_progress(conversion_id: str, **data):
    payload = {
        'ts': datetime.datetime.now().isoformat(),
        **data
    }
    with open(progress_path(conversion_id), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False)

# --------------------------------------------------------------------------------------
# Helpers: normalização de nomes e utilidades de layout
# --------------------------------------------------------------------------------------
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize('NFKD', str(s))
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()

def get_layout_by_names_in_master(master, names):
    """Procura CustomLayout por nomes (normalizados), aceitando match exato ou parcial."""
    try:
        if master.CustomLayouts.Count == 0:
            return None
        targets = [_norm(n) for n in names]
        for i in range(1, master.CustomLayouts.Count + 1):
            cl = master.CustomLayouts(i)
            name = _norm(getattr(cl, "Name", f"layout_{i}"))
            for t in targets:
                if name == t or t in name:
                    return cl
    except Exception:
        pass
    return None

# Palavras-chave que indicam “slide de encerramento”
REF_KEYWORDS = ['refer', 'crédit', 'credito', 'bibliograf', 'fontes', 'agradec']

def collect_no_title_indices(pres):
    """Captura os índices de slides que NÃO têm placeholder de título (no arquivo ORIGINAL)."""
    idxs = set()
    for i in range(1, pres.Slides.Count + 1):
        try:
            _ = pres.Slides(i).Shapes.Title  # se houver placeholder, não lança exceção
            has_title = True
        except Exception:
            has_title = False
        if not has_title:
            idxs.add(i)
    return idxs

def slide_has_keywords(slide, keywords):
    """Verifica se qualquer caixa de texto do slide contém alguma keyword (normalizado)."""
    try:
        for j in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(j)
            if getattr(shp, "HasTextFrame", 0) and shp.TextFrame.HasText:
                txt = _norm(shp.TextFrame.TextRange.Text)
                if any(kw in txt for kw in keywords):
                    return True
    except Exception:
        pass
    return False

def find_sem_secao_layout(master):
    """Procura o layout SEM_SEÇÃO por variações de nome normalizadas."""
    names = ['sem_seção', 'sem secao', 'sem_sessao', 'sem-sessao', 'sem-secao', 'sem sessao']
    return get_layout_by_names_in_master(master, names)

def normalize_layouts_with_sem_secao_fallback(pres, candidates_idx, also_last_n=3, use_keywords=True):
    """
    Aplica SEM_SEÇÃO apenas:
      - nos slides que estavam sem placeholder de título no ORIGINAL (candidates_idx)
      - e que tenham keywords típicas de final (se use_keywords=True), OU estejam entre os últimos N.
    Sem fallback genérico (não muda se SEM_SEÇÃO não existir).
    """
    try:
        if pres is None or pres.Slides.Count == 0 or not candidates_idx:
            return

        master = pres.Designs(1).SlideMaster if pres.Designs.Count >= 1 else pres.SlideMaster
        sem_secao = find_sem_secao_layout(master)
        if sem_secao is None:
            # Sem layout “SEM_SEÇÃO”: não força nada para evitar efeitos colaterais
            return

        last_start = max(1, pres.Slides.Count - also_last_n + 1)

        for i in sorted(candidates_idx):
            if i < 1 or i > pres.Slides.Count:
                continue
            s = pres.Slides(i)

            allow = (i >= last_start)
            if use_keywords and not allow:
                allow = slide_has_keywords(s, REF_KEYWORDS)

            if not allow:
                continue

            try:
                s.CustomLayout = sem_secao
                s.FollowMasterBackground = True
            except Exception:
                pass
    except Exception:
        pass

# --------------------------------------------------------------------------------------
# Conversão via PowerPoint (COM) com callback de progresso
# --------------------------------------------------------------------------------------
def convert_presentations(template_path: str, presentations_folder: str, output_folder: str, progress_cb=None):
    """
    Abre cada .ppt/.pptx da pasta 'presentations_folder', aplica o template e salva em 'output_folder'.
    progress_cb(stage, **kwargs) se fornecido, recebe atualizações (current_file, converted_count, total_files, stage).
    Retorna (converted_files, error_message).
    """
    template_path_abs = os.path.abspath(template_path)
    presentations_folder_abs = os.path.abspath(presentations_folder)
    output_folder_abs = os.path.abspath(output_folder)

    if not os.path.exists(template_path_abs):
        return [], f"Template não encontrado: {template_path_abs}"

    try:
        arquivos_encontrados = os.listdir(presentations_folder_abs)
    except Exception as e:
        return [], f"Não foi possível listar a pasta de apresentações: {e}"

    arquivos_ppt = [f for f in arquivos_encontrados if f.lower().endswith(('.ppt', '.pptx'))]
    if not arquivos_ppt:
        return [], "Nenhum arquivo PowerPoint encontrado no ZIP"

    # Inicializa COM e PowerPoint
    try:
        pythoncom.CoInitialize()
        pp = win32com.client.Dispatch("PowerPoint.Application")
        # pp.Visible = 0  # opcional
    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return [], f"Falha ao iniciar PowerPoint/COM: {e}"

    converted_files = []
    total_files = len(arquivos_ppt)

    try:
        for arquivo in arquivos_ppt:
            pres = None
            try:
                caminho_antigo_abs = os.path.abspath(os.path.join(presentations_folder_abs, arquivo))
                if progress_cb:
                    progress_cb(stage='opening', current_file=arquivo, converted_count=len(converted_files), total_files=total_files)

                pres = pp.Presentations.Open(caminho_antigo_abs, ReadOnly=0, Untitled=0, WithWindow=0)

                # (1) Snapshot dos slides SEM título no ORIGINAL
                orig_no_title_idx = collect_no_title_indices(pres)

                # (2) Aplicar template
                if progress_cb:
                    progress_cb(stage='applying_template', current_file=arquivo, converted_count=len(converted_files), total_files=total_files)
                pres.ApplyTemplate(template_path_abs)

                # (3) Ajuste conservador: apenas candidatos + keywords/últimos N
                normalize_layouts_with_sem_secao_fallback(pres, orig_no_title_idx, also_last_n=3, use_keywords=True)

                # Salvar
                if progress_cb:
                    progress_cb(stage='saving', current_file=arquivo, converted_count=len(converted_files), total_files=total_files)

                caminho_novo_abs = os.path.abspath(os.path.join(output_folder_abs, arquivo))
                pres.SaveAs(caminho_novo_abs)

                converted_files.append(arquivo)

            except pywintypes.com_error as ce:
                if progress_cb:
                    progress_cb(stage='error', current_file=arquivo, converted_count=len(converted_files), total_files=total_files, error=str(ce))
            except Exception as e:
                if progress_cb:
                    progress_cb(stage='error', current_file=arquivo, converted_count=len(converted_files), total_files=total_files, error=str(e))
            finally:
                try:
                    if pres is not None:
                        pres.Close()
                        pres = None
                except Exception:
                    pass

    finally:
        try:
            pp.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return converted_files, None

# --------------------------------------------------------------------------------------
# Fluxo assíncrono: thread de conversão com progresso
# --------------------------------------------------------------------------------------
def run_conversion_async(conversion_id: str, template_path: Path, zip_path: Path, presentations_folder: Path, output_folder: Path, output_zip_path: Path):
    try:
        write_progress(conversion_id, status='processing', current_file=None, converted_count=0, total_files=0)

        # Checagem prévia de consistência
        ok, arquivos_ppt, outros = zip_contains_only_ppt(zip_path)
        if not ok:
            write_progress(conversion_id, status='error', error='ZIP não contém .ppt/.pptx', other_files=outros)
            log_conversion('conversion_error', conversion_id, error='ZIP sem PPT/PPTX', other_files=outros)
            return

        # Extrai ZIP
        with zipfile.ZipFile(str(zip_path), 'r') as zf:
            zf.extractall(str(presentations_folder))

        # Callback de progresso que escreve JSON
        def progress_cb(stage, **kwargs):
            write_progress(conversion_id, status='processing', stage=stage, **kwargs)

        # Converte
        converted_files, error = convert_presentations(str(template_path), str(presentations_folder), str(output_folder), progress_cb=progress_cb)

        if error:
            write_progress(conversion_id, status='error', error=str(error))
            log_conversion('conversion_error', conversion_id, error=str(error))
            return

        if not converted_files:
            write_progress(conversion_id, status='error', error='Nenhum arquivo PowerPoint encontrado')
            log_conversion('conversion_error', conversion_id, error='Nenhum arquivo PowerPoint encontrado')
            return

        # Zip de saída
        with zipfile.ZipFile(str(output_zip_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            for file in converted_files:
                file_path = output_folder / file
                zf.write(str(file_path), arcname=file)

        write_progress(conversion_id, status='done', current_file=None, converted_count=len(converted_files), total_files=len(converted_files), converted_files=converted_files)
        log_conversion('conversion_done', conversion_id, total=len(converted_files), files=converted_files)

    except Exception as e:
        write_progress(conversion_id, status='error', error=str(e))
        log_conversion('conversion_error', conversion_id, error=str(e))

# --------------------------------------------------------------------------------------
# Rotas
# --------------------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/status')
def status():
    powerpoint_available = False
    try:
        pythoncom.CoInitialize()
        try:
            pp = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint_available = True
            pp.Quit()
        except Exception:
            powerpoint_available = False
    except Exception:
        powerpoint_available = False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return jsonify({
        'status': 'online',
        'powerpoint_available': powerpoint_available
    })

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Verifica envio
        if 'template' not in request.files or 'presentations' not in request.files:
            return jsonify({'success': False, 'error': 'Template e arquivo ZIP são obrigatórios'}), 400

        template_file = request.files['template']
        presentations_file = request.files['presentations']

        if template_file.filename == '' or presentations_file.filename == '':
            return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'}), 400

        # Valida extensões
        if not is_template_file(template_file.filename):
            return jsonify({'success': False, 'error': 'Template deve ser .ppt ou .pptx'}), 400

        if not is_zip_file(presentations_file.filename):
            return jsonify({'success': False, 'error': 'Apresentações devem estar em um arquivo .zip'}), 400

        # Diretório da conversão
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        conversion_id = f"conversion_{timestamp}"
        conversion_folder = UPLOAD_FOLDER / conversion_id
        conversion_folder.mkdir(parents=True, exist_ok=True)

        # Log início
        try:
            log_conversion('conversion_start', conversion_id, template=template_file.filename, zip=presentations_file.filename)
        except Exception:
            pass

        # Salva template e zip
        template_filename = secure_filename(template_file.filename)
        template_path = conversion_folder / template_filename
        template_file.save(str(template_path))

        zip_filename = secure_filename(presentations_file.filename)
        zip_path = conversion_folder / zip_filename
        presentations_file.save(str(zip_path))

        # Pastas auxiliares
        presentations_folder = conversion_folder / 'presentations'
        presentations_folder.mkdir(parents=True, exist_ok=True)
        output_folder = DOWNLOAD_FOLDER / conversion_id
        output_folder.mkdir(parents=True, exist_ok=True)
        output_zip_path = DOWNLOAD_FOLDER / f"{conversion_id}_convertidos.zip"

        # Progresso inicial
        write_progress(conversion_id, status='queued', current_file=None, converted_count=0, total_files=0)

        # Inicia thread assíncrona
        t = Thread(target=run_conversion_async, args=(conversion_id, template_path, zip_path, presentations_folder, output_folder, output_zip_path), daemon=True)
        t.start()

        # Retorna imediatamente com o conversion_id
        return jsonify({'success': True, 'conversion_id': conversion_id})

    except Exception as e:
        try:
            log_conversion('conversion_error', 'unknown', error=str(e))
        except Exception:
            pass
        return jsonify({'success': False, 'error': f'Erro interno: {str(e)}'}), 500

@app.route('/progress/<conversion_id>')
def get_progress(conversion_id):
    p = progress_path(conversion_id)
    if not p.exists():
        return jsonify({'status': 'unknown'}), 404
    try:
        data = json.loads(p.read_text(encoding='utf-8'))
        return jsonify(data)
    except Exception as e:
        return jsonify({'status': 'error', 'error': str(e)}), 500

@app.route('/download/<conversion_id>')
def download_file(conversion_id):
    try:
        zip_filename = f"{conversion_id}_convertidos.zip"
        zip_path = DOWNLOAD_FOLDER / zip_filename

        if not zip_path.exists():
            return jsonify({'error': 'Arquivo não encontrado'}), 404

        resp = send_file(str(zip_path), as_attachment=True, download_name=f"apresentacoes_convertidas_{conversion_id}.zip")

        # Log de download
        try:
            log_conversion('download', conversion_id)
        except Exception:
            pass

        # Cleanup após envio
        def _cleanup():
            try:
                base_upload = UPLOAD_FOLDER / conversion_id
                base_download = DOWNLOAD_FOLDER / conversion_id
                # remove pastas temporárias e o ZIP
                for p in [base_upload, base_download]:
                    shutil.rmtree(p, ignore_errors=True)
                try:
                    os.remove(str(zip_path))
                except Exception:
                    pass
                # remove progresso
                try:
                    os.remove(str(progress_path(conversion_id)))
                except Exception:
                    pass
            except Exception as e:
                try:
                    log_conversion('cleanup_error', conversion_id, error=str(e))
                except Exception:
                    pass

        resp.call_on_close(_cleanup)
        return resp

    except Exception as e:
        return jsonify({'error': f'Erro no download: {str(e)}'}), 500

# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
if __name__ == '__main__':
    print("\n🔍 VERSÃO DEBUG ATIVA 🔍")
    print("Backend funcionando!")
    print("Acesse: http://localhost:5000")
    print("NOTA: Esta versão mostra logs MUITO detalhados no terminal")
    print("============================================================")
    app.run(host='0.0.0.0', port=5000, debug=True)
