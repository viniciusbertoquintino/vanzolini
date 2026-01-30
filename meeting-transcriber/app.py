from pathlib import Path
from datetime import datetime
import time
import queue
import logging
import random

from streamlit_webrtc import WebRtcMode, webrtc_streamer
import streamlit as st

import pydub
from openai import OpenAI, RateLimitError, APIError, APIConnectionError, APITimeoutError
from dotenv import load_dotenv, find_dotenv

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

PASTA_ARQUIVOS = Path(__file__).parent / 'arquivos'
PASTA_ARQUIVOS.mkdir(exist_ok=True)

PROMPT = '''
Faça o resumo do texto delimitado por #### 
O texto é a transcrição de uma reunião.
O resumo deve contar com os principais assuntos abordados.
O resumo deve ter no máximo 300 caracteres.
O resumo deve estar em texto corrido.
No final, devem ser apresentados todos acordos e combinados 
feitos na reunião no formato de bullet points.

O formato final que eu desejo é:

Resumo reunião:
- escrever aqui o resumo.

Acordos da Reunião:
- acrodo 1
- acordo 2
- acordo 3
- acordo n

texto: ####{}####
'''


_ = load_dotenv(find_dotenv())


def salva_arquivo(caminho_arquivo, conteudo):
    """Salva arquivo com encoding UTF-8"""
    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        f.write(conteudo)

def le_arquivo(caminho_arquivo):
    """
    Lê arquivo tentando múltiplos encodings para compatibilidade.
    Tenta UTF-8 primeiro, depois Windows-1252 e ISO-8859-1 como fallback.
    """
    if not caminho_arquivo.exists():
        return ''
    
    # Lista de encodings para tentar (em ordem de preferência)
    encodings = ['utf-8', 'windows-1252', 'iso-8859-1', 'latin-1']
    
    for encoding in encodings:
        try:
            with open(caminho_arquivo, 'r', encoding=encoding) as f:
                conteudo = f.read()
                logger.debug(f"Arquivo {caminho_arquivo.name} lido com encoding {encoding}")
                return conteudo
        except UnicodeDecodeError:
            # Tenta próximo encoding
            continue
        except Exception as e:
            # Outros erros (permissão, etc)
            logger.error(f"Erro ao ler arquivo {caminho_arquivo}: {e}")
            return ''
    
    # Se todos os encodings falharem, tenta com errors='replace' como último recurso
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8', errors='replace') as f:
            logger.warning(f"Arquivo {caminho_arquivo.name} lido com substituição de caracteres inválidos")
            return f.read()
    except Exception as e:
        logger.error(f"Erro crítico ao ler arquivo {caminho_arquivo}: {e}")
        return ''

def listar_reunioes():
    lista_reunioes = PASTA_ARQUIVOS.glob('*')
    lista_reunioes = list(lista_reunioes)
    lista_reunioes.sort(reverse=True)
    reunioes_dict = {}
    for pasta_reuniao in lista_reunioes:
        data_reuniao = pasta_reuniao.stem
        ano, mes, dia, hora, min, seg = data_reuniao.split('_')
        reunioes_dict[data_reuniao] = f'{ano}/{mes}/{dia} {hora}:{min}:{seg}'
        titulo = le_arquivo(pasta_reuniao / 'titulo.txt')
        if titulo != '':
            reunioes_dict[data_reuniao] += f' - {titulo}'
    return reunioes_dict


# OPENAI UTILS =====================
# Configuração do cliente com timeout
client = OpenAI(
    timeout=60.0,  # Timeout de 60 segundos
    max_retries=3,  # Retry automático do cliente
)

def transcreve_audio(caminho_audio, max_retries=3, base_delay=1.0):
    """
    Transcreve áudio usando Whisper API com retry e tratamento de erros.
    
    Args:
        caminho_audio: Caminho para o arquivo de áudio
        max_retries: Número máximo de tentativas
        base_delay: Delay base para backoff exponencial
    
    Returns:
        str: Texto transcrito
        
    Raises:
        Exception: Se todas as tentativas falharem
    """
    for tentativa in range(max_retries):
        try:
            with open(caminho_audio, "rb") as audio_file:
                response = client.audio.transcriptions.create(
                    file=audio_file,
                    model="whisper-1",
                    timeout=60.0  # Timeout específico para esta chamada
                )
            
            if not response or not hasattr(response, 'text'):
                raise ValueError("Resposta inválida da API de transcrição")
            
            logger.info(f"Transcrição bem-sucedida após {tentativa + 1} tentativa(s)")
            return response.text
            
        except RateLimitError as e:
            # Rate limit: espera mais tempo antes de retry
            wait_time = base_delay * (2 ** tentativa) + random.uniform(0, 1)  # Backoff exponencial com jitter
            if tentativa < max_retries - 1:
                logger.warning(f"Rate limit atingido. Tentativa {tentativa + 1}/{max_retries}. Aguardando {wait_time:.2f}s...")
                time.sleep(wait_time)
            else:
                logger.error(f"Rate limit após {max_retries} tentativas: {e}")
                st.error("Limite de requisições excedido. Tente novamente em alguns instantes.")
                raise
                
        except (APIConnectionError, APITimeoutError) as e:
            # Erros de conexão/timeout: retry com backoff
            wait_time = base_delay * (2 ** tentativa) + random.uniform(0, 1)
            if tentativa < max_retries - 1:
                logger.warning(f"Erro de conexão/timeout. Tentativa {tentativa + 1}/{max_retries}. Aguardando {wait_time:.2f}s...")
                time.sleep(wait_time)
            else:
                logger.error(f"Erro de conexão após {max_retries} tentativas: {e}")
                st.error("Erro de conexão com a API. Verifique sua internet e tente novamente.")
                raise
                
        except APIError as e:
            # Outros erros da API: não retry para erros de cliente
            logger.error(f"Erro da API: {e}")
            st.error(f"Erro ao transcrever áudio: {e}")
            raise
            
        except Exception as e:
            # Erros inesperados
            logger.error(f"Erro inesperado na transcrição: {e}", exc_info=True)
            if tentativa < max_retries - 1:
                wait_time = base_delay * (2 ** tentativa)
                time.sleep(wait_time)
            else:
                st.error(f"Erro ao transcrever áudio: {e}")
                raise
    
    raise Exception("Falha ao transcrever áudio após todas as tentativas")


def gerar_resposta_openai(prompt, max_retries=3, base_delay=1.0):
    """
    Gera resposta usando Chat Completions API com retry e tratamento de erros.
    
    Args:
        prompt: Texto do prompt
        max_retries: Número máximo de tentativas
        base_delay: Delay base para backoff exponencial
    
    Returns:
        str: Resposta gerada
        
    Raises:
        Exception: Se todas as tentativas falharem
    """
    for tentativa in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",  # Modelo correto e atualizado
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em resumir reuniões."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                timeout=60.0  # Timeout específico para esta chamada
            )
            
            # Validação da resposta
            if not response or not response.choices:
                raise ValueError("Resposta inválida da API")
            
            resposta_texto = response.choices[0].message.content
            if not resposta_texto:
                raise ValueError("Resposta vazia da API")
            
            logger.info(f"Resposta gerada com sucesso após {tentativa + 1} tentativa(s)")
            return resposta_texto
            
        except RateLimitError as e:
            # Rate limit: espera mais tempo antes de retry
            wait_time = base_delay * (2 ** tentativa) + random.uniform(0, 1)  # Backoff exponencial com jitter
            if tentativa < max_retries - 1:
                logger.warning(f"Rate limit atingido. Tentativa {tentativa + 1}/{max_retries}. Aguardando {wait_time:.2f}s...")
                time.sleep(wait_time)
            else:
                logger.error(f"Rate limit após {max_retries} tentativas: {e}")
                st.error("Limite de requisições excedido. Tente novamente em alguns instantes.")
                raise
                
        except (APIConnectionError, APITimeoutError) as e:
            # Erros de conexão/timeout: retry com backoff
            wait_time = base_delay * (2 ** tentativa) + random.uniform(0, 1)
            if tentativa < max_retries - 1:
                logger.warning(f"Erro de conexão/timeout. Tentativa {tentativa + 1}/{max_retries}. Aguardando {wait_time:.2f}s...")
                time.sleep(wait_time)
            else:
                logger.error(f"Erro de conexão após {max_retries} tentativas: {e}")
                st.error("Erro de conexão com a API. Verifique sua internet e tente novamente.")
                raise
                
        except APIError as e:
            # Outros erros da API: não retry para erros de cliente
            logger.error(f"Erro da API: {e}")
            st.error(f"Erro ao gerar resposta: {e}")
            raise
            
        except Exception as e:
            # Erros inesperados
            logger.error(f"Erro inesperado ao gerar resposta: {e}", exc_info=True)
            if tentativa < max_retries - 1:
                wait_time = base_delay * (2 ** tentativa)
                time.sleep(wait_time)
            else:
                st.error(f"Erro ao gerar resposta: {e}")
                raise
    
    raise Exception("Falha ao gerar resposta após todas as tentativas")


# TAB GRAVA REUNIÃO =====================

def adiciona_chunck_audio(frames_de_audio, audio_chunck):
    for frame in frames_de_audio:
        sound = pydub.AudioSegment(
            data=frame.to_ndarray().tobytes(),
            sample_width=frame.format.bytes,
            frame_rate=frame.sample_rate,
            channels=len(frame.layout.channels),
        )
        audio_chunck += sound
    return audio_chunck

def tab_grava_reuniao():
    webrtx_ctx = webrtc_streamer(
        key='recebe_audio',
        mode=WebRtcMode.SENDONLY,
        audio_receiver_size=1024,
        media_stream_constraints={'video': False, 'audio': True},
    )

    if not webrtx_ctx.state.playing:
        return

    container = st.empty()
    container.markdown('Comece a falar')
    pasta_reuniao = PASTA_ARQUIVOS / datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
    pasta_reuniao.mkdir()

    ultima_trancricao = time.time()
    audio_completo = pydub.AudioSegment.empty()
    audio_chunck = pydub.AudioSegment.empty()
    transcricao = ''

    while True:
        if webrtx_ctx.audio_receiver:
            try:
                frames_de_audio = webrtx_ctx.audio_receiver.get_frames(timeout=1)
            except queue.Empty:
                time.sleep(0.1)
                continue
            audio_completo = adiciona_chunck_audio(frames_de_audio, audio_completo)
            audio_chunck = adiciona_chunck_audio(frames_de_audio, audio_chunck)
            if len(audio_chunck) > 0:
                audio_completo.export(pasta_reuniao / 'audio.mp3')
                agora = time.time()
                if agora - ultima_trancricao > 5:
                    ultima_trancricao = agora
                    audio_chunck.export(pasta_reuniao / 'audio_temp.mp3')
                    try:
                        transcricao_chunck = transcreve_audio(pasta_reuniao / 'audio_temp.mp3')
                        transcricao += transcricao_chunck
                        salva_arquivo(pasta_reuniao / 'transcricao.txt', transcricao)
                        container.markdown(transcricao)
                        audio_chunck = pydub.AudioSegment.empty()
                    except Exception as e:
                        logger.error(f"Erro ao transcrever chunk de áudio: {e}")
                        st.warning(f"Erro ao transcrever: {e}. Continuando gravação...")
        else:
            break


# TAB SELEÇÃO REUNIÃO =====================
def tab_selecao_reuniao():
    reunioes_dict = listar_reunioes()
    if len(reunioes_dict) > 0:
        reuniao_selecionada = st.selectbox('Selecione uma reunião',
                                        list(reunioes_dict.values()))
        st.divider()
        reuniao_data = [k for k, v in reunioes_dict.items() if v == reuniao_selecionada][0]
        pasta_reuniao = PASTA_ARQUIVOS / reuniao_data
        if not (pasta_reuniao / 'titulo.txt').exists():
            st.warning('Adicione um titulo')
            titulo_reuniao = st.text_input('Título da reunião')
            st.button('Salvar',
                      on_click=salvar_titulo,
                      args=(pasta_reuniao, titulo_reuniao))
        else:
            titulo = le_arquivo(pasta_reuniao / 'titulo.txt')
            transcricao = le_arquivo(pasta_reuniao / 'transcricao.txt')
            resumo = le_arquivo(pasta_reuniao / 'resumo.txt')
            if resumo == '':
                with st.spinner('Gerando resumo...'):
                    try:
                        gerar_resumo(pasta_reuniao)
                        resumo = le_arquivo(pasta_reuniao / 'resumo.txt')
                    except Exception as e:
                        logger.error(f"Erro ao gerar resumo: {e}")
                        st.error(f"Erro ao gerar resumo: {e}")
                        resumo = "Erro ao gerar resumo. Tente novamente."
            st.markdown(f'## {titulo}')
            st.markdown(f'{resumo}')
            st.markdown(f'Transcricao: {transcricao}')
        
def salvar_titulo(pasta_reuniao, titulo):
    salva_arquivo(pasta_reuniao / 'titulo.txt', titulo)

def gerar_resumo(pasta_reuniao):
    transcricao = le_arquivo(pasta_reuniao / 'transcricao.txt')
    if not transcricao or transcricao.strip() == '':
        raise ValueError("Transcrição vazia. Não é possível gerar resumo.")
    resumo = gerar_resposta_openai(PROMPT.format(transcricao))
    salva_arquivo(pasta_reuniao / 'resumo.txt', resumo)


# TAB IMPORTAR GOOGLE MEET =====================
def tab_importar_google_meet():
    """
    Aba para importar transcrições do Google Meet.
    """
    st.markdown("### 📥 Importar do Google Meet")
    st.info("Esta funcionalidade está sendo desenvolvida!")
    st.markdown("Em breve você poderá importar transcrições do Google Meet aqui.")

# MAIN =====================
def main():
    st.header('Bem-vindo ao MeetGPT 🎙️', divider=True)
    tab_gravar, tab_selecao, tab_importar = st.tabs([
    'Gravar Reunião', 
    'Ver transcrições salvas',
    'Importar do Google Meet'  # ← Nova aba!
])
    with tab_importar:
        tab_importar_google_meet()
    with tab_gravar:
        tab_grava_reuniao()
    with tab_selecao:
        tab_selecao_reuniao()

if __name__ == '__main__':
    main()