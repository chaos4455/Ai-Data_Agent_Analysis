import os
import time
import pandas as pd
import markdown
import google.generativeai as genai
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from datetime import datetime
from gtts import gTTS
import re

# 🔑 Configura a chave da API
API_KEY = 'suaapikeydogooglegemini-obtenhanogoogledatastudio'
genai.configure(api_key=API_KEY)

# 🧠 Nome do modelo
NOME_MODELO = "gemini-1.5-pro-latest"

# 📧 Configurações do SMTP
SMTP_HOST = 'x'
SMTP_PORT = 465
SMTP_USERNAME = 'x'
SMTP_PASSWORD = 'x'
EMAIL_DESTINATARIOS = ['x', 'x']
EMAIL_REMETENTE = 'x'

# ⚙️ Função para configurar a geração de texto
def configurar_geracao(temperatura=0.8, top_p=0.95, top_k=64, max_tokens=8096):
    return {
        "temperature": temperatura,
        "top_p": top_p,
        "top_k": top_k,
        "max_output_tokens": max_tokens,
        "response_mime_type": "text/plain",
    }

# 💬 Função para enviar uma mensagem para o modelo
def enviar_mensagem(sessao, mensagem):
    resposta = sessao.send_message([mensagem])
    return resposta.text

# 📑 Função para criar HTML estilizado para o e-mail
def criar_html_markdown(markdown_texto):
    # Converte o Markdown em HTML
    html_conteudo = markdown.markdown(markdown_texto, extensions=['extra', 'smarty'])
    
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; color: #333; margin: 0; padding: 0; }}
            .container {{ max-width: 800px; margin: 0 auto; padding: 20px; }}
            .header {{ background-color: #4F46E5; color: #fff; padding: 20px; text-align: center; }}
            .footer {{ background-color: #F3F4F6; color: #6B7280; padding: 10px; text-align: center; font-size: 14px; }}
            .content {{ background-color: #FFFFFF; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            h1 {{ font-size: 24px; margin: 0; }}
            h2 {{ font-size: 20px; color: #4F46E5; }}
            pre {{ background-color: #F9FAFB; padding: 15px; border-radius: 4px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Replika AI Solutions</h1>
                <p>Relatório de Análise de Suporte Técnico</p>
            </div>
            <div class="content">
                <h2>Relatório Gerado</h2>
                <div>{html_conteudo}</div>
            </div>
            <div class="footer">
                <p>Obrigado por usar Replika AI Solutions!</p>
                <p>Você está recebendo este e-mail porque está inscrito em nossos relatórios.</p>
                <p><a href="mailto:{EMAIL_REMETENTE}" style="color: #4F46E5;">Entre em contato</a></p>
            </div>
        </div>
    </body>
    </html>
    """
    return html

# 🎙️ Função para converter texto em áudio e salvar o arquivo
def converter_texto_para_audio(texto, nome_arquivo):
    try:
        # Remove caracteres indesejados e aspas
        texto = re.sub(r"[\'\"`´#*]", '', texto)
        
        # Converte texto em áudio
        tts = gTTS(text=texto, lang='pt')
        tts.save(nome_arquivo)
    except Exception as e:
        print(f"Erro ao converter texto em áudio: {e}")

# 📧 Função para enviar e-mail com anexo MP3
def enviar_email(subject, html_body, anexo_mp3):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
    msg['Subject'] = subject

    # Adiciona corpo HTML ao e-mail
    msg.attach(MIMEText(html_body, 'html'))

    # Adiciona o anexo MP3
    try:
        with open(anexo_mp3, 'rb') as f:
            audio = MIMEAudio(f.read(), _subtype="mp3")
            audio.add_header('Content-Disposition', 'attachment', filename=os.path.basename(anexo_mp3))
            msg.attach(audio)
    except Exception as e:
        print(f"Erro ao adicionar anexo MP3: {e}")

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(msg['From'], EMAIL_DESTINATARIOS, msg.as_string())
        print(f"Email enviado com sucesso para {EMAIL_DESTINATARIOS}")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

# 🔍 Função para processar as planilhas e enviar relatórios
def processar_planilhas():
    pasta_planilhas = 'planilha-agente'
    for arquivo in os.listdir(pasta_planilhas):
        if arquivo.endswith('.csv') or arquivo.endswith('.xlsx'):
            caminho_arquivo = os.path.join(pasta_planilhas, arquivo)
            try:
                if arquivo.endswith('.csv'):
                    df = pd.read_csv(caminho_arquivo)
                elif arquivo.endswith('.xlsx'):
                    df = pd.read_excel(caminho_arquivo, sheet_name=None)
                    df = pd.concat(df.values(), ignore_index=True)

                # 🚀 Inicia a sessão de chat com o modelo
                sessao_chat = genai.GenerativeModel(
                    model_name=NOME_MODELO,
                    generation_config=configurar_geracao(),
                ).start_chat(history=[])

                # ✨ Analisar Dados
                texto_dados = df.to_csv(index=False)
                prompt = (
                    "analise os dados da planilha,a apresente relatórios, kpis, indicadores, insights, gargalos, oriente ações, oriente o time, etc..."
                    "responda extremamente longo, adote uma persona, crie uma personalidade a cada resposta, explore e mostre suas capacidades como ia pra analisar dados e trazer insights, apresente sempre os dados em minimos detalhes, mais de 55 insights, use icones, emojis, resposta bem longo todas as respostas, analise os dados, cruze os dados Os dados da planilha são de suporte técnico, ajude a analisar e tomar decisões rápidas. "
                    "Crie relatórios detalhados com insights, estratégias e recomendações. "
                    "Utilize a análise de dados para identificar problemas e sugerir melhorias. "
                    "Forneça relatórios em formato Markdown com emojis e ícones para clareza."
                    f"\n\nDados da planilha:\n\n{texto_dados}"
                    f""
                )

                resposta = enviar_mensagem(sessao_chat, prompt)
                html_relatorio = criar_html_markdown(resposta)
                
                # Salva o MP3 da resposta
                nome_arquivo_audio = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.mp3"
                converter_texto_para_audio(resposta, nome_arquivo_audio)
                
                # Enviar e-mail com o relatório e anexo MP3
                enviar_email(
                    f"Relatório de Suporte Técnico - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 
                    html_relatorio, 
                    nome_arquivo_audio
                )

            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")

# 🕰️ Executa o script a cada 1 minuto
while True:
    processar_planilhas()
    time.sleep(60)
