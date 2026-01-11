#!/usr/bin/env python3
"""Envia notificações (email e WhatsApp) após extração do SIGA"""

import smtplib
import json
import urllib.request
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from datetime import datetime

def carregar_config():
    config_path = Path(__file__).parent / ".email-config"
    config = {}
    with open(config_path) as f:
        for line in f:
            if '=' in line:
                key, value = line.strip().split('=', 1)
                config[key] = value
    return config

def carregar_resumo():
    resumo_path = Path(__file__).parent / "output" / "resumo_ultimo.json"
    with open(resumo_path) as f:
        return json.load(f)

def formatar_email(resumo, alertas=None):
    total = resumo['total_geral']
    ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

    # Cor da ocupação baseada no valor
    cor_ocupacao = "#ff4444" if ocupacao >= 80 else "#70AD47"

    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h2 {{ color: #4472C4; }}
            table {{ border-collapse: collapse; width: 100%; max-width: 600px; }}
            th, td {{ border: 1px solid #ddd; padding: 10px; text-align: center; }}
            th {{ background-color: #4472C4; color: white; }}
            .total {{ background-color: #D6DCE5; font-weight: bold; }}
            .ocupacao {{ font-size: 24px; color: {cor_ocupacao}; font-weight: bold; }}
            .alerta-box {{ background-color: #fff3cd; border: 2px solid #ff4444; border-radius: 8px; padding: 15px; margin: 15px 0; }}
            .alerta-titulo {{ color: #ff4444; font-size: 18px; font-weight: bold; margin-bottom: 10px; }}
            .alerta-item {{ color: #856404; margin: 5px 0; }}
            .alta {{ background-color: #ffcccc; color: #cc0000; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h2>Relatório de Vagas - Colégio Elo</h2>
        <p>Extração realizada em: <strong>{resumo['data_extracao'][:16].replace('T', ' ')}</strong></p>
    """

    # Adiciona box de alertas se houver
    if alertas and len(alertas) > 0:
        html += """
        <div class="alerta-box">
            <div class="alerta-titulo">🚨 ALERTAS DE OCUPAÇÃO ALTA (≥80%)</div>
        """
        for alerta in alertas:
            html += f'<div class="alerta-item">{alerta}</div>'
        html += "</div>"

    html += f"""
        <h3>Resumo Geral</h3>
        <p class="ocupacao">Ocupação: {ocupacao}%</p>
        <p>Matriculados: <strong>{total['matriculados']}</strong> de <strong>{total['vagas']}</strong> vagas</p>
        <p>Disponíveis: <strong>{total['disponiveis']}</strong> | Novatos: <strong>{total['novatos']}</strong> | Veteranos: <strong>{total['veteranos']}</strong></p>

        <h3>Por Unidade</h3>
        <table>
            <tr>
                <th>Unidade</th>
                <th>Vagas</th>
                <th>Matriculados</th>
                <th>Disponíveis</th>
                <th>Ocupação</th>
            </tr>
    """

    for unidade in resumo['unidades']:
        nome = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)
        classe_ocup = 'class="alta"' if ocup >= 80 else ''
        html += f"""
            <tr>
                <td>{nome}</td>
                <td>{t['vagas']}</td>
                <td>{t['matriculados']}</td>
                <td>{t['disponiveis']}</td>
                <td {classe_ocup}>{ocup}%</td>
            </tr>
        """

    html += f"""
            <tr class="total">
                <td>TOTAL</td>
                <td>{total['vagas']}</td>
                <td>{total['matriculados']}</td>
                <td>{total['disponiveis']}</td>
                <td>{ocupacao}%</td>
            </tr>
        </table>

        <p style="margin-top: 20px;">
            <a href="https://brunaviegas-siga-vagas-dashboard.streamlit.app"
               style="background-color: #4472C4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
                Ver Dashboard Completo
            </a>
        </p>

        <p style="color: #888; font-size: 12px; margin-top: 30px;">
            Email automático enviado pelo SIGA Vagas Extractor
        </p>
    </body>
    </html>
    """

    return html

def verificar_alertas(resumo):
    """Verifica se há alertas de ocupação BAIXA (quanto maior ocupação, melhor)"""
    alertas = []

    # Verifica ocupação geral
    total = resumo['total_geral']
    ocupacao_geral = round(total['matriculados'] / total['vagas'] * 100, 1)
    if ocupacao_geral < 50:
        alertas.append(f"❄️ OCUPAÇÃO GERAL CRÍTICA: {ocupacao_geral}%")
    elif ocupacao_geral < 70:
        alertas.append(f"⚠️ OCUPAÇÃO GERAL BAIXA: {ocupacao_geral}%")

    # Verifica ocupação por unidade
    for unidade in resumo['unidades']:
        nome = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)
        if ocup < 50:
            alertas.append(f"❄️ {nome}: {ocup}% (CRÍTICO)")
        elif ocup < 70:
            alertas.append(f"⚠️ {nome}: {ocup}% (Atenção)")

        # Verifica por segmento
        for seg, vals in unidade['segmentos'].items():
            ocup_seg = round(vals['matriculados'] / vals['vagas'] * 100, 1)
            if ocup_seg < 50:
                alertas.append(f"❄️ {nome} - {seg}: {ocup_seg}%")

    return alertas

def enviar_email():
    config = carregar_config()
    resumo = carregar_resumo()

    email_remetente = config['EMAIL']
    senha = config['SENHA']
    email_destino = config.get('EMAIL_DESTINO', email_remetente)

    total = resumo['total_geral']
    ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

    # Verifica alertas
    alertas = verificar_alertas(resumo)
    tem_alerta = len(alertas) > 0

    msg = MIMEMultipart('alternative')

    # Subject com alerta se necessário
    if tem_alerta:
        msg['Subject'] = f"❄️ ALERTA: Ocupação baixa detectada! ({ocupacao}%) - {datetime.now().strftime('%d/%m/%Y')}"
    elif ocupacao >= 90:
        msg['Subject'] = f"🔥 SIGA Vagas: Excelente! {ocupacao}% ocupação - {datetime.now().strftime('%d/%m/%Y')}"
    elif ocupacao >= 80:
        msg['Subject'] = f"✨ SIGA Vagas: Muito bom! {ocupacao}% ocupação - {datetime.now().strftime('%d/%m/%Y')}"
    else:
        msg['Subject'] = f"📊 SIGA Vagas: {total['matriculados']} matriculados ({ocupacao}% ocupação) - {datetime.now().strftime('%d/%m/%Y')}"

    msg['From'] = email_remetente
    msg['To'] = email_destino

    html = formatar_email(resumo, alertas)
    msg.attach(MIMEText(html, 'html'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_remetente, senha)
            server.sendmail(email_remetente, email_destino, msg.as_string())
        print(f"Email enviado para {email_destino}")
        return True
    except Exception as e:
        print(f"Erro ao enviar email: {e}")
        return False

def enviar_whatsapp():
    """Envia mensagem de alerta via WhatsApp usando CallMeBot"""
    config = carregar_config()

    # Verifica se WhatsApp está configurado
    if 'WHATSAPP_PHONE' not in config or 'WHATSAPP_APIKEY' not in config:
        print("WhatsApp não configurado")
        return False

    resumo = carregar_resumo()
    alertas = verificar_alertas(resumo)

    total = resumo['total_geral']
    ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

    # Monta mensagem
    msg = f"📊 *SIGA Vagas - {datetime.now().strftime('%d/%m/%Y')}*\n\n"
    msg += f"*Ocupação Geral:* {ocupacao}%\n"
    msg += f"*Matriculados:* {total['matriculados']} / {total['vagas']}\n"
    msg += f"*Disponíveis:* {total['disponiveis']}\n\n"

    # Resumo por unidade
    msg += "*Por Unidade:*\n"
    for unidade in resumo['unidades']:
        nome = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)
        emoji = "🔴" if ocup >= 80 else "🟢"
        msg += f"{emoji} {nome}: {ocup}% ({t['matriculados']}/{t['vagas']})\n"

    # Alertas
    if alertas:
        msg += f"\n🚨 *ALERTAS (≥80%):*\n"
        for alerta in alertas[:10]:  # Limita a 10 alertas
            msg += f"{alerta}\n"

    msg += f"\n🔗 Dashboard: brunaviegas-siga-vagas-dashboard.streamlit.app"

    # Envia via CallMeBot
    phone = config['WHATSAPP_PHONE']
    apikey = config['WHATSAPP_APIKEY']

    url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={urllib.parse.quote(msg)}&apikey={apikey}"

    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=30) as response:
            result = response.read().decode('utf-8')
            print(f"WhatsApp enviado: {result}")
            return True
    except Exception as e:
        print(f"Erro ao enviar WhatsApp: {e}")
        return False

if __name__ == "__main__":
    enviar_email()
    # enviar_whatsapp()  # Desativado temporariamente
