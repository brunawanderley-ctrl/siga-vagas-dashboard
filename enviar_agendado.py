#!/usr/bin/env python3
"""
Envia relatórios agendados baseado na configuração do dashboard.
Execute via cron ou manualmente.
"""

import json
import smtplib
import urllib.request
import urllib.parse
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime, timedelta
from io import BytesIO

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"

def carregar_config_email():
    """Carrega configuração de email"""
    config_path = BASE_DIR / ".email-config"
    config = {}
    if config_path.exists():
        with open(config_path) as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key] = value
    return config

def carregar_schedule():
    """Carrega configuração de agendamento"""
    schedule_path = OUTPUT_DIR / "schedule_config.json"
    if schedule_path.exists():
        with open(schedule_path) as f:
            return json.load(f)
    return {}

def carregar_dados():
    """Carrega dados do relatório"""
    with open(OUTPUT_DIR / "resumo_ultimo.json") as f:
        resumo = json.load(f)
    with open(OUTPUT_DIR / "vagas_ultimo.json") as f:
        vagas = json.load(f)
    return resumo, vagas

def deve_enviar_agora(schedule):
    """Verifica se deve enviar o relatório agora"""
    if not schedule.get('ativo', False):
        return False

    agora = datetime.now()
    hora_agendada = schedule.get('hora', 6)
    minuto_agendado = schedule.get('minuto', 0)
    frequencia = schedule.get('frequencia', 'Diário')

    # Verifica se está na janela de horário (15 minutos de tolerância)
    hora_atual = agora.hour
    minuto_atual = agora.minute

    if hora_atual != hora_agendada:
        return False

    if abs(minuto_atual - minuto_agendado) > 15:
        return False

    # Verifica frequência
    if frequencia == "Diário":
        return True
    elif frequencia == "Semanal":
        dia_semana_config = schedule.get('dia_semana', 0)
        return agora.weekday() == dia_semana_config
    elif frequencia == "Mensal":
        dia_mes_config = schedule.get('dia_mes', 1)
        return agora.day == dia_mes_config

    return False

def gerar_html_relatorio(resumo, vagas, tipo_relatorio):
    """Gera HTML do relatório"""
    total = resumo['total_geral']
    ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

    # Cor da ocupação baseada no valor
    cor_ocupacao = "#22c55e" if ocupacao >= 70 else ("#f97316" if ocupacao >= 50 else "#ef4444")

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{tipo_relatorio} - Colégio Elo</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 30px; background: #f8fafc; }}
            .container {{ max-width: 700px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; }}
            .header {{ background: linear-gradient(135deg, #1e4976 0%, #2563eb 100%); color: white; padding: 30px; }}
            .header h1 {{ margin: 0; font-size: 24px; }}
            .header p {{ margin: 10px 0 0 0; opacity: 0.9; }}
            .content {{ padding: 30px; }}
            .kpi-grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-bottom: 25px; }}
            .kpi {{ background: #f8fafc; border-radius: 10px; padding: 20px; text-align: center; }}
            .kpi-value {{ font-size: 32px; font-weight: 700; color: {cor_ocupacao}; }}
            .kpi-label {{ font-size: 12px; color: #64748b; text-transform: uppercase; margin-top: 5px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background: #1e4976; color: white; padding: 12px; text-align: center; font-size: 13px; }}
            td {{ border: 1px solid #e2e8f0; padding: 10px; text-align: center; font-size: 13px; }}
            tr:nth-child(even) {{ background: #f8fafc; }}
            .ocupacao-alta {{ color: #22c55e; font-weight: bold; }}
            .ocupacao-media {{ color: #f97316; font-weight: bold; }}
            .ocupacao-baixa {{ color: #ef4444; font-weight: bold; }}
            .footer {{ text-align: center; padding: 20px; color: #94a3b8; font-size: 12px; border-top: 1px solid #e2e8f0; }}
            .btn {{ display: inline-block; background: #2563eb; color: white; padding: 12px 24px; border-radius: 8px; text-decoration: none; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>📊 {tipo_relatorio}</h1>
                <p>Colégio Elo • {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
            </div>

            <div class="content">
                <div class="kpi-grid">
                    <div class="kpi">
                        <div class="kpi-value">{ocupacao}%</div>
                        <div class="kpi-label">Ocupação Geral</div>
                    </div>
                    <div class="kpi">
                        <div class="kpi-value" style="color: #2563eb;">{total['matriculados']:,}</div>
                        <div class="kpi-label">Matriculados</div>
                    </div>
                    <div class="kpi">
                        <div class="kpi-value" style="color: #64748b;">{total['vagas']:,}</div>
                        <div class="kpi-label">Vagas Totais</div>
                    </div>
                    <div class="kpi">
                        <div class="kpi-value" style="color: #f97316;">{total['disponiveis']:,}</div>
                        <div class="kpi-label">Disponíveis</div>
                    </div>
                </div>

                <h3 style="color: #1e4976; margin-top: 25px;">Por Unidade</h3>
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

        classe_ocup = 'ocupacao-alta' if ocup >= 70 else ('ocupacao-media' if ocup >= 50 else 'ocupacao-baixa')

        html += f"""
                    <tr>
                        <td style="text-align: left;"><strong>{nome}</strong></td>
                        <td>{t['vagas']}</td>
                        <td>{t['matriculados']}</td>
                        <td>{t['disponiveis']}</td>
                        <td class="{classe_ocup}">{ocup}%</td>
                    </tr>
        """

    html += f"""
                </table>

                <div style="text-align: center; margin-top: 25px;">
                    <a href="https://brunaviegas-siga-vagas-dashboard.streamlit.app" class="btn">
                        Ver Dashboard Completo
                    </a>
                </div>
            </div>

            <div class="footer">
                <p>Relatório automático enviado pelo SIGA Vagas Dashboard</p>
                <p>Colégio Elo © {datetime.now().year}</p>
            </div>
        </div>
    </body>
    </html>
    """

    return html

def gerar_texto_whatsapp(resumo, tipo_relatorio):
    """Gera texto para WhatsApp"""
    total = resumo['total_geral']
    ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

    msg = f"📊 *{tipo_relatorio}*\n"
    msg += f"Colégio Elo • {datetime.now().strftime('%d/%m/%Y')}\n\n"

    msg += f"*Ocupação Geral:* {ocupacao}%\n"
    msg += f"*Matriculados:* {total['matriculados']} / {total['vagas']}\n"
    msg += f"*Disponíveis:* {total['disponiveis']}\n\n"

    msg += "*Por Unidade:*\n"
    for unidade in resumo['unidades']:
        nome = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)
        emoji = "🟢" if ocup >= 70 else ("🟡" if ocup >= 50 else "🔴")
        msg += f"{emoji} {nome}: {ocup}%\n"

    msg += f"\n🔗 Dashboard: brunaviegas-siga-vagas-dashboard.streamlit.app"

    return msg

def enviar_email(config_email, html_content, tipo_relatorio):
    """Envia email com o relatório"""
    email_remetente = config_email.get('EMAIL')
    senha = config_email.get('SENHA')
    email_destino = config_email.get('EMAIL_DESTINO', email_remetente)

    if not email_remetente or not senha:
        print("❌ Configuração de email incompleta")
        return False

    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"📊 {tipo_relatorio} - Colégio Elo - {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'] = email_remetente
    msg['To'] = email_destino

    msg.attach(MIMEText(html_content, 'html'))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_remetente, senha)
            server.sendmail(email_remetente, email_destino, msg.as_string())
        print(f"✅ Email enviado para {email_destino}")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar email: {e}")
        return False

def enviar_whatsapp(config_email, texto):
    """Envia mensagem via WhatsApp usando CallMeBot"""
    phone = config_email.get('WHATSAPP_PHONE')
    apikey = config_email.get('WHATSAPP_APIKEY')

    if not phone or not apikey:
        print("⚠️ WhatsApp não configurado")
        return False

    url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={urllib.parse.quote(texto)}&apikey={apikey}"

    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=30) as response:
            result = response.read().decode('utf-8')
            print(f"✅ WhatsApp enviado: {result}")
            return True
    except Exception as e:
        print(f"❌ Erro ao enviar WhatsApp: {e}")
        return False

def registrar_envio(tipo_relatorio, canais):
    """Registra o envio no histórico"""
    historico_path = OUTPUT_DIR / "envios_historico.json"
    historico = []

    if historico_path.exists():
        try:
            with open(historico_path) as f:
                historico = json.load(f)
        except:
            historico = []

    historico.append({
        'data': datetime.now().isoformat(),
        'tipo': tipo_relatorio,
        'canais': canais,
        'sucesso': True
    })

    # Mantém apenas os últimos 100 registros
    historico = historico[-100:]

    with open(historico_path, 'w') as f:
        json.dump(historico, f, indent=2, ensure_ascii=False)

def main():
    print("=" * 50)
    print(f"⏰ Verificando agendamento - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)

    # Carrega configurações
    schedule = carregar_schedule()

    if not schedule:
        print("⚠️ Nenhum agendamento configurado")
        return

    if not schedule.get('ativo', False):
        print("🔴 Agendamento desativado")
        return

    # Verifica se deve enviar
    if not deve_enviar_agora(schedule):
        freq = schedule.get('frequencia', 'Diário')
        hora = schedule.get('hora', 6)
        minuto = schedule.get('minuto', 0)
        print(f"⏳ Agendado para {freq} às {hora:02d}:{minuto:02d}")
        print("   Não é hora de enviar ainda.")
        return

    print("🚀 Iniciando envio agendado...")

    # Carrega dados
    try:
        resumo, vagas = carregar_dados()
    except FileNotFoundError:
        print("❌ Arquivos de dados não encontrados")
        return

    tipo_relatorio = schedule.get('tipo_relatorio', 'Resumo Executivo')
    config_email = carregar_config_email()
    canais_enviados = []

    # Gera conteúdo
    html_content = gerar_html_relatorio(resumo, vagas, tipo_relatorio)
    texto_whatsapp = gerar_texto_whatsapp(resumo, tipo_relatorio)

    # Envia por email
    if schedule.get('enviar_email', True):
        if enviar_email(config_email, html_content, tipo_relatorio):
            canais_enviados.append('email')

    # Envia por WhatsApp
    if schedule.get('enviar_whatsapp', False):
        if enviar_whatsapp(config_email, texto_whatsapp):
            canais_enviados.append('whatsapp')

    # Registra envio
    if canais_enviados:
        registrar_envio(tipo_relatorio, canais_enviados)
        print(f"\n✅ Envio concluído! Canais: {', '.join(canais_enviados)}")
    else:
        print("\n⚠️ Nenhum envio realizado")

if __name__ == "__main__":
    main()
