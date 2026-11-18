from datetime import date, datetime, timedelta
import os
import pandas as pd
import streamlit as st
from streamlit_calendar import calendar
import win32com.client as win32

# Configuração da página - Layout Wide & Moderno (DEVE SER O PRIMEIRO COMANDO)
st.set_page_config(
    page_title="Agenda Corporativa - Evomobile",
    page_icon="💜",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================
# SISTEMA DE AUTENTICAÇÃO (SENHA DA DIRETORIA)
# ==========================================
def check_password():
  """Retorna True se o usuário inseriu a senha correta."""
  if st.session_state.get("password_correct", False):
    return True

  st.markdown("""
        <div style='text-align: center; padding: 40px 0;'>
            <h1>🔒 EVOMOBILE | ACESSO RESTRITO</h1>
            <p style='color: #64748b;'>Digite a senha da diretoria para acessar a agenda corporativa.</p>
        </div>
    """, unsafe_allow_html=True)

  col_l1, col_l2, col_l3 = st.columns([1, 1, 1])
  with col_l2:
    with st.form("login_form"):
      senha_digitada = st.text_input(
          "Senha de Acesso:", type="password", key="input_senha_diretoria"
      )
      submit_login = st.form_submit_button("Entrar no Sistema")

      if submit_login:
        # Senha de acesso da diretoria
        if senha_digitada == "evomobile2026!":
          st.session_state["password_correct"] = True
          st.rerun()
        else:
          st.error("❌ Senha incorreta. Acesso negado.")

  return False


# Se a senha não estiver correta, interrompe e mostra apenas a tela de login
if not check_password():
  st.stop()
# ==========================================


# ==========================================
# INÍCIO DO APP ORIGINAL (APÓS LOGIN CORRETO)
# ==========================================

# Mapeamento Oficial de Consultores e E-mails da Evomobile
CONSULTORES_EMAILS = {
    "Bruno Lima": "bruno.lima@evomobile.com.br",
    "Suelen": "suellen.morgado@evomobile.com.br",
    "Claudia": "claudia.cunha@evomobile.com.br",
    "Mallu": "marilu.marcal@evomobile.com.br",
    "Deisy": "deisy.rodrigues@evomobile.com.br",
    "Beatriz": "beatriz@evomobile.com.br",
}
LISTA_CONSULTORES = list(CONSULTORES_EMAILS.keys())

DIAS_SEMANA = {
    "Monday": "Segunda-feira",
    "Tuesday": "Terça-feira",
    "Wednesday": "Quarta-feira",
    "Thursday": "Quinta-feira",
    "Friday": "Sexta-feira",
    "Saturday": "Sábado",
    "Sunday": "Domingo",
}
MESES = {
    "January": "Janeiro",
    "February": "Fevereiro",
    "March": "Março",
    "April": "Abril",
    "May": "Maio",
    "June": "Junho",
    "July": "Julho",
    "August": "Agosto",
    "September": "Setembro",
    "October": "Outubro",
    "November": "Novembro",
    "December": "Dezembro",
}


def formatar_data_br(dt):
  dia_semana = DIAS_SEMANA.get(dt.strftime("%A"), dt.strftime("%A"))
  return f"{dia_semana}, {dt.strftime('%d')} de {MESES.get(dt.strftime('%B'), dt.strftime('%B'))}"


# Estilização CSS Moderna (Glassmorphism & Cores Corporativas)
st.markdown(
    """
    <style>
        .stApp { background-color: #f8f9fc; color: #1e293b; }
        .hero-banner {
            background: linear-gradient(135deg, #660099 0%, #8c19c5 100%);
            padding: 30px 40px;
            border-radius: 20px;
            color: white;
            box-shadow: 0 10px 25px rgba(102, 0, 153, 0.2);
            margin-bottom: 30px;
        }
        .hero-banner h1 { margin: 0; font-size: 2.2rem; font-weight: 800; color: #ffffff; }
        .hero-banner p { margin: 8px 0 0 0; font-size: 1.1rem; color: #f3e8ff; }
        
        .metric-card {
            background: white;
            padding: 20px;
            border-radius: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            border-left: 5px solid #660099;
            text-align: center;
        }
        
        .stTabs [data-baseweb="tab-list"] {
            background-color: #ffffff;
            padding: 8px;
            border-radius: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            gap: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            color: #660099;
            font-weight: 700;
            padding: 10px 24px;
            border-radius: 10px;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            color: #ffffff !important;
            background-color: #660099 !important;
        }
        
        .card-reuniao {
            background: white;
            padding: 22px;
            border-radius: 16px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.04);
            margin-bottom: 18px;
            border-left: 6px solid #660099;
            transition: transform 0.2s;
        }
        .card-concluida { border-left-color: #10b981 !important; background-color: #f0fdf4; }
        .card-cancelada { border-left-color: #ef4444 !important; background-color: #fef2f2; opacity: 0.75; }
        
        .badge-status {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
        }
        .tag-agendada { background-color: #f3e8ff; color: #660099; }
        .tag-concluida { background-color: #d1fae5; color: #065f46; }
        .tag-cancelada { background-color: #fee2e2; color: #991b1b; }
        
        .stButton>button {
            background-color: #660099;
            color: white;
            border-radius: 10px;
            border: none;
            font-weight: 600;
            height: 42px;
            width: 100%;
            box-shadow: 0 4px 10px rgba(102, 0, 153, 0.2);
        }
        .stButton>button:hover { background-color: #4c0073; color: white; }
    </style>
""",
    unsafe_allow_html=True,
)

FILE_PATH = "reunioes.csv"

if not os.path.exists(FILE_PATH) or os.stat(FILE_PATH).st_size == 0:
  df = pd.DataFrame(
      columns=[
          "id",
          "titulo",
          "consultor",
          "email_consultor",
          "data",
          "hora_inicio",
          "link",
          "status",
      ]
  )
  df.to_csv(FILE_PATH, index=False)


def carregar_dados():
  try:
    df_carregado = pd.read_csv(FILE_PATH)
    df_carregado["data"] = df_carregado["data"].astype(str)
    if "status" not in df_carregado.columns:
      df_carregado["status"] = "Agendada"
    if "email_consultor" not in df_carregado.columns:
      df_carregado["email_consultor"] = ""
    return df_carregado
  except:
    return pd.DataFrame(
        columns=[
            "id",
            "titulo",
            "consultor",
            "email_consultor",
            "data",
            "hora_inicio",
            "link",
            "status",
        ]
    )


def salvar_dados(df_novo):
  df_novo.to_csv(FILE_PATH, index=False)


# Função para disparo de e-mail via Outlook do Windows
def enviar_email_consultor(destinatario, consultor, titulo, data, hora, link):
  try:
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    mail.To = destinatario
    mail.Subject = (
        f"🚀 Nova Reunião Agendada: {titulo} ({data} às {hora})"
    )

    mail.Body = f"""Olá, {consultor}!

Você tem uma nova reunião agendada no sistema da Evomobile:

📌 Cliente / Assunto: {titulo}
📅 Data: {data}
⏰ Horário: {hora}
🔗 Link de Acesso: {link}

Por favor, acesse o link no horário marcado. Bom trabalho!

Atenciosamente,
Gestão Evomobile / Vivo Empresas
"""
    mail.Send()
    return True
  except Exception as e:
    print(f"Erro ao enviar e-mail via Outlook: {e}")
    return False


# --- HEADER MODERNO ---
st.markdown(
    """
    <div class='hero-banner'>
        <h1>⚡ EVOMOBILE | AGENDA CORPORATIVA</h1>
        <p>Gestão inteligente de compromissos e reuniões comerciais</p>
    </div>
""",
    unsafe_allow_html=True,
)

df_reunioes = carregar_dados()

# --- PAINEL DE FILTROS AVANÇADOS (Mês, Consultor e Status) ---
with st.expander(
    "🔍 Filtros Avançados (Consultor, Mês e Status)", expanded=True
):
  col_f1, col_f2, col_f3 = st.columns(3)

  with col_f1:
    consultor_busca = st.selectbox(
        "Filtrar por Consultor:", ["Todos os Consultores"] + LISTA_CONSULTORES
    )
  with col_f2:
    meses_disponiveis = ["Todos os Meses"] + list(MESES.values())
    mes_busca = st.selectbox("Filtrar por Mês:", meses_disponiveis, index=0)
  with col_f3:
    status_busca = st.selectbox(
        "Filtrar por Status:", ["Todos", "Agendada", "Concluída", "Cancelada"], index=0
    )

# Aplicação dos Filtros
df_exibicao = df_reunioes.copy()

if not df_exibicao.empty:
  if consultor_busca != "Todos os Consultores":
    df_exibicao = df_exibicao[df_exibicao["consultor"] == consultor_busca]

  if mes_busca != "Todos os Meses":
    mes_num = [k for k, v in MESES.items() if v == mes_busca][0]
    mes_str = f"-{list(MESES.keys()).index(mes_num)+1:02d}-"
    df_exibicao = df_exibicao[df_exibicao["data"].str.contains(mes_str)]

  if status_busca != "Todos":
    df_exibicao = df_exibicao[df_exibicao["status"] == status_busca]

# --- MÉTRICAS RÁPIDAS (KPIs) ---
total_reunioes = len(df_exibicao)
concluidas = (
    len(df_exibicao[df_exibicao["status"] == "Concluída"])
    if not df_exibicao.empty
    else 0
)
canceladas = (
    len(df_exibicao[df_exibicao["status"] == "Cancelada"])
    if not df_exibicao.empty
    else 0
)
agendadas = (
    len(df_exibicao[df_exibicao["status"] == "Agendada"])
    if not df_exibicao.empty
    else 0
)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
  st.markdown(
      f"<div class='metric-card'><h3>📊 Total</h3><h2>{total_reunioes}</h2></div>",
      unsafe_allow_html=True,
  )
with col_m2:
  st.markdown(
      f"<div class='metric-card' style='border-left-color: #660099;'><h3>📌"
      f" Agendadas</h3><h2>{agendadas}</h2></div>",
      unsafe_allow_html=True,
  )
with col_m3:
  st.markdown(
      f"<div class='metric-card' style='border-left-color: #10b981;'><h3>✅"
      f" Concluídas</h3><h2>{concluidas}</h2></div>",
      unsafe_allow_html=True,
  )
with col_m4:
  st.markdown(
      f"<div class='metric-card' style='border-left-color: #ef4444;'><h3>❌"
      f" Canceladas</h3><h2>{canceladas}</h2></div>",
      unsafe_allow_html=True,
  )

st.write("")

# --- PREPARAÇÃO DOS EVENTOS DO CALENDÁRIO ---
calendar_events = []
if not df_exibicao.empty:
  for _, row in df_exibicao.iterrows():
    try:
      start_time = f"{row['data']}T{row['hora_inicio']}:00"
      dt_start = datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S")
      end_time = (dt_start + timedelta(hours=1)).strftime("%Y-%m-%dT%H:%M:%S")

      cor_status = "#660099"
      if row["status"] == "Concluída":
        cor_status = "#10b981"
      elif row["status"] == "Cancelada":
        cor_status = "#ef4444"

      calendar_events.append({
          "title": (
              f"📌 {row['hora_inicio']} - {row['titulo']}"
              f" ({row['consultor']})"
          ),
          "start": start_time,
          "end": end_time,
          "backgroundColor": cor_status,
          "borderColor": cor_status,
      })
    except:
      pass

# Abas do App
tabs = ["📅 Calendário Mensal", "📋 Cronograma e Validação", "➕ Agendar Nova Reunião"]
aba_selecionada = st.tabs(tabs)

# --- TAB 1: CALENDÁRIO MENSAL ---
with aba_selecionada[0]:
  st.markdown("### 📅 Visão Geral do Mês")
  calendar_options = {
      "headerToolbar": {
          "left": "today prev,next",
          "center": "title",
          "right": "dayGridMonth,timeGridWeek",
      },
      "initialView": "dayGridMonth",
      "locale": "pt-br",
      "firstDay": 0,
  }
  calendar(events=calendar_events, options=calendar_options)

# --- TAB 2: CRONOGRAMA E VALIDAÇÃO ---
with aba_selecionada[1]:
  st.markdown("### 🗓️ Validação, Gestão e Exclusão de Reuniões")

  if df_exibicao.empty:
    st.info("Nenhuma reunião encontrada com os filtros selecionados.")
  else:
    df_ordenado = df_exibicao.sort_values(by=["data", "hora_inicio"])

    for idx, row in df_ordenado.iterrows():
      classe_card = (
          "card-concluida"
          if row["status"] == "Concluída"
          else ("card-cancelada" if row["status"] == "Cancelada" else "")
      )
      tag_status = (
          "tag-concluida"
          if row["status"] == "Concluída"
          else ("tag-cancelada" if row["status"] == "Cancelada" else "tag-agendada")
      )

      try:
        dt_obj = datetime.strptime(row["data"], "%Y-%m-%d")
        data_formatada = formatar_data_br(dt_obj)
      except:
        data_formatada = row["data"]

      st.markdown(
          f"""
            <div class='card-reuniao {classe_card}'>
                <div style='display: flex; justify-content: space-between; align-items: center;'>
                    <span style='font-weight: 700; color: #660099; font-size: 18px;'>⏰ {row['hora_inicio']} - {row['titulo']}</span>
                    <span class='badge-status {tag_status}'>{row['status']}</span>
                </div>
                <div style='margin-top: 8px; color: #475569; font-size: 14px;'>
                    📅 <b>Data:</b> {data_formatada} | 👤 <b>Consultor:</b> {row['consultor']} ({row['email_consultor']})
                </div>
                <div style='margin-top: 4px; font-size: 14px;'>
                    🔗 <a href='{row['link']}' target='_blank' style='color:#660099; font-weight:bold;'>Acessar Link da Reunião</a>
                </div>
            </div>
            """,
          unsafe_allow_html=True,
      )

      # Botões de Ação vinculados diretamente ao índice único (idx) do DataFrame principal
      col_v1, col_v2, col_v3, col_v4 = st.columns(4)
      with col_v1:
        if st.button("✅ Concluir", key=f"btn_conc_{idx}"):
          df_reunioes.loc[idx, "status"] = "Concluída"
          salvar_dados(df_reunioes)
          st.rerun()
      with col_v2:
        if st.button("❌ Cancelar", key=f"btn_canc_{idx}"):
          df_reunioes.loc[idx, "status"] = "Cancelada"
          salvar_dados(df_reunioes)
          st.rerun()
      with col_v3:
        if st.button("🔄 Resetar", key=f"btn_reset_{idx}"):
          df_reunioes.loc[idx, "status"] = "Agendada"
          salvar_dados(df_reunioes)
          st.rerun()
      with col_v4:
        if st.button("🗑️ Excluir", key=f"btn_del_{idx}"):
          df_reunioes = df_reunioes.drop(idx)
          salvar_dados(df_reunioes)
          st.rerun()

# --- TAB 3: CADASTRO COM ENVIO DE E-MAIL ---
with aba_selecionada[2]:
  st.markdown("### ➕ Agendar Nova Reunião e Notificar Consultor")
  st.write(
      "Preencha os campos abaixo. Ao salvar, a reunião será registrada e um"
      " e-mail será disparado automaticamente pelo seu Outlook para o"
      " consultor."
  )

  with st.form("form_novo_agendamento"):
    col_nc1, col_nc2 = st.columns(2)
    with col_nc1:
      novo_titulo = st.text_input("Título / Nome do Cliente:")
      novo_consultor = st.selectbox(
          "Consultor Responsável:", LISTA_CONSULTORES
      )
      nova_data = st.date_input("Data da Reunião:", value=date.today())
    with col_nc2:
      nova_hora = st.time_input(
          "Horário de Início:", value=datetime.strptime("09:00", "%H:%M").time()
      )
      novo_link = st.text_input(
          "Link da Reunião (Teams / Meet / Zoom):", value="https://"
      )
      st.info(
          f"📧 E-mail de destino automático:"
          f" **{CONSULTORES_EMAILS[novo_consultor]}**"
      )

    botao_agendar = st.form_submit_button(
        "📅 Agendar e Enviar E-mail via Outlook"
    )

    if botao_agendar:
      if novo_titulo.strip() == "":
        st.error("Por favor, preencha o Título / Nome do Cliente.")
      else:
        email_destino = CONSULTORES_EMAILS[novo_consultor]
        data_str = nova_data.strftime("%Y-%m-%d")
        hora_str = nova_hora.strftime("%H:%M")

        # Disparar e-mail via Outlook do Windows
        enviar_email_consultor(
            email_destino, novo_consultor, novo_titulo, data_str, hora_str, novo_link
        )

        # Gerar ID único
        novo_id = str(int(datetime.now().timestamp() * 1000))
        nova_linha = pd.DataFrame([{
            "id": novo_id,
            "titulo": novo_titulo,
            "consultor": novo_consultor,
            "email_consultor": email_destino,
            "data": data_str,
            "hora_inicio": hora_str,
            "link": novo_link,
            "status": "Agendada",
        }])

        df_reunioes = pd.concat([df_reunioes, nova_linha], ignore_index=True)
        salvar_dados(df_reunioes)

        st.success(
            f"Reunião agendada com sucesso para **{novo_consultor}**! E-mail"
            f" enviado para `{email_destino}`."
        )
        st.rerun()