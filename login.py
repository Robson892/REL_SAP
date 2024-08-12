import streamlit as st
import importlib

# Definindo uma chave de estado para o login
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# Função para mostrar a tela de login
def show_login():
    st.title('Tela de Login')
    
    username = st.text_input('Nome de usuário')
    password = st.text_input('Senha', type='password')
    # Estilizando o botão de login com HTML/CSS
    st.markdown("""
        <style>
        div.stButton > button:first-child {
            background-color: blue;
            color: white;
            height: 3em;
            width: 10em;
            border-radius:10px;
            border:2px solid #ffffff;
            font-size:20px;
            font-weight: bold;
            margin-top: 20px;
        }
        div.stButton > button:hover {
            background-color: darkblue;
            color: white;
            border:2px solid #ffffff;
        }
        </style>
        """, unsafe_allow_html=True)
    if st.button('Login'):
        if username == "usuario" and password == "senha123":
            st.session_state['logged_in'] = True
            st.experimental_rerun()  # Força o recarregamento da aplicação imediatamente após o login
        else:
            st.error('Nome de usuário ou senha incorretos')

# Função para carregar a página principal do arquivo `rel_sap.py`
def show_rel_sap():
    rel_sap = importlib.import_module('rel_sap')
    rel_sap.sap()  # Chamando a função sap() definida em rel_sap.py

# Controle de navegação entre as páginas
if st.session_state['logged_in']:
    show_rel_sap()
else:
    show_login()
