import pandas as pd
import pyautogui
import time
import keyboard  # Biblioteca para detectar tecla pressionada
import sys  # Para evitar o fechamento do CMD imediatamente

# Carregar a planilha
arquivo_excel = "*********"
try:
    planilha = pd.read_excel(arquivo_excel, dtype={"Renda_Bruta": str, "Comissao": str})
except Exception as e:
    print(f"Erro ao abrir o arquivo Excel: {e}")
    sys.exit()

# Valores fixos
valores_fixos = {
    "CNPJ_do_Locador": "*********",
    "Nome_do_Locador": " *********",
    "Endereco_do_imovel": "*********",
    "CEP": "*********"
}

coordenadas_fixas = {
    "CNPJ_do_Locador": (175, 144),
    "Nome_do_Locador": (335, 144),
    "Endereco_do_imovel": (184, 296),
    "CEP": (484, 337)
}

# Dicionário que associa colunas do Excel às coordenadas na tela
campos = {
    "CPF_do_Locatario": (174, 198),
    "Nome_do_Locatario": (339, 200),
    "Data_do_Contrato": (331, 238),
    "Renda_Bruta": (323,613),
    "Comissao": (463, 616)
}

# Coordenadas para preenchimento de UF e Município
coordenada_uf_1 = (195, 336)  
coordenada_uf_2 = (175, 444)
coordenada_municipio_1 = (374, 337)
coordenada_municipio_2 = (422, 458)
numero_cliques_municipio = 104 
coordenada_municipio_final = (361, 459)

# Coordenadas dos botões
botao_ok = (705, 118)  
botao_incluir = (705, 118)  

# Tempo antes de começar (para focar na janela correta)
time.sleep(5)

# Percorrer cada cliente (linha do Excel)
for index, linha in planilha.iterrows():
    if keyboard.is_pressed("esc"):  # Verifica se a tecla "Esc" foi pressionada
        print("Execução interrompida pelo usuário.")
        sys.exit()

    try:
        # Preencher CNPJ do Locador caractere por caractere
        pyautogui.click(coordenadas_fixas["CNPJ_do_Locador"])
        time.sleep(0.5)  # Pequena pausa para garantir que o campo está focado

        # Apagar qualquer texto pré-existente (caso o campo não esteja vazio)
        pyautogui.hotkey("ctrl", "a")  # Seleciona tudo
        time.sleep(0.5)
        pyautogui.press("backspace")  # Apaga o conteúdo

        # Digitar caractere por caractere com pausas
        for char in valores_fixos["CNPJ_do_Locador"].strip():  # Remove espaços extras
            pyautogui.write(char)
            time.sleep(0.1)  # Pequena pausa para evitar falhas na digitação

        time.sleep(0.5)
        
        # Preencher Nome do Locador caractere por caractere
# Preencher Nome do Locador caractere por caractere
        pyautogui.click(coordenadas_fixas["Nome_do_Locador"])
        time.sleep(0.5)  # Pequena pausa para garantir que o campo está focado

        # Apagar qualquer texto pré-existente (caso o campo não esteja vazio)
        pyautogui.hotkey("ctrl", "a")  # Seleciona tudo
        time.sleep(0.1)
        pyautogui.press("backspace")  # Apaga o conteúdo

        # Digitar caractere por caractere com pausas
        for char in valores_fixos["Nome_do_Locador"].strip():  # Remove espaços extras
            pyautogui.write(char)
            time.sleep(0.01)  # Pequena pausa para evitar falhas na digitação

        time.sleep(0.5)  # Pequena pausa antes do próximo campo


        # Preencher Endereço normalmente
        pyautogui.click(coordenadas_fixas["Endereco_do_imovel"])
        pyautogui.write(valores_fixos["Endereco_do_imovel"])
        time.sleep(0.01)

        # Preencher CEP caractere por caractere
        pyautogui.click(coordenadas_fixas["CEP"])
        for char in valores_fixos["CEP"]:
            pyautogui.write(char)
            time.sleep(0.01)

        # Preencher CPF do Locatário (dígito por dígito)
        pyautogui.click(campos["CPF_do_Locatario"])
        for char in str(linha["CPF_do_Locatario"]):
            pyautogui.write(char)
            time.sleep(0.01)
        
        # Preencher Nome do Locatário
        pyautogui.click(campos["Nome_do_Locatario"])
        pyautogui.write(str(linha["Nome_do_Locatario"]) if pd.notna(linha["Nome_do_Locatario"]) else "")
        time.sleep(0.5)
        
        # Preencher Data do Contrato formatada corretamente
        pyautogui.click(campos["Data_do_Contrato"])
        data_contrato = linha["Data_do_Contrato"]
        if pd.notna(data_contrato):
            if isinstance(data_contrato, pd.Timestamp):
                data_contrato = data_contrato.strftime('%d/%m/%Y')
            pyautogui.write(data_contrato)
        time.sleep(0.5)

       # Preencher Renda Bruta caractere por caractere
        pyautogui.click(campos["Renda_Bruta"])
        renda_bruta = str(linha["Renda_Bruta"]) if pd.notna(linha["Renda_Bruta"]) else ""
        renda_bruta = renda_bruta.replace(".", ",")
        for char in renda_bruta:
            pyautogui.write(char)
            time.sleep(0.5)
        time.sleep(0.5)

        # Preencher Comissão caractere por caractere
        pyautogui.click(campos["Comissao"])
        comissao = str(linha["Comissao"]) if pd.notna(linha["Comissao"]) else ""
        comissao = comissao.replace(".", ",")
        for char in comissao:
            pyautogui.write(char)
            time.sleep(0.5)
        time.sleep(0.1)

        # Preencher UF
        pyautogui.click(coordenada_uf_1)
        time.sleep(0.05)
        pyautogui.click(coordenada_uf_2)
        time.sleep(0.8)

        # Preencher Município
        pyautogui.click(coordenada_municipio_1)
        time.sleep(0.05)
        pyautogui.click(coordenada_municipio_final)
        time.sleep(0.05)

        # Clicar no tipo de imóvel
        pyautogui.click(457, 240)

        # Clicar no botão OK para enviar o formulário
        pyautogui.click(botao_ok)
        time.sleep(2)

        # Clicar no botão Incluir
        pyautogui.click(botao_incluir)
        time.sleep(1)

    except Exception as e:
        print(f"Erro ao preencher formulário para o cliente {index + 1}: {e}")
        sys.exit()

print("Preenchimento concluído!")
