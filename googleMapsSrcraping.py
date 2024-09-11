import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from fpdf import FPDF
import time

# Função para garantir que o texto seja codificado corretamente
def safe_text(text):
    return text.encode('latin-1', 'replace').decode('latin-1')

# Configurar o Selenium WebDriver
options = webdriver.ChromeOptions()
# Remover a linha abaixo para não executar em modo headless
# options.add_argument('--headless')  # Executa o Chrome em modo headless (sem abrir o navegador)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# URL do site com as avaliações
url = 'https://www.google.com/maps/place/Hospital+Erasto+Gaertner/@-25.4402428,-49.2480384,15z/data=!4m18!1m9!3m8!1s0x94dce45382833a69:0xbcacc1ec0b2350db!2sHospital+S%C3%A3o+Marcelino+Champagnat!8m2!3d-25.4359584!4d-49.246138!9m1!1b1!16s%2Fg%2F11gzk88zh!3m7!1s0x94dce51c1ddbe78f:0x88fdecec6e651de7!8m2!3d-25.4533705!4d-49.238806!9m1!1b1!16s%2Fg%2F1yfjkzz97?entry=ttu&g_ep=EgoyMDI0MDkwOC4wIKXMDSoASAFQAw%3D%3D'

# Acessar a página
driver.get(url)
time.sleep(5)  # Esperar o carregamento inicial da página

# Definir o limite de avaliações
limite_avaliacoes = 8  # Coloque aqui o número máximo de avaliações que você deseja capturar

# Lista para armazenar todas as avaliações
avaliacoes = []

# Loop para continuar o scroll até que o número máximo de avaliações seja atingido
while len(avaliacoes) < limite_avaliacoes:
    # Scroll para carregar mais avaliações
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)  # Aguardar o carregamento após o scroll

    # Contar o número de avaliações visíveis
    novas_avaliacoes = driver.find_elements(By.CLASS_NAME, 'jftiEf')
    
    # Adicionar novas avaliações à lista
    for avaliacao in novas_avaliacoes:
        if avaliacao not in avaliacoes:
            avaliacoes.append(avaliacao)
    
    # Verificar se novas avaliações foram carregadas
    if len(novas_avaliacoes) == 0:
        break  # Parar o loop se não houver mais avaliações sendo carregadas

# Criar um objeto PDF
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()

# Definir a fonte
pdf.set_font("Arial", size=12)

# Contador de avaliações processadas
contador_avaliacoes = 0

# Percorrer as avaliações e extrair as informações (limitando ao número desejado)
for avaliacao in avaliacoes[:limite_avaliacoes]:  # Limitar ao número máximo de avaliações
    try:
        # Clique no botão "Mais" se ele estiver presente
        try:
            mais_button = avaliacao.find_element(By.CLASS_NAME, 'w8nwRe')  # Classe do botão "Mais"
            mais_button.click()
            time.sleep(1)  # Esperar um momento para o texto expandido ser carregado
        except:
            pass  # Se não houver botão "Mais", siga em frente

        # Nome do paciente
        nome = avaliacao.find_element(By.CLASS_NAME, 'd4r55').text.strip()
    except:
        nome = 'Nome não encontrado'

    try:
        # Nota do paciente
        nota = avaliacao.find_element(By.CLASS_NAME, 'kvMYJc').get_attribute('aria-label')
    except:
        nota = 'Nota não encontrada'

    try:
        # Comentário do paciente (depois do clique no "Mais", se houver)
        comentario = avaliacao.find_element(By.CLASS_NAME, 'MyEned').text.strip()
    except:
        comentario = 'Comentário não encontrado'

    # Escrever no PDF (convertendo o texto para o encoding correto)
    pdf.cell(200, 10, txt=safe_text(f"Nome: {nome}"), ln=True, align='L')
    pdf.cell(200, 10, txt=safe_text(f"Nota: {nota}"), ln=True, align='L')
    pdf.multi_cell(200, 10, txt=safe_text(f"Comentário: {comentario}"), align='L')
    pdf.cell(200, 10, txt=safe_text('-' * 40), ln=True, align='L')

    # Incrementar o contador de avaliações processadas
    contador_avaliacoes += 1

# Caminho para salvar o arquivo na área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "avaliacoes.pdf")

# Salvar o PDF na área de trabalho
pdf.output(desktop_path)
print(f"PDF salvo com sucesso em: {desktop_path}")

# Imprimir o total de avaliações processadas
print(f"Total de avaliações processadas: {contador_avaliacoes}")

# Fechar o navegador
driver.quit()