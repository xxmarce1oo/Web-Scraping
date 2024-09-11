import os
import requests
from bs4 import BeautifulSoup
from fpdf import FPDF

# Função para garantir que o texto seja codificado corretamente
def safe_text(text):
    return text.encode('latin-1', 'replace').decode('latin-1')

# URL do site com as avaliações
url = 'https://www.doctoralia.com.br/clinicas/hospital-marcelino-champagnat#facility-opinion-stats'

# Cabeçalho HTTP
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Fazer a requisição à página
response = requests.get(url, headers=headers)

# Verificar se a requisição foi bem-sucedida
if response.status_code == 200:
    # Criar o objeto BeautifulSoup para análise do HTML
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Encontrar todas as avaliações dentro do elemento "media opinion text-break"
    avaliacoes = soup.find_all('div', class_='media opinion text-break')
    
    # Criar um objeto PDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Definir a fonte
    pdf.set_font("Arial", size=12)
    
    # Percorrer cada avaliação e extrair as informações
    for avaliacao in avaliacoes:
        # Nome do paciente usando itemprop="name"
        nome = avaliacao.find('span', itemprop='name').text.strip() if avaliacao.find('span', itemprop='name') else 'Nome não encontrado'
        
        # Nota do paciente (ajustar com a classe correta, se necessário)
        nota = avaliacao.find('span', class_='star-rating__rating').text.strip() if avaliacao.find('span', class_='star-rating__rating') else 'Nota não encontrada'
        
        # Comentário do paciente usando itemprop="description"
        comentario = avaliacao.find('p', itemprop='description').text.strip() if avaliacao.find('p', itemprop='description') else 'Comentário não encontrado'

        # Escrever no PDF (convertendo o texto para o encoding correto)
        pdf.cell(200, 10, txt=safe_text(f"Nome: {nome}"), ln=True, align='L')
        pdf.cell(200, 10, txt=safe_text(f"Nota: {nota}"), ln=True, align='L')
        pdf.multi_cell(200, 10, txt=safe_text(f"Comentário: {comentario}"), align='L')
        pdf.cell(200, 10, txt=safe_text('-' * 40), ln=True, align='L')
    
    # Caminho para salvar o arquivo na área de trabalho
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "avaliacoes.pdf")
    
    # Salvar o PDF na área de trabalho
    pdf.output(desktop_path)
    print(f"PDF salvo com sucesso em: {desktop_path}")

else:
    print(f"Erro ao acessar a página: {response.status_code}")