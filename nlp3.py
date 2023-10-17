import pdfplumber
import re
from unidecode import unidecode
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import RSLPStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

stop_words = set(stopwords.words('portuguese'))
stemmer = RSLPStemmer()

def preprocess_text(text):
    text = text.lower()
    text = unidecode(text)
    text = re.sub(r'[^a-zA-Z0-9 ]', ' ', text)
    
    tokens = word_tokenize(text)
    tokens = [word for word in tokens if word not in stop_words]
    tokens = [stemmer.stem(word) for word in tokens]
    
    return ' '.join(tokens)

data_from_lcp_116 = []

def extract_codes_from_lcp_116():
    global data_from_lcp_116
    if data_from_lcp_116: 
        return data_from_lcp_116

    caminho_pdf = rf"C:\Users\user\Documents\Lcp 116.pdf"
    
    texto_total = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_total += texto + '\n'
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
        return []

    texto_total = unidecode(texto_total).upper().replace(".", "").replace(",", "")
    pattern = re.compile(r'(\d+(?:\.\d+)?)\s-\s(.*?)(?=\d+\s-\s|\Z)', re.DOTALL)
    matches = pattern.findall(texto_total)

    for code, description in matches:
        description = preprocess_text(description)
        data_from_lcp_116.append({"codigo": code, "descricao": description})

    return data_from_lcp_116

vectorizer = TfidfVectorizer(max_df=0.6, min_df=1, ngram_range=(1,2), stop_words=list(stop_words))

def get_jaccard_sim(str1, str2):
    a = set(str1.split())
    b = set(str2.split())
    c = a.intersection(b)
    return float(len(c)) / (len(a) + len(b) - len(c))

def get_best_match_code(num_docto, id_solicitacao):
    caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
    
    texto_total = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_total += texto + '\n'
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
        return None

    # Pré-processando o texto do PDF da nota
    text_from_pdf = preprocess_text(texto_total)

    # Coletando descrições do LCP 116
    data = extract_codes_from_lcp_116()
    descriptions = [item['descricao'] for item in data]

    # Filtrando os tokens com base nas descrições do LCP 116
    tokens_from_pdf = text_from_pdf.split()
    filtered_tokens = [token for token in tokens_from_pdf if any(token in desc for desc in descriptions)]

    if not filtered_tokens:
        print("Nenhum termo relevante encontrado no PDF.")
        return None

    filtered_text_from_pdf = ' '.join(filtered_tokens)
    # print(filtered_text_from_pdf)

    # Comparando a similaridade do texto filtrado com as descrições do LCP 116
    tfidf_matrix = vectorizer.fit_transform([filtered_text_from_pdf] + descriptions)
    cosine_similarities = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:]).flatten()
    
    # similaridade Jaccard
    jaccard_scores = [get_jaccard_sim(filtered_text_from_pdf, desc) for desc in descriptions]
    avg_scores = [0.6*cos + 0.4*jac for cos, jac in zip(cosine_similarities, jaccard_scores)]
    
    best_match_idx = avg_scores.index(max(avg_scores))
    
    return data[best_match_idx]['codigo']


def get_child_codes(parent_code):
    return [item for item in extract_codes_from_lcp_116() if item['codigo'].startswith(parent_code) and len(item['codigo']) == len(parent_code) + 2]


def refine_best_match_code(num_docto, id_solicitacao):
    best_code = get_best_match_code(num_docto, id_solicitacao)
    
    # Se for um código "pai" com apenas 1 ou 2 caracteres
    if len(best_code) in [1, 2]:
        # Pegue somente os filhos desse código
        data_subset = get_child_codes(best_code)
        
        # E faça uma reanálise somente com os filhos
        descriptions_subset = [item['descricao'] for item in data_subset]
        
        # Verifique se existe uma descrição válida para processar
        if not descriptions_subset:
            print("Não foram encontradas descrições válidas para o código pai.")
            return None
        
        tfidf_matrix_subset = vectorizer.fit_transform([preprocess_text(num_docto)] + descriptions_subset)
        cosine_similarities_subset = cosine_similarity(tfidf_matrix_subset[0:1], tfidf_matrix_subset[1:]).flatten()
        
        # Similaridade Jaccard
        jaccard_scores_subset = [get_jaccard_sim(preprocess_text(num_docto), desc) for desc in descriptions_subset]
        avg_scores_subset = [0.6*cos + 0.4*jac for cos, jac in zip(cosine_similarities_subset, jaccard_scores_subset)]
        
        best_match_idx_subset = avg_scores_subset.index(max(avg_scores_subset))
        
        return data_subset[best_match_idx_subset]['codigo']
    
    else:
        # Se já for um código filho, retorne diretamente
        return best_code

# result = refine_best_match_code('1', '00')
# print(result)
