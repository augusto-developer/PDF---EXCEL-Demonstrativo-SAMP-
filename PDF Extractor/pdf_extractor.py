import pdfplumber
import pandas as pd

def extract_all_text_from_pdf(pdf_path, output_excel):
    """Extrai todo o texto de um PDF e salva em um arquivo Excel"""
    all_text = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    all_text.append({"Página": page_number + 1, "Conteúdo": text})
        
        if all_text:
            df = pd.DataFrame(all_text)
            df.to_excel(output_excel, index=False, engine='openpyxl')
            return True
        return False
    
    except Exception as e:
        raise Exception(f"Erro na extração do PDF: {str(e)}")