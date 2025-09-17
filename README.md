# Projeto_certificado
Este projeto em python automatiza a criação de certificados de qualidade. Ele extrai dados de PDFs (NF e SAP) e de uma planilha Excel, preenche um template Word e o junta com os certificados correspondentes, gerando um único PDF final nomeado com o cliente e o número da NF. Sua função é agilizar a emissão de documentação.

# ==================================================================================
import pdfplumber
import PyPDF2
import re
import os
import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate

def extrair_dados_nf_pdf(caminho_pdf):
    """
    Extrai os dados principais de um PDF de Nota Fiscal.
    """
    dados = {
        'nota_fiscal': None, 'data_emissao': None, 'pedido_cliente': None,
        'razao_social': None, 'itens_produto': [], 'req_interna': None
    }
    texto_completo = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(x_tolerance=2)
                if texto_pagina: texto_completo += texto_pagina + "\n"
        
        match = re.search(r'Nº\s+(\d+)', texto_completo)
        if match: dados['nota_fiscal'] = match.group(1).strip()

        match = re.search(r'(\d{2}/\d{2}/\d{4})', texto_completo)
        if match: dados['data_emissao'] = match.group(1).strip()
        
        match = re.search(r'PEDIDO DE COMPRA\s*-?\s*(\d+)', texto_completo, re.DOTALL)
        if match: dados['pedido_cliente'] = match.group(1).strip()

        match = re.search(r'Pedido de venda\.\s*(\d+)', texto_completo)
        if match: dados['req_interna'] = match.group(1).strip()
        
        bloco_produtos_match = re.search(r'DADOS DO PRODUTO / SERVIÇO(.*?)CÁLCULO DO ISSQN', texto_completo, re.DOTALL)
        if bloco_produtos_match:
            texto_produtos = bloco_produtos_match.group(1)
            padrao_item = re.compile(r'(E\d+[A-Z0-9]+)\s+(.*?)\s+(\d{4}\.\d{2}\.\d{2})[^\n]*\n(.*?ITEM\s+(\d+))', re.DOTALL)
            matches = padrao_item.finditer(texto_produtos)
            for match in matches:
                descricao_parte1 = match.group(2).strip()
                descricao_parte2 = match.group(4).strip()
                descricao_final = f"{descricao_parte1} {descricao_parte2}"
                dados['itens_produto'].append({'descricao_pdf': descricao_final})
    except Exception as e:
        print(f"ERRO ao processar o PDF da NF '{caminho_pdf}': {e}")
    return dados

def extrair_dados_sap(caminho_pdf):
    """
    Extrai o nome do cliente e a lista de unidades metálicas do PDF do SAP.
    (Versão final com Regex baseada na estrutura: 5 letras e 2/3 números)
    """
    dados_sap = {'razao_social': None, 'unidades_metalicas': []}
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = "".join([p.extract_text(x_tolerance=2) or "" for p in pdf.pages])

        
        match = re.search(r'CLIENTE/CLIENT[^\n]*\n([^\d]+)', texto_completo, re.IGNORECASE)
        if match:
            dados_sap['razao_social'] = match.group(1).strip()

        
        unidades_brutas = re.findall(r'\b([A-Za-z]{5}\d{2,3})\b', texto_completo)
        
        
        if unidades_brutas:
            dados_sap['unidades_metalicas'] = list(dict.fromkeys(unidades_brutas))

    except Exception as e:
        print(f"AVISO: Não foi possível ler o PDF do SAP '{caminho_pdf}': {e}")
    return dados_sap

def salvar_docx_como_pdf(docx_path, pdf_path):
    """
    Converte um arquivo .docx para .pdf.
    """
    try:
        convert(docx_path, pdf_path)
        print("-> Documento Word convertido para PDF.")
        return True
    except Exception as e:
        print(f"ERRO ao converter Word para PDF: {e}")
        return False

def fundir_pdfs(lista_arquivos, pdf_saida):
    """
    Une múltiplos arquivos PDF em um único arquivo.
    """
    try:
        merger = PyPDF2.PdfMerger()
        for pdf in lista_arquivos:
            if os.path.exists(pdf): 
                merger.append(pdf)
            else:
                print(f"AVISO: O arquivo de certificado '{pdf}' não foi encontrado para ser fundido.")
        merger.write(pdf_saida)
        merger.close()
        print(f"SUCESSO: Arquivo final salvo em: '{pdf_saida}'")
        return True
    except Exception as e:
        print(f"ERRO ao fundir os PDFs: {e}")
        return False

if __name__ == "__main__":
    
    NOME_ARQUIVO_NF = "NF.pdf"
    NOME_ARQUIVO_SAP = "SAP.pdf"
    
    
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base_dir = os.getcwd()

    caminho_pdf_nf = os.path.join(base_dir, NOME_ARQUIVO_NF)
    caminho_pdf_sap = os.path.join(base_dir, NOME_ARQUIVO_SAP)
    caminho_planilha_certificados = os.path.join(base_dir, "certificados.xlsx")
    pasta_dos_certificados = os.path.join(base_dir, "certificados")
    caminho_template_word = os.path.join(base_dir, "branco.docx")
    caminho_word_saida = os.path.join(base_dir, "documento_preenchido_temp.docx")
    caminho_pdf_saida_word = os.path.join(base_dir, "documento_preenchido_temp.pdf")
    
    print(">>> INICIANDO PROCESSAMENTO...")
    
    try:
       
        dados_finais = extrair_dados_nf_pdf(caminho_pdf_nf)
        dados_sap = extrair_dados_sap(caminho_pdf_sap)
        
        if not dados_finais or not dados_finais.get('itens_produto'):
            raise Exception("A extração de itens da Nota Fiscal falhou ou não encontrou itens.")
        
        dados_finais['razao_social'] = dados_sap.get('razao_social')
        unidades_sap = dados_sap.get('unidades_metalicas', [])

        print(f"-> SUCESSO: {len(dados_finais['itens_produto'])} descrições de itens encontradas na NF.")
        print(f"-> SUCESSO: {len(unidades_sap)} Unidades Metálicas únicas encontradas no SAP.")

        
        df = pd.read_excel(caminho_planilha_certificados, 
                           dtype={
                               'Cód. Unid. Metálica': str, 
                               'Número Certificado': str
                           })

        lista_unidades_final = []
        certificados_a_anexar = []
        arquivos_na_pasta = os.listdir(pasta_dos_certificados)

        for codigo in unidades_sap:
            info_planilha = df[df['Cód. Unid. Metálica'] == codigo]
            
            if not info_planilha.empty:
                primeira_ocorrencia = info_planilha.iloc[0]
                desc = primeira_ocorrencia['Desc. Unid. Metálica']
                lista_unidades_final.append({'desc': desc})
                
                cert_numero_base = primeira_ocorrencia['Número Certificado']
                
                if pd.notna(cert_numero_base) and str(cert_numero_base).strip():
                    cert_numero_base = str(cert_numero_base).strip()
                    prefixo_busca = cert_numero_base.zfill(10)
                    
                    certificado_encontrado = False
                    for nome_arquivo in arquivos_na_pasta:
                        if nome_arquivo.startswith(prefixo_busca):
                            cert_path = os.path.join(pasta_dos_certificados, nome_arquivo)
                            if cert_path not in certificados_a_anexar:
                                certificados_a_anexar.append(cert_path)
                            certificado_encontrado = True
                            break
                    
                    if not certificado_encontrado:
                        print(f"AVISO: Certificado '{prefixo_busca}' para a unidade '{codigo}' listado na planilha, mas arquivo PDF não encontrado na pasta 'certificados'.")
                else:
                    print(f"AVISO: Unidade '{codigo}' encontrada na planilha, mas não possui um número de certificado associado.")
            
            else:
                desc = f"{codigo} (Não cadastrado na planilha)"
                lista_unidades_final.append({'desc': desc})
                print(f"AVISO: Unidade '{codigo}' do SAP não foi encontrada na planilha 'certificados.xlsx'.")

        dados_finais['unidades_metalicas'] = lista_unidades_final
        
        print(f"-> Encontrados {len(certificados_a_anexar)} certificados únicos para anexar.")
        
        
        doc = DocxTemplate(caminho_template_word)
        doc.render(dados_finais)
        doc.save(caminho_word_saida)
        print("-> Template preenchido com sucesso.")

      
        if salvar_docx_como_pdf(caminho_word_saida, caminho_pdf_saida_word):
            nf_numero = dados_finais.get('nota_fiscal', 'SEM_NUMERO')
            
         
            nome_cliente_completo = dados_finais.get('razao_social')
            primeiro_nome_cliente = "Cliente" 
            if nome_cliente_completo and nome_cliente_completo.strip():
                primeiro_nome_cliente = nome_cliente_completo.strip().split(' ')[0]
            caminho_pdf_saida_final = os.path.join(base_dir, f"{primeiro_nome_cliente} NF {nf_numero}.pdf")
            
            lista_arquivos_fundir = [caminho_pdf_saida_word] + certificados_a_anexar
            fundir_pdfs(lista_arquivos_fundir, caminho_pdf_saida_final)
            
            try: 
                os.remove(caminho_word_saida)
                os.remove(caminho_pdf_saida_word)
                print("-> Arquivos temporários removidos.")
            except OSError as e:
                print(f"AVISO: Não foi possível remover arquivos temporários: {e}")
                pass
            
    except FileNotFoundError as e:
         print(f"\n>>> ERRO CRÍTICO DE ARQUIVO NÃO ENCONTRADO: '{e.filename}'. Verifique se todos os arquivos necessários (NF.pdf, SAP.pdf, certificados.xlsx, branco.docx) estão na pasta.")
    except KeyError as e:
        print(f"\n>>> ERRO CRÍTICO DE COLUNA NÃO ENCONTRADA: {e}. Verifique se os nomes das colunas na planilha 'certificados.xlsx' estão corretos. Devem ser 'Cód. Unid. Metálica', 'Desc. Unid. Metálica' e 'Número Certificado'.")
    except Exception as e:
        print(f"\n>>> ERRO INESPERADO NO FLUXO PRINCIPAL: {type(e).__name__} - {e}")
        
    print("\n>>> FIM.")
