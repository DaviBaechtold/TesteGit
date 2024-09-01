import os
import tempfile
from azure.keyvault.secrets import SecretClient
from azure.identity import DefaultAzureCredential
import requests
import fitz  # PyMuPDF
from PIL import Image
import io
import base64
import http.client
import json
from flask import Flask, request, jsonify, send_file
import pythoncom  # Import necessário para inicializar COM
from io import BytesIO
import uuid
from docx2pdf import convert

app = Flask(__name__)

# Configuração do Azure Key Vault
keyVaultName = "brf-kv-servicenow-dev"
KVUri = f"https://{keyVaultName}.vault.azure.net"

credential = DefaultAzureCredential()
client = SecretClient(vault_url=KVUri, credential=credential)

# Recuperando segredos do Key Vault
username = client.get_secret("username").value
password = client.get_secret("password").value

# Endpoint base
base_url = "https://brfdev.service-now.com/api/now/attachment/{sys_id}/file"

# Funções e rotas Flask
def send_to_servicenow_async(sn_url, sn_payload, auth):
    try:
        print("Iniciando envio para ServiceNow...\n")
        response = requests.put(sn_url, json=sn_payload, auth=auth, timeout=30)  # Adicionando timeout de 30 segundos
        print(f"Envio para ServiceNow concluído com status: {response.status_code}\n")
    except requests.Timeout:
        error_message = "Erro: Timeout ao tentar enviar dados para o ServiceNow após 30 segundos.\n"
        print(error_message)
        sn_payload['u_erro_message'] = error_message
    except Exception as e:
        error_message = f"Erro ao enviar dados para ServiceNow: {str(e)}\n"
        print(error_message)
        sn_payload['u_erro_message'] = error_message

@app.route('/process', methods=['POST'])
def process_document():
    pythoncom.CoInitialize()  # Inicializa o COM
    print("Iniciando processamento do documento...\n")
    try:
        data = request.get_json()
        sys_id = data.get('Attach_id')
        change_id = data.get('Change_id')
        file_name = data.get('Attach_name')  # Nome do arquivo para diferenciar ETS e ESF

        if not sys_id:
            print("Erro: sys_id não fornecido.")
            return jsonify({"error": "sys_id not provided"}), 400

        if not change_id:
            print("Erro: Change SysId não fornecido.")
            return jsonify({"error": "Change SysId not provided"}), 400

        if not file_name:
            print("Erro: Nome do arquivo não fornecido.")
            return jsonify({"error": "file_name not provided"}), 400

        url = base_url.format(sys_id=sys_id)

        print(f"Recebido dados: sys_id={sys_id}, change_id={change_id}, file_name={file_name}\n")

        print(f"Fazendo requisição GET para URL: {url}\n")
        response = requests.get(url, auth=(username, password))

        # Verificando se a requisição foi bem-sucedida
        if response.status_code == 200:
            content_type = response.headers['Content-Type']
            print(f"Requisição bem-sucedida. Tipo de conteúdo: {content_type}\n")
            
            # Usando BytesIO para armazenar o conteúdo em memória
            file_in_memory = BytesIO(response.content)

            if 'application/pdf' in content_type:
                print("PDF recebido.\n")
                
                # Processando o PDF diretamente da memória
                process_pdf(file_in_memory, change_id, file_name)
                return jsonify({"message": "Data processing started."})

            elif 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in content_type:
                print("DOCX recebido.\n")
                
                # Criando um diretório temporário
                with tempfile.TemporaryDirectory() as tmp_dir:
                    # Salvando o DOCX em um arquivo temporário dentro do diretório
                    tmp_docx_path = os.path.join(tmp_dir, f"{uuid.uuid4()}.docx")
                    with open(tmp_docx_path, 'wb') as tmp_docx:
                        tmp_docx.write(file_in_memory.getvalue())

                    # Convertendo DOCX para PDF em arquivo temporário dentro do diretório
                    tmp_pdf_path = os.path.join(tmp_dir, f"{uuid.uuid4()}.pdf")
                    convert(tmp_docx_path, tmp_pdf_path)
                    print(f"DOCX convertido para PDF em: {tmp_pdf_path}\n")
                    
                    # Lendo o PDF temporário para memória
                    with open(tmp_pdf_path, 'rb') as pdf_file:
                        pdf_in_memory = BytesIO(pdf_file.read())

                    # Processando o PDF gerado em memória
                    process_pdf(pdf_in_memory, change_id, file_name)

                return jsonify({"message": "Data processing started."})
            else:
                print("Erro: Tipo de arquivo não suportado.\n")
                return jsonify({"error": "Unsupported file type"}), 400
        else:
            error_message = f"Erro ao recuperar o documento. Status code: {response.status_code}\n"
            print(error_message)
            return jsonify({"error": error_message}), 500
    except Exception as e:
        error_message = f"Erro durante o processamento do documento: {str(e)}\n"
        print(error_message)
        return jsonify({"error": error_message}), 500
    finally:
        pythoncom.CoUninitialize()  # Finaliza o COM
        print("[FLASK] COM finalizado.\n")

def generate_payload(document_type, image_urls):
    if document_type == "ETS":
        return json.dumps({
            "messages": [
                {
                    "role": "system",
                    "content": "You are an assistant specialized in evaluating test evidence documents (ETS) for developments within the ServiceNow environment. Your task is to evaluate these documents based on best development practices and the coherence of the documentation, and provide a score from 0 to 10. Specifically, you should check the following aspects:\n\n1. **Test Planning and Execution:** Check the detailed test scenarios, execution steps, and evidence of tests. Ensure that screenshots and other evidence are clear and appropriately referenced.\n2. **Identification:** Verify if the document properly identifies the test cases, test environment, and test results.\n3. **General Guidelines:** Ensure that the document adheres to the general guidelines for creating test evidence documentation, highlighting the necessary ETS for correct implementation.\n\nUse the provided example documents to guide your evaluation."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Please analyze the test evidence document (ETS) provided."
                        }
                    ] + image_urls
                }
            ],
            "max_tokens": 1000,
            "stream": False
        })
    elif document_type == "ESF":
        return json.dumps({
            "messages": [
                {
                    "role": "system",
                    "content": "You are an assistant specialized in evaluating functional specification documents (ESF) for developments within the ServiceNow environment. Your task is to evaluate these documents based on best development practices and the coherence of the documentation, and provide a score from 0 to 10. Specifically, you should check the following aspects:\n\n1. **Identification:** Verify if the document properly identifies the project, author, date, and process/module.\n2. **Revision History:** Check if the document includes a detailed revision history with dates, summaries of changes, and the person who made the changes.\n3. **Use Case Specification:** Evaluate the clarity and completeness of use case specifications, including preconditions, postconditions, main workflows, alternative flows, exceptions, and business rules.\n4. **Components Description:** Ensure that all components, such as client scripts, UI policies, and flow designers, are well described with proper links to their ServiceNow implementations.\n5. **Integrations:** Confirm that the document details any system integrations, the purpose of the integrations, and the systems involved.\n6. **Layout and Field Definitions:** Verify que field definitions include field names, types, sizes, and valid values, along with additional information.\n7. **Translations:** Ensure that all necessary translations are provided and are accurate.\n\nUse the provided example documents to guide your evaluation."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Please analyze the functional specification document (ESF) provided."
                        }
                    ] + image_urls
                }
            ],
            "max_tokens": 1000,
            "stream": False
        })
    else:
        raise ValueError("Unsupported document type")

def process_pdf(file_in_memory, change_id, file_name):
    print("Iniciando processamento do PDF.\n")
    
    # Identificar o tipo de documento baseado no nome do arquivo
    document_type = "ETS" if "ETS" in file_name else "ESF"
    
    # Abrindo o PDF com PyMuPDF a partir da memória
    pdf_document = fitz.open(stream=file_in_memory, filetype="pdf")
    
    # Lista para armazenar as URLs de imagens
    image_urls = []
    
    # Iterando sobre cada página do PDF
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        
        # Convertendo a página para uma imagem PNG
        img_data = pix.tobytes()
        img = Image.open(io.BytesIO(img_data))
        
        # Convertendo a imagem PNG para base64
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
        
        # Criando a URL data:image/png;base64
        image_url = f"data:image/png;base64,{img_base64}"
        image_urls.append({
            "type": "image_url",
            "image_url": {
                "url": image_url
            }
        })

    print("Processamento do PDF concluído. Iniciando envio ao Azure OpenAI...\n")
    
    # Gerar o payload com base no tipo de documento
    payload = generate_payload(document_type, image_urls)
    
    headers = {
        'Content-Type': 'application/json',
        'api-key': '4ae6980dac9c474cb88ac78437d1619e'
    }
    
    try:
        # Fazendo a requisição POST para o serviço do Azure OpenAI
        conn = http.client.HTTPSConnection("brf-openai-devlifecycle-lab.openai.azure.com")
        conn.request("POST", "/openai/deployments/gpt-4/chat/completions?api-version=2023-03-15-preview", payload, headers)
        res = conn.getresponse()
        data = res.read()
        response_content = json.loads(data.decode('utf-8'))["choices"][0]["message"]["content"]  # Extrai apenas o "content"
        error_message = ""
    except Exception as e:
        response_content = ""
        error_message = f"Erro ao se comunicar com o Azure OpenAI: {str(e)}\n"
        print(error_message)

    print("Resposta do Azure OpenAI recebida. Iniciando envio ao ServiceNow...\n")

    # Preparando os dados para enviar ao ServiceNow
    sn_url = f"https://brfdev.service-now.com/api/brff/brf_azure_openai/description/{change_id}"
    ets_description = json.dumps({"response": response_content}) if document_type == "ETS" else ""
    esf_description = json.dumps({"response": response_content}) if document_type == "ESF" else ""

    sn_payload = {
        "ets_description": ets_description,
        "esf_description": esf_description,
        "u_erro_message": error_message
    }

    # Enviando a requisição POST de forma assíncrona
    send_to_servicenow_async(sn_url, sn_payload, (username, password))
    print("Envio ao ServiceNow iniciado.\n")

if __name__ == '__main__':
    print("Iniciando o servidor Flask...\n")
    app.run(host='0.0.0.0', port=6969, debug=True)