import win32com.client
import os

PASTA_DESTINO = r"C:\temp"
NOME_PASTA = "ImpressCN"
NOME_STORE = "lucas.com.br - Outlook"

os.makedirs(PASTA_DESTINO, exist_ok=True)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def encontrar_pasta(pasta_pai, nome):
    for i in range(1, pasta_pai.Folders.Count + 1):
        pasta = pasta_pai.Folders.Item(i)

        if pasta.Name.lower().strip() == nome.lower().strip():
            return pasta

        encontrada = encontrar_pasta(pasta, nome)
        if encontrada:
            return encontrada

    return None

store_encontrado = None

for i in range(1, outlook.Stores.Count + 1):
    store = outlook.Stores.Item(i)
    print("Verificando store:", store.DisplayName)

    if store.DisplayName.lower().strip() == NOME_STORE.lower().strip():
        store_encontrado = store
        break

if not store_encontrado:
    raise Exception("Store não encontrado.")

print("Store encontrado:", store_encontrado.DisplayName)

root = store_encontrado.GetRootFolder()
pasta_ImpressCN = encontrar_pasta(root, NOME_PASTA)

if not pasta_ImpressCN:
    raise Exception("Pasta ImpressCN NÃO encontrada dentro do store.")

print("Pasta encontrada:", pasta_ImpressCN.Name)

#Filtragem de emails não lidos
mensagens = pasta_ImpressCN.Items
mensagens = mensagens.Restrict("[Unread] = true")
mensagens.Sort("[ReceivedTime]", True)

print("Total de não lidos:", mensagens.Count)

lista_msgs = [mensagens.Item(i) for i in range(1, mensagens.Count + 1)]

for msg in lista_msgs:
    try:
        if msg.Attachments.Count > 0:
            for j in range(1, msg.Attachments.Count + 1):
                anexo = msg.Attachments.Item(j)

                if anexo.FileName.lower().endswith(".pdf"):
                    caminho = os.path.join(PASTA_DESTINO, anexo.FileName)

                    contador = 1
                    nome_original = anexo.FileName

                    while os.path.exists(caminho):
                        nome_sem_ext = os.path.splitext(nome_original)[0]
                        ext = os.path.splitext(nome_original)[1]
                        novo_nome = f"{nome_sem_ext}_{contador}{ext}"
                        caminho = os.path.join(PASTA_DESTINO, novo_nome)
                        contador += 1

                    anexo.SaveAsFile(caminho)
                    print("📄 PDF salvo:", caminho)

        msg.UnRead = False
        msg.Save()

    except Exception as e:
        print("Erro ao processar e-mail:", e)

print("Processo finalizado com sucesso.")