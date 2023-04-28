from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import requests

url = 'https://teamcocomco.sharepoint.com/'
username = 'daniel.buitrago@teamco.com.co'
password = 'rockero22051067R'
client_id = '2c2dc263-7a85-49fe-a0e5-d7096a2c382d'
client_secret = 'tUF8Q~Axza4yjx4A6zgOK1_S0ptc3_j9uxun3c-a'
#resource = 'https://teamcocomco.sharepoint.com/sites/teamcocom'

#AUTENTICACION
ctx_auth = AuthenticationContext(url)
if ctx_auth.acquire_token_for_app(client_id, client_secret):
    # Crear el contexto del cliente de SharePoint
    ctx = ClientContext(url, ctx_auth)
    ctx.execute_query()
    print("Connected to SharePoint Online site: {0}".format(url))
else:
    print(ctx_auth.get_last_error())

# DESCARGA DE ARCHIVO
file_url = "https://teamcocomco.sharepoint.com/sites/teamcocom/Documentos compartidos/Libro1.xls"
response = requests.get(file_url)
with open('storage/descargado.xls', 'wb') as output_file:
    output_file.write(response.content)
    #myfile = (ctx.web.get_file_by_server_relative_url(file_url).download(output_file).execute_query())
print("File downloaded successfully.")

#OBTENER DATOS DE UNA LISTA
from office365.sharepoint.lists.list import List
# Creación del objeto List
list_url = 'https://teamcocomco.sharepoint.com/sites/teamcocom/Lists/viaticos/'
my_list = List.from_url(list_url, ctx)
items = my_list.get_items()
ctx.load(items)
ctx.execute_query()
for item in items:
    print(item.properties['Title'])



'''
# Obtener la lista de tareas
list_name = 'your-list-name'
list_obj = ctx.web.lists.get_by_title("viaticos")

items=list_obj.get_items()
ctx.load(items)
ctx.execute_query()
# Imprimir los elementos de la lista
for item in items:
    print(item.properties)


# Crear un nuevo elemento en la lista de tareas
item_properties = {'Title': 'New Task', 'Status': 'Not Started'}
list_item = list_obj.add_item(item_properties)
ctx.execute_query()
# Imprimir el ID del nuevo elemento
print('New item ID: {0}'.format(list_item.id))



#DESCARGA DE ARCHIVOS
with open("C:/Users/David Buitrago/Downloads/Test4/storage/", "wb") as local_file: 
    myfile = (context.web.get_file_by_server_relative_path(file_url).download(local_file).execute_query())
print("[Ok] file has been downloaded into: {0}".format(download_path))

    file_url = "https://teamcocomco.sharepoint.com/sites/teamcocom/Documentos compartidos/Libro1.xlsx"
    file = File.from_url(file_url)
    ctx.load(file)
    ctx.execute_query()
    response = File.open_binary(ctx, file.serverRelativeUrl)
    with open('descargado1.xlsx', 'wb') as output_file:
        output_file.write(response.content)

'''