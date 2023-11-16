import pymongo

from helper import borrar_archivo_excel, check_month_cancelled, crear_excel_con_formato, enviar_correo

# Si existe el archivo excel local lo borro 
borrar_archivo_excel("UsuariosRechazados.xlsx")

# Configura la conexión a la base de datos MongoDB
client = pymongo.MongoClient(f"mongodb+srv://inakielhaiek:Rocketfy123@cluster0.hnjrxcm.mongodb.net/")

# # Selecciona la base de datos
db = client["Rocketfy"]

# # Accede a una colección
collection = db["Clients"]
# result = collection.find({})
all_clients_id = collection.distinct("order_vendor_dbname")
with open ("archivos unicos.txt", "a") as unicos:
    unicos.write(str(all_clients_id) + "\n")

crear_excel_con_formato()
check_client_failled = False
for client_id in all_clients_id:
# for indice in range(5):
    # client_id = all_clients_id[indice]

    client_failed = collection.find(
        {"order_vendor_dbname": client_id,
        "shipping_status": {"$in": ["returned", "cancelled"]}
        })

    # Crea un diccionario para contar elementos por mes
    elementos_por_mes = {}
    # Itera a través de los client_failed y cuenta elementos por mes
    for resultado in client_failed:
        with open ("verArray.txt", "a") as archivo:
            archivo.write(str(resultado) + "\n")

        mes_envio = resultado["shipping_date"].month
        existe_mes = elementos_por_mes.get(mes_envio)
        if existe_mes is None:
            elementos_por_mes[mes_envio] = {
              "month": mes_envio,
              "cantidad" : 0,
              "shipping_list" : []  
            }

        elementos_por_mes[mes_envio]["cantidad"] += 1
        elementos_por_mes[mes_envio]["shipping_list"].append(resultado["shipping_id"])
    # Verifica si hay algún mes con 3 o más elementos
    client_status = check_month_cancelled(elementos_por_mes, resultado)
    # print(client_status)
    if client_status == True:
         check_client_failled = True
if check_client_failled == True:
        body_email = f"Estos son los usuarios que tuvieron 3 compras o mas rechazadas o canceladas en un mismo mes. Adjunto txt con los usuarios, cuantas compras realizaron, en que mes las realizaron y el id de cada compra. Saludos"
        # PONER EL EMAIL DEL DESTINATARIO
        enviar_correo("poner email aca","Notificacion", body_email, "UsuariosRechazados.xlsx")
else :
      body_email = f"No hemos detectado usuarion con 3 o mas compras canceladas o rechazadas en un mismo mes. Saludos"
    #   PONER EL EMAIL DEL DESTINATARIO
      enviar_correo("poner email aca","Notificacion", body_email, "no attachment")