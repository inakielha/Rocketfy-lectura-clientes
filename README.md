# Analizador de Compras Canceladas o Rechazadas -Rocketfy

## Descripción
Este programa se encarga de analizar una base de datos MongoDB que contiene más de 630,000 registros de clientes que realizaron compras. El objetivo es identificar usuarios que hayan realizado 3 o más compras rechazadas o canceladas en un mismo mes.

## Funcionalidades Clave
- Lee más de 630,000 registros de una base de datos MongoDB en unos minutos.
- Filtra usuarios que han realizado compras canceladas o rechazadas.
- Identifica aquellos usuarios que superan la cantidad esperada de compras rechazadas o canceladas en un mes.
- Genera un informe en formato Excel con los usuarios identificados.
- Envía el informe por correo electrónico a los interesados.

## Requisitos
Antes de ejecutar el programa, asegúrate de:
- Tener acceso a la base de datos MongoDB.
- Configurar tus credenciales de MongoDB en el script.
- Habilitar la autenticación de dos factores (2FA) para el envío de correos electrónicos.

## Instalación
1. Clona este repositorio en tu máquina local.
   ```bash
   git clone https://github.com/inakielha/Rocketfy-lectura-clientes.git)https://github.com/inakielha/Rocketfy-lectura-clientes.git
   cd prueba tecnica
   python db.py
