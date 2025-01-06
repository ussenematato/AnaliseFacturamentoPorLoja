# AnaliseFacturamentoPorLoja
 Analise de facturamento por lojas com envio automatico de relatórios

# Bibliotecas
1. Instalar o pandas

O pandas é uma biblioteca popular para manipulação de dados. Para instalá-la, use o seguinte comando no terminal ou prompt de comando:

pip install pandas
Após a instalação, você pode importá-la normalmente no código:
import pandas as pd

2. Instalar o pywin32

A biblioteca win32com.client faz parte do pacote pywin32, que permite interagir com aplicações do Windows, como Excel, Word, etc.
Para instalá-lo, use o seguinte comando:

pip install pywin32

Depois, pode importar o win32com.client no código assim:
import win32com.client as win32

# Enviar um email com o relatório
email.To = "destinatario@example.com"  # E-mail do destinatário
