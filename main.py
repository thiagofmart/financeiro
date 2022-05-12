import pdfkit
from datetime import date
from . import schemas




class Compras():
    def gerar_pedido_compra(self, pedido: schemas.PedidoCompra):
        path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        options = {
            'margin-top': '0cm',
            'margin-right': '0cm',
            'margin-bottom': '0cm',
            'margin-left': '0cm',
            'encoding': "UTF-8",
            }

        config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
        pdfkit.from_file('./templates/pedido_SOLAR.html', "out3.pdf", configuration=config, options=options)
        pass
    def enviar_pedido_compra(self, pedido_id):
        pass
