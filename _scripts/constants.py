from _database import schemas
from tools import convert_to_base64
import os




empresas = {
        'SOLAR': schemas.Enteprise(
            id= '06.330.557/0001-77',
            formal_name='SOLAR AR CONDICIONADO LTDA.',
            informal_name= 'SOLAR AR CONDICIONADO',
            address= 'Av. Antonio Munhoz Bonilha',
            number= '543',
            complement= '545',
            district='Vila Palmeiras',
            city='São Paulo',
            county='SP',
            country='Brazil',
            postal_code='02.725-055',
            phone= '(11) 3951-5407',
            site= 'www.solarar.com.br',
            main_email= 'solar@solarar.com.br',
            purchase_email=schemas.Email(
                email='compras@solarar.com.br',
                senha='Solar@2022',
                assinatura=''),
            bill_receiver_email='solar.nf@solarar.com.br, solarar@nfe.omie.com.br',
            bill_sender_email=schemas.Email(
                email='faturamento@solarar.com.br',
                senha='Sol@r0201#2020',
                assinatura=''),
            financial_email=schemas.Email(
                email='financeiro@solarar.com.br',
                senha='Sol@rFin#2020',
                assinatura=''),
            additional_content = {
                'regime_tributario':'lucro real',
                'ie':'116.849.891.114',
                'im':'3.333.176-6',
            },
            logo="data:image/jpeg;base64,"+convert_to_base64('./img/LOGO SOLAR.png'),

        ),
        'CONTRUAR':schemas.Enteprise(
            id= '19.105.718/0001-70',
            formal_name='CONTRUAR AR CONDICIONADO LTDA.',
            informal_name= 'CONTRUAR AR CONDICIONADO',
            address= 'Rua Adão Ribeiro',
            number= '33',
            complement= '',
            district='Jardim Primavera',
            city='São Paulo',
            county='SP',
            country='Brazil',
            postal_code='02.755-070',
            phone= '(11) 3951-5407',
            site='',
            main_email='contruar@contruar.com.br',
            purchase_email=schemas.Email(
                email='compras@solarar.com.br',
                senha='Solar@2022',
                assinatura=''),
            bill_receiver_email='contruar.nf@contruar.com.br, contruarar@nfe.omie.com.br',
            bill_sender_email=schemas.Email(
                email='faturamento@solarar.com.br',
                senha='Sol@r0201#2020',
                assinatura=''),
            financial_email=schemas.Email(
                email='financeiro@solarar.com.br',
                senha='Sol@rFin#2020',
                assinatura=''),
            additional_content = {
                'regime_tributario':'simples nacional',
                'ie':'143.702.008.111',
                'im':'4.858.987-0',
            },
            logo="data:image/jpeg;base64,"+convert_to_base64('./img/LOGO CONTRUAR.png'),
            ),
        }

