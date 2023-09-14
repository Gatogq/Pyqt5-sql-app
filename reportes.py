import pandas as pd
import sys
import os
import warnings
from custom_functions import query_to_df, append_df_to_google_sheets 

warnings.simplefilter('ignore')
if getattr(sys,'frozen',False):
    b_dir = sys._MEIPASS
else:
    b_dir = os.path.dirname(os.path.abspath(__file__))

service_account_key = os.path.join(b_dir,'data','aut-1777-81fea8b3057a.json') 
sheet_key = '1CJFbMqtffrvkFHCdsCQ0bVTDB024M-EQEmAAFLRIjDM'
server = '192.168.21.163\IBDMX'
user = 'UsrAuditoria'
password = 'U$r@ud1t0r14MX'


def reporte_traspasos(month,year,sheet_name):
    
    query = f""" SELECT dtraspaso.sproducto AS Material, cproducto.sdescripcion AS Descripcion, dtraspaso.ncantidadtraspaso AS Cantidad,
             cproducto.sumedida as UnidadMedida, dtraspaso.spuntoventaorigen AS PuntoVentaOrigen, dtraspaso.spuntoventa AS PuntoVentaDestino,
             rprecio.ncosto AS CostoUnitario,(rprecio.ncosto * dtraspaso.ncantidadtraspaso) AS Importe, dtraspaso.dtraspaso AS Fecha,
             month(dtraspaso.dtraspaso) AS Mes
             FROM pvf_prodsc.dbo.dtraspaso
             INNER JOIN pvf_prodsc.dbo.cproducto
             ON dtraspaso.sproducto = cproducto.sproducto
             INNER JOIN pvf_prodsc.dbo.rprecio
             ON dtraspaso.sproducto = rprecio.sproducto
             WHERE
             rprecio.nzona=0
             AND
             (month(dtraspaso) = {month}
             AND
             year(dtraspaso)={year})
             AND
             dtraspaso.spuntoventa like 'F%'
             AND
             dtraspaso.spuntoventa NOT IN ('F022','F035','F060','F083','F086','F094','F096','F257',
                                           'F128','F141','F148','F163','F169','F173','F180','F266',
                                           'F186','F190','F196','F198','F224','F230','F235','F274','F279','F000')
             AND
             dtraspaso.spuntoventaorigen NOT IN ('F022','F035','F060','F083','F086','F094','F096','F257',
                                           'F128','F141','F148','F163','F169','F173','F180','F266',
                                           'F186','F190','F196','F198','F224','F230','F235','F274','F279','F000')
        """

    df = query_to_df(server,user,password,query)
    df['Fecha'] = df['Fecha'].astype(str)
    df['Index'] = range(1, len(df) + 1)
    append_df_to_google_sheets(df,service_account_key,sheet_key,sheet_name)

    return len(df), sheet_key

if __name__ == "__main__":
    pass
    
