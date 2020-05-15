__author__ = '3288715'

import networkx as nx

print 'Lectura de Oferta'

excel_dir = wkbk_dir()

oferta = {i[0]:{'region':i[1],
                'vuelo':i[2],
                'empresa':i[3],
                'ruta':i[4],
                'sentido':i[5],
                'origen':i[6],
                'destino':i[7],
                'material':i[8],
                'modalidad':i[9],
                'costo':i[9]
                } for i in Cell('Dicc Rutas', 2, 1).table}

frecuencias = {i[0]:{j:i[j] for j in range(1,52)} for i in Cell('Frecuencias',2 ,1 ).table}
disponibles = {i[0]:{j:i[j] for j in range(1,52)} for i in Cell('Disponibles',2 ,1 ).table}

modelo = {i:{} for i in range(1,52)}

for wk in modelo.iterkeys():
    modelo[wk]['red'] = nx.MultiDiGraph(data=((oferta[i]['origen'],
                                               oferta[i]['destino'],
                                                {
                                                 'llave':i,
                                                 'material':oferta[i]['material']
                                                }) for i in frecuencias if frecuencias[i][wk] > 0))


for i,j in modelo.iteritems():
    print i, j['red'].number_of_edges()

raw_input('presione cualquier tecla para continuar')


