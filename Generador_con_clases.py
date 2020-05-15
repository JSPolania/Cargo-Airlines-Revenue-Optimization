__author__ = '3288715'

import revenue_tools as rm
from time import time
from itertools import groupby

tti = time()

excel_dir = wkbk_dir()

ini = Cell('Parametros',3,3).value
fin  = Cell('Parametros',4,3).value + 1

print 'Lecttura de oferta'

oferta = {i[0]:{'vuelo':i[1],
                'origen':i[6],
                'destino':i[7],
                'material':i[8],
                'modalidad':i[9],
                } for i in Cell('Dicc Rutas', 2, 1).table}

frecuencias = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Frecuencias',2 ,1 ).table}
costos = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Costos',2,1).table}

print 'lectura de demanda'

od = {i[0]:{'commodity':i[1],
            'origen':i[2],
            'destino':i[3]
            } for i in Cell('Demanda',2,1).table}

demanda = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Demanda',2,1).table}
escalas = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Escalas',2,1).table}
ex_postas = {i[0]:{j:rm.Apertura(i[j+3]) for j in range(ini,fin)} for i in Cell('Ex Postas',2,1).table}
ex_vlos = {i[0]:{j:rm.Apertura(i[j+3]) for j in range(ini,fin)} for i in Cell('Ex vuelos',2,1).table}
modalidad = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Modalidad',2,1).table}
segmentos = []
segtramo = []


'Lectura de segmentos anterior'

oldseg = {(i[0],i[1],i[2]):i[5] for i in Cell('Segmentos',2,1).table}

def rest_manual(wk,commodity, desc):
    if (wk,commodity,desc) in oldseg:
        return oldseg[(wk,commodity,desc)]
    else: return 'NA'

for wk in range(ini,fin):

    tiwk = time()

    faltantes = []

    print '******  Iniciando generacion de segmentos wk ' + str(wk)

    model = rm.Network(wk)
    for vlo in oferta.iterkeys():
        if frecuencias[vlo][wk] > 0:
            model.agregar_arco(oferta[vlo]['origen'],
                            oferta[vlo]['destino'],
                            vlo,
                            oferta[vlo]['material'],
                            oferta[vlo]['vuelo'],
                            oferta[vlo]['modalidad'],
                            costos[vlo][wk]) #oferta[vlo]['costo']

    model.hashmap_atos()
    ttl_paths = 0
    for c in (i for i in od.iterkeys() if demanda[i] > 0):
        pathcount = 0
        for path in model.segmentos(od[c]['origen'],od[c]['destino'], escalas[c][wk],ex_postas[c][wk],ex_vlos[c][wk],modalidad[c][wk] ):
            if len(path) > 0:
                pathcount +=1
                ttl_paths +=1
                segmentos.append([wk,c,'/'.join([i[2] for i in path]) ,'PATH' + str(pathcount).zfill(2), sum(i[3]['costo'] for i in path), rest_manual(wk,c,'/'.join([i[2] for i in path]))])
                legcount = 0
                for tramo in path:
                    legcount += 1
                    segtramo.append([wk,c,'/'.join([i[2] for i in path]), tramo[2],'LEG' + str(legcount).zfill(2)])
        if pathcount == 0:
            faltantes.append([c])

    tfwk = time()
    print str(ttl_paths) + ' segmentos generados , ' + str(len(faltantes)) + ' commodities faltantes, ' + '%.3f' % (tfwk - tiwk,) + ' segundos'
    for i in faltantes: print i

segmentos.sort(key=lambda x:(x[1],x[3],x[0]))
segtramo.sort(key=lambda x:(x[0],x[1],x[2],x[3]))

    #print model.nombre, model.red.number_of_edges(), model.red['SCL']['MAD']['SCL-MAD-Base']

'''_____________________________________________rutina para rescatar las restricciones manuales______________________'''


Cell('Segmentos',2,1).table_range.clear()
Cell('Segmentos',2,1).table = segmentos
Cell('Seg_tramo',2,1).table_range.clear()
Cell('Seg_tramo',2,1).table = segtramo

ttf = time()

print 'tiempo total de generacion de segmentos ' + '%.2f' % (ttf-tti,) + ' segundos'

raw_input('presione cualquier tecla para continuar')