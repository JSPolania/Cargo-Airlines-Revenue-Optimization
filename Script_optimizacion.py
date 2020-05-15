__author__ = 'Stark Industries'

import gurobipy as gr
from time import time

excel_dir = wkbk_dir()

tti = time()

tleci = time()

print 'lectura de oferta'

outpst = [] #lista con el output de producto_segmento_tramo
outtramo = [] #lista con el output para el modelo pruducto
outprod =  [] #list con el output del modelo producto
outseg = [] # lista para el output de modelo segmento

ini = Cell('Parametros',3,3).value
fin  = Cell('Parametros',4,3).value + 1
var_resultados = Cell('Parametros',2,3).value

outpst_reader = Cell('Parametros',7,3).value
outseg_reader = Cell('Parametros',8,3).value
outprod_reader = Cell('Parametros',9,3).value
outtramo_reader = Cell('Parametros',10,3).value
method_reader = Cell('Parametros',13,3).value
method_vol = Cell('Parametros',15,3).value



'''______________________________________________________Oferta____________________________________________________'''

oferta = {i[0]:{'vuelo':i[1],
                'origen':i[6],
                'destino':i[7],
                'material':i[8],
                'modalidad':i[9],
                'ruta':i[4],
                'distancia':float(i[11]),
                'region':i[2],
                'empresa':i[3],
                'sentido':i[5],
                } for i in Cell('Dicc Rutas', 2, 1).table}

frecuencias = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Frecuencias',2 ,1 ).table}
dispeso = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Disp. Peso',2,1).table}
disvol = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Disp. Vol',2,1).table}
costos = {i[0]:{j:i[j] for j in range(ini,fin)} for i in Cell('Costos',2,1).table}

'''_____________________________________________________demanda____________________________________________________'''

print 'lectura de demanda'
od = {i[0]:{'commodity':i[1],
            'origen':i[2],
            'destino':i[3]
            } for i in Cell('Demanda',2,1).table}

demanda = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Demanda',2,1).table}
densidad = {i[0]:{j:i[j+3] * 0.001 for j in range(ini,fin)} for i in Cell('Densidad',2,1).table}
allotment = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Allotment',2,1).table}
tarifa = {i[0]:{j:i[j+3] for j in range(ini,fin)} for i in Cell('Tarifa',2,1).table}


'''_______________________________________________________Segmentos________________________________________________'''

print 'lectura de segmentos'

iterseg, orderpath, cto, mrest  = gr.multidict({(i[0],i[1],i[2]):[i[3],i[4],i[5]] for i in Cell('Segmentos',2,1).table})
iterseg = gr.tuplelist(iterseg)
itertram = gr.tuplelist(Cell('Seg_Tramo',2,1).table)

oldopt = {} #arreglo para guardar los resutlados anteriores qu serviran como input para el siguiente modelo

tlecf = time()

'''______________________________________________________algoritmo de optimizacion_________________________________'''


topti = time()

for wk in range(ini,fin):

    topii = time()

    print '...iniciado optimizacion wk ' + str(wk)
    m = gr.Model(str(wk))
    m.params.OutputFlag = var_resultados
    flow = {}
    '''_______________________________________________seteo de variables____________________________________________'''

    for _, c, path in sorted(iterseg.select(wk,'*','*')):
        flow[c,path] = m.addVar(lb=0, obj=tarifa[c][wk] - cto[wk, c,path],vtype=gr.GRB.CONTINUOUS, name='flow_%s_%s' % (c,path))
    m.update()

    '''_______________________________________restricciones de demanda y allotment__________________________________'''

    rdemand = {}
    rallot =  {}
    for c in od.iterkeys():
        temporal = [[tc,path] for _, tc, path in iterseg.select(wk,c,'*')]
        if len(temporal) > 0:
            rdemand[c] = m.addConstr(gr.quicksum(flow[i[0],i[1]] for i in temporal) <= demanda[c][wk],name= 'DDa_%s' % c )
            rallot[c] = m.addConstr(gr.quicksum(flow[i[0],i[1]] for i in temporal) >= allotment[c][wk], name='Allot_%s' % c)
    m.update()

    '''___________________________________________Restricciones de peso y volumen__________________________________'''

    rpeso = {}
    rvol = {}
    for tramo in oferta.iterkeys():
        temporal = [[tc,path] for _, tc, path, leg, idleg in itertram.select(wk,'*','*',tramo)]
        if temporal > 0:
            rpeso[tramo] = m.addConstr(gr.quicksum(flow[i[0],i[1]] for i in temporal) <= dispeso[tramo][wk] * frecuencias[tramo][wk], name='WR_%s' % tramo)
            if method_vol == 1:
                rvol[tramo] = m.addConstr(gr.quicksum(flow[i[0],i[1]] / densidad[i[0]][wk]  for i in temporal) <= disvol[tramo][wk] * frecuencias[tramo][wk], name='VR_%s' % tramo)
    m.update()

    '''______________________________________________restricciones manuales_______________________________________'''

    rmanual = {}
    for _, c, path in iterseg.select(wk,'*','*'):
        if mrest[wk,c,path] != 'NA':
            rmanual[c,path] = m.addConstr(flow[c,path] == mrest[wk,c,path], name='MR_%s' % c + '_' + path)
    m.update()
    '''______________________________________________Seteo del solver_____________________________________________'''

    m.modelSense = gr.GRB.MAXIMIZE #inddicar al modelo que es de maximizacion
    m.params.method = method_reader
    #para cada variable, setear el el resultado anterior como una posible solucion del nuevo modelo

    for _, c, path in iterseg.select(wk,'*','*'):
        if (wk-1,c,path) in oldopt:
            flow[c,path].PStart = oldopt[wk-1,c,path]
    m.update()

    m.optimize()

    topif = time()
    '''______________________________________________escritura de info_____________________________________________'''

    if m.status == gr.GRB.status.OPTIMAL:

        for _, c, path in iterseg.select(wk,'*','*'):
            oldopt[(wk,c,path)] = flow[(c,path)].x

        print 'solucion encontrada: optimo=',int(m.objval),'variables=', m.NumVars, 'Metodo', m.params.method, 'tiempo:', '%.3f' % (topif - topii,),' seg'

        '''______________________________________________________funciones_______________________________________'''

        def primersegmento(peso, orden): #funcion para sacar el peso solo del primer tramo
            if orden == 'LEG01':
                return peso
            else:
                return 0

        def return_demand(llave, tipo):
            if llave in rdemand:
                if tipo == 'slack': return rdemand[llave].slack
                elif tipo == 'pi': return rdemand[llave].pi
                elif tipo == 'limitsup': return rdemand[llave].SARHSUp
                elif tipo == 'limitinf': return rdemand[llave].SARHSLow
                elif tipo == 'RHS': return rdemand[llave].RHS
            else: pass

        def return_allotment(llave, tipo):
            if llave in rallot:
                if tipo == 'slack': return rallot[llave].slack
                elif tipo == 'pi': return rallot[llave].pi
                elif tipo == 'limitsup': return rallot[llave].SARHSUp
                elif tipo == 'limitinf': return rallot[llave].SARHSLow
                elif tipo == 'RHS': return rallot[llave].RHS
            else: pass

        def return_rmanual(wk,commodity, IDpath, tipo):
            if mrest[wk,c,path] != 'NA':
                if tipo == 'restriccion': return mrest[wk,c,path]
                elif tipo == 'slack': return rmanual[c,path].slack
                elif tipo == 'pi': return rmanual[c,path].pi
                elif tipo == 'limitsup': return rmanual[c,path].SARHSUp
                elif tipo == 'limitinf': return rmanual[c,path].SARHSLow

        def return_peso(wk,leg,tipo):
            if leg in rpeso:
                if tipo == 'slack': return rpeso[leg].slack
                elif tipo == 'pi': return rpeso[leg].pi
                elif tipo == 'limitsup': return rpeso[leg].SARHSUp
                elif tipo == 'limitinf': return rpeso[leg].SARHSLow

        def return_vol(wk,leg,tipo):
            if method_vol == 1:
                if leg in rvol:
                    if tipo == 'slack': return rvol[leg].slack
                    elif tipo == 'pi': return rvol[leg].pi
                    elif tipo == 'limitsup': return rvol[leg].SARHSUp
                    elif tipo == 'limitinf': return rvol[leg].SARHSLow
            else: return 0

        def factor_red(twk, tcommodity, tpath, tramo):
            if oferta[tramo]['modalidad'] == 'Belly' or oferta[tramo]['modalidad'] == 'CAO':
                return oferta[tramo]['distancia'] / sum(oferta[leg]['distancia'] for _,_,_,leg,_ in itertram.select(twk, tcommodity, tpath,'*','*') if oferta[leg]['modalidad'] == 'Belly' or oferta[leg]['modalidad'] == 'CAO')
            else: return 0

        '''_________________________________________modelo producto-segmento-Tramo_______________________________'''

        if outpst_reader == 1: #if para saber si usuario desea output de PST
            for _, c, path, idleg, orden in itertram.select(wk,'*','*','*','*'):
                if flow[c,path].x > 0:
                    outpst.append([wk,                                                                                     #semana
                                c,                                                                                         #ID Commodity
                                path,                                                                                      #ID Path
                                idleg,                                                                                     #ID tramo
                                od[c]['commodity'],                                                                        #descipcion commodity
                                od[c]['origen'],                                                                           #origen guia
                                od[c]['destino'],                                                                          #destino guia
                                orden,                                                                                     #numero del tramo en el segmento
                                oferta[idleg]['vuelo'],                                                                    #numero de vuelo
                                oferta[idleg]['region'],                                                                   #Region
                                oferta[idleg]['empresa'],                                                                  #empresa
                                oferta[idleg]['origen'],                                                                   #origen tramo
                                oferta[idleg]['destino'],                                                                  #destino tramo
                                oferta[idleg]['ruta'],                                                                     #ruta
                                oferta[idleg]['sentido'],                                                                  #sentido
                                oferta[idleg]['material'],                                                                 #material
                                oferta[idleg]['modalidad'],                                                                #modalidad
                                flow[c,path].x,                                                                            #peso tramo
                                flow[c,path].x / densidad[c][wk],                                                          #volumen tramo
                                primersegmento(flow[c,path].x,orden),                                                      #peso segmento
                                primersegmento(flow[c,path].x,orden) / densidad[c][wk],                                    #peso segmento
                                factor_red(wk,c,path,idleg),                                                               #factor red
                                oferta[idleg]['distancia'],                                                                #distancia
                                flow[c,path].x * oferta[idleg]['distancia'],                                               #RTK
                                tarifa[c][wk] * flow[c,path].x * factor_red(wk,c,path,idleg),                              #ingreso tramo
                                tarifa[c][wk] * flow[c,path].x,                                                            #ingreso red
                                tarifa[c][wk] * primersegmento(flow[c,path].x, orden),                                     #ingreso segmento
                                costos[idleg][wk] * flow[c,path].x,                                                        #costo
                                ((tarifa[c][wk] * factor_red(wk,c,path,idleg)) - costos[idleg][wk]) * flow[c,path].x       #margen contribucion
                    ])

        '''___________________________________________Modelo Segmento______________________________________________'''

        if outseg_reader == 1: #if para saber si usuario desesa ouput de segmento
            for _,c,path in iterseg.select(wk,'*','*'):
                dist_total = sum(oferta[x]['distancia'] for _,_,_,x,_ in itertram.select(wk,c,path,'*','*'))
                #costo_total = sum(costos[x][wk] for _,_,_,x,_ in itertram.select(wk,c,path,'*','*'))
                outseg.append([wk,                                                                                         #week
                               c,                                                                                          #ID commodity
                               path,                                                                                       #ID PAth
                               orderpath[wk,c ,path],                                                                      #routing
                               od[c]['commodity'],                                                                         #descipcion commodity
                               od[c]['origen'],                                                                            #origen guia
                               od[c]['destino'],                                                                           #destino guia
                               cto[wk,c,path],                                                                                #Costo unitario
                               flow[c,path].x,                                                                             #peso segmento
                               flow[c,path].x /densidad[c][wk],                                                            #volumen segmento
                               flow[c,path].RC,                                                                            #costo reducido
                               flow[c,path].VBasis,                                                                        #Vbasis
                               flow[c,path].SAObjLow,                                                                      #limite inferior objetivo
                               flow[c,path].SAObjUp,                                                                       #limite superior objetivo
                               return_rmanual(wk,c,path,'restriccion'),                                                    #restriccion
                               return_rmanual(wk,c,path,'slack'),                                                          #restriccion slack
                               return_rmanual(wk,c,path,'pi'),                                                             #restriccion precio sombra
                               return_rmanual(wk,c,path,'limitinf'),                                                       #restriccion limit inf
                               return_rmanual(wk,c,path,'limitsup'),                                                       #restriccion limit sup
                               dist_total,                                                                                 #distancia total recorrida
                               dist_total * flow[c,path].x,                                                                #RTK
                               flow[c,path].x * tarifa[c][wk],                                                             #ingreso segmento
                               cto[wk,c,path] * flow[c,path].x,                                                               #costo total
                               (tarifa[c][wk] - cto[wk,c,path]) * flow[c,path].x,                                             #margen de contribucion
                               len([x for _,_,_,_,x in itertram.select(wk,c,path,'*','*')])                                #cantidad de tramos
                ])

        '''___________________________________________Modelo producto_____________________________________________'''

        if outprod_reader == 1: #if para saber si usuario desea output de producto
            for c in od.iterkeys():
                peso_total = sum(flow[i,j].x for _,i,j in iterseg.select(wk,c,'*'))
                costo_total = sum(costos[tt][wk] * flow[tc,tp].x for _,tc,tp,tt,_ in itertram.select(wk,c,'*','*','*') if flow[tc,tp].x > 0)
                outprod.append([wk,                                                                                        #Semana
                                c,                                                                                         #ID Commodity
                                od[c]['commodity'],                                                                        #commodity
                                od[c]['origen'],                                                                           #origen
                                od[c]['destino'],                                                                          #Destino
                                demanda[c][wk],                                                                            #demanda
                                return_demand(c,'slack'),                                                                  #cuanta demanda quedo insatisfecha
                                return_demand(c,'pi'),                                                                     #precio sombra
                                return_demand(c,'limitinf'),                                                               #limite superior de la restriccion
                                return_demand(c,'limitsup'),                                                               #limite ingerior de la restriccion
                                allotment[c][wk],                                                                          #allotment
                                return_allotment(c,'slack'),                                                               #cuanto alloemtn sobro
                                return_allotment(c,'pi'),                                                                  #precio sombra del alloment
                                return_allotment(c,'limitinf'),                                                            #limite superior del allotment
                                return_allotment(c,'limitsup'),                                                            #limite ingerior del allotment
                                densidad[c][wk],                                                                           #densidad
                                peso_total,                                                                                #peso producto
                                peso_total * densidad[c][wk],                                                              #volumen total
                                peso_total * tarifa[c][wk],                                                                #ingreso total
                                sum(oferta[tt]['distancia'] * flow[tc,tp].x for _,tc,tp,tt,_ in itertram.select(wk,c,'*','*','*') if flow[tc,tp].x > 0), #RTK
                                costo_total,                                                                               #costo
                                (peso_total * tarifa[c][wk]) - costo_total,                                                #MgC
                                demanda[c][wk] * tarifa[c][wk],                                                            #ingreso planif
                                peso_total * return_demand(c,'pi')
            ])

        '''_________________________________________Modelo Tramo_________________________________________________'''

        if outtramo_reader == 1: #if para saber si usuario desea output de tramo
            for tramo in oferta.iterkeys():
                peso_trans = sum(flow[tc,tp].x for _,tc,tp,_,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0)
                outtramo.append([wk,                                                                                       #week
                                 tramo,                                                                                    #ID tramo
                                 oferta[tramo]['vuelo'],                                                                   #numero vuelo
                                 oferta[tramo]['region'],                                                                  #region
                                 oferta[tramo]['empresa'],                                                                 #empresa
                                 oferta[tramo]['origen'],                                                                  #origen
                                 oferta[tramo]['destino'],                                                                 #destino
                                 oferta[tramo]['ruta'],                                                                    #ruta
                                 oferta[tramo]['sentido'],                                                                 #sentido
                                 oferta[tramo]['material'],                                                                #material
                                 oferta[tramo]['modalidad'],                                                               #modalidad
                                 sum(flow[tc,tp].x for _,tc,tp,_,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0), #Peso
                                 return_peso(wk,tramo,'slack'),                                                            #slack restriccion peso
                                 return_peso(wk,tramo,'pi'),                                                               #precio sombre restriccion peso
                                 return_peso(wk,tramo,'limitinf'),                                                         #limite inferior restriccion peso
                                 return_peso(wk,tramo,'limitsup'),                                                         #limite superior restriccion peso
                                 sum(flow[tc,tp].x / densidad[tc][wk] for _,tc,tp,_,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0),   #
                                 return_vol(wk,tramo,'slack'),                                                             #slack restriccion vol
                                 return_vol(wk,tramo,'pi'),                                                                #precio sombre restriccion vol
                                 return_vol(wk,tramo,'limitinf'),                                                          #limite inferior restriccion vol
                                 return_vol(wk,tramo,'limitsup'),                                                          #limite superior restriccion vol
                                 oferta[tramo]['distancia'],                                                               #distancia
                                 sum(flow[tc,tp].x * tarifa[tc][wk]  * factor_red(wk,tc,tp,tt) for _,tc,tp,tt,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0), #ingreso tramo
                                 sum(flow[tc,tp].x * tarifa[tc][wk] for _,tc,tp,_,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0), #ingreso segmento
                                 peso_trans * oferta[tramo]['distancia'],                                                  #RTK
                                 peso_trans * costos[tramo][wk],                                                           #Costo
                                 sum((tarifa[tc][wk]  * factor_red(wk,tc,tp,tt) - costos[tramo][wk] ) * flow[tc,tp].x  for _,tc,tp,tt,_ in itertram.select(wk,'*','*',tramo,'*') if flow[tc,tp].x > 0), #ingreso tramo
                                 peso_trans * return_peso(wk,tramo,'pi')                                                     #precio sombra peso ponderado
                                 ])


    elif m.status == gr.GRB.status.INF_OR_UNBD:
        print 'No se puede optimizar'
        m.computeIIS()
        print 'las siguientes restricciones no se pudieron satisfacer'
        for c in m.getConstrs():
            if c.IISConstr:
                print('%s' % c.constrName)

toptf = time()

trwi = time()

Cell('Resultados PST',2,1).table_range.clear()
Cell('Resultados PST',2,1).table = outpst

Cell('Resultados Prod',2,1).table_range.clear()
Cell('Resultados Prod',2,1).table = outprod

Cell('Resultados segmento',2,1).table_range.clear()
Cell('Resultados segmento',2,1).table = outseg

Cell('Resultados tramo',2,1).table_range.clear()
Cell('Resultados tramo',2,1).table = outtramo

trwf = time()

ttf = time()


print 'tiempo de lectura:', '%.3f' % (tlecf - tleci,),' seg'
print 'tiempo de optimizacion:', '%.3f' % (toptf - topti,),' seg'
print 'tiempo de escritura:', '%.3f' % (trwf - trwi,),' seg'
print 'tiempo total:', '%.3f' % (ttf - tti,),' seg'

raw_input('presione cualquier tecla para continuar')