__author__ = '3288715'

import networkx as nx

class Network:

    def __init__(self, nombre):
        self.nombre = nombre
        self.red = nx.MultiDiGraph()

    def agregar_arco(self, origen, destino, llave, material, nvlo, modo, cto):
        self.red.add_edge(origen, destino, key=llave, material=material, nvlo=nvlo, modalidad=modo, costo=cto)

    def hashmap_atos(self):
        self._atos = {i: {'pasadas':0} for i in self.red.nodes_iter()}

    def segmentos(self, source, target, max_escalas,ex_atos,ex_vuelos, Modo):

        for i in self._atos.keys():
            self._atos[i]['pasadas'] = 0 #reinicalizar los contadores de pasadas para no repetir aeropuerto
        self._atos[source]['pasadas'] +=1

        self._visited = [source]
        self._stack = [((u,v,key,data) for u,v,key,data in self.red.edges(source,data=True,keys=True))]   #anexar el iterador de los nodos a los cuales se llega desde el origen a la pila
        self._Arcos = []
        self._escalas = 0

        while self._stack:
            self._child = next(self._stack[-1], None)
            if self._child is None:
                self._stack.pop()
                self._poped = self._visited.pop()
                if self._Arcos:
                    self._arcpoped = self._Arcos.pop()
                    if self._arcpoped[3]['modalidad'] == 'Belly' or self._arcpoped[3]['modalidad'] == 'CAO':
                        self._escalas -= 1
                    self._atos[self._arcpoped[1]]['pasadas'] -=1

            else:

                if self._child[3]['modalidad'] == 'Belly' or self._child[3]['modalidad'] == 'CAO':
                    self._escalas += 1
                self._atos[self._child[1]]['pasadas'] +=1

                if ((self._escalas < max_escalas) or (self._escalas == max_escalas and self._child[1] == target)) and \
                        (self._atos[self._child[1]]['pasadas'] <=1) and \
                        (Modalidad(Modo,self._child[3]['modalidad']) == True) and \
                        (self._child[1] not in ex_atos) and \
                        (self._child[2] not in ex_vuelos):
                    if self._child[1] == target:
                        self._Arcos.append(self._child)
                        yield self._Arcos
                        self._arcpoped = self._Arcos.pop()
                        if self._arcpoped[3]['modalidad'] == 'Belly' or self._arcpoped[3]['modalidad'] == 'CAO':
                            self._escalas -= 1
                        self._atos[self._arcpoped[1]]['pasadas'] -=1

                    elif self._child[1] not in self._visited:
                        self._visited.append(self._child[1])
                        self._Arcos.append(self._child)
                        self._stack.append(((u,v,key, data) for u,v,key,data in self.red.edges(self._child[1],data=True, keys=True)))

                else:
                    if self._child[3]['modalidad'] == 'Belly' or self._child[3]['modalidad'] == 'CAO':
                        self._escalas -= 1
                    self._atos[self._child[1]]['pasadas'] -=1

def Apertura(x):
    return x.split(':')



def Modalidad(commodity, vuelo):
    if commodity == 'Belly' and vuelo == 'CAO':return False
    elif commodity == 'CAO' and vuelo == 'Belly':return False
    else: return True