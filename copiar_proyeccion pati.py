__author__ = 'Tato'

arreglo = CellRange('LLENADO',(6, 8), (10000, 66)).table

final = []

for x in range(0,len(arreglo)):
    if arreglo[x][5] == 'TONS'  and \
    arreglo[x][4] != 'DEMANDA' and \
    arreglo[x][4] != None:

        for j in range(0,53):

            final.append(arreglo[x][0:5] + [j, arreglo[x][6 + j],arreglo[x+1][6 + j], arreglo[x+2][6 + j],arreglo[x+3][6 + j]])

Cell('copia',2,1).table_range.clear()
Cell('copia',2,1).table = final

raw_input('fin')
