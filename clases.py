import pandas as pd
from openpyxl import load_workbook
import os.path,time

class Direccion():
    def __init__(self,direccion):
        self.direccion=direccion


class ArchivosAsis(Direccion):
    def hacelista(self):
        
            
            #Devuelve los datos de las columnas que pedimos 
            archivoexel = pd.read_excel(self.direccion)
            columnas=[str(archivoexel.columns[0])]
            listaPanda=archivoexel[columnas]
            listaPanda=listaPanda[str(archivoexel.columns[0])].tolist()#pasa de data frame a una lista                   
            print (listaPanda)
            listaPanda.remove('Arturo Muñoz Rodríguez')
            

            listaAsistencia=[]
            for x in listaPanda:
                if x not in listaAsistencia:
                    listaAsistencia.append(x)
                    
            return listaAsistencia




class ArchivoListaOf(Direccion):
    def hacelista(self):
        
            #Devuelve los datos de las columnas que pedimos 
            archivoexel = pd.read_excel(self.direccion)
            columnas=[str(archivoexel.columns[2])]
            listaPanda=archivoexel[columnas]
            listaPanda=listaPanda[str(archivoexel.columns[2])].tolist()#pasa de data frame a una lista                   
            return listaPanda
    


class  Exportalista(Direccion):
    def __init__(self,lista,fechaDirectorioIn,direccion):
        super().__init__(direccion) 
        self.lista=lista
        self.fechaDirectorioIn=fechaDirectorioIn

    def exporta(self):
        try:
            archivoexel = pd.read_excel(self.direccion)
            print(time.ctime(os.path.getctime(self.fechaDirectorioIn)))

            for i in range(len(archivoexel.columns)):
                print(i)
                if i==0:
                    continue
                elif ((type(archivoexel.iloc[5,i])==int or type(archivoexel.iloc[5,i])==str)==False ):
                    for x in range(len(self.lista)):
                        archivoexel.iloc[(x+4),i]=self.lista[x]
                    archivoexel.iloc[2,i]=time.ctime(os.path.getmtime(self.fechaDirectorioIn))
                    print(archivoexel)
                    archivoexel.to_excel(self.direccion,index=False)
                    break
                    
                
            
        except:
            MessageBox.showwarning("Programa", "Error Reintentar") # título, mensaje






class ListaAsteriscosUnos():
    def __init__(self,listaDia,listaOf):
        self.listaDia=listaDia
        self.listaOf=listaOf

    def crealista(self):
        listaFinal=[1 for i in range(len(self.listaOf))]
        
        for i in self.listaOf:
            if i in self.listaDia:
                indice=self.listaOf.index(i)
                listaFinal[indice]='*'
        return listaFinal
       
        
class Direccion():
    def __init__(self,direccion):
        self.direccion=direccion

    


              



def listaClase(lista):
    return lista.hacelista()


        