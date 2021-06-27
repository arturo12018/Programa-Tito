from  clases  import  *
from tkinter import filedialog
from tkinter import messagebox as MessageBox


resultado= "yes"
while(resultado == "yes"):
    ficheroListaTeams=filedialog.askopenfilename(title="Abrir",initialdir="D:\Documentos",filetypes=(("Ficheros de Excel (*.xlsx)","*.xlsx"),("Ficheros de Texto (*.txt)","*.txt")))
    print(ficheroListaTeams)

    listaModificar="D:\Documentos\ProgramaTito\dosdosTeamsC.xlsx"
    listaAsistencia=ArchivosAsis(str(ficheroListaTeams))
    listaOficial=ArchivoListaOf(listaModificar)

    try:
        listaAs=listaClase(listaAsistencia)
        listaAs.sort()#pone en oredn alfabetico
        print(listaAs)
        print(len(listaAs))
        listaOf=listaClase(listaOficial)
        
            

        listaTemp,listaOf=listaOf,[]
        for i in listaTemp:
            if type(i)==float or i=='Arturo Muñoz Rodríguez':
                continue
            else:
                listaOf.append(i)    
        for i in range(3):
            del  listaOf[0]
        print(listaOf) 
        print(len(listaOf))
        

        listaExel=ListaAsteriscosUnos(listaAs,listaOf)
        listaParaExel=listaExel.crealista()
        print(listaParaExel)

        

        if listaParaExel.count(1)==len(listaOf):
        
            
            MessageBox.showwarning("Programa", "Lista de alumno equivocada") # título, mensaje
            
        
        else:
            listaFinal=Exportalista(listaParaExel,str(ficheroListaTeams),listaModificar)
            listaFinal.exporta()
            MessageBox.showinfo("Programa", "Finalizado con exito") # título, mensaje
            resultado = MessageBox.askquestion("Salir", "¿Desea cargar otro archivo?")

            
                
            
        

        

    except:
        MessageBox.showwarning("Programa", "Archivo no encontrado y/o error de archivo") # título, mensaje
        resultado = MessageBox.askquestion("Salir", "¿Quiere cargar otro archivo?")

       
        




