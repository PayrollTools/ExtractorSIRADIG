# -*- coding: utf-8 -*-
"""
Created on Sat Apr 17 11:12:08 2021

@author: Jorge E. Arrieta
"""

from xml.dom import minidom
import csv

def ex_cuil(archivo):
    xmlDoc = minidom.parse(archivo)
    presentacion_ = xmlDoc.getElementsByTagName('presentacion')
    
    dat_p=[]
            
    for nodo_hijo in presentacion_:
        
        dat_p.clear()
        
        periodo=''
        cuit=''
              
        if nodo_hijo.getElementsByTagName('periodo'):
           periodo = nodo_hijo.getElementsByTagName('periodo')
           
        if nodo_hijo.getElementsByTagName('empleado'):
            
            for mini_hijo in nodo_hijo.getElementsByTagName('empleado'):
                
                if mini_hijo.getElementsByTagName('cuit'):
                   cuit = mini_hijo.getElementsByTagName('cuit')
                
        if periodo!="":
            dat_p.append(periodo[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if cuit!="":
            dat_p.append(cuit[0].childNodes[0].data)
        else:
            dat_p.append(" ") 

    return dat_p

def ex_presentacion(archivo,file_):
    xmlDoc = minidom.parse(archivo)
    presentacion_ = xmlDoc.getElementsByTagName('presentacion')
    
    dat_p=[]
    
    cvs_f = open(file_,'a',newline="")
    doc_csv_w= csv.writer(cvs_f)
            
    for nodo_hijo in presentacion_:
        
        dat_p.clear()
        
        periodo=''
        nroPresentacion=''
        fechaPresentacion=''
        cuit=''
        tipoDoc=''
        apellido=''
        nombre=''
        provincia=''
        cp=''
        localidad=''
        calle=''
        nro=''
        piso=''
              
        if nodo_hijo.getElementsByTagName('periodo'):
           periodo = nodo_hijo.getElementsByTagName('periodo')
           
           
        if nodo_hijo.getElementsByTagName('nroPresentacion'):
           nroPresentacion = nodo_hijo.getElementsByTagName('nroPresentacion')
           
        
        if nodo_hijo.getElementsByTagName('fechaPresentacion'):
           fechaPresentacion = nodo_hijo.getElementsByTagName('fechaPresentacion')
           
           
        if nodo_hijo.getElementsByTagName('empleado'):
            
            for mini_hijo in nodo_hijo.getElementsByTagName('empleado'):
                
                if mini_hijo.getElementsByTagName('cuit'):
                   cuit = mini_hijo.getElementsByTagName('cuit')
                   
                if mini_hijo.getElementsByTagName('tipoDoc'):
                   tipoDoc = mini_hijo.getElementsByTagName('tipoDoc')
                   
                if mini_hijo.getElementsByTagName('apellido'):
                   apellido = mini_hijo.getElementsByTagName('apellido')
                   
                if mini_hijo.getElementsByTagName('nombre'):
                   nombre = mini_hijo.getElementsByTagName('nombre')
                   
                if mini_hijo.getElementsByTagName('direccion'):
                    
                   for _hijo in mini_hijo.getElementsByTagName('direccion'):
                       
                       if _hijo.getElementsByTagName('provincia'):
                          provincia = _hijo.getElementsByTagName('provincia')
                          
                       if _hijo.getElementsByTagName('cp'):
                          cp = _hijo.getElementsByTagName('cp')
                          
                       if _hijo.getElementsByTagName('localidad'):
                          localidad = _hijo.getElementsByTagName('localidad')
                
                       if _hijo.getElementsByTagName('calle'):
                          calle = _hijo.getElementsByTagName('calle')
                          
                       if _hijo.getElementsByTagName('nro'):
                          nro = _hijo.getElementsByTagName('nro')
                          
                       if _hijo.getElementsByTagName('piso'):
                          piso = _hijo.getElementsByTagName('piso')
                
        if len(str(periodo)) > 1:
            dat_p.append(periodo[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if len(str(nroPresentacion)) > 1:
            dat_p.append(nroPresentacion[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if len(str(fechaPresentacion)) > 1:
            dat_p.append(fechaPresentacion[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if len(str(cuit)) > 1:
            dat_p.append(cuit[0].childNodes[0].data)
        else:
            dat_p.append(" ") 
            
        if len(str(tipoDoc)) > 1:
            dat_p.append(tipoDoc[0].childNodes[0].data)
        else:
            dat_p.append(" ")  
            
        if len(str(apellido)) > 1:
            dat_p.append(apellido[0].childNodes[0].data)
        else:
            dat_p.append(" ")    
            
        try:
            if len(str(nombre)) > 1:
                dat_p.append(nombre[0].childNodes[0].data)
        except:
            dat_p.append(" ")  
            
        if len(str(provincia)) > 1:
            dat_p.append(provincia[0].childNodes[0].data)
        else:
            dat_p.append(" ")   
            
        if len(str(cp)) > 1:
            dat_p.append(cp[0].childNodes[0].data)
        else:
            dat_p.append(" ") 
            
        if len(str(localidad)) > 1:
            dat_p.append(localidad[0].childNodes[0].data)
        else:
            dat_p.append(" ")    
            
        if len(str(calle)) > 1:
            dat_p.append(calle[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if len(str(nro)) > 1:
            dat_p.append(nro[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        if piso!="":
            dat_p.append(piso[0].childNodes[0].data)
        else:
            dat_p.append(" ")
            
        doc_csv_w.writerow(dat_p)
    cvs_f.close()    
            

def ex_deduccion(archivo,file_):
    
    xmlDoc = minidom.parse(archivo)
    deduc_ = xmlDoc.getElementsByTagName('deduccion')
    
    dat_=[]
    
    cvs_f = open(file_,'a',newline="")
    doc_csv_w= csv.writer(cvs_f)
    
    for nodo_hijo in deduc_:
        
        dat_.clear()
        
        deduccion=''
        tipoDoc=''
        nroDoc=''
        denominacion=''
        descBasica=''
        descAdicional=''
        montoTotal=''
                                        
        dat_.extend(ex_cuil(archivo))
                                        
        if nodo_hijo.getAttribute('tipo'):
            deduccion = nodo_hijo.getAttribute('tipo')
                                
        if nodo_hijo.getElementsByTagName('tipoDoc'):
            tipoDoc = nodo_hijo.getElementsByTagName('tipoDoc')
                        
        if nodo_hijo.getElementsByTagName('nroDoc'):
            nroDoc = nodo_hijo.getElementsByTagName('nroDoc')
                        
        if nodo_hijo.getElementsByTagName('denominacion'):
            denominacion = nodo_hijo.getElementsByTagName('denominacion')
            
        if nodo_hijo.getElementsByTagName('descBasica'):
            descBasica = nodo_hijo.getElementsByTagName('descBasica')
                        
        if nodo_hijo.getElementsByTagName('descAdicional'):
            descAdicional = nodo_hijo.getElementsByTagName('descAdicional')
                                   
        if nodo_hijo.getElementsByTagName('montoTotal'):
            montoTotal = nodo_hijo.getElementsByTagName('montoTotal')
            
        if deduccion!="":
            dat_.append(deduccion)
        else:
            dat_.append(" ")
            
        if tipoDoc!= "":
            dat_.append(tipoDoc[0].childNodes[0].data)
        else:
            dat_.append(" ")
        
        if len(str(nroDoc)) > 1:
            dat_.append(nroDoc[0].childNodes[0].data)
        else:
            dat_.append(" ")
            
        if len(str(denominacion)) > 1:
            dat_.append(denominacion[0].childNodes[0].data)
        else:
            dat_.append(" ")
            
        if len(str(descBasica)) > 1:
            dat_.append(descBasica[0].childNodes[0].data)
        else:
            dat_.append(" ")
        
        if len(str(descAdicional)) > 1:
            dat_.append(descAdicional[0].childNodes[0].data)
        else:
            dat_.append(" ")
            
        if len(str(montoTotal)) > 1:
            dat_.append(montoTotal[0].childNodes[0].data)
        else:
            dat_.append(" ")
            
        doc_csv_w.writerow(dat_)
    cvs_f.close()
            
def ex_retPerPago (archivo,file_): #Verificado
    xmlDoc = minidom.parse (archivo)
    retPago_ = xmlDoc.getElementsByTagName('retPerPago')
    
    dat_=[]
    
    cvs_f = open(file_,'a',newline='')
    doc_csv_w= csv.writer(cvs_f)
    
    for nodo_hijo in retPago_:
        
        dat_.clear()
        
        retPerPago=''
        tipoDoc=''
        nroDoc=''
        denominacion=''
        descBasica=''
        descAdicional=''
        montoTotal='' 
        
        dat_.extend(ex_cuil(archivo))
        
        if nodo_hijo.getAttribute('tipo'):
            retPerPago = nodo_hijo.getAttribute('tipo')
            
        if nodo_hijo.getElementsByTagName('tipoDoc'):
            tipoDoc = nodo_hijo.getElementsByTagName('tipoDoc')
            
        if nodo_hijo.getElementsByTagName('nroDoc'):
            nroDoc = nodo_hijo.getElementsByTagName('nroDoc')
            
        if nodo_hijo.getElementsByTagName('denominacion'):
            denominacion = nodo_hijo.getElementsByTagName('denominacion')
            
        if nodo_hijo.getElementsByTagName('descBasica'):
            descBasica = nodo_hijo.getElementsByTagName('descBasica')
            
        if nodo_hijo.getElementsByTagName('descAdicional'):
            descAdicional = nodo_hijo.getElementsByTagName('descAdicional')
            
        if nodo_hijo.getElementsByTagName('montoTotal'):
            montoTotal = nodo_hijo.getElementsByTagName('montoTotal')
            
        if retPerPago!= '':
            dat_.append(retPerPago)
        else:
            dat_.append(' ')
            
        if tipoDoc != '':
            dat_.append(tipoDoc[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        if len(str(nroDoc)) > 1:
            dat_.append(nroDoc[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        if len(str(denominacion)) > 1:
            dat_.append(denominacion[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        if len(str(descBasica)) > 1:
            dat_.append(descBasica[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        if len(str(descAdicional)) > 1:
            dat_.append(descAdicional[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if len(str(montoTotal)) > 1:
            dat_.append(montoTotal[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        doc_csv_w.writerow(dat_)
    
    cvs_f.close()
            
def ex_ganLiqOtrosEmpEnt(archivo,file_): #Verificado
    
    xmlDoc = minidom.parse(archivo)
    empEnt_ = xmlDoc.getElementsByTagName('empEnt')
    
    cvs_f = open(file_,'a',newline='')
    doc_csv_w= csv.writer(cvs_f)
    
    dat_=[]
    dat_p=[] 
    
    for nodo_hijo in empEnt_:
        
        dat_.clear()
        
        dat_.extend(ex_cuil(archivo))
            
        cuit=''
        denominacion=''
        mes=''
        obraSoc=''
        segSoc=''
        sind=''
        ganBrut=''
        retGan=''
        retribNoHab=''
        ajuste=''
        exeNoAlc=''
        sac=''
        horasExtGr=''
        horasExtEx=''
        matDid=''
        gastosMovViat=''
    
            
        if nodo_hijo.getElementsByTagName('cuit'):
            cuit = nodo_hijo.getElementsByTagName('cuit')
                
        if nodo_hijo.getElementsByTagName('denominacion'):
            denominacion = nodo_hijo.getElementsByTagName('denominacion')
            
        if len(str(cuit)) > 1:
            dat_.append(cuit[0].childNodes[0].data)
        else:
            dat_.append(' ')
                
        if len(str(denominacion)) > 1:
            dat_.append(denominacion[0].childNodes[0].data)
        else:
            dat_.append(' ')    
                
        for mini_hijo in nodo_hijo.getElementsByTagName('ingAp'):
                
            if mini_hijo.getAttribute('mes'):
                mes = mini_hijo.getAttribute('mes')
                                   
            if mini_hijo.getElementsByTagName('obraSoc'):
                obraSoc = mini_hijo.getElementsByTagName('obraSoc')
                       
            if mini_hijo.getElementsByTagName('segSoc'):
                segSoc = mini_hijo.getElementsByTagName('segSoc')
                       
            if mini_hijo.getElementsByTagName('sind'):
                sind = mini_hijo.getElementsByTagName('sind')   
                       
            if mini_hijo.getElementsByTagName('ganBrut'):
                ganBrut = mini_hijo.getElementsByTagName('ganBrut')
                       
            if mini_hijo.getElementsByTagName('retGan'):
                retGan = mini_hijo.getElementsByTagName('retGan')  
                       
            if mini_hijo.getElementsByTagName('retribNoHab'):
                retribNoHab = mini_hijo.getElementsByTagName('retribNoHab')   
                    
            if mini_hijo.getElementsByTagName('ajuste'):
                ajuste = mini_hijo.getElementsByTagName('ajuste')   
                       
            if mini_hijo.getElementsByTagName('exeNoAlc'):
                exeNoAlc = mini_hijo.getElementsByTagName('exeNoAlc')  
                       
            if mini_hijo.getElementsByTagName('sac'):
                sac = mini_hijo.getElementsByTagName('sac')
                       
            if mini_hijo.getElementsByTagName('horasExtGr'):
                horasExtGr = mini_hijo.getElementsByTagName('horasExtGr')
                       
            if mini_hijo.getElementsByTagName('horasExtEx'):
                horasExtEx = mini_hijo.getElementsByTagName('horasExtEx') 
                       
            if mini_hijo.getElementsByTagName('matDid'):
                matDid = mini_hijo.getElementsByTagName('matDid')
                       
            if mini_hijo.getElementsByTagName('gastosMovViat'):
                gastosMovViat = mini_hijo.getElementsByTagName('gastosMovViat')   
            
            dat_p.extend(dat_)
            
            if mes != '':
                dat_p.append(mes)
            else:
                dat_p.append()
                    
            if len(str(obraSoc)) > 1:
                dat_p.append(obraSoc[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(segSoc)) > 1:
                dat_p.append(segSoc[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(sind)) > 1:
                dat_p.append(sind[0].childNodes[0].data) 
            else:
                dat_p.append(' ')  
                    
            if len(str(ganBrut)) > 1:
                dat_p.append(ganBrut[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(retGan)) > 1:
                dat_p.append(retGan[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(retribNoHab)) > 1:
                dat_p.append(retribNoHab[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(ajuste)) > 1:
                dat_p.append(ajuste[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(exeNoAlc)) > 1:
                dat_p.append(exeNoAlc[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(sac)) > 1:
                dat_p.append(sac[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(horasExtGr)) > 1:
                dat_p.append(horasExtGr[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(horasExtEx)) > 1:
                dat_p.append(horasExtEx[0].childNodes[0].data) 
            else:
                dat_p.append(' ') 
                    
            if len(str(matDid)) > 1:
                dat_p.append(matDid[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                    
            if len(str(gastosMovViat)) > 1:
                dat_p.append(gastosMovViat[0].childNodes[0].data) 
            else:
                dat_p.append(' ')
                
            doc_csv_w.writerow(dat_p)
            
            dat_p.clear()
            
    cvs_f.close()
                  
def ex_cargaFamilia(archivo,file_):#Verificado
    xmlDoc = minidom.parse(archivo)
    cargaFamilia_ = xmlDoc.getElementsByTagName('cargaFamilia')
    
    dat_=[]
    
    cvs_f = open(file_,'a', newline='')
    doc_csv_w= csv.writer(cvs_f)
        
    for nodo_hijo in cargaFamilia_: 
        
        dat_.clear()
        
        tipoDoc=''
        nroDoc=''
        apellido=''
        nombre=''
        fechaNac=''
        mesDesde=''
        mesHasta=''
        parentesco=''
        vigenteProximosPeriodos=''
        fechaLimite=''
        porcentajeDeduccion=''
        
        dat_.extend(ex_cuil(archivo))
        
        if nodo_hijo.getElementsByTagName('tipoDoc'):
           tipoDoc = nodo_hijo.getElementsByTagName('tipoDoc')
           
        if nodo_hijo.getElementsByTagName('nroDoc'):
           nroDoc = nodo_hijo.getElementsByTagName('nroDoc')
           
        if nodo_hijo.getElementsByTagName('apellido'):
           apellido = nodo_hijo.getElementsByTagName('apellido')
           
        if nodo_hijo.getElementsByTagName('nombre'):
           nombre = nodo_hijo.getElementsByTagName('nombre')
           
        if nodo_hijo.getElementsByTagName('fechaNac'):
           fechaNac = nodo_hijo.getElementsByTagName('fechaNac')
           
        if nodo_hijo.getElementsByTagName('mesDesde'):
           mesDesde = nodo_hijo.getElementsByTagName('mesDesde')
           
        if nodo_hijo.getElementsByTagName('mesHasta'):
           mesHasta = nodo_hijo.getElementsByTagName('mesHasta')
           
        if nodo_hijo.getElementsByTagName('parentesco'):
           parentesco = nodo_hijo.getElementsByTagName('parentesco')
           
        if nodo_hijo.getElementsByTagName('vigenteProximosPeriodos'):
           vigenteProximosPeriodos = nodo_hijo.getElementsByTagName('vigenteProximosPeriodos')
           
        if nodo_hijo.getElementsByTagName('fechaLimite'):
           fechaLimite = nodo_hijo.getElementsByTagName('fechaLimite')   
           
        if nodo_hijo.getElementsByTagName('porcentajeDeduccion'):
           porcentajeDeduccion = nodo_hijo.getElementsByTagName('porcentajeDeduccion')
           
        if len(str(tipoDoc)) > 1:
            dat_.append(tipoDoc[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if len(str(nroDoc)) > 1:
            dat_.append(nroDoc[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if len(str(apellido)) > 1:
            dat_.append(apellido[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        try:
            if len(str(nombre)) > 1 :
                dat_.append(nombre[0].childNodes[0].data)
        except: 
            dat_.append(' ')
            
            
        if len(str(fechaNac)) > 1:
            dat_.append(fechaNac[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if mesDesde!='':
            dat_.append(mesDesde[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if mesHasta != '':
            dat_.append(mesHasta[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if parentesco!='':
            dat_.append(parentesco[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if vigenteProximosPeriodos != '':
            dat_.append(vigenteProximosPeriodos[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if len(str(fechaLimite)) > 1:
            dat_.append(fechaLimite[0].childNodes[0].data)
        else:
            dat_.append(' ')    
            
        if len(str(porcentajeDeduccion)) > 1:
            dat_.append(porcentajeDeduccion[0].childNodes[0].data)
        else:
            dat_.append(' ')
        
        doc_csv_w.writerow(dat_)
    cvs_f.close()
           
def ex_agenteRetencion(archivo, file_): #Verificado
    
    xmlDoc = minidom.parse(archivo)
    agenteRetencion_ = xmlDoc.getElementsByTagName('agenteRetencion')
    
    dat_=[]
    
    cvs_f = open(file_,'a',newline='')
    doc_csv_w= csv.writer(cvs_f)
    
    for nodo_hijo in agenteRetencion_:
        
        cuit=''
        denominacion=''
        
        dat_.clear()
        dat_.extend(ex_cuil(archivo))
        
        if nodo_hijo.getElementsByTagName('cuit'):
           cuit = nodo_hijo.getElementsByTagName('cuit')
           
           
        if nodo_hijo.getElementsByTagName('denominacion'):
           denominacion = nodo_hijo.getElementsByTagName('denominacion')
           
           
        if len(str(cuit)) > 1:
            dat_.append(cuit[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        if len(str(denominacion)) > 1:
            dat_.append(denominacion[0].childNodes[0].data)
        else:
            dat_.append(' ')
            
        doc_csv_w.writerow(dat_)
    cvs_f.close()
           
def ex_datosAdicionales(archivo, file_):
    xmlDoc = minidom.parse(archivo)
    datosAdicionales_ = xmlDoc.getElementsByTagName('datoAdicional')
    
    dat_=[]
    
    cvs_f = open(file_,'a',newline='')
    doc_csv_w= csv.writer(cvs_f)
    
    for nodo_hijo in datosAdicionales_:
        
        nombre=''
        mesDesde=''
        mesHasta=''
        valor=''
        
        dat_.clear()
        dat_.extend(ex_cuil(archivo))
                        
        if nodo_hijo.getAttribute('nombre'):
           nombre= nodo_hijo.getAttribute('nombre')
                    
        if nodo_hijo.getAttribute('mesDesde'):
           mesDesde= nodo_hijo.getAttribute('mesDesde')
           
        if nodo_hijo.getAttribute('mesHasta'):
           mesHasta= nodo_hijo.getAttribute('mesHasta')
           
        if nodo_hijo.getAttribute('valor'):  
           valor= nodo_hijo.getAttribute('valor')
           
        if len(str(nombre)) > 1:
            dat_.append(nombre)
        else:
            dat_.append(' ')
        
        if mesDesde != '':
           dat_.append(mesDesde)
        else:
           dat_.append(' ')
           
        if mesHasta != '':
            dat_.append(mesHasta)
        else:
            dat_.append(' ')
            
        if valor != '':
            dat_.append(valor)
        else:
            dat_.append(' ')
         
        doc_csv_w.writerow(dat_)
        
    cvs_f.close()