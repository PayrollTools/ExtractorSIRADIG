# -*- coding: utf-8 -*-

"""
Created on Sat Apr 17 11:12:08 2021

@author: Jorge E. Arrieta
"""

import wx
import os
import xml_
import csv
import webbrowser

class vtn_acerca(wx.Frame):
    
    def __init__(self,parent,title):
        
        estilo=(wx.CLOSE_BOX|wx.CAPTION)

        wx.Frame.__init__(self,parent=parent,title=title,size=(270,200),style=estilo)
        
        icono = wx.Icon('Icono.ico')
        
        self.SetIcon(icono)
        
        wx.StaticText(self,label= "Extractor SIRADIG Versión 1.0.2",pos=(8,11))
        wx.StaticText(self,label="Actualizado a versión SIRADIG 1.14",pos=(8,35))
        wx.StaticText(self,label="Desarrollado por: Jorge Eduardo Arrieta",pos=(8,60))
        wx.StaticText(self,label="Año, 2022 - Buenos Aires, Argentina",pos=(8,84))
        
        btnCerrar = wx.Button(self,-1,label='Cerrar',pos=(148,116),size=(78,32))

        self.Bind(wx.EVT_CLOSE,self.onClose)
        self.Bind(wx.EVT_BUTTON,self.onClose,btnCerrar)
                        
    def onClose(self,event):
        frame.Enable(True)                
        self.Destroy()   


class principal (wx.Frame):

    def __init__(self,parent,title):

        estilo=(wx.CLOSE_BOX|wx.CAPTION)

        wx.Frame.__init__(self,parent=parent,title=title,size=(620,360),style=estilo)
        
        icono = wx.Icon('Icono.ico')
        
        self.SetIcon(icono)

        #Labels

        wx.StaticText(self,label= "Listado XML a procesar:",pos=(8,11))
        wx.StaticText(self,label="Ubicación archivo de salida:",pos=(8,149))
        wx.StaticText(self,label="Proceso a ejecutar:",pos=(8,212))

        #ListBox
               
        self.lboxDirecciones = wx.ListBox(self,-1,pos=(8,36),size=(581,83),style = wx.LB_SINGLE)
        #self.txtDirecciones = wx.TextCtrl(self,-1,pos=(8,36),size=(581,83),style=wx.TE_READONLY|wx.TE_MULTILINE)

        #Cajas de texto

        self.txtDireccion = wx.TextCtrl(self,-1,pos=(8,175),size=(493,28),style=wx.TE_READONLY)

        #Botones

        btnSeleccion = wx.Button(self,-1,label='Seleccionar',pos=(516,175),size=(78,28))
        btnAcerca = wx.Button(self,-1,label='Acerca de...',pos=(334,278),size=(78,28))
        btnCerrar = wx.Button(self,-1,label='Cerrar',pos=(423,278),size=(78,28))
        btnProcesar = wx.Button(self,-1,label='Procesar',pos=(511,278),size=(78,28))
        btnClear = wx.Button(self,-1,label='Borrar selección',pos=(334,125),size=(121   ,28))
        btnUbicacion = wx.Button(self,-1,label='Ubicación XML',pos=(468,125),size=(121,28))

        #ComboBox

        procesos =  ['Agente de retención',
                    'Cargas de Familia',
                    'Deducciones',
                    'Ganancias liquidadas por otros empleadores o entidades',
                    'Retenciones, percepciones y pagos a cuenta',
                    'Datos adicionales',
                    'Presentación']

        self.cbxSeleccion = wx.ComboBox(self,-1,choices=procesos,pos=(8,237),size=(581,28), style=wx.CB_READONLY)

        
        self.Bind(wx.EVT_BUTTON,self.OnProcesar,btnProcesar)
        self.Bind(wx.EVT_BUTTON,self.OnLimpiar,btnClear)
        self.Bind(wx.EVT_BUTTON,self.onSeleccionar,btnSeleccion)
        self.Bind(wx.EVT_BUTTON,self.onUbicacion,btnUbicacion)
        self.Bind(wx.EVT_BUTTON,self.onAcerca,btnAcerca)
        self.Bind(wx.EVT_CLOSE,self.onClickCerrar)
        self.Bind(wx.EVT_BUTTON,self.onClickCerrar,btnCerrar)
        
        
    def OnLimpiar(self,event):
        self.lboxDirecciones.Clear()
        
    def onAcerca(self,event):
        vtnAcerca = vtn_acerca(self, 'Información')
        vtnAcerca.SetBackgroundColour("lightgray")
        vtnAcerca.Center()
        self.Enable(False)
        vtnAcerca.Show()
        
                
    def onUbicacion (self,event):
        msg = 'Seleccionar directorio XML'
        dlgUbicacion = wx.DirDialog(self,msg)
        dlgUbicacion.ShowModal()
        
        directorio = dlgUbicacion.GetPath()
        
        if directorio != "":
            
            archivos = [f.path for f in os.scandir(directorio) if f.name.endswith('.xml')]
            
            x = 0
            
            for i in archivos:
                
                self.lboxDirecciones.Append(i)
                
                x= x+1
                                
            e = str(len(archivos))
                                   
            wx.MessageBox(f"Se encontraron {e} archivos")
            
        else:
            
            wx.MessageBox("No se selecciono un directorio")

    def onClickCerrar(self,event):
        
        msj = wx.MessageDialog(self,caption = 'Advertencia',message = '¿Esta seguro de querer cerrar el programa?',style=wx.YES_NO)
        resultado = msj.ShowModal()
        
        if resultado == wx.ID_YES:
            self.Destroy()

    def onSeleccionar(self,event):
        #Definición de variables para dlgBox
        msg='Seleccionar directorio de salida'
        dlgBox = wx.DirDialog(self,msg)
        dlgBox.ShowModal()
        obtdir=dlgBox.GetPath()
        if obtdir=="":
           wx.MessageBox("Falta seleccionar el directorio")
        else: 
           self.txtDireccion.SetValue(obtdir)

        dlgBox.Destroy()

    def verificar_datos(self):
        procesar = True
                
        if self.txtDireccion.GetValue()=='':
            wx.MessageBox('Falta especificar directorio de salida.')
            procesar=False
        
        if self.lboxDirecciones.GetCount()== 0:
            wx.MessageBox('No hay archivos XML a procesar')
            procesar=False

        if self.cbxSeleccion.GetValue()=='':
            wx.MessageBox('Falta seleccionar el proceso a ejecutar')
            procesar=False
        
        return procesar  
    
    def procesar_xml(self,op,file_):
        
        i=0
        
        if op == 1:
            
            #Agente de retención  OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'cuit','denominacion'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
        
        if op == 2:
            
            #Ajustes
            
            pass
        
        if op == 3:
            
            #Cargas de familia OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'tipoDoc','nroDoc','apellido','nombre','fechaNac','mesDesde',
                          'mesHasta','parentesco','vigenteProximosPeriodos','fechaLimite',
                          'porcentajeDeduccion'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
        
        if op == 4:
            
            #Deducciones 
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'deduccion','tipoDoc','nroDoc','denominacion','descBasica',
                          'descAdicional','montoTotal'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
                    
        if op == 5:
            
            #Ganancias otros empleadores OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'cuit','denominacion','mes','obraSoc','segSoc','sind',
                          'ganBrut','retGan','retribNoHab','ajuste','exeNoAlc',
                          'sac','horasExtGr','horasExtEx','matDid','gastosMovViat'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
        
        if op == 6:
            
            #Pagos a cuenta  OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'retPerPago','tipoDoc','nroDoc','denominacion','descBasica',
                          'descAdicional','montoTotal '] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
        
        if op == 7:
            
            #Datos adicionales OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'CUIL',
                          'nombre','mesDesde','mesHasta','valor'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()
            
        if op == 8:
            
            #Presentación OK
            
            cvs_f = open(file_,'a',newline="")
            doc_csv_w= csv.writer(cvs_f)
    
            encabezado = ['periodo', 'nroPresentacion', 'fechaPresentacion', 'cuit',
                          'tipoDoc','apellido', 'nombre','provincia','cp','localidad',
                          'calle','nro','piso'] 
            
            doc_csv_w.writerow(encabezado)
            cvs_f.close()    
       
        while i < self.lboxDirecciones.GetCount():
                
              try:
                  
                  if op == 1:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_agenteRetencion(self.lboxDirecciones.GetString(i),file_)
                
                  if op == 2:
                     pass
                
                  if op == 3:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_cargaFamilia(self.lboxDirecciones.GetString(i),file_)
                    
                  if op == 4:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_deduccion(self.lboxDirecciones.GetString(i),file_)                         
                                
                  if op == 5:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_ganLiqOtrosEmpEnt(self.lboxDirecciones.GetString(i),file_)
                    
                  if op == 6:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_retPerPago(self.lboxDirecciones.GetString(i),file_)
                    
                  if op == 7:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_datosAdicionales(self.lboxDirecciones.GetString(i),file_)
                     
                  if op == 8:
                      self.lboxDirecciones.SetSelection(i)
                      xml_.ex_presentacion(self.lboxDirecciones.GetString(i),file_)   
                    
                  i = i +1
                    
              except:
                  wx.MessageBox(f'Proceso interrumpido en archivo: {self.lboxDirecciones.GetSelection()}')
                  break
                
        wx.MessageBox(f'Se terminaron de procesar {i} archivos xml.')

    def OnProcesar(self,event):
        
        validez = self.verificar_datos()
        
        if validez == True:
            
            file_ = self.txtDireccion.GetLineText(0) + "\Deducciones.csv"
            
            if self.cbxSeleccion.GetValue()== 'Agente de retención':
                op_=1
                file_ = self.txtDireccion.GetLineText(0) + "\Agente de retencion.csv"
                
            if self.cbxSeleccion.GetValue()== 'Ajustes':
                op_=2
                file_ = self.txtDireccion.GetLineText(0) + "\Ajustes.csv"
                
            if self.cbxSeleccion.GetValue()== 'Cargas de Familia':
                op_=3
                file_ = self.txtDireccion.GetLineText(0) + "\Cargas de familia.csv"
                    
            if self.cbxSeleccion.GetValue()== 'Deducciones':
                op_=4
                file_ = self.txtDireccion.GetLineText(0) + "\Deducciones.csv"                         
                                
            if self.cbxSeleccion.GetValue()== 'Ganancias liquidadas por otros empleadores o entidades':
                op_=5
                file_ = self.txtDireccion.GetLineText(0) + "\Ganancias OE.csv"
                    
            if self.cbxSeleccion.GetValue()== 'Retenciones, percepciones y pagos a cuenta':
                op_=6
                file_ = self.txtDireccion.GetLineText(0) + "\Pagos a cta.csv"
                    
            if self.cbxSeleccion.GetValue()== 'Datos adicionales':
                op_=7
                file_ = self.txtDireccion.GetLineText(0) + "\Datos adicionales.csv"
                
            if self.cbxSeleccion.GetValue()== 'Presentación':
                op_=8
                file_ = self.txtDireccion.GetLineText(0) + "\Presentación.csv"    
            
            if os.path.isfile(file_) == True:
                texto= """EL archivo ya existe en la ubicación especificada. 
                ¿Desea eliminar el archivo y continuar con la operación?"""
                
                mensaje = wx.MessageDialog(parent= self, caption= 'Advertencia', message = texto, style= wx.YES_NO|wx.CANCEL)
                
                result = mensaje.ShowModal()
                
                if result == wx.ID_YES:
                    os.remove(file_)
                    self.procesar_xml(op_,file_)
                elif result == wx.ID_NO:
                    wx.MessageBox('Se aborto la operación')
            
            else:
                self.procesar_xml(op_,file_)
        
        
                
                
                
if __name__=="__main__":
    app=wx.App()
    frame = principal(None,"Payroll tools - Extractor SIRADIG")
    frame.SetBackgroundColour("lightgray")
    frame.Center()
    frame.Show()
    app.MainLoop()