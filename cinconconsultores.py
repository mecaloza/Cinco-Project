# -*- coding: utf-8 -*-
"""
Created on Thu Nov 28 14:28:41 2019

@author: user
"""
from selenium import webdriver
import wx
import time
import openpyxl
from webscraping import asignar_nro_proceso, get_the_web
from get_cities import get_cities_entities_web, make_cities_entities_dictionary
import os


DB = openpyxl.load_workbook('Database-Process.xlsx')

sheet = DB['Hoja1']

class MyFrame(wx.Frame):
    
    def OnKeyDown(self, event):
        """quit if user press q or Esc"""
        if event.GetKeyCode() == 27 or event.GetKeyCode() == ord('Q'): #27 is Esc
            self.Close(force=True)
            
        else:
            event.Skip()
 
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, "Software Legal", size=(1200, 700))  
        self.Bind(wx.EVT_KEY_UP, self.OnKeyDown)
        
        try:
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(
                image_file, 
                wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            self.panel = wx.StaticBitmap(
                self, -1, bmp1, (0, 0))
            
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
            
            
        button = wx.Button(self.panel, id=wx.ID_ANY, label="Ingresar Proceso" ,pos=(900, 100), size=(200, 50))
        button.Bind(wx.EVT_BUTTON, self.Ingresarproceso)
        
        button2 = wx.Button(self.panel, id=wx.ID_ANY, label="Consultar Proceso" ,pos=(900, 150), size=(200, 50))
        button2.Bind(wx.EVT_BUTTON, self.BtnConsultaProceso)
        
        button3 = wx.Button(self.panel, id=wx.ID_ANY, label="Actualizar información proceso" ,pos=(900, 200), size=(200, 50))
        button3.Bind(wx.EVT_BUTTON, self.onButton3)
        
        btn_asignar_procesos = wx.Button(self.panel, id=wx.ID_ANY, label="Ident. Nro Proceso" ,pos=(900, 250), size=(200, 50))
        btn_asignar_procesos.Bind(wx.EVT_BUTTON, self.onBtn_asignar_procesos)
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)

 
    #-------------Button Functions-----------------#
    def Ingresarproceso(self, event):
        secondWindow = window2(parent=self.panel)
        secondWindow.Show()

    def BtnConsultaProceso(self, event): 
        consultawindow=ww_Consultar_Proceso(parent=self.panel)
        consultawindow.Show()

        
    def onButton3(self, event):
        print ("Button pressed!")
        
    def onBtn_asignar_procesos(self, event):
        asignar_nro_proceso()
        
    #-------------Button Functions-----------------#    

        
class window2(wx.Frame):
    
    ciudades_entidades=make_cities_entities_dictionary()
    title = "Ingresar Proceso"
    
    def __init__(self,parent):
        ciudades_entidades=make_cities_entities_dictionary()
        wx.Frame.__init__(self,parent, -1,'Ingresar Proceso', size=(1200,700))   

        try:
            
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(
                image_file, 
                wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            self.panel = wx.StaticBitmap(
                self, -1, bmp1, (0, 0))
           
       
          
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        
        self.SetIcon(ico)
        self.lblname1 = wx.StaticText(self.panel, label="Ciudad:", pos=(600, 50))
        self.lblname1.SetBackgroundColour("white")
        self.Ciudad=wx.ComboBox(self.panel, choices=ciudades_entidades[1],id=wx.ID_ANY, pos=(700,45),size=(100, -1))
        #self.Ciudad = wx.TextCtrl(self.panel, size=(100, -1),pos=(700, 45))
        self.Ciudad.Bind(wx.EVT_COMBOBOX, self.get_entidades)
    
        self.lblname10 = wx.StaticText(self.panel, label="Entidad:", pos=(600, 100))
        self.lblname10.SetBackgroundColour("white")
        self.Entidades=wx.ComboBox(self.panel, choices=ciudades_entidades[0]['MEDELLIN '], pos=(700,95),size=(100, -1))      

        self.lblname2 = wx.StaticText(self.panel, label="Jurisdicción:",pos=(600, 150))
        self.lblname2.SetBackgroundColour("white")
        self.Jurisdi = wx.TextCtrl(self.panel, size=(100, -1),pos=(700, 145))
        
        self.lblname3 = wx.StaticText(self.panel, label="Por Nombre o \npor razon Social:",pos=(600, 250))
        self.lblname3.SetBackgroundColour("white")
        self.razon = wx.TextCtrl(self.panel, size=(100, -1),pos=(700, 245))
        
        self.lblname4 = wx.StaticText(self.panel, label="Tipo Sujeto:",pos=(600, 350))
        self.lblname4.SetBackgroundColour("white")
        self.Tipsuj = wx.ComboBox(self.panel, choices=['Demandado','Demandante'], size=(100, -1),pos=(700, 345))
        
        self.lblname5 = wx.StaticText(self.panel, label="Responsable:",pos=(600, 450))
        self.lblname5.SetBackgroundColour("white")
        self.Responsable = wx.TextCtrl(self.panel, size=(100, -1),pos=(700, 445))
        
        self.lblname6 = wx.StaticText(self.panel, label="Tipo de Persona:",pos=(600, 550))
        self.lblname6.SetBackgroundColour("white")
        self.Tipopersona = wx.ComboBox(self.panel, choices=['Natural','Juridica'], size=(100, -1),pos=(700, 545))
        
        self.lblname7 = wx.StaticText(self.panel, label="Nombre:",pos=(900, 50))
        self.lblname7.SetBackgroundColour("white")
        self.Nombre= wx.TextCtrl(self.panel, size=(100, -1),pos=(1050, 45))
        
        self.lblname8 = wx.StaticText(self.panel, label="Fecha de Radicación:",pos=(900, 150))
        self.lblname8.SetBackgroundColour("white")
        self.Fechara = wx.TextCtrl(self.panel, size=(100, -1),pos=(1050, 145))
        
        self.lblname9 = wx.StaticText(self.panel, label="Cedula:",pos=(900, 250))
        self.lblname9.SetBackgroundColour("white")
        self.Cedula = wx.TextCtrl(self.panel, size=(100, -1),pos=(1050, 245))
        
        self.lblname10 = wx.StaticText(self.panel, label="Apoderado:",pos=(900, 350))
        self.lblname10.SetBackgroundColour("white")
        self.Apoderado = wx.TextCtrl(self.panel, size=(100, -1),pos=(1050, 345))
        
        button = wx.Button(self.panel, id=wx.ID_ANY, label="Crear Proceso" ,pos=(900, 400), size=(200, 100))
        button.Bind(wx.EVT_BUTTON, self.Crearproceso)
        
        button = wx.Button(self.panel, id=wx.ID_ANY, label="Cancelar" ,pos=(900, 550), size=(200, 100))
        button.Bind(wx.EVT_BUTTON, self.OnCloseWindow)
        
        self.SetBackgroundColour(wx.Colour(100,100,100))
        self.Centre()
        self.Show()
        
    def OnCloseWindow(self, event):
        self.Destroy()

    def get_entidades(self):
        global Ciudad
        Ciudad = self.Ciudad.GetValue()
        return Ciudad
        
    def Crearproceso(self, event):
        
        Nproce = 1
        
        while (sheet.cell(row = Nproce, column = 1).value != None) :
          Nproce = Nproce + 1
          print(Nproce)

        Ciudad = self.Ciudad.GetValue()
        sheet.cell(row = Nproce, column = 2).value = Ciudad
        
        Jurisdi = self.Jurisdi.GetValue()
        sheet.cell(row = Nproce, column = 3).value = Jurisdi
        
        Entidad = self.Entidades.GetValue()
        sheet.cell(row = Nproce, column = 3).value = Entidad
        
        razon= self.razon.GetValue()
        sheet.cell(row = Nproce, column = 4).value = razon
        
        Tipsuj = self.Tipsuj.GetValue()
        sheet.cell(row = Nproce, column = 5).value = Tipsuj 
        
        Responsable = self.Responsable.GetValue()
        sheet.cell(row = Nproce, column = 6).value = Responsable
        
        Tipopersona = self.Tipopersona.GetValue()
        sheet.cell(row = Nproce, column = 7).value = Tipopersona
        
        Nombre  = self.Nombre.GetValue()
        sheet.cell(row = Nproce, column = 8).value = Nombre
        
        Fechara  = self.Fechara.GetValue()
        sheet.cell(row = Nproce, column = 9).value = Fechara
        
        Cedula= self.Cedula.GetValue()
        sheet.cell(row = Nproce, column = 10).value = Cedula
        
        Apoderado = self.Apoderado.GetValue()
        sheet.cell(row = Nproce, column = 11).value =Apoderado
         
        sheet.cell(row = Nproce  , column = 1).value = Nproce 
        
        
#        filepath = "C:\Users\user\.spyder-py3\" 

        DB.save('Database-Process.xlsx')

class ww_Consultar_Proceso(wx.Frame):
    

    def __init__(self,parent):
        wx.Frame.__init__(self,parent, -1,'Consultar Proceso', size=(1200,700))   

        try:
            
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(
                image_file, 
                wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            self.panel = wx.StaticBitmap(
                self, -1, bmp1, (0, 0))
           
       
          
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        
        self.lblname1 = wx.StaticText(self.panel, label="Ingrese Numero de Proceso", pos=(830, 250))
        self.lblname1.SetBackgroundColour("white")
        self.numero_consulta=wx.TextCtrl(self.panel, size=(300, -1),pos=(750, 270))

        btn_consultar = wx.Button(self.panel, id=wx.ID_ANY, label="Consultar" ,pos=(840, 300), size=(100, 30))
        btn_consultar.Bind(wx.EVT_BUTTON, self.Consultar_Excel)
        
    def Consultar_Excel(self,event):
        
        numero_consulta=self.numero_consulta.GetValue()
        print(os.getcwd())
        workbook_path=os.getcwd()+'/Procesos/'+ numero_consulta + '.xlsx'
        os.startfile(workbook_path)
        
class MyApp(wx.App):
    
    def OnInit(self):
        self.frame= MyFrame()
        self.frame.Show()
        return True       
 
# Run the program     
app=MyApp()
app.MainLoop()
del app