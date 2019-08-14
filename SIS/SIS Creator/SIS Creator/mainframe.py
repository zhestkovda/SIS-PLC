import config
import text

import dbase
import template
import generate

import widgets

import os
import wx

class MainFrame(wx.Dialog):
    def __init__(self, *args, **kwds):
        wx.Dialog.__init__(self, *args, **kwds)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(wx.Panel(self, -1, size=(-1, config.border)), 0)
        
        h_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h_sizer.Add(wx.Panel(self, -1, size=(config.border, -1)), 0)
        
        v_sizer = wx.BoxSizer(wx.VERTICAL)

        # path to xls file
        v_sizer.Add(wx.StaticText(self, -1, text.MainPathToXls), 0)
        v_sizer.Add(wx.Panel(self, -1, size=(-1, config.border)), 0)
        
        h1_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.txtXLSPath = wx.TextCtrl(self, -1, "")
        h1_sizer.Add(self.txtXLSPath, 1, flag = wx.ALL|wx.EXPAND)
        h1_sizer.Add(wx.Panel(self, -1, size=(config.border, -1)), 0)
        
        self.btnSelectXls = wx.Button(self, -1, text.btnSelectXLS, size=(22, -1))
        h1_sizer.Add(self.btnSelectXls, 0)
        
        v_sizer.Add(h1_sizer, 0, flag = wx.ALL|wx.EXPAND)
        v_sizer.Add(wx.Panel(self, -1, size=(-1, config.border)), 0)

        # notebook
        self.notebook = wx.Notebook(self, -1, style = wx.BK_DEFAULT)

        self.sh1 = wx.Panel(self.notebook, -1)
        self.sh2 = wx.Panel(self.notebook, -1)
        self.notebook.AddPage(self.sh1, text.MainGeneralXLSSettings)
        self.notebook.AddPage(self.sh2, text.MainGeneralFHXSettings)
        
#===============================================================================
#       Sheet 1: template settings
#===============================================================================
        h_sh1_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), 0)
        
        v_sh1_sizer = wx.BoxSizer(wx.VERTICAL)
        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border)), 0)
                
        v_sh1_sizer.Add(widgets.SmallCenterHeaderPanel(self.sh1, style=wx.BORDER_SUNKEN, label=text.MainGeneralXLSSettings), 0, flag = wx.ALL|wx.EXPAND)
        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border)), 0)
        
        gbs_sizer = wx.GridBagSizer(6, 3)
        
        # number of items
        gbs_sizer.Add(wx.StaticText(self.sh1, -1, text.MainNOI), (0,0), flag = wx.ALIGN_CENTER_VERTICAL)
        
        self.txtNOI = wx.TextCtrl(self.sh1, -1, "", size = (238, -1))
        
        gbs_sizer.Add(self.txtNOI, (0, 2), flag = wx.ALL|wx.EXPAND)
                      
        # default area
        gbs_sizer.Add(wx.StaticText(self.sh1, -1, text.MainDefaultArea), (1,0), flag = wx.ALIGN_CENTER_VERTICAL)
        
        self.txtDefArea = wx.TextCtrl(self.sh1, -1, "", size = (238, -1))
        
        gbs_sizer.Add(self.txtDefArea, (1, 2), flag = wx.ALL|wx.EXPAND)
        
        # default status opts
        gbs_sizer.Add(wx.StaticText(self.sh1, -1, text.MainDefaultStatusOpts), (2, 0), flag = wx.ALIGN_CENTER_VERTICAL)
        
        self.chStatOpts = wx.Choice(self.sh1, -1, choices = text.XLSSheetVTRStOptsList)
        
        gbs_sizer.Add(self.chStatOpts, (2, 2), flag = wx.ALL|wx.EXPAND)
        v_sh1_sizer.Add(gbs_sizer, 0, flag = wx.ALL|wx.EXPAND)

        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border + 4)), 0)
                
        # default bypass opts
        v_sh1_sizer.Add(wx.StaticText(self.sh1, -1, text.MainDefaultBypassOpts), 0)
        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border)), 0)
        
        self.cbByp1 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[0])
        self.cbByp2 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[1])
        self.cbByp3 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[2])
        self.cbByp4 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[3])
        self.cbByp5 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[4])
        self.cbByp6 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[5])
        self.cbByp7 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[6])
        self.cbByp8 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[7])
        self.cbByp9 = wx.CheckBox(self.sh1, -1, text.XLSSheetVTRBypOptsList[8])
        
        gbs1_sizer = wx.GridBagSizer(2, 2)
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(2*config.border, -1)), (0, 0))
        gbs1_sizer.Add(self.cbByp1, (0, 1))        
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (1, 0))        
        gbs1_sizer.Add(self.cbByp2, (1, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (2, 0))
        gbs1_sizer.Add(self.cbByp3, (2, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (3, 0))
        gbs1_sizer.Add(self.cbByp4, (3, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (4, 0))
        gbs1_sizer.Add(self.cbByp5, (4, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (5, 0))
        gbs1_sizer.Add(self.cbByp6, (5, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (6, 0))
        gbs1_sizer.Add(self.cbByp7, (6, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (7, 0))
        gbs1_sizer.Add(self.cbByp8, (7, 1))
        gbs1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), (8, 0))
        gbs1_sizer.Add(self.cbByp9, (8, 1))

        v_sh1_sizer.Add(gbs1_sizer, 0, flag = wx.ALL)        
        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border)), 0)
        
        # external bypass
        self.cbExtBypPerm = wx.CheckBox(self.sh1, -1, text.MainExtBypass)
        v_sh1_sizer.Add(self.cbExtBypPerm, 0)
                             
        v_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(-1, config.border)), 0)
        h_sh1_sizer.Add(v_sh1_sizer, 0, flag = wx.ALL|wx.EXPAND)
        h_sh1_sizer.Add(wx.Panel(self.sh1, -1, size=(config.border, -1)), 0)
        
        self.sh1.SetSizerAndFit(h_sh1_sizer)
        
#===============================================================================
#       Sheet 2: fhx settings
#===============================================================================

        h_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(config.border, -1)), 0)
        
        v_sh2_sizer = wx.BoxSizer(wx.VERTICAL)
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(358, config.border)), 0)
                
        v_sh2_sizer.Add(widgets.SmallCenterHeaderPanel(self.sh2, style=wx.BORDER_SUNKEN, label=text.MainGeneralFHXSettings), 0, flag = wx.ALL|wx.EXPAND)
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)

        self.radioEng = wx.RadioButton(self.sh2, -1, text.MainDeltaVVersEng, style = wx.RB_GROUP)
        self.radioRus = wx.RadioButton(self.sh2, -1, text.MainDeltaVVersRus)

        gbs_sh2_sizer = wx.GridBagSizer(6, 3)
        
        # language
        gbs_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainDeltaVVers), (0, 0), flag = wx.ALIGN_CENTER_VERTICAL) 
                
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(self.radioEng, 0, flag = wx.ALL|wx.EXPAND)        
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(6*config.border, -1)), 0)
        h1_sh2_sizer.Add(self.radioRus, 0, flag = wx.ALL|wx.EXPAND)        
        gbs_sh2_sizer.Add(h1_sh2_sizer, (0, 2))

        # bypass permit name

        gbs_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainBypassPermName), (1,0), flag = wx.ALIGN_CENTER_VERTICAL)
        
        self.txtBPName = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        gbs_sh2_sizer.Add(self.txtBPName, (1, 2), flag = wx.ALL|wx.EXPAND)
        
        # bypass permit link

        gbs_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainBypassPermRef), (2,0), flag = wx.ALIGN_CENTER_VERTICAL)
        
        self.txtBPRef = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        gbs_sh2_sizer.Add(self.txtBPRef, (2, 2), flag = wx.ALL|wx.EXPAND)

        v_sh2_sizer.Add(gbs_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND)

        # Generate Areas
        self.cbArea = wx.CheckBox(self.sh2, -1, text.MainGenerateArea)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbArea, 0, flag = wx.ALL|wx.EXPAND)

        # Generate SLS 
        self.cbSLS = wx.CheckBox(self.sh2, -1, text.MainGenerateSLS)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbSLS, 0, flag = wx.ALL|wx.EXPAND)        

        # Domain Name
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, -1)), 0)
        
        self.txtDomainName = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        h1_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainGenerateDomainName), 0, flag = wx.ALIGN_CENTER_VERTICAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, -1)), 1, flag = wx.ALL|wx.EXPAND)
        
        h1_sh2_sizer.Add(self.txtDomainName, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND) 

        # Namur
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, -1)), 0)
        
        self.cbNamur = wx.CheckBox(self.sh2, -1, text.MainGenerateNamur)
        h1_sh2_sizer.Add(self.cbNamur, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND) 
                
        # Overrange
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, -1)), 0)
        
        self.txtOverrange = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        h1_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainGenerateOverange), 0, flag = wx.ALIGN_CENTER_VERTICAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, -1)), 1, flag = wx.ALL|wx.EXPAND)
        
        h1_sh2_sizer.Add(self.txtOverrange, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND) 

        # Underrange
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, -1)), 0)
        
        self.txtUnderrange = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        h1_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainGenerateUnderrange), 0, flag = wx.ALIGN_CENTER_VERTICAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, -1)), 1, flag = wx.ALL|wx.EXPAND)
        
        h1_sh2_sizer.Add(self.txtUnderrange, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND) 

        # Linefault detect
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, -1)), 0)
        
        self.cbLF = wx.CheckBox(self.sh2, -1, text.MainGenerateLF)
        h1_sh2_sizer.Add(self.cbLF, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND)         
                
        # autocalculate voter name
        self.cbName = wx.CheckBox(self.sh2, -1, text.MainAutocalcNames)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbName, 0, flag = wx.ALL|wx.EXPAND)

        # autocalculate depct
        
        self.cbDecpt = wx.CheckBox(self.sh2, -1, text.MainAutocalcDecpt)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbDecpt, 0, flag = wx.ALL|wx.EXPAND)
                
        # use external bypasses
        self.cbExtByp = wx.CheckBox(self.sh2, -1, text.MainUseExtBypass)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbExtByp, 0, flag = wx.ALL|wx.EXPAND)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)

        # Trip Hys
        h1_sh2_sizer = wx.BoxSizer(wx.HORIZONTAL)
       
        self.txtTripHys = wx.TextCtrl(self.sh2, -1, "", size = (229, -1))
        
        h1_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainGenerateTripHys), 0, flag = wx.ALIGN_CENTER_VERTICAL)
        h1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, -1)), 1, flag = wx.ALL|wx.EXPAND)
        
        h1_sh2_sizer.Add(self.txtTripHys, 0, flag = wx.ALL|wx.EXPAND)
        
        v_sh2_sizer.Add(h1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND)         
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)

        # ai bad if limited
        
        self.cbAIBadIfLimited = wx.CheckBox(self.sh2, -1, text.MainAIOpt)

        v_sh2_sizer.Add(self.cbAIBadIfLimited, 0, flag = wx.ALL|wx.EXPAND)
        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        
        # options for DO
        v_sh2_sizer.Add(wx.StaticText(self.sh2, -1, text.MainDefaultDOOpts), 0)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border/2)), 0)

        self.cbDO1 = wx.CheckBox(self.sh2, -1, text.MainDOOptsList[0])
        self.cbDO2 = wx.CheckBox(self.sh2, -1, text.MainDOOptsList[1])
        self.cbDO3 = wx.CheckBox(self.sh2, -1, text.MainDOOptsList[2])       
       
        gbs1_sh2_sizer = wx.GridBagSizer(2, 2)
        gbs1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, 18)), (0, 0))
        gbs1_sh2_sizer.Add(self.cbDO1, (0, 1))        
        gbs1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, 18)), (1, 0))        
        gbs1_sh2_sizer.Add(self.cbDO2, (1, 1))
        gbs1_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(2*config.border, 18)), (2, 0))
        gbs1_sh2_sizer.Add(self.cbDO3, (2, 1))
       
        v_sh2_sizer.Add(gbs1_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND)

        self.cbGenerateFB = wx.CheckBox(self.sh2, -1, text.MainGenerateFrameBorder)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        v_sh2_sizer.Add(self.cbGenerateFB, 0, flag = wx.ALL|wx.EXPAND)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)

        v_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(-1, config.border)), 0)
        h_sh2_sizer.Add(v_sh2_sizer, 0, flag = wx.ALL|wx.EXPAND)
        h_sh2_sizer.Add(wx.Panel(self.sh2, -1, size=(config.border, -1)), 0)

        self.sh2.SetSizerAndFit(h_sh2_sizer)
                
        v_sizer.Add(self.notebook, 1, flag = wx.ALL|wx.EXPAND)
                
        v_sizer.Add(wx.Panel(self, -1, size=(-1, config.border)), 0)
        
        b_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btnGenerateTemplate = wx.Button(self, -1, text.btnGenerateTemplate)
        self.btnGenerateFHX = wx.Button(self, -1, text.btnGenerateFHX)
        
        b_sizer.Add(self.btnGenerateTemplate, 0)
        b_sizer.Add(wx.Panel(self, -1, size=(config.border, -1)), 1)
        b_sizer.Add(self.btnGenerateFHX, 0)
        
        v_sizer.Add(b_sizer, 0, flag = wx.ALL|wx.EXPAND)        
        
        h_sizer.Add(v_sizer, 1, flag = wx.ALL|wx.EXPAND)
        h_sizer.Add(wx.Panel(self, -1, size=(config.border, -1)), 0)
        
        sizer.Add(h_sizer, 1, flag = wx.ALL|wx.EXPAND)
        sizer.Add(wx.Panel(self, -1, size=(-1, config.border)), 0)

        self.setInitValue()

        self.Bind(wx.EVT_BUTTON, self.clckSelectXLS, self.btnSelectXls)
        self.Bind(wx.EVT_BUTTON, self.clckGenerateTemplate, self.btnGenerateTemplate)
        self.Bind(wx.EVT_BUTTON, self.clckGenerateFHX, self.btnGenerateFHX)
        
        self.Bind(wx.EVT_TEXT, self.checkXLSPath, self.txtXLSPath)
        
        self.Bind(wx.EVT_TEXT, self.checkOverrange, self.txtOverrange)
        self.Bind(wx.EVT_TEXT, self.checkUnderrange, self.txtUnderrange)
        self.Bind(wx.EVT_TEXT, self.checkTripHys, self.txtTripHys)
        
        self.Bind(wx.EVT_CHECKBOX, self.processSLS, self.cbSLS)
        
        self.SetSizerAndFit(sizer)
        
    def setInitValue(self):
        xls_data = dbase.GetXLSData()
        
        self.txtNOI.SetValue(str(xls_data[1]))
        self.txtDefArea.SetValue(xls_data[2])
        self.chStatOpts.SetStringSelection(text.XLSSheetVTRStOptsList[xls_data[3]])
        
        if len(xls_data[4].split('1')) > 1:
            self.cbByp1.SetValue(True)
        if len(xls_data[4].split('2')) > 1:
            self.cbByp2.SetValue(True)
        if len(xls_data[4].split('3')) > 1:
            self.cbByp3.SetValue(True)
        if len(xls_data[4].split('4')) > 1:
            self.cbByp4.SetValue(True)
        if len(xls_data[4].split('5')) > 1:
            self.cbByp5.SetValue(True)
        if len(xls_data[4].split('6')) > 1:
            self.cbByp6.SetValue(True)
        if len(xls_data[4].split('7')) > 1:
            self.cbByp7.SetValue(True)
        if len(xls_data[4].split('8')) > 1:
            self.cbByp8.SetValue(True)
        if len(xls_data[4].split('9')) > 1:
            self.cbByp9.SetValue(True)
        
        self.cbExtBypPerm.SetValue(xls_data[5])
        
        fhx_data = dbase.GetFHXData()
        
        self.txtXLSPath.SetValue(fhx_data[1])
        
        if fhx_data[2] == 0:
            self.radioEng.SetValue(True)
        else:
            self.radioRus.SetValue(True)
        
        self.txtBPName.SetValue(fhx_data[3])
        self.txtBPRef.SetValue(fhx_data[4])
        self.cbName.SetValue(fhx_data[5])
        self.cbDecpt.SetValue(fhx_data[6])
        self.cbExtByp.SetValue(fhx_data[7])
        
        if len(fhx_data[8].split('1')) > 1:
            self.cbDO1.SetValue(True)
        if len(fhx_data[8].split('2')) > 1:
            self.cbDO2.SetValue(True)
        if len(fhx_data[8].split('3')) > 1:
            self.cbDO3.SetValue(True)
        
        self.cbGenerateFB.SetValue(fhx_data[9])
        
        self.setVisibility()
    
    def checkXLSPath(self, event): 
        if os.path.isfile(self.txtXLSPath.GetValue()) and len(self.txtXLSPath.GetValue().split('.xlsx')) > 1:
            self.txtXLSPath.SetForegroundColour(wx.BLACK) 
            self.txtXLSPath.Refresh()
        else:
            self.txtXLSPath.SetForegroundColour(wx.RED) 
            self.txtXLSPath.Refresh()

    def checkOverrange(self, event):
        flag = True
        
        try:
            val = float(self.txtOverrange.GetValue())
            if val < -25 or val > 125:
                flag = False
        except:
            flag = False

        if flag:
            self.txtOverrange.SetForegroundColour(wx.BLACK) 
            self.txtOverrange.Refresh()
        else:
            self.txtOverrange.SetForegroundColour(wx.RED) 
            self.txtOverrange.Refresh()            

    def checkUnderrange(self, event):
        flag = True
        
        try:
            val = float(self.txtUnderrange.GetValue())
            if val < -25 or val > 125:
                flag = False
        except:
            flag = False

        if flag:
            self.txtUnderrange.SetForegroundColour(wx.BLACK) 
            self.txtUnderrange.Refresh()
        else:
            self.txtUnderrange.SetForegroundColour(wx.RED) 
            self.txtUnderrange.Refresh()   
    
    def checkTripHys(self, event):
        flag = True
        
        try:
            val = float(self.txtTripHys.GetValue())
            if val < 0 or val > 50:
                flag = False
        except:
            flag = False

        if flag:
            self.txtTripHys.SetForegroundColour(wx.BLACK) 
            self.txtTripHys.Refresh()
        else:
            self.txtTripHys.SetForegroundColour(wx.RED) 
            self.txtTripHys.Refresh()

    def clckSelectXLS(self, event):
        dlg = wx.FileDialog(self, message=text.MsgOpenXLSFile, defaultDir=os.getcwd(), defaultFile="", wildcard=config.xlsWildcard, style=wx.OPEN)

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            
            self.txtXLSPath.SetValue(path)

    def processSLS(self, event):
        self.setVisibility()
       
    def setVisibility(self):
        if self.cbSLS.GetValue():
            self.txtDomainName.Enable()            
            self.cbNamur.Enable()
            self.txtOverrange.Enable()
            self.txtUnderrange.Enable() 
            self.cbLF.Enable()
        else:
            self.txtDomainName.Disable()
            self.cbNamur.Disable()
            self.txtOverrange.Disable()
            self.txtUnderrange.Disable() 
            self.cbLF.Disable()

    def clckGenerateTemplate(self, event):
        try:
            noi = int(self.txtNOI.GetValue())
        except:
            noi = 0
            
        area = self.txtDefArea.GetValue()
        stat_opts = self.chStatOpts.GetStringSelection()
        
        strByp = ''

        if self.cbByp1.GetValue():
            strByp += '1,'
        if self.cbByp2.GetValue():
            strByp += '2,'
        if self.cbByp3.GetValue():
            strByp += '3,'
        if self.cbByp4.GetValue():
            strByp += '4,'                                 
        if self.cbByp5.GetValue():
            strByp += '5,'        
        if self.cbByp6.GetValue():
            strByp += '6,'
        if self.cbByp7.GetValue():
            strByp += '7,'            
        if self.cbByp8.GetValue():
            strByp += '8,'
        if self.cbByp9.GetValue():
            strByp += '9,'
        
        if len(strByp) > 0:
            strByp = strByp[0:len(strByp)-1]
                                            
        if self.cbExtBypPerm.GetValue():
            ext_byp_perm = "Yes"
        else:
            ext_byp_perm = "No" 

        dlg = wx.FileDialog(self, message=text.MsgSaveXLSFile, defaultDir=os.getcwd(), defaultFile="", wildcard=config.xlsWildcard, style=wx.SAVE)

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            
            template.create(path, noi, area, stat_opts, strByp, ext_byp_perm)
        
    def clckGenerateFHX(self, event):
        dlg = wx.FileDialog(self, message=text.MsgSaveFHXFile, defaultDir=os.getcwd(), defaultFile="", wildcard=config.fhxWildcard, style=wx.SAVE)

        if dlg.ShowModal() == wx.ID_OK:
            c_path = dlg.GetPath()
            c_xls_path = self.txtXLSPath.GetValue()
            
            if self.radioRus.GetValue():
                c_lang = True
            else:
                c_lang = False
            
            c_bp_name = self.txtBPName.GetValue()
            c_bp_link = self.txtBPRef.GetValue()
            c_decpt = self.cbDecpt.GetValue()
            c_name = self.cbName.GetValue()
            c_ext_byp = self.cbExtByp.GetValue()
            
            c_do_list = [self.cbDO1.GetValue(), self.cbDO2.GetValue(), self.cbDO3.GetValue()]
            
            c_fb = self.cbGenerateFB.GetValue()
            
            c_cbArea = self.cbArea.GetValue()
            c_cbSLS = self.cbSLS.GetValue()
            c_txtDomain = self.txtDomainName.GetValue()
            c_cbNamur = self.cbNamur.GetValue()
            c_txtOverrange = self.txtOverrange.GetValue()
            c_txtUnderrange = self.txtUnderrange.GetValue()
            c_cbLF = self.cbLF.GetValue()
            c_txtTripHys = self.txtTripHys.GetValue()
            c_cbAIBadIfLimited = self.cbAIBadIfLimited.GetValue()
            
            generate.main_logic(c_path, c_xls_path, c_lang, c_bp_name, c_bp_link, c_name, c_decpt, c_ext_byp, c_do_list, c_fb, c_cbArea, c_cbSLS, c_txtDomain, c_cbNamur, c_txtOverrange, c_txtUnderrange, c_cbLF, c_txtTripHys, c_cbAIBadIfLimited)
