# -*- coding: cp1251 -*-

import wx
import wx.lib.newevent

HEADER_TEXT_COLOR = wx.SYS_COLOUR_HIGHLIGHTTEXT
HEADER_PANEL_BACKGROUND = wx.SYS_COLOUR_HIGHLIGHT
TEXT_ERROR = 'RED'

BORDER = 6

(StateChangedEvent, EVT_STATE_CHANGED) = wx.lib.newevent.NewEvent()


class DbObjectState(dict):
    def __init__(self, obj):
        dict.__init__(self)
        
        for c in obj.c:
            val = getattr(obj, c.key)
            self[c.key] = val


class BoldStaticText( wx.StaticText ):
    def __init__( self, *args, **kwds ):
        
        try:
            delta = kwds.pop('delta')
        except KeyError:
            delta = 0
        
        color = kwds.pop('color', None)
        
        wx.StaticText.__init__( self, *args, **kwds )
        font = self.GetFont()
        font.SetWeight( wx.BOLD )
        font.SetPointSize( font.GetPointSize() + delta )
        if color:
            self.SetForegroundColour(color)
        self.SetFont( font )


class HeaderPanel(wx.Panel):
    def __init__(self, *args, **kwds):
        self.label = kwds.pop('label')
        
        wx.Panel.__init__(self, *args, **kwds)
        
        self.sizer=wx.BoxSizer(wx.VERTICAL)
        
        background = wx.SystemSettings.GetColour(HEADER_PANEL_BACKGROUND)
        self.SetBackgroundColour(background)
        
        self.text = BoldStaticText(self, -1, self.label, delta=3, 
            color=wx.SystemSettings.GetColour(HEADER_TEXT_COLOR))
        
        self.sizer.Add(self.text, flag=wx.ALL, border=BORDER)
        self.SetSizerAndFit(self.sizer)
    
    def setLabel(self, label):
        if self.label != label:
            self.text.SetLabel(label)
            self.label = label
        
class SmallHeaderPanel(wx.Panel):
    def __init__(self, *args, **kwds):
        self.label = kwds.pop('label')
        
        wx.Panel.__init__(self, *args, **kwds)
        
        self.sizer=wx.BoxSizer(wx.VERTICAL)
        
        #background = wx.SystemSettings.GetColour(HEADER_PANEL_BACKGROUND)
        background = (0, 80, 165, 255)
        self.SetBackgroundColour(background)
        
        self.text = BoldStaticText(self, -1, self.label, delta=1,
            color=wx.SystemSettings.GetColour(HEADER_TEXT_COLOR))
        
        self.sizer.Add(self.text, flag=wx.ALL, border=BORDER)
        self.SetSizerAndFit(self.sizer)
    
    def setLabel(self, label):
        if self.label != label:
            self.text.SetLabel(label)
            self.label = label

class SmallCenterHeaderPanel(wx.Panel):
    def __init__(self, *args, **kwds):
        self.label = kwds.pop('label')
        
        wx.Panel.__init__(self, *args, **kwds)
        
        self.sizer=wx.BoxSizer(wx.VERTICAL)
        
        #background = wx.SystemSettings.GetColour(HEADER_PANEL_BACKGROUND)
        background = (0, 80, 165, 255)
        self.SetBackgroundColour(background)
        
        self.text = BoldStaticText(self, -1, self.label, delta=1,
            color=wx.SystemSettings.GetColour(HEADER_TEXT_COLOR))
        
        self.sizer.Add(self.text, flag=wx.ALL|wx.CENTRE, border=BORDER)
        self.SetSizerAndFit(self.sizer)
    
    def setLabel(self, label):
        if self.label != label:
            self.text.SetLabel(label)
            self.label = label


class MediumHeaderPanel(wx.Panel):
    def __init__(self, *args, **kwds):
        self.label = kwds.pop('label')
        
        wx.Panel.__init__(self, *args, **kwds)
        
        self.sizer=wx.BoxSizer(wx.VERTICAL)
        
        background = wx.SystemSettings.GetColour(HEADER_PANEL_BACKGROUND)
        self.SetBackgroundColour(background)
        
        self.text = BoldStaticText(self, -1, self.label, delta=3,
            color=wx.SystemSettings.GetColour(HEADER_TEXT_COLOR))
        
        self.sizer.Add(self.text, flag=wx.ALL|wx.CENTRE, 
            border=BORDER)
        self.SetSizerAndFit(self.sizer)
    
    def setLabel(self, label):
        if self.label != label:
            self.text.SetLabel(label)
            self.label = label
    
class IntegerEdit(wx.TextCtrl):
    def __init__(self, *args, **kwds):
        
        self.minValue = kwds.pop('min_value')
        self.maxValue = kwds.pop('max_value')
        self.title = kwds.pop('title')
        
        wx.TextCtrl.__init__(self, *args, **kwds)
        
        self.valid = True
        
        self.Bind(wx.EVT_TEXT, self.processTextChange)
    
    def processTextChange(self, event):
        val = self.getValue()
        
        if self.valid:
            self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNTEXT))
            self.Refresh()
        else:
            self.SetForegroundColour(TEXT_ERROR)
            self.Refresh()
        
        wx.PostEvent(self, StateChangedEvent(valid=self.valid, value=val))
        
    def setValue(self, val):
        self.SetValue( str(val) )
    
    def getValue(self):
        try:
            val = int(self.GetValue())
            
            if val < self.minValue or val > self.maxValue:
                self.valid = False
            else:
                self.valid = True
        except ValueError:
            self.valid = False
            val = None
        return val
    
    def disable(self):
        self.Disable()
        self.SetValue('')
        
    def enable(self):
        self.Enable()


class FloatEdit(wx.TextCtrl):
    def __init__(self, *args, **kwds):
        
        self.minValue = kwds.pop('min_value')
        self.maxValue = kwds.pop('max_value')
        self.title = kwds.pop('title')
        
        wx.TextCtrl.__init__(self, *args, **kwds)
        
        self.valid = True
        
        self.Bind(wx.EVT_TEXT, self.processTextChange)
    
    def processTextChange(self, event):
        val = self.getValue()
        
        if self.valid:
            self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNTEXT))
            self.Refresh()
        else:
            self.SetForegroundColour(TEXT_ERROR)
            self.Refresh()
        
        wx.PostEvent(self, StateChangedEvent(valid=self.valid, value=val))
        
    def setValue(self, val):
        self.SetValue( '%1.1f'%val )
    
    def getValue(self):
        try:
            val = float(self.GetValue())
            val = round(val, 1)
            
            if val < self.minValue or val > self.maxValue:
                self.valid = False
            else:
                self.valid = True
        except ValueError:
            self.valid = False
            val = None
        return val
    
    def disable(self):
        self.Disable()
        self.SetValue('')
        
    def enable(self):
        self.Enable()


class ChoiceEdit(wx.Choice):
    def __init__(self, *args, **kwds):
        self.choices = kwds['choices']
        self.title = kwds.pop('title')
        self.valid = True
        
        wx.Choice.__init__(self, *args, **kwds)
        
        self.Bind(wx.EVT_CHOICE, self.processChoice)
    
    def processChoice(self, event):
        wx.PostEvent(self, StateChangedEvent(valid=True, value=self.getValue()) )
    
    def getValue(self):
        return self.GetSelection()
    
    def setValue(self, val):
        self.SetSelection(val)
    
    def disable(self):
        self.SetSelection(0)
        self.Disable()
    
    def enable(self):
        self.Enable()

class StringEdit(wx.TextCtrl):
    def __init__(self, *args, **kwds):
        
        self.maxLength = kwds.pop('max_length')
        self.title = kwds.pop('title')
        
        wx.TextCtrl.__init__(self, *args, **kwds)
        
        self.valid = True
        
        self.Bind(wx.EVT_TEXT, self.processTextChange)
    
    def processTextChange(self, event):
        name = self.getValue()
        
        if self.valid:
            self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNTEXT))
            self.Refresh()
        else:
            self.SetForegroundColour(TEXT_ERROR)
            self.Refresh()
        
        wx.PostEvent(self, StateChangedEvent(valid=self.valid, value=name))
        
    def setValue(self, name):
        self.SetValue(name)
    
    def getValue(self):
        name = self.GetValue()            
        if len(name) > self.maxLength:
            self.valid = False
        else:
            self.valid = True

        if self.valid == True:
            return name
    
    def disable(self):
        self.Disable()
        self.SetValue('')
        
    def enable(self):
        self.Enable()   

class DateEdit(wx.DatePickerCtrl):
    def __init__(self, *args, **kwds):
        kwds['style'] = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY
        wx.DatePickerCtrl.__init__(self, *args, **kwds)

        import datetime as dt
        c = dt.datetime.now()
        #print c
        #self.SetValue(c)
        self.valid = True