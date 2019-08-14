import dbase
import config
import mainframe
import text

import wx

def main():  
    app = wx.App()

    dbase.CheckDB()
    dlg = mainframe.MainFrame(None, title=text.MainTitle)
   
    dlg.CenterOnScreen()
    dlg.ShowModal()
    dlg.Destroy()

    app.MainLoop()

if __name__ == '__main__':
    main()