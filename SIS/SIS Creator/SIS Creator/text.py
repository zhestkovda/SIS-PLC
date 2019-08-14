# -*- coding: cp1251 -*-

MainTitle = u'SIS Creator'

MainGeneralXLSSettings = u'Template Settings'
MainGeneralFHXSettings = u'FHX Settings'

MainPathToXls = u'Path to excel spreadsheet:'
MainDeltaVVers = u'DeltaV Language:'
MainDeltaVVersEng = u'English'
MainDeltaVVersRus = u'Russian'
MainNOI = u'Number of items:'
MainDefaultArea = u'Default Area:'
MainDefaultStatusOpts = u'Default Status Opts:'
MainDefaultBypassOpts = u'Default Bypass Opts:'
MainExtBypass = u'Use external bypass permit'

MainBypassPermName = u'Bypass Permit Name:'
MainBypassPermRef = u'Bypass Permit Ref:'
MainGenerateArea = u'Generate Areas' 
MainGenerateSLS = u'Generate Channels Configuration'
MainGenerateDomainName = u'Domain Name:'
MainGenerateNamur = u'Enable Namur for AI channels'
MainGenerateLF = u'Enable Linefault Detect for DI/DO channels'
MainGenerateOverange = u'Overrange for AI:'
MainGenerateUnderrange = u'Underrange for AI:'
MainAutocalcNames = u"Use DST names for Function Blocks"
MainAutocalcDecpt = u'Calculate Decpt automatically'
MainUseExtBypass = u'Use external bypasses'
MainGenerateTripHys = u'LSAVTR Trip Hysteresis:'
MainAIOpt = u'Enable "Bad if Limited" option for AI'
MainDefaultDOOpts = u'Default options for LSDO:'
MainDOOptsList = [u'Enable detection based on CAS_IN_D status', u'Enable detection based on output channel status', u'Enable detection based on PV_D status']
MainGenerateFrameBorder = u'Generate Frame'

btnSelectXLS = "..."
btnGenerateTemplate = u'Generate Template'
btnGenerateFHX = u'Generate FHX'

MsgOpenXLSFile = u'Select Excel File'
MsgSaveXLSFile = u'Save Template File'
MsgSaveFHXFile = u'Save FHX File'

defaultArea = u'AREA_A'
defaultChType = u'UNDEFINED_CHAN'
defaultChDesc = u'Undefined Channel Type'
SISDomainName = u'SIS_DOMAIN'


FHXRev = [u'Rev', u'Ревизия']
FHXDate = [u'Date', u'Дата']
FHXAuthor = [u'Author', u'Автор']
FHXComments = [u'Comments', u'Комментарии']

XLSSheetTitle = u'Template'

XLSSheetChannelsGeneral = u'Channels Information'

XLSSheetN = u'№'

XLSSheetSLS = u'SLS'
XLSSheetChType = u'Type' 
XLSSheetChannel = u'Channel'
XLSSheetChDesc = u'Channel Description'

XLSSheetCHTypeList = [u'AI', u'HART AI', u'HART AO', u'DI', u'DO']
XLSSheetCHTypeDVList = [u'AI_LS_CHAN', u'AI_HART_LS_CHAN', u'AO_HART_LS_CHAN', u'DI_LS_CHAN', u'DO_LS_CHAN']

XLSSheetModulesGeneral = u'Modules Information'
XLSSheetArea = u'Area'
XLSSheetModName = u'Module Name'
XLSSheetModDesc = u'Module Description'

XLSSheetFBGeneral = u'Functional Block LSAI/LSDI/LSDO'
XLSSheetFBName = u'Name'
XLSSheetFBType = u'Type'
XLSSheetDST = u'DST'
XLSSheetEU0 = u'0'
XLSSheetEU100 = u'100'
XLSSheetUnits = u'Units'
XLSSheetDecpt = u'Decpt'

XLSSheetFBTypeList = ['AI', 'DI', 'DO']

XLSSheetVTRGeneral = u'Functional Block LSAVTR/LSDVTR'
XLSSheetVTRName = u'Name'
XLSSheetVTRType = u'Type' 
XLSSheetVTRIn = u'In'
XLSSheetVTRNum2Trip = u'Num2Trip'
XLSSheetVTRDetType = u'Detect Type'
XLSSheetVTRPreTripLim = u'P Trip'
XLSSheetVTRTripLim = u'Trip'
XLSSheetVTRStatusOpts = u'Status Opts'
XLSSheetVTRBypOpts = u'Bypass Opts'
XLSSheetVTRBypassed = u'Bypass Permit'

XLSSheetVTRTypeList = ['AVTR', 'DVTR']
XLSSheetVTRDetTypeList = ['Greater Than', 'Less Than']
XLSSheetVTRDetTypeListRus = [u'Больше Чем', u'Меньше Чем']
XLSSheetVTRStOptsList = [u'', u'Always Use Value', u'Will Not Vote if Bad', u'Vote to Trip if Bad']
XLSSheetVTRStOptsListRus = [u'Всегда Использовать Значение', u'Не Будет Голосовать если Плохой', u'Голосовать за Защиту если Плохой']
XLSSheetVTRBypOptsList = [u'A maintenance bypass reduces the number to trip', u'Multiple maintenance bypasses are allowed', u'Maintenance bypass timeout is for indication only', u'Startup bypass preset is allowed while active', u'Startup bypass expires upon stabilization', u'Reminder applies to startup bypass', u'Startup bypass duration is event based', u'Bypass permit is not required to bypass', u'Bypass permit control should be visible in operator interface']
XLSSheetVTRBypYesNoList = [u'No', u'Yes']