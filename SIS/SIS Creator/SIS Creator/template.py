# -*- coding: cp1251 -*-

import text
import xlsxwriter

def create(c_path, c_num, area, stat_opts, strByp, ext_byp_perm):
    
    wb = xlsxwriter.Workbook(c_path)
    ws = wb.add_worksheet(text.XLSSheetTitle)

    # set width
    ws.set_column('A:A', 3)
 
    ws.set_column('B:B', 13)
    ws.set_column('C:C', 13)
    ws.set_column('D:D', 13)
    ws.set_column('E:E', 8)
    ws.set_column('F:F', 40)
     
    ws.set_column('G:G', 13)    
    ws.set_column('H:H', 13.1)
    ws.set_column('I:I', 40)
    
    ws.set_column('J:J', 13)
    ws.set_column('K:K', 5)
    ws.set_column('L:L', 5)
    ws.set_column('M:M', 5)
    ws.set_column('N:N', 5)
    ws.set_column('O:O', 6)

    ws.set_column('P:P', 13)
    ws.set_column('Q:Q', 7)
    ws.set_column('R:R', 5)
    ws.set_column('S:S', 9)
    ws.set_column('T:T', 13)
    ws.set_column('U:U', 6)
    ws.set_column('V:V', 6)
    ws.set_column('W:W', 18)
    ws.set_column('X:X', 13)
    ws.set_column('Y:Y', 13)
    
    ws.set_row(0, 20)
    
    # Add a format for the header cells.
    main_header_format = wb.add_format({
        'border': 1,
        'bg_color': '#92D050',
        'bold': True,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'indent': 0,
        })    

    header_format = wb.add_format({
        'border': 1,
        'bg_color': '#C6EFCE',
        'bold': True,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'indent': 0,
        }) 
    
    align_center_format = wb.add_format({
        'align': 'center',
        'valign': 'vcenter',
        }) 
    
    # general information
    ws.merge_range('A1:F1', text.XLSSheetChannelsGeneral, main_header_format)    
    
    ws.write('A2', text.XLSSheetN, header_format)
    ws.write('B2', text.XLSSheetSLS, header_format)
    ws.write('C2', text.XLSSheetDST, header_format)
    ws.write('D2', text.XLSSheetChType, header_format)
    ws.write('E2', text.XLSSheetChannel, header_format)
    ws.write('F2', text.XLSSheetChDesc, header_format)
    
    
    ws.merge_range('G1:I1', text.XLSSheetModulesGeneral, main_header_format)
    ws.write('G2', text.XLSSheetArea, header_format)
    ws.write('H2', text.XLSSheetModName, header_format)
    ws.write('I2', text.XLSSheetModDesc, header_format)
    
    # fb information
    ws.merge_range('J1:O1', text.XLSSheetFBGeneral, main_header_format)    
    
    ws.write('J2', text.XLSSheetFBName, header_format)
    ws.write('K2', text.XLSSheetFBType, header_format)
    ws.write('L2', text.XLSSheetEU0, header_format)
    ws.write('M2', text.XLSSheetEU100, header_format)    
    ws.write('N2', text.XLSSheetUnits, header_format)
    ws.write('O2', text.XLSSheetDecpt, header_format)    

    # vtr information
    ws.merge_range('P1:Y1', text.XLSSheetVTRGeneral, main_header_format)    
    
    ws.write('P2', text.XLSSheetVTRName, header_format)
    ws.write('Q2', text.XLSSheetVTRType, header_format)
    ws.write('R2', text.XLSSheetVTRIn, header_format)
    ws.write('S2', text.XLSSheetVTRNum2Trip, header_format)
    ws.write('T2', text.XLSSheetVTRDetType, header_format)
    ws.write('U2', text.XLSSheetVTRPreTripLim, header_format)
    ws.write('V2', text.XLSSheetVTRTripLim, header_format)
    ws.write('W2', text.XLSSheetVTRStatusOpts, header_format)  
    ws.write('X2', text.XLSSheetVTRBypOpts, header_format)    
    ws.write('Y2', text.XLSSheetVTRBypassed, header_format)
    
    # fill values     
    for i in xrange(3, 3 + c_num):
        # current item
        ws.write('A' + str(i), str(i-2), align_center_format)
        
        # ch type
        ws.data_validation('D' + str(i), {'validate': 'list',
                                  'source': text.XLSSheetCHTypeList})        
        
        ws.write('D' + str(i), '', align_center_format)
        
        # channel
        ws.data_validation('E' + str(i), {'validate': 'integer',
                                 'criteria': 'between',
                                 'minimum': 1,
                                 'maximum': 16,
                                 'error_title': 'Input value is not valid!',
                                 'error_message': 'It should be an integer between 1 and 16'})        
        
        ws.write('E' + str(i), '', align_center_format)
        
        # area
        ws.write('G' + str(i), area, align_center_format)
        
        # fb type
        ws.data_validation('K' + str(i), {'validate': 'list',
                                  'source': text.XLSSheetFBTypeList})
        # scale 
        ws.data_validation('L' + str(i), {'validate': 'decimal',
                                 'criteria': 'between',
                                 'minimum': -999999.999,
                                 'maximum': 999999.999})
        
        ws.data_validation('M' + str(i), {'validate': 'decimal',
                                 'criteria': 'between',
                                 'minimum': -999999.999,
                                 'maximum': 999999.999})

        ws.data_validation('O' + str(i), {'validate': 'integer',
                                 'criteria': 'between',
                                 'minimum': 0,
                                 'maximum': 5,
                                 'error_title': 'Input value is not valid!',
                                 'error_message': 'It should be an integer between 1 and 5'})   
        
        ws.write('L' + str(i), '', align_center_format)
        ws.write('M' + str(i), '', align_center_format)
        ws.write('N' + str(i), '', align_center_format)
        ws.write('O' + str(i), '', align_center_format)
 
        # vtr type        
        ws.data_validation('Q' + str(i), {'validate': 'list',
                                  'source': text.XLSSheetVTRTypeList})

        ws.write('Q' + str(i), '', align_center_format)
        
        # vtr in
        ws.data_validation('R' + str(i), {'validate': 'integer',
                                 'criteria': 'between',
                                 'minimum': 1,
                                 'maximum': 16,
                                 'error_title': 'Input value is not valid!',
                                 'error_message': 'It should be an integer between 1 and 16'})

        ws.write('R' + str(i), '', align_center_format)
        
        # vtr num2trip
        ws.data_validation('S' + str(i), {'validate': 'integer',
                                 'criteria': 'between',
                                 'minimum': 1,
                                 'maximum': 16,
                                 'error_title': 'Input value is not valid!',
                                 'error_message': 'It should be an integer between 1 and 16'})                                 
        
        ws.write('S' + str(i), '', align_center_format)
        
        # vtr detect type
        ws.data_validation('T' + str(i), {'validate': 'list',        
                                  'source': text.XLSSheetVTRDetTypeList})        

        ws.write('T' + str(i), '', align_center_format)

        # trip & pre trip

        ws.data_validation('U' + str(i), {'validate': 'decimal',
                                 'criteria': 'between',
                                 'minimum': -999999.999,
                                 'maximum': 999999.999})
        
        ws.data_validation('V' + str(i), {'validate': 'decimal',
                                 'criteria': 'between',
                                 'minimum': -999999.999,
                                 'maximum': 999999.999})
        
        ws.write('U' + str(i), '', align_center_format)
        ws.write('V' + str(i), '', align_center_format)

        # status opts
        ws.data_validation('W' + str(i), {'validate': 'list',
                                  'source': text.XLSSheetVTRStOptsList})
        
        ws.write('W' + str(i), stat_opts)
        
        # byp opts & comment        
        c_str = 'List of possible states:\n'        
        for q in xrange(0, len(text.XLSSheetVTRBypOptsList)):
            c_str += str(q+1) + ' - ' + text.XLSSheetVTRBypOptsList[q] + '\n'

        ws.write_comment('X' + str(i), c_str, {'x_scale': 2.42, 'y_scale': 1.96})

        ws.write('X' + str(i), strByp)


        # external bypass permit
        ws.data_validation('Y' + str(i), {'validate': 'list',
                                  'source': text.XLSSheetVTRBypYesNoList})
        
        ws.write('Y' + str(i), ext_byp_perm, align_center_format)

    wb.close()