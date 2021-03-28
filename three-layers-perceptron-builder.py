""" ============================================================================
= LibreOffice Calc Three Layers Perceptron Builder version 1.0.0               =
= Copyrights (C) 2021 Velbazhd Software LLC                                    =
=                                                                              =
= developed by Todor Balabanov ( todor.balabanov@gmail.com )                   =
= Sofia, Bulgaria                                                              =
=                                                                              =
= This program is free software: you can redistribute it and/or modify         =
= it under the terms of the GNU General Public License as published by         =
= the Free Software Foundation, either version 3 of the License, or            =
= (at your option) any later version.                                          =
=                                                                              =
= This program is distributed in the hope that it will be useful,              =
= but WITHOUT ANY WARRANTY; without even the implied warranty of               =
= MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                =
= GNU General Public License for more details.                                 =
=                                                                              =
= You should have received a copy of the GNU General Public License            =
= along with this program. If not, see <http://www.gnu.org/licenses/>.         =
=                                                                              =
============================================================================ """

from __future__ import unicode_literals
from threading import Thread
from time import sleep


def ThreadWorker(desktop):
    sheet = XSCRIPTCONTEXT.getDesktop().getCurrentComponent().CurrentController.ActiveSheet
    
    input_size = int(sheet.getCellRangeByName("C7").getValue())
    hidden_size = int(sheet.getCellRangeByName("C8").getValue())
    output_size = int(sheet.getCellRangeByName("C9").getValue())
    example_step = max(max(input_size, hidden_size), output_size) + 2
    total_values = int(sheet.getCellRangeByName("C11").getValue())
    example_start = int(sheet.getCellRangeByName("C12").getValue())

    ''' Scale input. '''
    for t in range(1, total_values + 1):
        sheet.getCellRangeByName("E" + str(t)).setValue(sheet.getCellRangeByName("$C$4").getValue() + 
        (sheet.getCellRangeByName("$C$5").getValue() - sheet.getCellRangeByName("$C$4").getValue()) * ((sheet.getCellRangeByName("A" + str(t)).getValue() - 
        sheet.getCellRangeByName("$C$1").getValue()) / (sheet.getCellRangeByName("$C$2").getValue() - sheet.getCellRangeByName("$C$1").getValue())))
    
    ''' Example start index. '''
    x = 1
    
    for t in range(example_start, total_values - (input_size + output_size) + 1):
        ''' Report example number. '''
        sheet.getCellRangeByName("C12").setValue(t)
        
        ''' Setup biases. '''
        sheet.getCellRangeByName("G" + str(x)).setValue(1)
        sheet.getCellRangeByName("G" + str(x)).CellBackColor = (255 << 16 | 255 << 8 | 0)
        sheet.getCellRangeByName("H" + str(x)).setValue(1)
        sheet.getCellRangeByName("H" + str(x)).CellBackColor = (255 << 16 | 255 << 8 | 0)
        sheet.getCellRangeByName("I" + str(x)).setValue(1)
        sheet.getCellRangeByName("I" + str(x)).CellBackColor = (255 << 16 | 255 << 8 | 0)
        sheet.getCellRangeByName("J" + str(x)).setValue(1)
        sheet.getCellRangeByName("J" + str(x)).CellBackColor = (255 << 16 | 255 << 8 | 0)

        ''' Input data loading. '''
        for i in range(1, input_size + 1):
            sheet.getCellRangeByName("G" + str(x + i)).setValue(sheet.getCellRangeByName("E" + str(t + i)).getValue())
            sheet.getCellRangeByName("G" + str(x + i)).CellBackColor = (255 << 16 | 0 << 8 | 0)
    
        ''' Setup hidden layer. '''
        wih = 1
        for h in range(1, hidden_size + 1):
            sum = ""
            for i in range(0, input_size + 1):
                sum = sum + "G" + str(x + i) + "*Q" + str(wih)
                wih = wih + 1
                if i < input_size:
                    sum = sum + " + "
            sheet.getCellRangeByName("H" + str(x + h)).setFormula("=TANH( " + sum + " )")
            sheet.getCellRangeByName("H" + str(x + h)).CellBackColor = (0 << 16 | 0 << 8 | 255)
        
        ''' Setup output layer. '''
        who = 1
        for o in range(1, output_size + 1):
            sum = ""
            for h in range(0, hidden_size + 1):
                sum = sum + "H" + str(x + h) + "*S" + str(who)
                who = who + 1
                if h < hidden_size:
                    sum = sum + " + "
            sheet.getCellRangeByName("I" + str(x + o)).setFormula("=TANH( " + sum + " )")
            sheet.getCellRangeByName("I" + str(x + o)).CellBackColor = (0 << 16 | 255 << 8 | 0)
        
        ''' Expected data loading. '''
        for e in range(1, output_size + 1):
            sheet.getCellRangeByName("J" + str(x + e)).setValue(sheet.getCellRangeByName("E" + str(t + e + input_size)).getValue())
            sheet.getCellRangeByName("J" + str(x + e)).CellBackColor = (0 << 16 | 127 << 8 | 0)
        
        ''' Network output error. '''
        for r in range(1, output_size + 1):
            sheet.getCellRangeByName("K" + str(x + r)).setFormula("= (J" + str(x + r) + "-I" + str(x + r) + ") * (J" + str(x + r) + "-I" + str(x + r) + ")")
            sheet.getCellRangeByName("K" + str(x + r)).CellBackColor = (0 << 16 | 255 << 8 | 255)

        x = x + example_step
        sleep(1)
    
    ''' Network total error. '''
    sheet.getCellRangeByName("M1").setFormula("= SQRT( SUM(K:K) / COUNT(K:K) )")
    sheet.getCellRangeByName("M1").CellBackColor = 0
    
    ''' Setup biases for the operational copy of the network. '''
    sheet.getCellRangeByName("P1").setValue(1)
    sheet.getCellRangeByName("P1").CellBackColor = (255 << 16 | 255 << 8 | 0)
    sheet.getCellRangeByName("R1").setValue(1)
    sheet.getCellRangeByName("R1").CellBackColor = (255 << 16 | 255 << 8 | 0)
    sheet.getCellRangeByName("T1").setValue(1)
    sheet.getCellRangeByName("T1").CellBackColor = (255 << 16 | 255 << 8 | 0)

    ''' Input data loading for the operational copy of the network. '''
    for i in range(2, input_size + 2):
        sheet.getCellRangeByName("O" + str(i)).CellBackColor = (191 << 16 | 0 << 8 | 0)
        sheet.getCellRangeByName("P" + str(i)).setFormula("=$C$4 + ($C$5 - $C$4) * ((O" + str(i) + " - $C$1) / ($C$2 - $C$1))")
        sheet.getCellRangeByName("P" + str(i)).CellBackColor = (255 << 16 | 0 << 8 | 0)
    
    ''' Setup hidden layer for the operational copy of the network. '''
    wih = 1
    for h in range(2, hidden_size + 2):
        sum = ""
        for i in range(1, input_size + 2):
            sum = sum + "P" + str(i) + "*Q" + str(wih)
            sheet.getCellRangeByName("Q" + str(wih)).CellBackColor = (255 << 16 | 0 << 8 | 255)
            wih = wih + 1
            if i < input_size + 1:
                sum = sum + " + "
        sheet.getCellRangeByName("R" + str(h)).setFormula("=TANH( " + sum + " )")
        sheet.getCellRangeByName("R" + str(h)).CellBackColor = (0 << 16 | 0 << 8 | 255)
        
    ''' Setup output layer for the operational copy of the network. '''
    who = 1
    for o in range(2, output_size + 2):
        sum = ""
        for h in range(1, hidden_size + 2):
            sum = sum + "R" + str(h) + "*S" + str(who)
            sheet.getCellRangeByName("S" + str(who)).CellBackColor = (255 << 16 | 0 << 8 | 255)
            who = who + 1
            if h < hidden_size + 1:
                sum = sum + " + "
        sheet.getCellRangeByName("T" + str(o)).setFormula("=TANH( " + sum + " )")
        sheet.getCellRangeByName("T" + str(o)).CellBackColor = (0 << 16 | 255 << 8 | 0)
        sheet.getCellRangeByName("U" + str(o)).setFormula("=$C$1 + ($C$2 - $C$1) * ((T" + str(o) + " - $C$4) / ($C$5 - $C$4))")
        sheet.getCellRangeByName("U" + str(o)).CellBackColor = (0 << 16 | 127 << 8 | 0)
    
    return


def BuildAnnModel():
    thread = Thread(target=ThreadWorker, args=(XSCRIPTCONTEXT.getDesktop(),))
    thread.start()
