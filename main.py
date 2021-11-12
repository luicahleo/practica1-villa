# -*- coding: utf-8 -*-
"""
Created on Fri Nov  5 10:29:38 2021

@author: Gabriel
"""

from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

name = 'practica1.xlsx'
excel_doc = openpyxl.load_workbook(name, data_only=True)
sheet = excel_doc['Hoja1']

a = Read_Excel_to_List(sheet, 'B2', 'B7')
b = Read_Excel_to_List(sheet, 'D2', 'D8')
Fabricas = Read_Excel_to_List(sheet, 'A2', 'A7')
Almacenes = Read_Excel_to_List(sheet, 'C2', 'C8')
c = Read_Excel_to_NesteDic(sheet, 'F1', 'J5')


def ejemplito():
    solver = pywraplp.Solver.CreateSolver('GLOP')

    x = {}
    rfab = {}
    ralm = {}

    for i in Fabricas:
        x[i] = {}
        for j in Almacenes:
            # para deshabilitar las 'x', para saltarnos estas 'x'
            if c[i][j] != 'x':
                x[i][j] = solver.NumVar(0, solver.infinity(), 'X%d;%d' % (i, j))
    print('Número de variables=', solver.NumVariables())


    for i in Fabricas:
        rfab[i] = solver.Add(sum(x[i][j] for j in Almacenes if c[i][j] != 'x') == a[i - 1], 'RF%d' % (i))

    for j in Almacenes:
        ralm[j] = solver.Add(sum(x[i][j] for i in Fabricas if c[i][j] != 'x') == b[j - 1], 'RA%d' % (j))

    print('Número de restricciones=', solver.NumConstraints())

    solver.Minimize(solver.Sum(c[i][j] * x[i][j] for i in Fabricas for j in Almacenes if c[i][j]!='x'))

    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        for i in Fabricas:
            for j in Almacenes:
                if c[i][j] != 'x':
                    print('X%d;%d = %d' %
                      (i, j, x[i][j].solution_value()))
        for i in Fabricas:
            for j in Almacenes:
                if c[i][j] != 'x':
                    print('CR%d;%d = %d' %
                      (i, j, x[i][j].ReducedCost()))
        for i in Fabricas:
            print('u%d=%d' %
                  (i, rfab[i].dual_value()))
        for j in Almacenes:
            print('v%d=%d' %
                  (j, ralm[j].dual_value()))
        print('Función objetivo =', solver.Objective().Value())
    else:
        print('El problema es inadmisible')

    Solu = {}
    for i in Almacenes:
        Solu[i] = {j: 0.0 for j in Almacenes}

    for i in Fabricas:
        for j in Almacenes:
            if c[i][j] != 'x':
                Solu[i][j] = c[i][j]
            else:
                Solu[i][j] = x[i][j].solution_value()

    Write_NesteDic_to_Excel(excel_doc, name, sheet, Solu, 'F10', 'M16')


ejemplito()







