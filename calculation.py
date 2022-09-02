from openpyxl import load_workbook
import xlwings as xw
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image


# input1 = st.number_input('Molar Mass of Monomer (g/mol)')
# input2 = st.number_input('Molar Mass of CTA (g/mol)')
# input3 = st.number_input('Molar Mass of Initiator (g/mol)')
# input4 = st.number_input('Initiator Ratio (to CTA)')
# input5 = st.number_input('Length of Polymer (# units)')
# input6 = st.number_input('Desired total mass of polymer (g)')
# input7 = st.number_input('Expected Conversion (%)')

def calculation(input1, input2, input3, input4, input5, input6, input7):
    workbook = load_workbook(filename='RAFT Calc.xlsx')
    sheet = workbook.active
    sheet["A2"] = float(input5)
    sheet["B2"] = float(input6)
    sheet["C2"] = float(input7)
    sheet["A5"] = float(input1)
    sheet["B5"] = float(input2)
    sheet["C5"] = float(input3)
    sheet["D5"] = float(input4)

    workbook.save(filename='output.xlsx')
    # Calculating
    ws = xw.Book("output.xlsx").sheets['Calculator']
    re = xw.Book("output.xlsx").sheets['please do not modify']

    # Selecting data from
    # a single cell
    results = {'Actual Mass of Monomer': [ws.range("B9").value, ws.range("B10").value],
               'Actual Mass of CTA': [ws.range("C9").value, ws.range("C10").value],
               'Actual Mass of Initiator': [ws.range("D9").value, ws.range("D10").value]}
    result = {'Molar Mass of Polymer (g/mol)': ws.range("A13").value}
    return (results, result)

