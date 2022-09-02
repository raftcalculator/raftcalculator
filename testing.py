from openpyxl import load_workbook
import xlwings as xw
import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image





workbook = load_workbook(filename='output.xlsx')
sheet = workbook.active
sheet["A2"] = int(input("num1"))
sheet["B2"] = int(input("num2"))

workbook.save(filename='output.xlsx')
# Calculating
ws = xw.Book("output.xlsx").sheets['Sheet1']

# Selecting data from
# a single cell
v1 = ws.range("C2").value

print("Result:", v1)

############# Display #############

