import openpyxl
from openpyxl.styles import Font, Color
import pandas as pd

df = pd.DataFrame()
df= pd.read_excel('/Users/jonathanoh/Desktop/python/Salary/position.xlsx', index_col=0)

df.loc[index][0]


