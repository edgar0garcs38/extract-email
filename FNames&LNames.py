import pandas as pd
import glob

# take first and last name from merged file
fNamelName = pd.read_excel('Clientes Sin Duplicar.xlsx', sheet_name='Sheet1')
# Step 1: Split at the first comma

split_names = fNamelName['Asegurado'].str.split(' ', n=2, expand=True)

# Step 2: Assign and convert to uppercase
fNamelName['Last Name'] = split_names[1].str.upper()
fNamelName['First Name'] = split_names[2].str.upper()

fNamelName.to_excel('CLIENTES CANCELADOS OCTUBRE 13 FN LN.xlsx', index=False)