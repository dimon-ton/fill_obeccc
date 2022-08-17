import pandas as pd

df = pd.read_excel('fill_obeccc.xlsx')

# email = df.get('e-mail')
# password = df.get('password')
print(df.password[0])


