import pandas as pd

data = {
    "Name": ["Ali", "Sara", "Omar", "Lena", "Jasper"],
    "Math": [88, 72, 95, 61, 78],
    "Science": [76, 85, 90, 55, 82],
    "English": [91, 68, 74, 80, 70],
}

df = pd.DataFrame(data)

df['Average'] = df[['Math','Science','English']].mean(axis=1).round=(1)
df['Grade'] = df["Average"].apply(lambda x: 'A' if x>=80 else 'B' if x>= 70 else 'C') 

df.to_excel('Grades.xlsx')
print('file created!')