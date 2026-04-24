import pandas as pd

# Raw 15-min counts
data = {
    "Approach":   ["Northbound", "Southbound", "Eastbound", "Westbound"],
    "7:00-7:15":  [312, 289, 450, 198],
    "7:15-7:30":  [387, 310, 523, 224],
    "7:30-7:45":  [401, 298, 498, 241],
    "7:45-8:00":  [356, 276, 481, 213],
}

df = pd.DataFrame(data)

#four 15-min columns
count_col = [ "7:00-7:15", "7:15-7:30", "7:30-7:45" ,"7:45-8:00"]

#Hourly volume = sum of all four intervals
df['Hourly Volume'] = df[count_col].sum(axis=1)

#Peak 15-min 
df['Peak 15-min'] = df[count_col].max(axis=1)

#PHF formula
df['PHF'] = (df['Hourly Volume'] / (4 * df['Peak 15-min'])).round(3)

#Marking if the phf is high
df['Status'] = df['PHF'].apply(lambda x: "OK" if x>=0.85 else "Peaky")

print(df)

df.to_excel('phf_raw.xlsx', index = False, engine = 'openpyxl')