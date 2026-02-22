
import pandas as pd

# Read first 10 rows without header to inspect structure
df = pd.read_excel('availability.xlsx', header=None, nrows=10)
print("Data Structure:")
print(df.to_string())

# Also print columns if we assume header=0
df_header = pd.read_excel('availability.xlsx', nrows=0)
print("\nColumns:")
print(list(df_header.columns))
