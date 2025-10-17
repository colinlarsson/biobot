
import pandas as pd

# Step 1: Set your file path
file_path = "EllvaS082125DO.csv"  # Replace with your CSV file name

# Step 2: Load the dataset
df = pd.read_csv(file_path)

# Step 3: Forward fill missing values
df_filled = df.fillna(method='ffill')

# Step 4: Downsample to every Nth row (change 25 to your desired interval)
df_downsampled = df_filled.iloc[::400, :]

# Step 5: Save the result to Excel
df_downsampled.to_excel("EllvaS082125DO_ffill_downsampled_script400.xlsx", index=False)

print("Data cleaned and downsampled successfully.")
