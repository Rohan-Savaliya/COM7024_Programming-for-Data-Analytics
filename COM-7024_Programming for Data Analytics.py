import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel("C:\Users\rksav\OneDrive\Desktop\assignment\Programming for data analytics\Sales Data_PDA_4052.xlsx")

print("Shape of dataset:", df.shape)
print("\nColumn names:")
print(df.columns.tolist())

print("\nFirst 5 rows:")
display(df.head())

print("\nDataset info:")
df.info()

print("\nMissing values in each column:")
print(df.isnull().sum())

# Set first row as header
df.columns = df.iloc[0]

# Remove the first row (since it's now header)
df = df[1:]

# Reset index
df.reset_index(drop=True, inplace=True)

# Show result
print("Fixed columns:")
print(df.columns)

print("\nFirst 5 rows after fixing:")
display(df.head())

# Convert date column
df['date'] = pd.to_datetime(df['date'])

# Convert value column to numeric
df['value_£'] = pd.to_numeric(df['value_£'])

# Check updated data types
print("Updated data types:")
print(df.dtypes)

print("\nDataset info after conversion:")
df.info()

# Check duplicate rows
duplicates = df.duplicated().sum()
print("Number of duplicate rows:", duplicates)

# Remove duplicates if any
df = df.drop_duplicates()

print("Shape after removing duplicates:", df.shape)

# Descriptive statistics for numerical column
print("Statistical Summary:")
display(df['value_£'].describe())

# Total sales per salesperson
sales_by_person = df.groupby('sales_person')['value_£'].sum().sort_values(ascending=False)

print("Total Sales by Salesperson:")
display(sales_by_person)

# Plot
sales_by_person.plot(kind='bar', figsize=(8,5))
plt.title("Total Sales by Salesperson")
plt.xlabel("Salesperson")
plt.ylabel("Total Sales (£)")
plt.xticks(rotation=45)
plt.show()

# Average sales by priority
priority_sales = df.groupby('priority')['value_£'].mean().sort_values(ascending=False)

print("Average Sales by Priority:")
display(priority_sales)


plt.figure(figsize=(8,5))
sns.boxplot(x='priority', y='value_£', data=df)
plt.title("Sales Distribution by Priority")
plt.xlabel("Priority")
plt.ylabel("Sales (£)")
plt.show()

# Correlation matrix
corr = df[['value_£']].corr()

print("Correlation Matrix:")
display(corr)

# Convert priority to numeric
priority_mapping = {
    'Low': 1,
    'Not Specified': 2,
    'Medium': 3,
    'High': 4,
    'Critical': 5
}

df['priority_encoded'] = df['priority'].map(priority_mapping)

# Correlation
corr = df[['value_£', 'priority_encoded']].corr()

print("Correlation between Priority and Sales:")
display(corr)

# Extract month
df['month'] = df['date'].dt.to_period('M')

# Monthly sales
monthly_sales = df.groupby('month')['value_£'].sum()

print("Monthly Sales:")
display(monthly_sales)

# Plot
monthly_sales.plot(figsize=(10,5))
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales (£)")
plt.xticks(rotation=45)
plt.show()

