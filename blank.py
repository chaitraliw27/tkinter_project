import pandas as pd

# Assuming df is your DataFrame and 'column_name' is the name of the column you're operating on
# Replace 'column_name' with the actual name of your column

# Define your function
def custom_function(s):
    if pd.isnull(s):
        return ''
    else:
        return ','.join(sorted(str(s).split(',')))

# Apply your function to the column
df['column_name'] = df['column_name'].apply(custom_function)

# Fill NaN values with blank
df['column_name'] = df['column_name'].fillna('')

# Now df['column_name'] will have blanks instead of NaN
