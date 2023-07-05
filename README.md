df1['Address'].str.extract(f'({"|".join(df2["Address"].tolist())})', flags=re.IGNORECASE)
