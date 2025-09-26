import camelot
import pandas as pd

# Step 1: Extract tables from the PDF using Camelot
tables = camelot.read_pdf(r"C:\Users\omart\Unilever\Codigos - Documents\Automatizaciones MT\Template\Tottus\Facturas pagadas a proveedor.PDF")

# Step 2: Concatenate all tables into a single DataFrame
df_all = pd.concat([table.df for table in tables], ignore_index=True)

# Step 3: Rename columns using the first row and drop the header row
df_all.columns = df_all.iloc[0]
df_all = df_all.drop(index=0).reset_index(drop=True)

# Step 4: Filter by values in column A
filter_values = ['Fact.Elect. Af. Emi','Ndd Afecta Elec. Rec','Ncr Ex Elect. Rec','Fac Afecta Elec. Rec',"Tipo Documento"]
filtered_df = df_all[df_all[df_all.columns[0]].isin(filter_values)]

# Step 5: Sort by custom order: VOUCHER first, then FPM, then others alphabetically
def sort_key(value):
    if value == 'Tipo Documento':
        return '0'
    elif value == 'Fac Afecta Elec. Rec':
        return '1'
    else:
        return '2' + value

filtered_df['sort_order'] = filtered_df[filtered_df.columns[0]].apply(sort_key)
filtered_df = filtered_df.sort_values(by='sort_order').drop(columns='sort_order')

# Step 6: Remove duplicate rows considering all columns
filtered_df = filtered_df.drop_duplicates()

# Step 7: Save the result to an Excel file
filtered_df.to_excel("CENCO_filtered_sorted_unique.xlsx", index=False)