import pandas as pd
import plotly.express as px
import os

#Path to Excel file
#This will need to be changed if using new data
file_path = r"C:\Users\eollomani\Documents\Primary BI  work.xlsx" #path to excel sheet
sheet_name = 'Table4'  #The specific sheet name you want to read


# Dictionary for full state names
state_full_name = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'CA': 'California',
    'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware', 'FL': 'Florida',
    'GA': 'Georgia', 'HI': 'Hawaii', 'ID': 'Idaho', 'IL': 'Illinois',
    'IN': 'Indiana', 'IA': 'Iowa', 'KS': 'Kansas', 'KY': 'Kentucky',
    'LA': 'Louisiana', 'ME': 'Maine', 'MD': 'Maryland', 'MA': 'Massachusetts',
    'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi', 'MO': 'Missouri',
    'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada', 'NH': 'New Hampshire',
    'NJ': 'New Jersey', 'NM': 'New Mexico', 'NY': 'New York',
    'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio', 'OK': 'Oklahoma',
    'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina',
    'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah',
    'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington', 'WV': 'West Virginia',
    'WI': 'Wisconsin', 'WY': 'Wyoming'
}


#Check if the sheet is being read correctly (use if running into errors)

try:
    #Check if the file exists
    if os.path.exists(file_path):
        #Load the specific sheet from the Excel file
        dataset = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        print("File loaded successfully")
        #print(dataset.head())  #Print the first few rows to ensure it's loaded correctly
    else:
        print(f"The file {file_path} does not exist.")
except PermissionError as e:
    print(f"PermissionError: {e}")
except Exception as e:
    print(f"An error occurred: {e}")

if 'dataset' in locals():
    dataset = pd.DataFrame({
        'Vendors': dataset['Name'],
        'States': dataset['States'],
        'Coverage': dataset['Coverage']
    })

    ds = dataset
    ds['Coverage'] = pd.to_numeric(ds['Coverage'])
    ds = ds.assign(States=ds['States'].str.split(', ')).explode('States').reset_index(drop=True)

    ds['state_full_name'] = ds['States'].map(state_full_name)
    ds = ds.sort_values(by=['States', 'Coverage'], ascending=[True, False])

    #use this if you have data for each state. 
    #ds['hover_text'] = ds.apply(lambda row: f"{row['Vendors']}: {row['Coverage']*100:.2f}%", axis=1)

    #use this if you do not have percentages for each state. This will only show the vendors for each state
    ds['hover_text'] = ds.apply(lambda row: f"{row['Vendors']}", axis=1)

    hover_data = ds.groupby('States')['hover_text'].apply('<br>'.join).reset_index()
    coverage_data = ds.groupby('States')['Coverage'].count().reset_index()

    ds_merge = pd.merge(coverage_data, hover_data, on='States')
    ds_merge['Coverage'] = ds_merge['Coverage'].round(2)
    ds_merge['state_full_name'] = ds_merge['States'].map(state_full_name)

    #map
    figure = px.choropleth(ds_merge,
                           locations='States',
                           locationmode='USA-states',
                           color='Coverage',
                           hover_name='state_full_name',
                           hover_data={'Coverage': False, 'hover_text': True},
                           scope='usa',
                           labels={'Coverage': 'Count of Primary Vendors'},
                           title='Vendors in the US',
                           color_continuous_scale=px.colors.sequential.Blues
    )

    figure.update_traces(hovertemplate='<b><span style="font-size: 16px;">%{customdata[0]}</span></b><br>%{customdata[1]}',
                         customdata=ds_merge[['state_full_name', 'hover_text']].values.tolist())

    
    #The title
    figure.update_layout(geo=dict(bgcolor='rgba(0,0,0,0)'), title={
        'text': 'Vendors in the US',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 40}
    })

    # Save the figure as an HTML file
    output_file = r'C:\Users\eollomani\Documents\vendors_in_us_by_state.html'
    figure.write_html(output_file)

    print(f"Figure saved to {output_file}")

    # Optionally, display the figure in a web browser
    import webbrowser
    webbrowser.open(output_file)