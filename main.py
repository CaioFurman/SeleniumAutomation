import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# The URL of the website containing the embedded Excel sheets
url = 'https://equipezgt.blogspot.com/p/liga-zgt-tabela-de-classificacao_16.html'

# Setting up the Selenium driver
print('Acessando o site...')
options = Options()
options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get(url)
index = 1
dfs = []

# Finding the links of the spreadsheets
links = driver.find_elements("xpath", "//iframe[contains(@src, 'spreadsheets')]")

# Spliting each Excel sheet
for link in links:
    result = str(link.get_attribute("src"))
    sheet_url = result
    print(f'Lendo categoria {index}...')
    index += 1
    try:
        # Read the data from the Google Sheets document
        df_list = pd.read_html(sheet_url, header=0, encoding='utf-8')

        if len(df_list) > 0:
            # Select the first data frame
            df = df_list[0]

            # Append the DataFrame to the list of DataFrames
            dfs.append(df)

        else:
            print("Nenhuma tabela encontrada no documento.")

    except Exception as e:
        print("Erro ocorreu:", str(e))

# Concatenate all the DataFrames into a single DataFrame
final_df = pd.concat(dfs, ignore_index=True)

# Create a new workbook
workbook = pd.ExcelWriter('zgt_tabelas.xlsx')

# Write the data to the new workbook
final_df.to_excel(workbook, index=False)

# Save the workbook
workbook.save()
print("Todas as tabelas foram salvas corretamente!")

# Clean up the Selenium driver
driver.quit()
