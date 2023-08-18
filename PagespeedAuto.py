import openpyxl
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def load_from_storage():
    try:
        with open('api_key.txt') as api_key:
            token = api_key.read().strip()
            return token
    except FileNotFoundError:
        print('No API key found, continuing without..')
        return None

token = load_from_storage()

#with open('pagespeed.txt') as pagespeedurls:
#    urls = [line.strip() for line in pagespeedurls]

urls = [
    "https://community.infineon.com",
    "https://community.nxp.com",
    "https://community.st.com",
    "https://e2e.ti.com",
    "https://forum.microchip.com",
    "https://devzone.nordicsemi.com",
    "https://stackoverflow.com",
    "https://community.cisco.com/",
    "https://community.atlassian.com/",
    "https://www.tableau.com/community"
]

data_list = []  # List to store data as dictionaries

for url in urls:
    if token is not None:
        api_url = f'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={url}&strategy=desktop&key={token}'
    else:
        api_url = f'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={url}&strategy=desktop'

    print(f'Requesting {api_url}...')
    response = requests.get(api_url)
    data = response.json()
    print(data)


    try:


        urlid = data['id'].split('?')[0]

        # Metrics from Lighthouse

        LCP = float(data["loadingExperience"]["metrics"]["LARGEST_CONTENTFUL_PAINT_MS"]["percentile"])/1000
        FID = float(data["loadingExperience"]["metrics"]["FIRST_INPUT_DELAY_MS"]["percentile"])
        CLS = float(data["loadingExperience"]["metrics"]["CUMULATIVE_LAYOUT_SHIFT_SCORE"]["percentile"])/100
        FCP = float(data["loadingExperience"]["metrics"]["FIRST_CONTENTFUL_PAINT_MS"]["percentile"])/1000
        INP = float(data['loadingExperience']['metrics']['INTERACTION_TO_NEXT_PAINT']['percentile'])
        TTI = data["lighthouseResult"]["audits"]["interactive"]["displayValue"]
        TBT = data["lighthouseResult"]["audits"]["total-blocking-time"]["displayValue"]
        TTFB = float(data["loadingExperience"]["metrics"]["EXPERIMENTAL_TIME_TO_FIRST_BYTE"]["percentile"])/1000
        performance_score = float(data['lighthouseResult']['categories']['performance']['score'])*100

        row = {
            'URL': urlid,
            'Largest Contentful Paint (LCP)': f'{LCP:.2f} s',
            'First Input Delay': f'{FID:.2f} ms',
            'Cumulative Layout Shift (CLS)': f'{CLS:.2f}',
            'First Contentful Paint  (FCP)': f'{FCP:.2f} s',
            'Interaction to Next Paint (INP)' : f'{INP:.2f} ms',
            'Time To Interactive (TTI)': TTI,
            'Total Blocking Time (TBT)': TBT,
            'Time to first Byte (TTFB)' : f'{TTFB:.2f} s',
            'Performance':  f'{performance_score:.2f} %'

        }

        data_list.append(row)

        print('row:', row)
        print(f'Performance: {performance_score}')
    except KeyError as e:
        print(f'<KeyError> One or more keys not found {url}. Error: {e}')
    except Exception as e:
        print(f'<Error> Failed to process {url}. Error: {e}')

# Create the DataFrame after the loop is completed
results = pd.DataFrame(data_list)

df = pd.DataFrame(data_list)
#results_transposed = results.transpose()
#results_transposed.to_csv('pagespeed-results-swapped.csv', header=False)

results.reset_index(drop=True, inplace=True)
benchmark_values = ['Benchmark', '< 2.5 s', '< 100 ms', '< 0.1', '< 1.8 s', '< 200 ms', '< 5 s', '< 200 ms','< 0.8 s','> = 75%']
results = pd.concat([pd.DataFrame([benchmark_values], columns=results.columns), results])
results_transposed = results.transpose()
print(results_transposed)
results_transposed.to_excel('Pagespeed_CommunityPerformance.xlsx')

excel_file_path = 'Pagespeed_CommunityPerformance.xlsx'
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active
worksheet.delete_rows(0)
modified_excel_file_path = 'Pagespeed_CommunityPerformance.xlsx'
workbook.save(modified_excel_file_path)#