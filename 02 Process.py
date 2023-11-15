import pandas as pd
from pathlib import Path
import glob
curDir = Path(__file__).parent

dataDir = curDir / 'download'

dfList = []
for fileP in glob.glob(str(dataDir / '*.xls*')):
    print("Reading: "+fileP)
    xl = pd.ExcelFile(fileP)
    sheet_names = set(xl.sheet_names)
    valid_names = set(['Sheet1','Data'])
    sheets = list(sheet_names.intersection(valid_names))
    for sheet in sheets:
        path = str(Path(fileP).stem).split('_')[0]

        df = pd.read_excel(fileP, sheet_name=sheet, header=None, usecols=[0,1], names=['DateTime','Value'])
        df['DateTime'] = pd.to_datetime(df['DateTime'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
        df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
        df.dropna(subset=["DateTime", "Value"], inplace=True)
        df.insert(0, 'Path', path)
        dfList.append(df)

result = pd.concat(dfList, ignore_index=True)
result_r = result.copy()
result_r = result_r.drop_duplicates(subset=['Path', 'DateTime'], keep='first')
result_r['DateTime'] = result_r['DateTime'].dt.round('5min')
result_f = result_r[result_r['DateTime'].dt.minute.eq(0)] # keep only the rows that are hourly-values. Remove any where the minute is not equal to 0.

# Idaho - Pacific NW Path Data note from Zach Zornes:
#   *BPA path data (-) is E-W from 1/1/1998 to 3/31/2017. Polarity reversed in 4/1/2017; WECC definition (-) is W-E
#      BPA data set from 1/1/1998 to 3/31/2017 inverted and displayed consistent wth WECC definition
result_f.loc[(result_f['Path'] == 'Idaho-PacificNW') & (result_f['DateTime'] < '2017-04-01'), 'Value'] *= -1

# Hemingway-SummerLake was measured as Midpoint-SummerLake until June 1st, 2010.
# Invert direction of Midpoint-SummerLake and relabel as Hemingway-SummerLake
result_f.loc[(result_f['Path'] == 'Midpoint-SummerLake'), 'Value'] *= -1
result_f.loc[(result_f['Path'] == 'Midpoint-SummerLake'), 'Path'] = 'Hemingway-SummerLake'

# Per Chelsea, Montana-PacificNW should be measured in the East-To-West direction as (+), but BPA measures (+) as out of BPA.
# Flip sign convention to match WECC path-rating catalog. (+ is East to West) 
result_f.loc[(result_f['Path'] == 'Montana-PacificNW'), 'Value'] *= -1

# Per Zach Zornes charts, negative values in BPA's recordings means East to West.
# Flip sign to match WECC-case measurement direction (+ = East to West)
result_f.loc[(result_f['Path'] == 'WestOfHatwai'), 'Value'] *= -1

# TotalWestSideLoad was recorded as TotalWestSideLoadAndNorthofJohnDay from 1998 to May 1st 2021
# result_f.loc[(result_f['Path'] == 'TotalWestSideLoadAndNorthofJohnDay'), 'Path'] = 'TotalWestSideLoad'

# SouthOfCuster was recorded as Monroe-EchoLake from 2004 to February 1st, 2013
# result_f.loc[(result_f['Path'] == 'Monroe-EchoLake'), 'Path'] = 'SouthOfCuster'

# SouthOfAllston was recorded as Allston-Keeler before March 1st, 2007.
# result_f.loc[(result_f['Path'] == 'Allston-Keeler'), 'Path'] = 'SouthOfAllston'



result_f.to_csv('download_vertical.csv', index=False)

result_horizontal = result_f.pivot(index='DateTime', columns='Path', values='Value')
result_horizontal.to_csv('download_horizontal.csv')

print('done')


