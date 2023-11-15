import pandas as pd
import numpy as np
from pathlib import Path
from urllib.request import urlretrieve
curDir = Path(__file__).parent

destDir = curDir / 'download'

if __name__=='__main__':
    print('Main')

    startYear = 2022 # BPA Data only goes back to 1996 at the earliest
    endYear = 2023

    flowgateInterties = {
        'Interties': [ 
              'AC'
             ,'BC'
             ,'DC'
             ,'ACDC'
             ]
        ,'Flowgates': [
             'Hemingway-SummerLake'
             ,'Midpoint-SummerLake' # The path was recorded as "Midpoint-SummerLake" before June 1st, 2010. Flip direction to aggregate with Hemingway-SummerLake
             ,'Idaho-PacificNW'
             ,'Reno-Alturas'
             ,'TotalWestSideLoad'
             ,'TotalWestSideLoadAndNorthofJohnDay' # TotalWestSideLoad was recorded as TotalWestSideLoadAndNorthofJohnDay from 1998 to May 1st 2021
             ,'PacePath20'
             ,'TOT.2'
             ,'WestOfBorahPath'
             ,'NorthOfEchoLake'
             ,'SouthOfCuster'
             ,'Monroe-EchoLake' # SouthOfCuster was recorded as Monroe-EchoLake from 2004 to February 1st, 2013
             ,'ColumbiaInjection'
             ,'JohnDayWind'
             ,'Montana-PacificNW'
             ,'NorthOfGrizzly'
             ,'NorthOfHanford'
             ,'Raver-Paul'
             ,'RockCreekWind'
             ,'SatsopInjection'
             ,'SouthOfAllston'
             ,'Allston-Keeler' # SouthOfAllston was recorded as Allston-Keeler before March 1st, 2007.
             ,'SouthOfBoundary'
             ,'WanapumInjection'
             ,'WestOfCascadesNorth'
             ,'WestOfCascadesSouth'
             ,'WestOfHatwai'
             ,'WestOfJohnDay'
             ,'WestOfLowerMonumental'
             ,'WestOfMcNary'
             ,'WestOfSlatt'
        ]
    }

    for flowgateIntertie in flowgateInterties:
        paths = flowgateInterties[flowgateIntertie]
        for path in paths:
            for year in range(startYear,endYear,1): 
                for month in range(1,13,1):
                    url = "https://transmission.bpa.gov/BUSINESS/Operations/Paths/{flowgateIntertie}/monthly/{path}/{year}/".format(**{'flowgateIntertie':flowgateIntertie, 'path': path, 'year':year, 'month':month})
                    filename = "{path}_{year}-{month:02d}.xls".format(**{'path':path, 'year':year, 'month':month})
                    filepath = destDir / filename

                    if(filepath.is_file()):
                        print(filename+" already exists. Skipping file.")
                        continue
                    try:
                        _, headers = urlretrieve(url + filename, str(filepath))
                        print("Downloading: "+filename)
                    except:
                        print("Did not find: "+filename)
                        try:
                            # Did not find XLS, try XLSX
                            _, headers = urlretrieve(url + filename +'x', str(filepath) + 'x')
                            print("Downloading: "+filename+'x')
                        except:
                            print("Did not find: "+filename+'x')

    print('Done')

