import pandas as pd
from os import listdir
from os.path import isfile, join

def getDataFrameFromFile(filename):
    df = pd.read_html(filename)
    df = df[1]
    df.columns = df.iloc[0]
    df.drop(df.index[0], inplace=True)
    return df

def processDataFrame(df):
    df.rename(columns={'Name': 'raw', 'Career Summary': 'last_company', 'Experience & Application Status': 'experience'}, inplace=True)

    d = df.raw.str.split('\n')
    df['name'] = d.str.get(0)
    df['age'] = d.str.get(1).replace('Age:','')
    df['versity'] = d.str.get(3)
    df['subject'] = d.str.get(4)
    df['matching_rate'] = d.str.get(5) 
    # df['phone'] = d.str.get(6) 

    d = df.experience.str.split('\n')
    df['salary'] = d.str.get(1)
    df['experience'] = d.str.get(0)
    df.head()


    headers = ['name', 
    #              'phone', 
                'age', 
                'versity',
                'subject',
                'last_company',
                'experience',
                'salary', 
                'matching_rate']
    df.drop(columns='raw', inplace=True)
    df = df.reindex(headers, axis=1)

    df.age = df.age.str.extract('(\d+\.?\d*)',expand=True).astype(float)
    df.last_company = df.last_company.str.split('\n').str.get(0)
    df.experience = df.experience.str.extract('(\d+)',expand=True).astype(int)
    df.salary = df.salary.str.extract('(\d+\,?\d*)',expand=True).get(0).str.replace(',','').astype(int)
    df.matching_rate = df.matching_rate.str.extract('(\d+)',expand=True).astype(int)
    
    return df

def getFilePaths(folder):
    f = [f for f in listdir(folder) if isfile(join(folder, f)) and f.endswith('.xls')]
    f.sort()
    print(f)
    return [folder + '/' + file for file in f]

def mergeFiles(folder):
    df = pd.DataFrame()
    for file in getFilePaths(folder):
        df = df.append(getDataFrameFromFile(file), ignore_index=True)
    df = processDataFrame(df)
    df.to_excel('dataset_%s.xlsx' % folder, index=False)



# mergeFiles('am')
processDataFrame(getDataFrameFromFile('am/Accounts+Manager_01.xls'))