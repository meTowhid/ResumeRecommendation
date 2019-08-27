import pandas as pd

def processFile(filename):
    # filename = 'IT+Engineer_2.xlsx'
    df = pd.read_excel(filename)

    df.rename(columns={'Name': 'raw', 'Career Summary': 'last_company', 'Experience & Application Status': 'experience', 'Unnamed: 3': 'salary'}, inplace=True)

    df['raw'][0].split('\n')

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
    df.head()

    df.age = df.age.str.extract('(\d+\.?\d*)',expand=True).astype(float)
    df.last_company = df.last_company.str.split('\n').str.get(0)
    df.experience = df.experience.str.extract('(\d+)',expand=True).astype(int)
    df.salary = df.salary.str.extract('(\d+\,?\d*)',expand=True).get(0).str.replace(',','').astype(int)
    df.matching_rate = df.matching_rate.str.extract('(\d+)',expand=True).astype(int)
    df.head()

    df.to_excel(filename.replace('.xlsx','_out.xlsx'), index=False, columns=headers)


def processAll():
    for f in ['IT/IT+Engineer_%i.xlsx'% i for i in range(1,18)]:
        processFile(f)

def mergeFiles():
    df = pd.DataFrame()
    for f in ['IT/IT+Engineer_%i_out.xlsx'% i for i in range(1,18)]:
        df = df.append(pd.read_excel(f), ignore_index=True)
    
    df.to_excel('dataset.xlsx', index=False)

processFile('HR-Executive.xlsx') 