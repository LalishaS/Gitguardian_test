import requests
import json
import pandas as pd

excelFile='gitLang/RepoList.xlsx'
newSheet='Repos'
excel_data=pd.read_excel(excelFile)
dataArray=pd.DataFrame(excel_data)
print(dataArray)
data=dataArray['Repo Name'].array
print("Data from Excel:\n",type(data))

dataList={}
testData=['Lang1 - perc1\nLang2 - perc2\nLang3 - perc3']
languageData=[]
data=[(str(item)).strip() for item in data]
dataList['Repo Name']=data





for index in range(0,dataArray['Repo Name'].count()):
    repo_name=dataArray['Repo Name'][index]
    print("Row num: ",index," - Repo name : ",repo_name)

    
    git_url="https://api.github.com/repos/P-Olympus/"+str(repo_name)+"/languages"
    response_bytes = requests.get(git_url, auth=('tharindu-olympus', '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'))
    response_dict=json.loads(response_bytes.content.decode('utf-8'))
    dict_len=len(response_dict)

    # Sum of lines
    total_count=0
    #Get sum
    for lang in response_dict.values():
        total_count+=int(lang)
    print("Total: ",total_count)
    #Calc percentage
    langText=""
    for lang,count in response_dict.items():
        percentage=(int(count))/total_count*100
        print(lang," : ",percentage,"%")
        dataStr=str(lang)+" : "+str(percentage)+"\n"
        langText+=dataStr
    languageData.append(langText)


dataList['Languages']=languageData
dataFrame=pd.DataFrame(data=dataList)

writer=pd.ExcelWriter(excelFile,engine='xlsxwriter')
dataFrame.to_excel(writer,sheet_name=newSheet,index=False)

#Set cell formatting to show newlines
workBook=writer.book
workSheet=writer.sheets[newSheet]
cellFormatWrapped=workBook.add_format({'text_wrap':True})
cellFormatNonWrapped=workBook.add_format({'text_wrap':False})
workSheet.set_column('A:A',cell_format=cellFormatNonWrapped)
workSheet.set_column('B:B',cell_format=cellFormatWrapped)
writer.save()
writer.close()