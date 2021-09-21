    dict={
        "A":1,"B":2,"C":3,"D":4,"E":5,"F":8,"G":3,"H":5,"I":1,"J":1,"K":2,"L":3,"M":4,"N":5,"O":7,"P":8,"Q":1,"R":2,
        "S":3,"T":4,"U":6,"V":6,"W":6,"X":5,"Y":1,"Z":7,
    }
    if(name==None):
        return 0
    name=name.upper()
    word=name
    names=name.split()
    s=""
    total=0
    a=[0]*len(names)
    for i in range(len(names)):
        s=names[i]
        for j in s:
            if j in dict:
                a[i]=a[i]+dict[j]
                total+=int(dict[j])
    for i in range(len(a)):
        a[i]=convert(a[i])
    while(len(a)!=11):
        a.append(0)
    a.append(total)
    return a
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    wb=xl.load_workbook(r'C:\Users\ASus\PycharmProjects\Code\Stocks_sheet.xlsx')
    sheet1=wb['Sheet1']
    print(sheet1.max_row)  
    k=[0]*12
    x=0
    for i in range(2,sheet1.max_row+1):
        k=[]
        f=[]
        data=sheet1.cell(i,2)
        data1=sheet1.cell(i,1)
        print(data.value,end=" ")
        k.extend(Traverse(data.value))
        print(k,end=" ")
        if x==0:
            f.append(data1.value)
            f.append(data.value)
            f.extend(k)
            df=pd.DataFrame([f],columns=['Symbol','Description','Name1','Name2','Name3','Name4','Name5','Name6','Name7','Name8','Name9','Name10','Name11','Total Value'])
            x=1
        else:
            f.append(data1.value)
            f.append(data.value)
            f.extend(k)
            to_append = f
            a_series = pd.Series(to_append, index = df.columns)
            df = df.append(a_series, ignore_index=True)
        print("Done")
    file='Book2.xlsx'
    df.to_excel(file)