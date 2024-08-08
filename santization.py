def data_sanitization(sample_data):
    newdata=[]
    i=0
    while(i<len(sample_data)-1):
        temp=[]
        if(sample_data[i][1]==sample_data[i+1][1] and sample_data[i][5]==sample_data[i+1][5]):
            for j in range(3):
                temp.append(sample_data[i][j])

            token=[]
            distance=[]
            date=[]
            token.append(sample_data[i][3])
            token.append(sample_data[i+1][3])
            distance.append(sample_data[i][4])
            distance.append(sample_data[i+1][4])
            date.append(sample_data[i][5])
            date.append(sample_data[i+1][5])
            temp.append(token)
            temp.append(distance)
            temp.append(date)
            newdata.append(temp)
            i+=1
        else:
            newdata.append(sample_data[i])
        i+=1
    return newdata



