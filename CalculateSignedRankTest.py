import xlrd
filename = 'Algo_Test.xlsx'
workbook = xlrd.open_workbook(filename)
worksheet = workbook.sheet_by_index(0)
data = []
try:
    for row in range(worksheet.nrows):
        eachRow = []
        for col in range(1,worksheet.ncols):
            eachRow.append(worksheet.cell_value(row,col))
        data.append(eachRow)
except Exception as e: 
    print(e)


def calculate_diff(l):
    '''Calculates differences for each pair data and adds in respective sublist'''
    line_count = 0
    try:
        for row in l:
            if line_count > 0:
                diff = row[0] - row[1]
                row.append(diff)
            line_count += 1
    except Exception as e:
        print(e)
    return l
    
def calculate_abs(l):
    '''Calculates absolute value of calculated differences
    for each pair data and updates difference value in respective sublist.
    Also adds new element to each sublist - sign'''
    line_count = 0
    try:
        for row in l:
            if line_count > 0:
                sign = 1 if row[2] > 0 else -1
                row.append(sign)
                row[2] = abs(row[2])
            line_count += 1
    except Exception as e: 
        print(e)
    return l

def calculate_rank(l): 
    '''
        Calculates rank for each calculated absolute value
    '''
    try:
        T = [each[2] for each in l[1:]]
        R = [0 for i in range(len(T))] 
        Temp = [(T[i], i) for i in range(len(T))] 

        Temp.sort(key=lambda x: x[0]) 
        (rank, n, i) = (1, 1, 0) 
        while i < len(T): 
            j = i 
            while j < len(T) - 1 and Temp[j][0] == Temp[j + 1][0]: 
                j += 1
            n = j - i + 1
            for j in range(n): 
                index = Temp[i+j][1] 
                R[index] = rank + (n - 1) * 0.5
                l[index+1].append((rank + (n - 1) * 0.5) * l[index+1][3])
            rank += n 
            i += n
    except Exception as e: 
        print(e)
    return l 

def calculate_testStats(l):
    try:
        positiveRanks = [each[4] for each in l[1:] if each[4] > 0]
        negativeRanks = [each[4] for each in l[1:] if each[4] < 0]
    except Exception as e: 
        print(e)
    return min(abs(sum(positiveRanks)),abs(sum(negativeRanks)))
    
def printResult(result):
    for eachCustomer in result:
        print("Signed Test Rank for ", eachCustomer[0], ' and ', eachCustomer[1], ' = ', eachCustomer[2])
        
def main():
    global data
    result = []
    try:
        for j in range(2,len(data[0])+1):
            l = [[row[0],row[i]] for row in data for i in range(j-1,j)]
            l = calculate_diff(l)
            '''Removing values with 0 differences'''
            m = list(filter(lambda each: each[2] != 0, l[1:]))
            m.insert(0,l[0])
            l = calculate_abs(m)
            l = calculate_rank(l)
            res = calculate_testStats(l)
            l[:1][0].append(res)
            result.append(l[:1][0])
    except Exception as e: 
        print(e)
    else:
        printResult(result)
  
if __name__ == "__main__": 
    main()
