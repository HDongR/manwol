

import os  # 경로 선택 등을 위한 모듈
import pandas as pd
import win32com.client as win32

#RESULT_DIR = 'C:\\Users\\fbtjd\\Downloads\\rst\\'
#CSV_PATH = 'C:\\Users\\fbtjd\\Downloads\\test.csv'
#HWP_PATH = 'C:\\Users\\fbtjd\\Downloads\\만월_동의서(param).hwp'

RESULT_DIR = ''
CSV_PATH = ''
HWP_PATH = ''

def main():
    # python 3에서는 print() 으로 사용합니다.
    print("Main Function") 
    print("check parameters.")
    print(os.getcwd()+"\\config.txt")
    f = open(os.getcwd()+"\\config.txt", 'r', encoding='UTF-8')
    while True:
        line = f.readline()
        if not line: break

        strings = line.split('=')
        if strings[0] == 'RESULT_DIR':
            print('RESULT_DIR OK', line)
            RESULT_DIR = strings[1].strip('\n')
        elif strings[0] == 'CSV_PATH':
            print('CSV_PATH OK', line)
            CSV_PATH = strings[1].strip('\n')
        elif strings[0] == 'HWP_PATH':
            print('HWP_PATH OK', line)    
            HWP_PATH = strings[1].strip('\n')
    f.close()

    data = pd.read_csv(CSV_PATH, encoding='UTF-8')
    
    dictions = {}
    
    colums = data.columns.values
    values = data.values;

    #print(data)
    #print(colums)

    for i in range(len(values)):
        d = values[i]
        d[9] = d[9].replace(" ","");
        key2 = d[9]
        if not d[9]:
            key2 = d[8]

        key = d[7] + '_' + key2

        if key in dictions:
            dList = dictions[key]
            dList.append(d)
            dictions[key] = dList
        else:
            dList = [d]
            dictions[key] = dList

    #print('result->', dictions)
    idx = 0
    for k in dictions.keys():
        
        if idx < 999999:
            dList = dictions[k]
            print(str(idx+1)+'/'+str(len(dictions.keys())) + ' ' + str(len(dList)) + 'row')

            hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            #4. 보안모듈 적용
            """ 보안모듈
            https://www.hancom.com/board/devdataView.do?board_seq=47&artcl_seq=4085&pageInfo.page=&search_text=
            HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules
            """
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

            hwp.Open(HWP_PATH, "HWP", "forceopen:true")
            #이름
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            option=hwp.HParameterSet.HFindReplace
            option.FindString = "param_name"
            option.ReplaceString = dList[0][7]
            option.IgnoreMessage = 1
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            #생년월일
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            option=hwp.HParameterSet.HFindReplace
            option.FindString = "param_birth"
            birthString = dList[0][9] #생년월일
            allNotBirthString = False

            dList[0][7] = dList[0][7].replace(" ", "")
            if dList[0][7] == '김동석':
                print('test')

            if not birthString:
                birthString = dList[0][8] #등록번호
                if not birthString:
                    allNotBirthString = True
            else: #생년월일 존재시 1990.04.30 년월일 사이에 점(.)붙이기
                if len(birthString) == 8:
                    year = birthString[0:4]
                    month = birthString[4:6]
                    day = birthString[6:8]
                    birthString = year + '.' + month + '.' + day

            option.ReplaceString = birthString
            option.IgnoreMessage = 1
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            
            #주소
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            option=hwp.HParameterSet.HFindReplace
            option.FindString = "param_addr"
            option.ReplaceString = dList[0][10]
            option.IgnoreMessage = 1
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
            
            allSameAddr = True #토지2개이상일 경우 모든 주소지가 같은지 체크
            preAddr = '';
            currAddr = '';

            for i in range(len(dList)):
                
                if i == 0:
                    preAddr = dList[i][10]
                else:
                    preAddr = dList[i-1][10]
                currAddr = dList[i][10]

                if preAddr != currAddr:
                    allSameAddr = False


                d = dList[i]
                #읍면
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = "param_em_" + str(i+1)
                option.ReplaceString = d[1]
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet) 
                
                #동리
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = "param_d_" + str(i+1)
                option.ReplaceString = d[2]
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

                #지번
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = "param_gibun_" + str(i+1)
                option.ReplaceString = d[4]
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

                #지목
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = "param_gimok_" + str(i+1)
                option.ReplaceString = d[5]
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

                #면적
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = "param_m_" + str(i+1)
                option.ReplaceString = d[6]
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)           

           


            find2Blank(hwp, len(dList))

            fileName = ''

            if not allSameAddr:
                filename = k + '_notSameAddr'
            else :
                filename = k

            if len(dList) > 7 :
                fileName = filename + '_upperRow.hwp'
            else :
                fileName = filename + '.hwp'
            hwp.SaveAs(os.path.join(RESULT_DIR, fileName))
            hwp.Quit()
            
        idx+=1

def find2Blank(hwp, len):
    maxRow = 7
    forLen = maxRow - len

    for i in range(0, forLen):
        #읍면
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "param_em_" + str(len+i+1)
        option.ReplaceString = ""
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet) 
        
        #동리
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "param_d_" + str(len+i+1)
        option.ReplaceString = ""
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

        #지번
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "param_gibun_" + str(len+i+1)
        option.ReplaceString = ""
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

        #지목
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "param_gimok_" + str(len+i+1)
        option.ReplaceString = ""
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

        #면적
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = "param_m_" + str(len+i+1)
        option.ReplaceString = ""
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

if __name__ == "__main__":
	main()