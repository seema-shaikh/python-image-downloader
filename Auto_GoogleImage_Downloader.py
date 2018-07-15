import os.path
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from tkinter import Tk
from tkFileDialog import askopenfilename
from os import path
from Tkinter import *
import Tkinter, tkFileDialog, Tkconstants
import time
import sys
import urllib2

print '\n'
print '-------------GOOGLE IMAGE DOWNLOADER----------------'
print '\n'
print '-------------Choose the excel file------------------'
Tk().withdraw()
filename = askopenfilename()
print "you have chosen "
print(filename)
#print(os.path.splitext("filename")[0])

print '\n'
print '-------------Now choose the ROOT folder-------------'
root = Tkinter.Tk()
root.withdraw()
dirName = tkFileDialog.askdirectory(parent=root, title='PLEASE SELECT A ROOT DIRECTORY TO SAVE SEARCH TERM FOLDERS')
saving_path = os.path.abspath(dirName)
print 'Your ROOT folder is: '
print(saving_path)

df = pd.read_excel(filename, sheetname='Sheet1')

print '-----------------------------------------------------'

search_term = []
search_term_id = []
image_size = []
number_of_images = []

readSelectedExcel = input('Want to run only specific rows? Please type Y or N:   ')

# start=0
# stop=0


idx = 0  # first row

# check if a column already exists

if 'Error Count' not in df.columns:
    df.insert(loc=idx, column='Error Count', value=" ",
              allow_duplicates=False)  # check if errorCount column has been created or not

#print(str(df))

if readSelectedExcel == Y:
    stt = input('Please enter starting row number')
    stp = input('Please enter stopping row number')
    start = stt - 2
    # stop=stp-1

    selected = df.iloc[stt - 2:stp - 1]
    search_term = selected['SEARCH TERM'].tolist()
    search_term_id = selected['SEARCH TERM ID'].tolist()
    image_size = selected['IMAGE SIZE'].tolist()
    number_of_images = selected['NUMBER OF IMAGES'].tolist()

    '''search_term = df['SEARCH TERM'].tolist()
    search_term_id=df['SEARCH TERM ID'].tolist()
    image_size = df['IMAGE SIZE'].tolist()
    number_of_images=df['NUMBER OF IMAGES'].tolist()'''

# print str(df)
# print '---------------------------'
# print str(df.iloc[0:5])
# first five rows hi execute karega

if readSelectedExcel == N:
    # selected=df.iloc[start:stop]
    start=0 
    # stop=len(search_term)-1
    search_term = df['SEARCH TERM'].tolist()
    search_term_id = df['SEARCH TERM ID'].tolist()
    image_size = df['IMAGE SIZE'].tolist()
    number_of_images = df['NUMBER OF IMAGES'].tolist()


def download_page(url):
    version = (3, 0)
    cur_version = sys.version_info
    if cur_version >= version:
        import urllib.request
        try:
            headers = {}
            headers[
                'User-Agent'] = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36"
            req = urllib.request.Request(url, headers=headers)
            resp = urllib.request.urlopen(req)
            respData = str(resp.read())
            return respData
        except Exception as e:
            print(str(e))
    else:
        import urllib2
        try:
            headers = {}
            headers[
                'User-Agent'] = "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.17 (KHTML, like Gecko) Chrome/24.0.1312.27 Safari/537.17"
            req = urllib2.Request(url, headers=headers)
            response = urllib2.urlopen(req)
            page = response.read()
            return page
        except:
            return "Page Not found"


def _images_get_next_item(s):
    start_line = s.find('rg_di')
    if start_line == -1:
        end_quote = 0
        link = "no_links"
        return link, end_quote
    else:
        start_line = s.find('"class="rg_meta"')
        start_content = s.find('"ou"', start_line + 1)
        end_content = s.find(',"ow"', start_content + 1)
        content_raw = str(s[start_content + 6:end_content - 1])
        return content_raw, end_content


def _images_get_all_items(page):
    items = []
    while True:
        item, end_content = _images_get_next_item(page)
        if item == "no_links":
            break
        else:
            items.append(item)
            # time.sleep(0.1)
            page = page[end_content:]
    return items


#myError = str(saving_path) + os.sep + 'errorlog.txt'

#t0 = time.time()

i = 0
# for i in range(start,stop+1):
try:
    
    while i < len(search_term):
        
        items = []
        # agar 4 daala starting value toh i ka value is 2---i.e (minus 2)
        iteration = "Item no.: " + str(i + 1) + " -->" + " Item name = " + str(search_term[i])
        print '\n'
        print (iteration)
        print ("Please wait for a while...")

        search_terms = search_term[i]
        search = search_terms.replace(' ', '%20')
        sid = search_term_id[i]
        try:
            myDir = str(saving_path) + os.sep + str(sid)
            os.makedirs(myDir)
        except OSError, e:
            if e.errno != 17:
                raise
            pass

        url = ''
        if image_size[i] == "Medium":
            url = 'https://www.google.co.in/search?q=' + search + '&tbm=isch&tbs=isz:m'
        if image_size[i] == "Large":
            url = 'https://www.google.co.in/search?q=' + search + '&tbm=isch&tbs=isz:l'

        if image_size[i] == "Any Size":
            url = 'https://www.google.co.in/search?q=' + search + '&tbm=isch'
        j = 0
        while j < len(image_size):
            '''url=''
            if image_size[j]=="large":
                url='https://www.google.co.in/search?q='+search+'&authuser=1&tbs=isz:m&tbm=isch&source=lnt&sa=X&ved=0ahUKEwiLhfuC-YbXAhUE2hoKHYavCa0QpwUIHQ&biw=1536&bih=760&dpr=1.25'
            if image_size[j]=="medium":
                url = 'https://www.google.co.in/search?q='+search+'&authuser=1&tbm=isch&source=lnt&tbs=isz:l&sa=X&ved=0ahUKEwi5lZf9-IbXAhXDthoKHR-ZDqsQpwUIHg&biw=1536&bih=760&dpr=1.25'
            if image_size[j]=="any size":
                url='https://www.google.co.in/search?q='+search+'&espv=2&biw=1366&bih=667&site=webhp&source=lnms&tbm=isch&sa=X&ei=XosDVaCXD8TasATItgE&ved=0CAcQ_AUoAg'
            '''
            raw_html = (download_page(url))
            # time.sleep(0.1)
            items = items + (_images_get_all_items(raw_html))
            j = j + 1
        l = 0
        while l < len(number_of_images):
            num = number_of_images[i]
            del items[num:]
            l = l + 1
        print ("Total Image Links = " + str(len(items)))

        #t1 = time.time()
        #total_time = t1 - t0
        #print("Total time taken: " + str(total_time) + " Seconds")
        print ("Downloading...")

        k = 0
        errorCount = 0

        #errornumber = open(myError, 'a')
        while (k < len(items)):
            from urllib2 import Request, urlopen
            from urllib2 import URLError, HTTPError

            try:
                req = Request(items[k], headers={
                    "User-Agent": "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.17 (KHTML, like Gecko) Chrome/24.0.1312.27 Safari/537.17"})
                response = urlopen(req, None, 15)

                # str(sid)+'_00'


                if (k >= 9 and k < 99):
                    
                    output_file = open(myDir + "/" + str(sid) + '_0' + str(k + 1) + ".jpg", 'wb')

                if k >= 99:
                    output_file = open(myDir + "/" + str(sid) + '_' + str(k + 1) + ".jpg", 'wb')

                if k < 9:
                    output_file = open(myDir + "/" + str(sid) + '_00' + str(k + 1) + ".jpg", 'wb')

                    # if k==99:
                    # output_file = open(myDir+"/"+str(sid)+'_'+str(k+1)+".jpg",'wb')

                data = response.read()
                output_file.write(data)
                response.close();

                print("completed ====> " + str(k + 1))

                k = k + 1;

            except IOError:
            
                try:

                    errorCount += 1
                    print("Error on image " + str(k + 1))

                ##################         print error values in df              ################
                # df['Error Count'].apply(lambda x: errorCount)  #wrong one

                finally:
                
                    df.loc[start, 'Error Count'] = errorCount
                    k = k + 1;
  
                # errornumber.write(str(sid) + ': ' + str(search_term[i-1]) + ": "+"IOError on image "+str(k+1) +'\n')

                

            except HTTPError as e:

                errorCount += 1
                print("HTTPError" + str(k))
                k = k + 1;
            except URLError as e:

                errorCount += 1
                print("URLError " + str(k))
                k = k + 1;
        if errorCount ==0:
            df.loc[start, 'Error Count'] = 0
            print("NO ERROR")

        if errorCount > 0:


            #errornumber.write(str(sid) + ': ' + str(search_term[i - 1]) + ":" + str(errorCount) + '\n')
        
            #errornumber.close()

            print("error count = " + str(errorCount))
        
            errorCount=0
        i = i + 1
        start = start + 1

        errorCount = 0

    #print(str(df))

finally:
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()

print("\n")
print("---------------Everything downloaded!!---------------")
print("Check the chosen ROOT folder and EXCEL Sheet")




