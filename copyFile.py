import os
import shutil

# getting the current working directory
src_dir = os.getcwd()

expressFolder = 'Z:\ExpressI\\5\\'

fileList = {
    'ARTRNRM.DBF',
    'OESOIT.DBF',
    'STCRD.DBF'

}
# printing current directory
#print(src_dir)
st = False
# copying the files
def copy():
    for i in fileList:
        try:
            shutil.copyfile(expressFolder + i, '.\\temp_db\\' + i)  # copy src to dst
            st = True
        except:
            st = False
            print('Copy Fail')

    # printing the list of new files
    print('Coppy Success')
    return st
