import pandas as pd
import os

#Import the required Libraries
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

#Create an instance of Tkinter frame
win= Tk()
# file_path = entry.get() #input()#'C:\App\Release_notes\AED_V3.016_S1.003_8130_ReleaseNote_Adaptations.xlsx'
file_path =  ''
save_path = ''
all_files = []
count=1

#Set the geometry of Tkinter frame
win.geometry("750x250")
win.title('GÐ')
# p1 = PhotoImage(file = 'C:\App\Release_notes\logo1.ico')
  
# Setting icon of master window
# win.iconphoto(False, p1, p1)

def display_text():
   global entry1, entry2, file_path, save_path
   string1 = entry1.get()
   string2 = entry2.get()
   file_path = string1 + '/'
   save_path = string2
   if len(string1)>0:
    for file in os.listdir(file_path):
        if file.endswith('.xlsx'):
            all_files.append(file)
    for i in all_files:
        release_noteAdaptation(i)
            # print(100*'=')
    messagebox.showinfo('info','>>>Task Done, Thank You!!!<<<')
    win.destroy()
   else:
    messagebox.showinfo('info','>>>Provide the FilePath!!!<<<')

   


# print('>>>Task Done, Thank You!!!<<<')

#Create an Entry widget to accept User Input
label1 = Label(win, text="Filepath:-", font=("Courier 20 bold")).place(x=5, y=5)
entry1= Entry(win, width=30, font=('Courier 20 italic') )
entry1.focus_set()
entry1.place(x=200, y=5)
#Create an Entry widget to accept User Input
label2 = Label(win, text="Savepath:-", font=("Courier 20 bold")).place(x=5, y=80)
entry2= Entry(win, width=30, font=('Courier 20 italic'))
# entry2.focus_set()
entry2.place(x=200,y=80)

#Create a Button to validate Entry Widget
ttk.Button(win, text= "Submit", width= 20 , command= display_text).place(x = 300, y=150)
    

#-------------------------------------------------------------------



def release_noteAdaptation(path):
    global count, file_path, save_path
    df = pd.DataFrame(pd.read_excel(file_path+path)) #, sheet_name='Sheet1'
    iso = os.path.basename(f"{file_path}{path}").split('/')[-1]
    r = df.iloc[4:25]
    df1 = pd.DataFrame(r)
    df1.insert(0,'Adaptation', iso[:3])
    df1.columns = ('Adaptation','Denomination', 'Emission', 'S/N Reading', 'S/N Comparision', 'Fitness', 'Remarks')
    df1.reset_index(inplace=True)
    df1.drop(['index','Remarks'], axis=1, inplace=True)
    for i in range(len(df1)):
        if (df1['S/N Reading'][i]) != 'YES' :
            df1['S/N Reading'][i] = 'NO'
        if (df1['S/N Comparision'][i]) != 'YES' :
            df1['S/N Comparision'][i] = 'NO'
        if (df1['Fitness'][i]) != 'YES' :
            df1['Fitness'][i] = 'NO'
        if type(df1['Emission'][i]) is not int :
            df1.drop(i, axis=0, inplace=True)
    # print(df1.head(20))
    if count==1:
        df1.to_csv(save_path+'\\ReleaseNoteAdaptation.csv', mode='w', index=False)    
    else:
        df1.to_csv(save_path+'\\ReleaseNoteAdaptation.csv', mode='a', header=False, index=False)
    count+=1

win.mainloop()






