import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
import pandas as pd
import warnings
import os
from tkinter import filedialog
import openpyxl
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

window = tk.Tk()
window.title("Atmiya Alumni Association") 
window.resizable(width=False, height=False)
window.rowconfigure(0, minsize=180, weight=1)
window.columnconfigure(0, minsize=110, weight=1)

image = Image.open("AUlogo.png")
resized_image= image.resize((100,100), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(resized_image)
img_label = tk.Label(image=photo,justify=CENTER)
img_label.grid(row=0, column=0, padx=10)

frm_entry = tk.Frame(master=window)
frm_entry.grid(row=0, column=1, padx=10)

lbl_1 = tk.Label(master=frm_entry, text="-- ATMIYA UNIVERSITY --",fg='Dark Blue',font=("Times New Roman", 30))
lbl_1.grid(row=0, column=0, sticky="sw")
lbl_1.grid(row=0, column=0, padx=0)

lbl_2 = tk.Label(master=frm_entry, text="~ Application developed by Brijraj Kacha, under the guidance of Dr. Ashish Kothari.",fg='Black',font=("Times New Roman", 15))
lbl_2.grid(row=2, column=0, sticky="sw")
lbl_2.grid(row=2, column=0, padx=0)

# ent_3 = tk.Entry(master=window,width=55,fg='Blue',font=("Times New Roman", 20))import
# ent_3.grid(row=5, column=1, sticky="e")

# lbl_4 = tk.Label(master=window, text="  | Select the Alumni Association data excel file \N{RIGHTWARDS BLACK ARROW}",fg='Green',font=("Times New Roman", 22))
# lbl_4.grid(row=5, column=1, sticky="n")
# lbl_4.grid(row=5, column=1, padx=0)

lbl_path = tk.Label(master=window, text="  Click here and select the Alumni Association data excel file : ",fg='black',font=("Times New Roman", 20))
lbl_path.grid(row=5, column=0,  sticky="nw")

lbl_result1 = tk.Label(master=window, text="  Total Programs with Minimum 1 Alumni Entry :   ",fg='black',font=("Times New Roman", 20))
lbl_result1.grid(row=6, column=0,  sticky="nw")
ent_4 = tk.Entry(master=window,width=60,fg='Dark Blue',font=("Times New Roman", 20))
ent_4.grid(row=6, column=1, sticky="e")

lbl_result2 = tk.Label(master=window, text="  Highest Entries of Alumni :   ",fg='black',font=("Times New Roman", 20))
lbl_result2.grid(row=7, column=0,  sticky="nw")
ent_5 = tk.Entry(master=window,width=60,fg='Dark Blue',font=("Times New Roman", 20))
ent_5.grid(row=7, column=1, sticky="e")

def browseFiles():
    ent_4.delete(0,tk.END)
    ent_5.delete(0,tk.END)
    file=filedialog.askopenfilename(filetypes=[("Excel File",'.xlsx')])
    df_selected = pd.read_excel(file)
    
    return df_selected

def AlumniSort(db):
    
        directory1 = os.getcwd()
        newpath = directory1+r'\Alumni_Entry Data'
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        
        #db=pd.read_excel(filepath)
        db.head()
        warnings.filterwarnings('ignore')
        def institute(inst1):
            if inst1=='Shri Manibhai & Smt Navalben Virani Science College':
                inst1='VSC'
            if inst1=='Atmiya University':
                inst1='AU'
            if inst1=='Atmiya Institute of Technology and Science':
                inst1='AITS'
            if inst1=='Atmiya Institute of Technology and Science for Diploma Studies':
                inst1='AITSDS'
            if inst1=='Gyanyagna College of Science & Management':
                inst1='GY'
            if inst1=='Atmiya Institute of Pharmacy':
                inst1='AIP'
            return inst1
        
        def programme(prog1):
            if prog1=='B.E./B.Tech Computer Engineering':
                prog1='BCE'
            if prog1=='M.E/M.Tech Computer Engineering':
                prog1='MCE'
            if prog1=='B.E./B.Tech Civil Engineering':
                prog1='BCivilE'
            if prog1=='M.E/M.Tech Civil Engineering':
                prog1='MCivilE'
            if prog1=='B.E./B.Tech Electronics & Communication Engineering':
                prog1='BECE'
            if prog1=='M.E./M.Tech Electronics & Communication Engineering':
                prog1='MECE'
            if prog1=='B.E./B.Tech Information Technology':
                prog1='BIT'
            if prog1=='B.E./B.Tech Electrical Engineering':
                prog1='BEE'
            if prog1=='M.E/M.Tech Electrical Engineering':
                prog1='MEE'
            if prog1=='B.E./ B.Tech Instrumentation and Control Engineering':
                prog1='BIC'
            if prog1=='M.E/M.Tech Instrumentation and Control Engineering':
                prog1='MIC'
            if prog1=='B.E./B.Tech Mechanical Engineering':
                prog1='BME'
            if prog1=='M.E/M.Tech Mechanical Engineering':
                prog1='MME'
            else:
                prog1 = prog1
            return prog1
        #------------------------------------------------------------------------------------------
        
        b=db["Select Program"].tolist()
        lst = []
        for i in b:
                lst.append(i)
        count={}
        for j in lst:
            if not j in count:
                count[j] = 1
            else:
                count[j] += 1
        #print(count)
        foet = ['FoET','B.E./B.Tech Computer Engineering','M.E/M.Tech Computer Engineering','B.E./B.Tech Civil Engineering',
                'M.E/M.Tech Civil Engineering','B.E./B.Tech Electronics & Communication Engineering',
                'M.E./M.Tech Electronics & Communication Engineering','B.E./B.Tech Information Technology',
                'B.E./B.Tech Electrical Engineering','M.E/M.Tech Electrical Engineering',
                'B.E./ B.Tech Instrumentation and Control Engineering','M.E/M.Tech Instrumentation and Control Engineering',
                'B.E./B.Tech Mechanical Engineering','M.E/M.Tech Mechanical Engineering',
                'Ph.D. Computer Engineering','Ph.D. Electrical Engineering','Ph.D. Electronics & Communication',
                'Ph.D. Computer Science']
        
        diploma= ['Diploma',"Diploma Automobile Engineering",
                  "Diploma Civil Engineering",
                  "Diploma Computer Engineering",
                  "Diploma Electrical Engineering",
                  "Diploma Mechanical Engineering",
                  "Diploma Electronics & Communication"]
        
        fos = ["FoS","B.Sc IT","B.Sc Microbiology","B.Sc. Bio Chemistry","B.Sc. Biotechnology",
               "B.Sc. Chemistry","B.Sc. Industrial Chemistry","B.Sc. Mathematics",
               "B.Sc. Physics","M.Sc Information Technology & Computer Application",
               "M.Sc Microbiology","M.Sc. Biotechnology","M.Sc. Chemistry",
               "M.Sc. Industrial Chemistry","M.Sc. IT","M.Sc. Mathematics",
               "M.Sc. Pharmaceutical Organic Chemistry",
               "5 years Integrated B.Sc.-M.Sc. Chemistry",
               "5 years Integrated B.Sc.-M.Sc. Mathematics",
               "5 years Integrated B.Sc.-M.Sc. Microbiology",
               "B.Voc. Applied Computer Technology",
               "B.Voc. Medical Laboratory and Molecular Diagnostic Technology","BCA",
               "B.Voc. Chemical Technology","B.Voc. Pharmaceutical Analysis & Quality Assurance",
               "Ph.D. Microbiology","Ph.D. Biotechnology","Ph.D. Chemistry","Ph.D. Mathematics","Ph.D. Science","PG DMLT"]
        
        fobc = ["FoBC","B.Com","B.Com (Honors)","BBA","BBA (Honors)","MBA","IMBA","BBA(EFB)",
                "B.Com (Logistics)","M.Com","Ph.D. Management","Ph.D. Commerce"]
        
        pharmacy = ["Pharmacy","B.Pharm","M.Pharm","Ph.D. Pharmacy"]
        
        mca = ["MCA"]
        
        fohss = ["FoHSS","BA English"]
        
        faculty = [foet,diploma,fos,fobc,pharmacy,mca,fohss]
        all_dict={}
        
        key=list(count.keys())
        val=list(count.values())
        all_total = 0
        
        for l in faculty:
            
            dept={}
            total=0
            print(l)
        
            for i in range(len(key)):
                for j in l:
                    if key[i]==j:
                        #print(key[i])
                        dept[key[i]]=val[i]
                        total+=val[i] 
            print(dept)
            df = pd.DataFrame(list(dept.items()),columns = ['Branch','Total Entries'])
        
           # len_df=len(df)
            df.loc[len(df.index)]=["Total",total]
            all_dict[l[0]]=total
        
            #datatoexcel = pd.ExcelWriter(l[0]+'.xlsx')  
            all_total += total
            df.to_excel(l[0]+'.xlsx')
            #df.save()
        
        main = pd.DataFrame(list(all_dict.items()),columns = ['Faculty','Total Entries'])
        main.loc[len(main.index)]=["Total",all_total]
        #maintoexcel = pd.ExcelWriter('All DEPT.xlsx')   
        main.to_excel('All DEPT.xlsx')
#       main.close() 
        
        #---------------------------------------------------------------------------------------
        
        DF1 = db.copy()
        DF1['Count']=1
        req_col = ['Passout Year (4 Digits - e.g. 2005)','Count']
        DF2 = DF1[req_col].copy()
        for i in range(len(DF2)):
            val = DF2['Passout Year (4 Digits - e.g. 2005)'].values[i]
            value = str(val)
            DF2['Passout Year (4 Digits - e.g. 2005)'].values[i] = value
        DF2.sort_values(by='Passout Year (4 Digits - e.g. 2005)', ascending=True)
        DF4 = DF2.groupby(['Passout Year (4 Digits - e.g. 2005)']).sum()
        #datatoexcel = pd.ExcelWriter('00_Yearwise.xls')
        DF4.to_excel('00_Yearwise.xlsx')
        
        total_count =0
        lst_percentage = []
        for j in range(len(DF4)):
            total_count = total_count + DF4['Count'].values[j]

        for j in range(len(DF4)):
            p = ((DF4['Count'].values[j])*100)/total_count
            p = round(p,3)
            lst_percentage.append(p)
        
        figure1 = plt.Figure(figsize=(10,4), dpi=100)
        ax1 = figure1.add_subplot(111)
        bar1 = FigureCanvasTkAgg(figure1, window)
        bar1.get_tk_widget().grid(row = 0, column=0, sticky="w")
        
        ax = DF4.plot(kind='bar',y='Count',color='RoyalBlue',ax=ax1)
        plt.title("Year Wise Alumniiii")
        plt.grid(b = True, color ='grey',
                linestyle ='-.', linewidth = 0.5,
                alpha = 0.9)
        plt.legend("")
        k=0
        for p in ax.patches:
            width = p.get_width() # will return the width of the rectangle
            height = p.get_height() #0
            x,y = p.get_xy()
            ax.annotate((str(lst_percentage[k])+"%"), (x + width/2, y + height+k), ha='center')
            k=k+1
            
        plt.show()
        
        #df1 = pd.read_excel(filepath)
        df1=db
        data=df1.copy()
        data.drop_duplicates(keep=False,inplace=True)
        inst_list = df1['Please Select Your Institution'].unique()
        prog_list = df1['Select Program'].unique()
        len_inst = len(inst_list)
        len_prog = len(prog_list)
        count=0
        highest = 0
        dict_a = {}
        for i in range(len_inst):
            inst1a = inst_list[i]
            inst1 = institute(inst1a)
            for j in range(len_prog):
                
                prog1a = prog_list[j]
                prog1 = programme(prog1a)
               # print(prog1a)
                
                for x in faculty:
                    if prog1a in x:
                        #print(x[0])
                        directory = x[0]
                        parent_dir = directory1+"/Alumni_Entry Data/"
                        #parent_dir = directory1 + newpath
                        path = os.path.join(parent_dir,directory)
                        
                        if not os.path.exists(path):
                            os.makedirs(path)
                            #mode = 0o666
                            #os.makedirs(path,mode,exist_ok=False)
                        fname = prog1+'_'+inst1+'.'+'xlsx'
                        pname = prog1+' '+inst1
                        fname3=path+'/'+fname
                        
                        AU_all = df1[df1['Please Select Your Institution']==inst1a]
                        AU_phy = AU_all[AU_all['Select Program']==prog1a]
                
                        len_AU_phy = len(AU_phy)
                    
                        if highest > len_AU_phy:
                            highest_entry = highest_entry
                    
                        else:
                            highest_entry = inst1+' '+prog1
                            highest = len_AU_phy
        
                        if len_AU_phy>0:
                            count=count+1
                            #print('Total Alumni Entry in ',pname,' = ',len_AU_phy)
                            fname2= path+fname
                            #print(fname2)
                            #datatoexcel = pd.ExcelWriter(fname3)
                            #fname3 = fname3.replace(' ', '_')
                            AU_phy.to_excel(fname3)
                            #AU_phy.close()
                   # print (os.getcwd()+"\\"+AU_phy)
        
                dict_a[pname]=len_AU_phy
        
        final = pd.DataFrame(list(dict_a.items()),columns = ['Faculty/Branch','Total Entries'])
       # finaltoexcel = pd.ExcelWriter('All DEPT-Faculty_final.xlsx')
        final.to_excel('All DEPT-Faculty_final.xlsx')
        #final.close()
        print('Total Programs with Minimum 1 Alumni Entry = ',count)
        print('Highest Entries of Alumni in ', highest_entry, ' = ', highest)
    
        
        return(count,highest_entry,highest)
    # except:
    #     return False

def btn():
   # path_of_file = str(ent_3.get())
    selected_df = browseFiles()
    # if path_of_file[0]==path_of_file[-1]=='"':
    #     path_of_file = path_of_file[1:-1]
    
    data_returned = AlumniSort(selected_df)
    print(data_returned)
    if data_returned != False:
        ent_4.insert(0,data_returned[0])
        ent5str = str(str(data_returned[1])+" = "+str(data_returned[2]))
        ent_5.insert(0,ent5str)
    else:
        ent_4.insert(0,"File not found : Enter the correct path !")
        ent_5.insert(0,"File not found : Enter the correct path !")
        
btn_convert = tk.Button(
                        master=window,
                        text="Select the file to run",bg='light grey',fg='black',
                        command = btn,font=("Times New Roman", 13)
                        )
#text="\N{RIGHTWARDS BLACK ARROW}"
button_explore = tk.Button(window,
						text = "Browse Files",
						command = browseFiles,font=("Times New Roman", 13))
#button_explore.grid(row=5,column=1,sticky="e")
btn_convert.grid(row=5, column=1, sticky="w")


window.mainloop()