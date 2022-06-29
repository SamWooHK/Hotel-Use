from tabula import read_pdf
import pandas as pd
import os
import datetime
import glob
import random

#date
today = datetime.date.today()
yesterday = today + datetime.timedelta(days=-1)
tomorrow = today + datetime.timedelta(days=1)

#find the manager report pdf
mgt=glob.glob(f"\\\\DESKTOP-09QEAO7\\Users\\shinjuku WS-4\\Desktop\\EHHS\\History\\Manager_{yesterday:%m%d}.pdf")[0]
print(f"Fetching data from {mgt}")

#convert into list
pdf_data=read_pdf(mgt,pages="all")
manager_p1=pdf_data[0].values.tolist()
manager_p2=pdf_data[1].values.tolist()
manager_p1.extend(manager_p2)

data_input=[
    'Rooms Occupied minus Comp and House Use',
    'Out of Order Rooms',
    'No Show Rooms',
    'Early Departure Rooms',
    'Late Reservation Cancellations for Today',
    'Room Revenue']

print("目が悪いのでこれをFLASH REPORT のヘルプに。。。\n")
for i in range(1,len(manager_p1)):

    try:
        if manager_p1[i][0] in data_input:
            print(f"{manager_p1[i][0]} : {manager_p1[i][1]}")
    except:
        pass

# アスキーアート
emoji=[
[
"♪",
" 　∧,,∧", 
".(´・ω・`)　♪", 
"　( つ　つ ",
"((（⌒　) ))", 
"　 し' ｕ ",
"", 
"♪　 ∧,,∧", 
"　∩´・ω・`)", 
"　ヽ　 ⊂ノ　♪", 
"(( （　 ⌒)　))", 
"　　ｕ し' ",
"", 
"　∧,,∧", 
"(・ω・｀) /')",  
"⊂⊂：::: _ノ彡",  
"　　し彡"
],
[
"|" ,
"|▁╱╱╱▁╱▔)" ,
"|⠀⠀⠀⠀⠀⠀⠀⠀⠀╭" ,
"|▊ ˳˳̊˳̊̊˳̊̊̊˳̊̊˳̊˳  ▊⠀⠀ ▏っつ",
"|⠀⠀◥◤⠀⠀⠀  ⠀▏っつ" ,
"|╲▁/\▁╱⠀ ╱" ,
"|╭━━━╯⠀╲" ,
"|╡⠀⠀⠀⠀⠀⠀⠀⠀▏ ",
"|╰━━━━╯⠀▏ "
"|⠀⠀⠀⠀⠀⠀⠀⠀⠀ ▏))"
],
[
" (,■)", 
"　ヽ(´ (ェ)｀)", 
"　　/        /ゝ  ドヤ♡", 
"     ノ￣￣ゝ"
],
[
"なんだ夢か ",
"      ∧_∧ ",
"　　 (　･ω･)　 ",
"　 ＿|　⊃／(＿＿_ ",
"／　└-(＿＿＿_／"
]
]

for row in emoji[random.randrange(len(emoji))]:
    print(row)

os.system("pause")