# -*- coding: utf-8 -*-
"""
Created on Wed Sep 15 18:22:32 2021

@author: allan
"""

import xlsxwriter
from tkinter import *

#Global Counts for each button
count1 = 0
def c1():
    global count1
    count1+=1
    k.config(text=count1)
    
count2 = 0
def c2():
    global count2
    count2+=1
    k1.config(text=count2)
    
count3 = 0
def c3():
    global count3
    count3+=1
    k2.config(text=count3)
    
count4 = 0
def c4():
    global count4
    count4+=1
    k3.config(text=count4)

count5 = 0
def c5():
    global count5
    count5+=1
    k4.config(text=count5)
    
count6 = 0
def c6():
    global count6
    count6+=1
    a.config(text=count6)
    
count7 = 0
def c7():
    global count7
    count7+=1
    a1.config(text=count7)

count8 = 0
def c8():
    global count8
    count8+=1
    a2.config(text=count8)
    
count9 = 0
def c9():
    global count9
    count9+=1
    a3.config(text=count9)
    
count10 = 0
def c10():
    global count10
    count10+=1
    d.config(text=count10)
    
count11 = 0
def c11():
    global count11
    count11+=1
    d1.config(text=count11)
    
count12 = 0
def c12():
    global count12
    count12+=1
    d2.config(text=count12)
    
count13 = 0
def c13():
    global count13
    count13+=1
    d3.config(text=count13)
    
count14 = 0
def c14():
    global count14
    count14+=1
    d4.config(text=count14)
    
#Player2
count1_2 = 0
def c1_2():
    global count1_2
    count1_2+=1
    k_2.config(text=count1_2)
    
count2_2 = 0
def c2_2():
    global count2_2
    count2_2+=1
    k1_2.config(text=count2_2)
    
count3_2 = 0
def c3_2():
    global count3_2
    count3_2+=1
    k2_2.config(text=count3_2)
    
count4_2 = 0
def c4_2():
    global count4_2
    count4_2+=1
    k3_2.config(text=count4_2)

count5_2 = 0
def c5_2():
    global count5_2
    count5_2+=1
    k4_2.config(text=count5_2)
    
count6_2 = 0
def c6_2():
    global count6_2
    count6_2+=1
    a_2.config(text=count6_2)
    
count7_2 = 0
def c7_2():
    global count7_2
    count7_2+=1
    a1_2.config(text=count7_2)

count8_2 = 0
def c8_2():
    global count8_2
    count8_2+=1
    a2_2.config(text=count8_2)
    
count9_2 = 0
def c9_2():
    global count9_2
    count9_2+=1
    a3_2.config(text=count9_2)
    
count10_2 = 0
def c10_2():
    global count10_2
    count10_2+=1
    d_2.config(text=count10_2)
    
count11_2 = 0
def c11_2():
    global count11_2
    count11_2+=1
    d1_2.config(text=count11_2)
    
count12_2 = 0
def c12_2():
    global count12_2
    count12_2+=1
    d2_2.config(text=count12_2)
    
count13_2 = 0
def c13_2():
    global count13_2
    count13_2+=1
    d3_2.config(text=count13_2)
    
count14_2 = 0
def c14_2():
    global count14_2
    count14_2+=1
    d4_2.config(text=count14_2)
    
#Player3
count1_3 = 0
def c1_3():
    global count1_3
    count1_3+=1
    k_3.config(text=count1_3)
    
count2_3 = 0
def c2_3():
    global count2_3
    count2_3+=1
    k1_3.config(text=count2_3)
    
count3_3 = 0
def c3_3():
    global count3_3
    count3_3+=1
    k2_3.config(text=count3_3)
    
count4_3 = 0
def c4_3():
    global count4_3
    count4_3+=1
    k3_3.config(text=count4_3)

count5_3 = 0
def c5_3():
    global count5_3
    count5_3+=1
    k4_3.config(text=count5_3)
    
count6_3 = 0
def c6_3():
    global count6_3
    count6_3+=1
    a_3.config(text=count6_3)
    
count7_3 = 0
def c7_3():
    global count7_3
    count7_3+=1
    a1_3.config(text=count7_3)

count8_3 = 0
def c8_3():
    global count8_3
    count8_3+=1
    a2_3.config(text=count8_3)
    
count9_3 = 0
def c9_3():
    global count9_3
    count9_3+=1
    a3_3.config(text=count9_3)
    
count10_3 = 0
def c10_3():
    global count10_3
    count10_3+=1
    d_3.config(text=count10_3)
    
count11_3 = 0
def c11_3():
    global count11_3
    count11_3+=1
    d1_3.config(text=count11_3)
    
count12_3 = 0
def c12_3():
    global count12_3
    count12_3+=1
    d2_3.config(text=count12_3)
    
count13_3 = 0
def c13_3():
    global count13_3
    count13_3+=1
    d3_3.config(text=count13_3)
    
count14_3 = 0
def c14_3():
    global count14_3
    count14_3+=1
    d4_3.config(text=count14_3)
    
#Player4
count1_4 = 0
def c1_4():
    global count1_4
    count1_4+=1
    k_4.config(text=count1_4)
    
count2_4 = 0
def c2_4():
    global count2_4
    count2_4+=1
    k1_4.config(text=count2_4)
    
count3_4 = 0
def c3_4():
    global count3_4
    count3_4+=1
    k2_4.config(text=count3_4)
    
count4_4 = 0
def c4_4():
    global count4_4
    count4_4+=1
    k3_4.config(text=count4_4)

count5_4 = 0
def c5_4():
    global count5_4
    count5_4+=1
    k4_4.config(text=count5_4)
    
count6_4 = 0
def c6_4():
    global count6_4
    count6_4+=1
    a_4.config(text=count6_4)
    
count7_4 = 0
def c7_4():
    global count7_4
    count7_4+=1
    a1_4.config(text=count7_4)

count8_4 = 0
def c8_4():
    global count8_4
    count8_4+=1
    a2_4.config(text=count8_4)
    
count9_4 = 0
def c9_4():
    global count9_4
    count9_4+=1
    a3_4.config(text=count9_4)
    
count10_4 = 0
def c10_4():
    global count10_4
    count10_4+=1
    d_4.config(text=count10_4)
    
count11_4 = 0
def c11_4():
    global count11_4
    count11_4+=1
    d1_4.config(text=count11_4)
    
count12_4 = 0
def c12_4():
    global count12_4
    count12_4+=1
    d2_4.config(text=count12_4)
    
count13_4 = 0
def c13_4():
    global count13_4
    count13_4+=1
    d3_4.config(text=count13_4)
    
count14_4 = 0
def c14_4():
    global count14_4
    count14_4+=1
    d4_4.config(text=count14_4)
    
#Player5
count1_5 = 0
def c1_5():
    global count1_5
    count1_5+=1
    k_5.config(text=count1_5)
    
count2_5 = 0
def c2_5():
    global count2_5
    count2_5+=1
    k1_5.config(text=count2_5)
    
count3_5 = 0
def c3_5():
    global count3_5
    count3_5+=1
    k2_5.config(text=count3_5)
    
count4_5 = 0
def c4_5():
    global count4_5
    count4_5+=1
    k3_5.config(text=count4_5)

count5_5 = 0
def c5_5():
    global count5_5
    count5_5+=1
    k4_5.config(text=count5_5)
    
count6_5 = 0
def c6_5():
    global count6_5
    count6_5+=1
    a_5.config(text=count6_5)
    
count7_5 = 0
def c7_5():
    global count7_5
    count7_5+=1
    a1_5.config(text=count7_5)

count8_5 = 0
def c8_5():
    global count8_5
    count8_5+=1
    a2_5.config(text=count8_5)
    
count9_5 = 0
def c9_5():
    global count9_5
    count9_5+=1
    a3_5.config(text=count9_5)
    
count10_5 = 0
def c10_5():
    global count10_5
    count10_5+=1
    d_5.config(text=count10_5)
    
count11_5 = 0
def c11_5():
    global count11_5
    count11_5+=1
    d1_5.config(text=count11_5)
    
count12_5 = 0
def c12_5():
    global count12_5
    count12_5+=1
    d2_5.config(text=count12_5)
    
count13_5 = 0
def c13_5():
    global count13_5
    count13_5+=1
    d3_5.config(text=count13_5)
    
count14_5 = 0
def c14_5():
    global count14_5
    count14_5+=1
    d4_5.config(text=count14_5)
    
#Player6
count1_6 = 0
def c1_6():
    global count1_6
    count1_6+=1
    k_6.config(text=count1_6)
    
count2_6 = 0
def c2_6():
    global count2_6
    count2_6+=1
    k1_6.config(text=count2_6)
    
count3_6 = 0
def c3_6():
    global count3_6
    count3_6+=1
    k2_6.config(text=count3_6)
    
count4_6 = 0
def c4_6():
    global count4_6
    count4_6+=1
    k3_6.config(text=count4_6)

count5_6 = 0
def c5_6():
    global count5_6
    count5_6+=1
    k4_6.config(text=count5_6)
    
count6_6 = 0
def c6_6():
    global count6_6
    count6_6+=1
    a_6.config(text=count6_6)
    
count7_6 = 0
def c7_6():
    global count7_6
    count7_6+=1
    a1_6.config(text=count7_6)

count8_6 = 0
def c8_6():
    global count8_6
    count8_6+=1
    a2_6.config(text=count8_6)
    
count9_6 = 0
def c9_6():
    global count9_6
    count9_6+=1
    a3_6.config(text=count9_6)
    
count10_6 = 0
def c10_6():
    global count10_6
    count10_6+=1
    d_6.config(text=count10_6)
    
count11_6 = 0
def c11_6():
    global count11_6
    count11_6+=1
    d1_6.config(text=count11_6)
    
count12_6 = 0
def c12_6():
    global count12_6
    count12_6+=1
    d2_6.config(text=count12_6)
    
count13_6 = 0
def c13_6():
    global count13_6
    count13_6+=1
    d3_6.config(text=count13_6)
    
count14_6 = 0
def c14_6():
    global count14_6
    count14_6+=1
    d4_6.config(text=count14_6)
    
#Player7
count1_7 = 0
def c1_7():
    global count1_7
    count1_7+=1
    k_7.config(text=count1_7)
    
count2_7 = 0
def c2_7():
    global count2_7
    count2_7+=1
    k1_7.config(text=count2_7)
    
count3_7 = 0
def c3_7():
    global count3_7
    count3_7+=1
    k2_7.config(text=count3_7)
    
count4_7 = 0
def c4_7():
    global count4_7
    count4_7+=1
    k3_7.config(text=count4_7)

count5_7 = 0
def c5_7():
    global count5_7
    count5_7+=1
    k4_7.config(text=count5_7)
    
count6_7 = 0
def c6_7():
    global count6_7
    count6_7+=1
    a_7.config(text=count6_7)
    
count7_7 = 0
def c7_7():
    global count7_7
    count7_7+=1
    a1_7.config(text=count7_7)

count8_7 = 0
def c8_7():
    global count8_7
    count8_7+=1
    a2_7.config(text=count8_7)
    
count9_7 = 0
def c9_7():
    global count9_7
    count9_7+=1
    a3_7.config(text=count9_7)
    
count10_7 = 0
def c10_7():
    global count10_7
    count10_7+=1
    d_7.config(text=count10_7)
    
count11_7 = 0
def c11_7():
    global count11_7
    count11_7+=1
    d1_7.config(text=count11_7)
    
count12_7 = 0
def c12_7():
    global count12_7
    count12_7+=1
    d2_7.config(text=count12_7)
    
count13_7 = 0
def c13_7():
    global count13_7
    count13_7+=1
    d3_7.config(text=count13_7)
    
count14_7 = 0
def c14_7():
    global count14_7
    count14_7+=1
    d4_7.config(text=count14_7)
    
    
#Player8
count1_8 = 0
def c1_8():
    global count1_8
    count1_8+=1
    k_8.config(text=count1_8)
    
count2_8 = 0
def c2_8():
    global count2_8
    count2_8+=1
    k1_8.config(text=count2_8)
    
count3_8 = 0
def c3_8():
    global count3_8
    count3_8+=1
    k2_8.config(text=count3_8)
    
count4_8 = 0
def c4_8():
    global count4_8
    count4_8+=1
    k3_8.config(text=count4_8)

count5_8 = 0
def c5_8():
    global count5_8
    count5_8+=1
    k4_8.config(text=count5_8)
    
count6_8 = 0
def c6_8():
    global count6_8
    count6_8+=1
    a_8.config(text=count6_8)
    
count7_8 = 0
def c7_8():
    global count7_8
    count7_8+=1
    a1_8.config(text=count7_8)

count8_8 = 0
def c8_8():
    global count8_8
    count8_8+=1
    a2_8.config(text=count8_8)
    
count9_8 = 0
def c9_8():
    global count9_8
    count9_8+=1
    a3_8.config(text=count9_8)
    
count10_8 = 0
def c10_8():
    global count10_8
    count10_8+=1
    d_8.config(text=count10_8)
    
count11_8 = 0
def c11_8():
    global count11_8
    count11_8+=1
    d1_8.config(text=count11_8)
    
count12_8 = 0
def c12_8():
    global count12_8
    count12_8+=1
    d2_8.config(text=count12_8)
    
count13_8 = 0
def c13_8():
    global count13_8
    count13_8+=1
    d3_8.config(text=count13_8)
    
count14_8 = 0
def c14_8():
    global count14_8
    count14_8+=1
    d4_8.config(text=count14_8)
    
#Player9
count1_9 = 0
def c1_9():
    global count1_9
    count1_9+=1
    k_9.config(text=count1_9)
    
count2_9 = 0
def c2_9():
    global count2_9
    count2_9+=1
    k1_9.config(text=count2_9)
    
count3_9 = 0
def c3_9():
    global count3_9
    count3_9+=1
    k2_9.config(text=count3_9)
    
count4_9 = 0
def c4_9():
    global count4_9
    count4_9+=1
    k3_9.config(text=count4_9)

count5_9 = 0
def c5_9():
    global count5_9
    count5_9+=1
    k4_9.config(text=count5_9)
    
count6_9 = 0
def c6_9():
    global count6_9
    count6_9+=1
    a_9.config(text=count6_9)
    
count7_9 = 0
def c7_9():
    global count7_9
    count7_9+=1
    a1_9.config(text=count7_9)

count8_9 = 0
def c8_9():
    global count8_9
    count8_9+=1
    a2_9.config(text=count8_9)
    
count9_9 = 0
def c9_9():
    global count9_9
    count9_9+=1
    a3_9.config(text=count9_9)
    
count10_9 = 0
def c10_9():
    global count10_9
    count10_9+=1
    d_9.config(text=count10_9)
    
count11_9 = 0
def c11_9():
    global count11_9
    count11_9+=1
    d1_9.config(text=count11_9)
    
count12_9 = 0
def c12_9():
    global count12_9
    count12_9+=1
    d2_9.config(text=count12_9)
    
count13_9 = 0
def c13_9():
    global count13_9
    count13_9+=1
    d3_9.config(text=count13_9)
    
count14_9 = 0
def c14_9():
    global count14_9
    count14_9+=1
    d4_9.config(text=count14_9)
    
#Player10
count1_10 = 0
def c1_10():
    global count1_10
    count1_10+=1
    k_10.config(text=count1_10)
    
count2_10 = 0
def c2_10():
    global count2_10
    count2_10+=1
    k1_10.config(text=count2_10)
    
count3_10 = 0
def c3_10():
    global count3_10
    count3_10+=1
    k2_10.config(text=count3_10)
    
count4_10 = 0
def c4_10():
    global count4_10
    count4_10+=1
    k3_10.config(text=count4_10)

count5_10 = 0
def c5_10():
    global count5_10
    count5_10+=1
    k4_10.config(text=count5_10)
    
count6_10 = 0
def c6_10():
    global count6_10
    count6_10+=1
    a_10.config(text=count6_10)
    
count7_10 = 0
def c7_10():
    global count7_10
    count7_10+=1
    a1_10.config(text=count7_10)

count8_10 = 0
def c8_10():
    global count8_10
    count8_10+=1
    a2_10.config(text=count8_10)
    
count9_10 = 0
def c9_10():
    global count9_10
    count9_10+=1
    a3_10.config(text=count9_10)
    
count10_10 = 0
def c10_10():
    global count10_10
    count10_10+=1
    d_10.config(text=count10_10)
    
count11_10 = 0
def c11_10():
    global count11_10
    count11_10+=1
    d1_10.config(text=count11_10)
    
count12_10 = 0
def c12_10():
    global count12_10
    count12_10+=1
    d2_10.config(text=count12_10)
    
count13_10 = 0
def c13_10():
    global count13_10
    count13_10+=1
    d3_10.config(text=count13_10)
    
count14_10 = 0
def c14_10():
    global count14_10
    count14_10+=1
    d4_10.config(text=count14_10)


#Main Window
window = Tk()
window.title("Game KDA Counter")
window.geometry("700x325")

#Labels

team_label = Label(window, text="Team")
team_label.grid(row=0,column=1)
player_label = Label(window, text="Player")
player_label.grid(row=0,column=2)
champ_label = Label(window, text="Champ")
champ_label.grid(row=0,column=3)
k_label = Label(window, text="K+0")
k_label.grid(row=0,column=4)
k1_label = Label(window, text="K+1")
k1_label.grid(row=0,column=5)
k2_label = Label(window, text="K+2")
k2_label.grid(row=0,column=6)
k3_label = Label(window, text="K+3")
k3_label.grid(row=0,column=7)
k4_label = Label(window, text="K+4")
k4_label.grid(row=0,column=8)
a_label = Label(window, text="A+0")
a_label.grid(row=0,column=9)
a1_label = Label(window, text="A+1")
a1_label.grid(row=0,column=10)
a2_label = Label(window, text="A+2")
a2_label.grid(row=0,column=11)
a3_label = Label(window, text="A+3")
a3_label.grid(row=0,column=12)
d_label = Label(window, text="D+0")
d_label.grid(row=0,column=13)
d1_label = Label(window, text="D+1")
d1_label.grid(row=0,column=14)
d2_label = Label(window, text="D+2")
d2_label.grid(row=0,column=15)
d3_label = Label(window, text="D+3")
d3_label.grid(row=0,column=16)
d4_label = Label(window, text="D+4")
d4_label.grid(row=0,column=17)


#KDA Count Buttons
#Player1
##94e09a
k = Button(window,text=str(count1))
k.config(command=c1)
k.config(bg='#94e09a')
k.grid(row=1,column=4)

k1 = Button(window,text=str(count2))
k1.config(command=c2)
k1.config(bg='#94e09a')
k1.grid(row=1,column=5)

k2 = Button(window,text=str(count3))
k2.config(command=c3)
k2.config(bg='#94e09a')
k2.grid(row=1,column=6)

k3 = Button(window,text=str(count4))
k3.config(command=c4)
k3.config(bg='#94e09a')
k3.grid(row=1,column=7)

k4 = Button(window,text=str(count5))
k4.config(command=c5)
k4.config(bg='#94e09a')
k4.grid(row=1,column=8)

#94d1e0
a = Button(window,text=str(count6))
a.config(command=c6)
a.config(bg='#94d1e0')
a.grid(row=1,column=9)

a1 = Button(window,text=str(count7))
a1.config(command=c7)
a1.config(bg='#94d1e0')
a1.grid(row=1,column=10)

a2 = Button(window,text=str(count8))
a2.config(command=c8)
a2.config(bg='#94d1e0')
a2.grid(row=1,column=11)

a3 = Button(window,text=str(count9))
a3.config(command=c9)
a3.config(bg='#94d1e0')
a3.grid(row=1,column=12)

#e09494
d = Button(window,text=str(count10))
d.config(command=c10)
d.config(bg='#e09494')
d.grid(row=1,column=13)

d1 = Button(window,text=str(count11))
d1.config(command=c11)
d1.config(bg='#e09494')
d1.grid(row=1,column=14)

d2 = Button(window,text=str(count12))
d2.config(command=c12)
d2.config(bg='#e09494')
d2.grid(row=1,column=15)

d3 = Button(window,text=str(count13))
d3.config(command=c13)
d3.config(bg='#e09494')
d3.grid(row=1,column=16)

d4 = Button(window,text=str(count14))
d4.config(command=c14)
d4.config(bg='#e09494')
d4.grid(row=1,column=17)

#Player2
k_2 = Button(window,text=str(count1_2))
k_2.config(command=c1_2)
k_2.config(bg='#94e09a')
k_2.grid(row=2,column=4)

k1_2 = Button(window,text=str(count2_2))
k1_2.config(command=c2_2)
k1_2.config(bg='#94e09a')
k1_2.grid(row=2,column=5)

k2_2 = Button(window,text=str(count3_2))
k2_2.config(command=c3_2)
k2_2.config(bg='#94e09a')
k2_2.grid(row=2,column=6)

k3_2 = Button(window,text=str(count4_2))
k3_2.config(command=c4_2)
k3_2.config(bg='#94e09a')
k3_2.grid(row=2,column=7)

k4_2 = Button(window,text=str(count5_2))
k4_2.config(command=c5_2)
k4_2.config(bg='#94e09a')
k4_2.grid(row=2,column=8)

#94d1e0
a_2 = Button(window,text=str(count6_2))
a_2.config(command=c6_2)
a_2.config(bg='#94d1e0')
a_2.grid(row=2,column=9)

a1_2 = Button(window,text=str(count7_2))
a1_2.config(command=c7_2)
a1_2.config(bg='#94d1e0')
a1_2.grid(row=2,column=10)

a2_2 = Button(window,text=str(count8_2))
a2_2.config(command=c8_2)
a2_2.config(bg='#94d1e0')
a2_2.grid(row=2,column=11)

a3_2 = Button(window,text=str(count9_2))
a3_2.config(command=c9_2)
a3_2.config(bg='#94d1e0')
a3_2.grid(row=2,column=12)

#e09494
d_2 = Button(window,text=str(count10_2))
d_2.config(command=c10_2)
d_2.config(bg='#e09494')
d_2.grid(row=2,column=13)

d1_2 = Button(window,text=str(count11_2))
d1_2.config(command=c11_2)
d1_2.config(bg='#e09494')
d1_2.grid(row=2,column=14)

d2_2 = Button(window,text=str(count12_2))
d2_2.config(command=c12_2)
d2_2.config(bg='#e09494')
d2_2.grid(row=2,column=15)

d3_2 = Button(window,text=str(count13_2))
d3_2.config(command=c13_2)
d3_2.config(bg='#e09494')
d3_2.grid(row=2,column=16)

d4_2 = Button(window,text=str(count14_2))
d4_2.config(command=c14_2)
d4_2.config(bg='#e09494')
d4_2.grid(row=2,column=17)

#Player3
k_3 = Button(window,text=str(count1_3))
k_3.config(command=c1_3)
k_3.config(bg='#94e09a')
k_3.grid(row=3,column=4)

k1_3 = Button(window,text=str(count2_3))
k1_3.config(command=c2_3)
k1_3.config(bg='#94e09a')
k1_3.grid(row=3,column=5)

k2_3 = Button(window,text=str(count3_3))
k2_3.config(command=c3_3)
k2_3.config(bg='#94e09a')
k2_3.grid(row=3,column=6)

k3_3 = Button(window,text=str(count4_3))
k3_3.config(command=c4_3)
k3_3.config(bg='#94e09a')
k3_3.grid(row=3,column=7)

k4_3 = Button(window,text=str(count5_3))
k4_3.config(command=c5_3)
k4_3.config(bg='#94e09a')
k4_3.grid(row=3,column=8)

#94d1e0
a_3 = Button(window,text=str(count6_3))
a_3.config(command=c6_3)
a_3.config(bg='#94d1e0')
a_3.grid(row=3,column=9)

a1_3 = Button(window,text=str(count7_3))
a1_3.config(command=c7_3)
a1_3.config(bg='#94d1e0')
a1_3.grid(row=3,column=10)

a2_3 = Button(window,text=str(count8_3))
a2_3.config(command=c8_3)
a2_3.config(bg='#94d1e0')
a2_3.grid(row=3,column=11)

a3_3 = Button(window,text=str(count9_3))
a3_3.config(command=c9_3)
a3_3.config(bg='#94d1e0')
a3_3.grid(row=3,column=12)

#e09494
d_3 = Button(window,text=str(count10_3))
d_3.config(command=c10_3)
d_3.config(bg='#e09494')
d_3.grid(row=3,column=13)

d1_3 = Button(window,text=str(count11_3))
d1_3.config(command=c11_3)
d1_3.config(bg='#e09494')
d1_3.grid(row=3,column=14)

d2_3 = Button(window,text=str(count12_3))
d2_3.config(command=c12_3)
d2_3.config(bg='#e09494')
d2_3.grid(row=3,column=15)

d3_3 = Button(window,text=str(count13_3))
d3_3.config(command=c13_3)
d3_3.config(bg='#e09494')
d3_3.grid(row=3,column=16)

d4_3 = Button(window,text=str(count14_3))
d4_3.config(command=c14_3)
d4_3.config(bg='#e09494')
d4_3.grid(row=3,column=17)

#Player4
k_4 = Button(window,text=str(count1_4))
k_4.config(command=c1_4)
k_4.config(bg='#94e09a')
k_4.grid(row=4,column=4)

k1_4 = Button(window,text=str(count2_4))
k1_4.config(command=c2_4)
k1_4.config(bg='#94e09a')
k1_4.grid(row=4,column=5)

k2_4 = Button(window,text=str(count3_4))
k2_4.config(command=c3_4)
k2_4.config(bg='#94e09a')
k2_4.grid(row=4,column=6)

k3_4 = Button(window,text=str(count4_4))
k3_4.config(command=c4_4)
k3_4.config(bg='#94e09a')
k3_4.grid(row=4,column=7)

k4_4 = Button(window,text=str(count5_4))
k4_4.config(command=c5_4)
k4_4.config(bg='#94e09a')
k4_4.grid(row=4,column=8)

#94d1e0
a_4 = Button(window,text=str(count6_4))
a_4.config(command=c6_4)
a_4.config(bg='#94d1e0')
a_4.grid(row=4,column=9)

a1_4 = Button(window,text=str(count7_4))
a1_4.config(command=c7_4)
a1_4.config(bg='#94d1e0')
a1_4.grid(row=4,column=10)

a2_4 = Button(window,text=str(count8_4))
a2_4.config(command=c8_4)
a2_4.config(bg='#94d1e0')
a2_4.grid(row=4,column=11)

a3_4 = Button(window,text=str(count9_4))
a3_4.config(command=c9_4)
a3_4.config(bg='#94d1e0')
a3_4.grid(row=4,column=12)

#e09494
d_4 = Button(window,text=str(count10_4))
d_4.config(command=c10_4)
d_4.config(bg='#e09494')
d_4.grid(row=4,column=13)

d1_4 = Button(window,text=str(count11_4))
d1_4.config(command=c11_4)
d1_4.config(bg='#e09494')
d1_4.grid(row=4,column=14)

d2_4 = Button(window,text=str(count12_4))
d2_4.config(command=c12_4)
d2_4.config(bg='#e09494')
d2_4.grid(row=4,column=15)

d3_4 = Button(window,text=str(count13_4))
d3_4.config(command=c13_4)
d3_4.config(bg='#e09494')
d3_4.grid(row=4,column=16)

d4_4 = Button(window,text=str(count14_4))
d4_4.config(command=c14_4)
d4_4.config(bg='#e09494')
d4_4.grid(row=4,column=17)

#Player5
k_5 = Button(window,text=str(count1_5))
k_5.config(command=c1_5)
k_5.config(bg='#94e09a')
k_5.grid(row=5,column=4)

k1_5 = Button(window,text=str(count2_5))
k1_5.config(command=c2_5)
k1_5.config(bg='#94e09a')
k1_5.grid(row=5,column=5)

k2_5 = Button(window,text=str(count3_5))
k2_5.config(command=c3_5)
k2_5.config(bg='#94e09a')
k2_5.grid(row=5,column=6)

k3_5 = Button(window,text=str(count4_5))
k3_5.config(command=c4_5)
k3_5.config(bg='#94e09a')
k3_5.grid(row=5,column=7)

k4_5 = Button(window,text=str(count5_5))
k4_5.config(command=c5_5)
k4_5.config(bg='#94e09a')
k4_5.grid(row=5,column=8)

#94d1e0
a_5 = Button(window,text=str(count6_5))
a_5.config(command=c6_5)
a_5.config(bg='#94d1e0')
a_5.grid(row=5,column=9)

a1_5 = Button(window,text=str(count7_5))
a1_5.config(command=c7_5)
a1_5.config(bg='#94d1e0')
a1_5.grid(row=5,column=10)

a2_5 = Button(window,text=str(count8_5))
a2_5.config(command=c8_5)
a2_5.config(bg='#94d1e0')
a2_5.grid(row=5,column=11)

a3_5 = Button(window,text=str(count9_5))
a3_5.config(command=c9_5)
a3_5.config(bg='#94d1e0')
a3_5.grid(row=5,column=12)

#e09494
d_5 = Button(window,text=str(count10_5))
d_5.config(command=c10_5)
d_5.config(bg='#e09494')
d_5.grid(row=5,column=13)

d1_5 = Button(window,text=str(count11_5))
d1_5.config(command=c11_5)
d1_5.config(bg='#e09494')
d1_5.grid(row=5,column=14)

d2_5 = Button(window,text=str(count12_5))
d2_5.config(command=c12_5)
d2_5.config(bg='#e09494')
d2_5.grid(row=5,column=15)

d3_5 = Button(window,text=str(count13_5))
d3_5.config(command=c13_5)
d3_5.config(bg='#e09494')
d3_5.grid(row=5,column=16)

d4_5 = Button(window,text=str(count14_5))
d4_5.config(command=c14_5)
d4_5.config(bg='#e09494')
d4_5.grid(row=5,column=17)

#Player6
k_6 = Button(window,text=str(count1_6))
k_6.config(command=c1_6)
k_6.config(bg='#94e09a')
k_6.grid(row=6,column=4)

k1_6 = Button(window,text=str(count2_6))
k1_6.config(command=c2_6)
k1_6.config(bg='#94e09a')
k1_6.grid(row=6,column=5)

k2_6 = Button(window,text=str(count3_6))
k2_6.config(command=c3_6)
k2_6.config(bg='#94e09a')
k2_6.grid(row=6,column=6)

k3_6 = Button(window,text=str(count4_6))
k3_6.config(command=c4_6)
k3_6.config(bg='#94e09a')
k3_6.grid(row=6,column=7)

k4_6 = Button(window,text=str(count5_6))
k4_6.config(command=c5_6)
k4_6.config(bg='#94e09a')
k4_6.grid(row=6,column=8)

#94d1e0
a_6 = Button(window,text=str(count6_6))
a_6.config(command=c6_6)
a_6.config(bg='#94d1e0')
a_6.grid(row=6,column=9)

a1_6 = Button(window,text=str(count7_6))
a1_6.config(command=c7_6)
a1_6.config(bg='#94d1e0')
a1_6.grid(row=6,column=10)

a2_6 = Button(window,text=str(count8_6))
a2_6.config(command=c8_6)
a2_6.config(bg='#94d1e0')
a2_6.grid(row=6,column=11)

a3_6 = Button(window,text=str(count9_6))
a3_6.config(command=c9_6)
a3_6.config(bg='#94d1e0')
a3_6.grid(row=6,column=12)

#e09494
d_6 = Button(window,text=str(count10_6))
d_6.config(command=c10_6)
d_6.config(bg='#e09494')
d_6.grid(row=6,column=13)

d1_6 = Button(window,text=str(count11_6))
d1_6.config(command=c11_6)
d1_6.config(bg='#e09494')
d1_6.grid(row=6,column=14)

d2_6 = Button(window,text=str(count12_6))
d2_6.config(command=c12_6)
d2_6.config(bg='#e09494')
d2_6.grid(row=6,column=15)

d3_6 = Button(window,text=str(count13_6))
d3_6.config(command=c13_6)
d3_6.config(bg='#e09494')
d3_6.grid(row=6,column=16)

d4_6 = Button(window,text=str(count14_6))
d4_6.config(command=c14_6)
d4_6.config(bg='#e09494')
d4_6.grid(row=6,column=17)

#Player7
k_7 = Button(window,text=str(count1_7))
k_7.config(command=c1_7)
k_7.config(bg='#94e09a')
k_7.grid(row=7,column=4)

k1_7 = Button(window,text=str(count2_7))
k1_7.config(command=c2_7)
k1_7.config(bg='#94e09a')
k1_7.grid(row=7,column=5)

k2_7 = Button(window,text=str(count3_7))
k2_7.config(command=c3_7)
k2_7.config(bg='#94e09a')
k2_7.grid(row=7,column=6)

k3_7 = Button(window,text=str(count4_7))
k3_7.config(command=c4_7)
k3_7.config(bg='#94e09a')
k3_7.grid(row=7,column=7)

k4_7 = Button(window,text=str(count5_7))
k4_7.config(command=c5_7)
k4_7.config(bg='#94e09a')
k4_7.grid(row=7,column=8)

#94d1e0
a_7 = Button(window,text=str(count6_7))
a_7.config(command=c6_7)
a_7.config(bg='#94d1e0')
a_7.grid(row=7,column=9)

a1_7 = Button(window,text=str(count7_7))
a1_7.config(command=c7_7)
a1_7.config(bg='#94d1e0')
a1_7.grid(row=7,column=10)

a2_7 = Button(window,text=str(count8_7))
a2_7.config(command=c8_7)
a2_7.config(bg='#94d1e0')
a2_7.grid(row=7,column=11)

a3_7 = Button(window,text=str(count9_7))
a3_7.config(command=c9_7)
a3_7.config(bg='#94d1e0')
a3_7.grid(row=7,column=12)

#e09494
d_7 = Button(window,text=str(count10_7))
d_7.config(command=c10_7)
d_7.config(bg='#e09494')
d_7.grid(row=7,column=13)

d1_7 = Button(window,text=str(count11_7))
d1_7.config(command=c11_7)
d1_7.config(bg='#e09494')
d1_7.grid(row=7,column=14)

d2_7 = Button(window,text=str(count12_7))
d2_7.config(command=c12_7)
d2_7.config(bg='#e09494')
d2_7.grid(row=7,column=15)

d3_7 = Button(window,text=str(count13_7))
d3_7.config(command=c13_7)
d3_7.config(bg='#e09494')
d3_7.grid(row=7,column=16)

d4_7 = Button(window,text=str(count14_7))
d4_7.config(command=c14_7)
d4_7.config(bg='#e09494')
d4_7.grid(row=7,column=17)

#Player8
k_8 = Button(window,text=str(count1_8))
k_8.config(command=c1_8)
k_8.config(bg='#94e09a')
k_8.grid(row=8,column=4)

k1_8 = Button(window,text=str(count2_8))
k1_8.config(command=c2_8)
k1_8.config(bg='#94e09a')
k1_8.grid(row=8,column=5)

k2_8 = Button(window,text=str(count3_8))
k2_8.config(command=c3_8)
k2_8.config(bg='#94e09a')
k2_8.grid(row=8,column=6)

k3_8 = Button(window,text=str(count4_8))
k3_8.config(command=c4_8)
k3_8.config(bg='#94e09a')
k3_8.grid(row=8,column=7)

k4_8 = Button(window,text=str(count5_8))
k4_8.config(command=c5_8)
k4_8.config(bg='#94e09a')
k4_8.grid(row=8,column=8)

#94d1e0
a_8 = Button(window,text=str(count6_8))
a_8.config(command=c6_8)
a_8.config(bg='#94d1e0')
a_8.grid(row=8,column=9)

a1_8 = Button(window,text=str(count7_8))
a1_8.config(command=c7_8)
a1_8.config(bg='#94d1e0')
a1_8.grid(row=8,column=10)

a2_8 = Button(window,text=str(count8_8))
a2_8.config(command=c8_8)
a2_8.config(bg='#94d1e0')
a2_8.grid(row=8,column=11)

a3_8 = Button(window,text=str(count9_8))
a3_8.config(command=c9_8)
a3_8.config(bg='#94d1e0')
a3_8.grid(row=8,column=12)

#e09494
d_8 = Button(window,text=str(count10_8))
d_8.config(command=c10_8)
d_8.config(bg='#e09494')
d_8.grid(row=8,column=13)

d1_8 = Button(window,text=str(count11_8))
d1_8.config(command=c11_8)
d1_8.config(bg='#e09494')
d1_8.grid(row=8,column=14)

d2_8 = Button(window,text=str(count12_8))
d2_8.config(command=c12_8)
d2_8.config(bg='#e09494')
d2_8.grid(row=8,column=15)

d3_8 = Button(window,text=str(count13_8))
d3_8.config(command=c13_8)
d3_8.config(bg='#e09494')
d3_8.grid(row=8,column=16)

d4_8 = Button(window,text=str(count14_8))
d4_8.config(command=c14_8)
d4_8.config(bg='#e09494')
d4_8.grid(row=8,column=17)

#Player9
k_9 = Button(window,text=str(count1_9))
k_9.config(command=c1_9)
k_9.config(bg='#94e09a')
k_9.grid(row=9,column=4)

k1_9 = Button(window,text=str(count2_9))
k1_9.config(command=c2_9)
k1_9.config(bg='#94e09a')
k1_9.grid(row=9,column=5)

k2_9 = Button(window,text=str(count3_9))
k2_9.config(command=c3_9)
k2_9.config(bg='#94e09a')
k2_9.grid(row=9,column=6)

k3_9 = Button(window,text=str(count4_9))
k3_9.config(command=c4_9)
k3_9.config(bg='#94e09a')
k3_9.grid(row=9,column=7)

k4_9 = Button(window,text=str(count5_9))
k4_9.config(command=c5_9)
k4_9.config(bg='#94e09a')
k4_9.grid(row=9,column=8)

#94d1e0
a_9 = Button(window,text=str(count6_9))
a_9.config(command=c6_9)
a_9.config(bg='#94d1e0')
a_9.grid(row=9,column=9)

a1_9 = Button(window,text=str(count7_9))
a1_9.config(command=c7_9)
a1_9.config(bg='#94d1e0')
a1_9.grid(row=9,column=10)

a2_9 = Button(window,text=str(count8_9))
a2_9.config(command=c8_9)
a2_9.config(bg='#94d1e0')
a2_9.grid(row=9,column=11)

a3_9 = Button(window,text=str(count9_9))
a3_9.config(command=c9_9)
a3_9.config(bg='#94d1e0')
a3_9.grid(row=9,column=12)

#e09494
d_9 = Button(window,text=str(count10_9))
d_9.config(command=c10_9)
d_9.config(bg='#e09494')
d_9.grid(row=9,column=13)

d1_9 = Button(window,text=str(count11_9))
d1_9.config(command=c11_9)
d1_9.config(bg='#e09494')
d1_9.grid(row=9,column=14)

d2_9 = Button(window,text=str(count12_9))
d2_9.config(command=c12_9)
d2_9.config(bg='#e09494')
d2_9.grid(row=9,column=15)

d3_9 = Button(window,text=str(count13_9))
d3_9.config(command=c13_9)
d3_9.config(bg='#e09494')
d3_9.grid(row=9,column=16)

d4_9 = Button(window,text=str(count14_9))
d4_9.config(command=c14_9)
d4_9.config(bg='#e09494')
d4_9.grid(row=9,column=17)

#Player10
k_10 = Button(window,text=str(count1_10))
k_10.config(command=c1_10)
k_10.config(bg='#94e09a')
k_10.grid(row=10,column=4)

k1_10 = Button(window,text=str(count2_10))
k1_10.config(command=c2_10)
k1_10.config(bg='#94e09a')
k1_10.grid(row=10,column=5)

k2_10 = Button(window,text=str(count3_10))
k2_10.config(command=c3_10)
k2_10.config(bg='#94e09a')
k2_10.grid(row=10,column=6)

k3_10 = Button(window,text=str(count4_10))
k3_10.config(command=c4_10)
k3_10.config(bg='#94e09a')
k3_10.grid(row=10,column=7)

k4_10 = Button(window,text=str(count5_10))
k4_10.config(command=c5_10)
k4_10.config(bg='#94e09a')
k4_10.grid(row=10,column=8)

#94d1e0
a_10 = Button(window,text=str(count6_10))
a_10.config(command=c6_10)
a_10.config(bg='#94d1e0')
a_10.grid(row=10,column=9)

a1_10 = Button(window,text=str(count7_10))
a1_10.config(command=c7_10)
a1_10.config(bg='#94d1e0')
a1_10.grid(row=10,column=10)

a2_10 = Button(window,text=str(count8_10))
a2_10.config(command=c8_10)
a2_10.config(bg='#94d1e0')
a2_10.grid(row=10,column=11)

a3_10 = Button(window,text=str(count9_10))
a3_10.config(command=c9_10)
a3_10.config(bg='#94d1e0')
a3_10.grid(row=10,column=12)

#e09494
d_10 = Button(window,text=str(count10_10))
d_10.config(command=c10_10)
d_10.config(bg='#e09494')
d_10.grid(row=10,column=13)

d1_10 = Button(window,text=str(count11_10))
d1_10.config(command=c11_10)
d1_10.config(bg='#e09494')
d1_10.grid(row=10,column=14)

d2_10 = Button(window,text=str(count12_10))
d2_10.config(command=c12_10)
d2_10.config(bg='#e09494')
d2_10.grid(row=10,column=15)

d3_10 = Button(window,text=str(count13_10))
d3_10.config(command=c13_10)
d3_10.config(bg='#e09494')
d3_10.grid(row=10,column=16)

d4_10 = Button(window,text=str(count14_10))
d4_10.config(command=c14_10)
d4_10.config(bg='#e09494')
d4_10.grid(row=10,column=17)

#Inputable Game Info
def open():
    #New Window
    top = Toplevel()
    top.title('Game Info')
    top.geometry("400x300")
    
    #Labels
    team_label = Label(top, text="Team")
    team_label.grid(row=0,column=0)
    team_label_1 = Label(top, text="Team")
    team_label_1.grid(row=5,column=0)

    player_label = Label(top, text="Player")
    player_label.grid(row=0,column=1)

    Champ_label = Label(top, text="Champ")
    Champ_label.grid(row=0,column=2)
    
    result_label = Label(top, text="Result")
    result_label.grid(row=2,column=0)
    result_label_1 = Label(top, text="Result")
    result_label_1.grid(row=7,column=0)
    
    gameid_label = Label(top, text="GameID")
    gameid_label.grid(row=0,column=3)
    
    savefile_label = Label(top, text="FileName")
    savefile_label.grid(row=2,column=3)
    
    #Inputable Boxes
    #GameID
    gameid = Entry(top, width=15)
    gameid.grid(row=1,column=3)
    
    #Result
    result_1 = Entry(top, width=15)
    result_1.grid(row=3,column=0)
    result_2 = Entry(top, width=15)
    result_2.grid(row=8,column=0)
    
    #FileName
    savefile = Entry(top, width=15)
    savefile.grid(row=3,column=3)
    
    #Player1
    team1 = Entry(top, width=15)
    team1.grid(row=1,column=0)
    player1 = Entry(top, width=15)
    player1.grid(row=1,column=1)
    champ1 = Entry(top, width=15)
    champ1.grid(row=1,column=2)
    
    #Player2
    player1_2 = Entry(top, width=15)
    player1_2.grid(row=2,column=1)
    champ1_2 = Entry(top, width=15)
    champ1_2.grid(row=2,column=2)
    
    #Player3
    player1_3 = Entry(top, width=15)
    player1_3.grid(row=3,column=1)
    champ1_3 = Entry(top, width=15)
    champ1_3.grid(row=3,column=2)
    
    #Player4
    player1_4 = Entry(top, width=15)
    player1_4.grid(row=4,column=1)
    champ1_4 = Entry(top, width=15)
    champ1_4.grid(row=4,column=2)
    
    #Player5
    player1_5 = Entry(top, width=15)
    player1_5.grid(row=5,column=1)
    champ1_5 = Entry(top, width=15)
    champ1_5.grid(row=5,column=2)
    
    #Player6
    team1_6 = Entry(top, width=15)
    team1_6.grid(row=6,column=0)
    player1_6 = Entry(top, width=15)
    player1_6.grid(row=6,column=1)
    champ1_6 = Entry(top, width=15)
    champ1_6.grid(row=6,column=2)
    
    #Player7
    player1_7 = Entry(top, width=15)
    player1_7.grid(row=7,column=1)
    champ1_7 = Entry(top, width=15)
    champ1_7.grid(row=7,column=2)

    #Player8
    player1_8 = Entry(top, width=15)
    player1_8.grid(row=8,column=1)
    champ1_8 = Entry(top, width=15)
    champ1_8.grid(row=8,column=2)

    #Player9
    player1_9 = Entry(top, width=15)
    player1_9.grid(row=9,column=1)
    champ1_9 = Entry(top, width=15)
    champ1_9.grid(row=9,column=2)
    
    #Player10
    player1_10 = Entry(top, width=15)
    player1_10.grid(row=10,column=1)
    champ1_10 = Entry(top, width=15)
    champ1_10.grid(row=10,column=2)

#Makes submit button enter the info
    def fclick():
        global t1
        t1 = team1.get()
        t1_l = Label(window, text=t1)
        t1_l.grid(row=1,column=1)
        t1_l.config(bg="#eda398")
        global p1
        p1 = player1.get()
        p1_l = Label(window, text=p1)
        p1_l.grid(row=1,column=2)
        p1_l.config(bg="#eda398")
        global c1
        c1 = champ1.get()
        c1_l = Label(window, text=c1)
        c1_l.grid(row=1,column=3)
        c1_l.config(bg="#eda398")
        #Player2
        global t1_2
        t1_2 = team1.get()
        t1_l_2 = Label(window, text=t1_2)
        t1_l_2.grid(row=2,column=1)
        t1_l_2.config(bg="#eda398")
        global p1_2
        p1_2 = player1_2.get()
        p1_l_2 = Label(window, text=p1_2)
        p1_l_2.grid(row=2,column=2)
        p1_l_2.config(bg="#eda398")
        global c1_2
        c1_2 = champ1_2.get()
        c1_l_2 = Label(window, text=c1_2)
        c1_l_2.grid(row=2,column=3)
        c1_l_2.config(bg="#eda398")
        #Player3
        global t1_3
        t1_3 = team1.get()
        t1_l_3 = Label(window, text=t1_3)
        t1_l_3.grid(row=3,column=1)
        t1_l_3.config(bg="#eda398")
        global p1_3
        p1_3 = player1_3.get()
        p1_l_3 = Label(window, text=p1_3)
        p1_l_3.grid(row=3,column=2)
        p1_l_3.config(bg="#eda398")
        global c1_3
        c1_3 = champ1_3.get()
        c1_l_3 = Label(window, text=c1_3)
        c1_l_3.grid(row=3,column=3)
        c1_l_3.config(bg="#eda398")
        #Player4
        global t1_4
        t1_4 = team1.get()
        t1_l_4 = Label(window, text=t1_4)
        t1_l_4.grid(row=4,column=1)
        t1_l_4.config(bg="#eda398")
        global p1_4
        p1_4 = player1_4.get()
        p1_l_4 = Label(window, text=p1_4)
        p1_l_4.grid(row=4,column=2)
        p1_l_4.config(bg="#eda398")
        global c1_4
        c1_4 = champ1_4.get()
        c1_l_4 = Label(window, text=c1_4)
        c1_l_4.grid(row=4,column=3)
        c1_l_4.config(bg="#eda398")
        #Player5
        global t1_5
        t1_5 = team1.get()
        t1_l_5 = Label(window, text=t1_5)
        t1_l_5.grid(row=5,column=1)
        t1_l_5.config(bg="#eda398")
        global p1_5
        p1_5 = player1_5.get()
        p1_l_5 = Label(window, text=p1_5)
        p1_l_5.grid(row=5,column=2)
        p1_l_5.config(bg="#eda398")
        global c1_5
        c1_5 = champ1_5.get()
        c1_l_5= Label(window, text=c1_5)
        c1_l_5.grid(row=5,column=3)
        c1_l_5.config(bg="#eda398")
        #Player6
        global t1_6
        t1_6 = team1_6.get()
        t1_l_6 = Label(window, text=t1_6)
        t1_l_6.grid(row=6,column=1)
        t1_l_6.config(bg="#90b8e8")
        global p1_6
        p1_6 = player1_6.get()
        p1_l_6 = Label(window, text=p1_6)
        p1_l_6.grid(row=6,column=2)
        p1_l_6.config(bg="#90b8e8")
        global c1_6
        c1_6 = champ1_6.get()
        c1_l_6 = Label(window, text=c1_6)
        c1_l_6.grid(row=6,column=3)
        c1_l_6.config(bg="#90b8e8")
        #Player7
        global t1_7
        t1_7 = team1_6.get()
        t1_l_7 = Label(window, text=t1_7)
        t1_l_7.grid(row=7,column=1)
        t1_l_7.config(bg="#90b8e8")
        global p1_7
        p1_7 = player1_7.get()
        p1_l_7 = Label(window, text=p1_7)
        p1_l_7.grid(row=7,column=2)
        p1_l_7.config(bg="#90b8e8")
        global c1_7
        c1_7 = champ1_7.get()
        c1_l_7 = Label(window, text=c1_7)
        c1_l_7.grid(row=7,column=3)
        c1_l_7.config(bg="#90b8e8")
        #Player8
        global t1_8
        t1_8 = team1_6.get()
        t1_l_8 = Label(window, text=t1_8)
        t1_l_8.grid(row=8,column=1)
        t1_l_8.config(bg="#90b8e8")
        global p1_8
        p1_8 = player1_8.get()
        p1_l_8 = Label(window, text=p1_8)
        p1_l_8.grid(row=8,column=2)
        p1_l_8.config(bg="#90b8e8")
        global c1_8
        c1_8 = champ1_8.get()
        c1_l_8 = Label(window, text=c1_8)
        c1_l_8.grid(row=8,column=3)
        c1_l_8.config(bg="#90b8e8")
        #Player9
        global t1_9
        t1_9 = team1_6.get()
        t1_l_9 = Label(window, text=t1_9)
        t1_l_9.grid(row=9,column=1)
        t1_l_9.config(bg="#90b8e8")
        global p1_9
        p1_9 = player1_9.get()
        p1_l_9 = Label(window, text=p1_9)
        p1_l_9.grid(row=9,column=2)
        p1_l_9.config(bg="#90b8e8")
        global c1_9
        c1_9 = champ1_9.get()
        c1_l_9 = Label(window, text=c1_9)
        c1_l_9.grid(row=9,column=3)
        c1_l_9.config(bg="#90b8e8")
        #Player10
        global t1_10
        t1_10 = team1_6.get()
        t1_l_10 = Label(window, text=t1_10)
        t1_l_10.grid(row=10,column=1)
        t1_l_10.config(bg="#90b8e8")
        global p1_10
        p1_10 = player1_10.get()
        p1_l_10 = Label(window, text=p1_10)
        p1_l_10.grid(row=10,column=2)
        p1_l_10.config(bg="#90b8e8")
        global c1_10
        c1_10 = champ1_10.get()
        c1_l_10 = Label(window, text=c1_10)
        c1_l_10.grid(row=10,column=3)
        c1_l_10.config(bg="#90b8e8")
        #Other
        global g_id
        g_id = gameid.get()
        g_id_1 = Label(window, text=g_id)
        g_id_1.grid(row=11,column=1)
        global f_name
        f_name = savefile.get()
        f_name_1 = Label(window, text=f_name)
        f_name_1.grid(row=12,column=1)
        global r_1
        r_1 = result_1.get()
        r_1_1 = Label(window, text=r_1)
        r_1_1.grid(row=13,column=1)
        global r_2
        r_2 = result_2.get()
        r_2_1 = Label(window, text=r_2)
        r_2_1.grid(row=14,column=1)

    submit = Button(top, text="Enter", command=fclick)
    submit.grid(row=11,column=1)
    
#Opens Game info    
btn = Button(window,text='Game Info', command=open)
btn.grid(row=0,column=0)


window.mainloop()


outWorkbook = xlsxwriter.Workbook(f_name + '.xlsx')
outSheet = outWorkbook.add_worksheet(f_name)

#Headers
outSheet.write("A1", "GameId")
outSheet.write("B1", "Team")
outSheet.write("C1", "Player")
outSheet.write("D1", "Champion")
outSheet.write("E1", "Role")
outSheet.write("F1", "R")
outSheet.write("G1", "K")
outSheet.write("H1", "D")
outSheet.write("I1", "A")
outSheet.write("J1", "K")
outSheet.write("K1", "K+1")
outSheet.write("L1", "K+2")
outSheet.write("M1", "K+3")
outSheet.write("N1", "K+4")
outSheet.write("O1", "A")
outSheet.write("P1", "A+1")
outSheet.write("Q1", "A+2")
outSheet.write("R1", "A+3")
outSheet.write("S1", "D")
outSheet.write("T1", "D+1")
outSheet.write("U1", "D+2")
outSheet.write("V1", "D+3")
outSheet.write("W1", "D+4")
outSheet.write("X1", "KDA")
outSheet.write("Y1", "KP")
outSheet.write("Z1", "KAV")
outSheet.write("AA1", "KV")
outSheet.write("AB1", "KV-DV")
outSheet.write("AC1", "DV")
outSheet.write("AD1", "AV")



#RolesList
roles = ['Top','Jungle','Mid','Bottom','Support']

#Game ID Write
for item in range(0,12):
    outSheet.write(item+1,0,g_id)
    
for item in range(len(roles)):
    outSheet.write(item+1,4,roles[item])
    outSheet.write(item+1,5,r_1)
    outSheet.write(item+6,4,roles[item])
    outSheet.write(item+6,5,r_2)

#Team
outSheet.write("B12",t1)
outSheet.write("C12","Team")
outSheet.write("F12",r_1)
outSheet.write("G12","=SUM(G2:G6)")
outSheet.write("H12","=SUM(H2:H6)")
outSheet.write("I12","=SUM(I2:I6)")
outSheet.write("J12","=SUM(J2:J6)")
outSheet.write("K12","=SUM(K2:K6)")
outSheet.write("L12","=SUM(L2:L6)")
outSheet.write("M12","=SUM(M2:M6)")
outSheet.write("N12","=SUM(N2:N6)")
outSheet.write("O12","=SUM(O2:O6)")
outSheet.write("P12","=SUM(P2:P6)")
outSheet.write("Q12","=SUM(Q2:Q6)")
outSheet.write("R12","=SUM(R2:R6)")
outSheet.write("S12","=SUM(S2:S6)")
outSheet.write("T12","=SUM(T2:T6)")
outSheet.write("U12","=SUM(U2:U6)")
outSheet.write("V12","=SUM(V2:V6)")
outSheet.write("W12","=SUM(W2:W6)")
outSheet.write("X12","=(G12+I12)/(IF(H12=0,1,H12))")
outSheet.write("Y12","=AVERAGE(Y2:Y6)")
outSheet.write("Z12", "=(J12*1)+(K12*0.9)+(L12*0.8)+(M12*0.7)+(N12*0.6)+(O12*0.5)+(P12*0.4)+(Q12*0.3)+(R12*0.2)")
outSheet.write("AA12", "=(J12*1)+(K12*0.9)+(L12*0.8)+(M12*0.7)+(N12*0.6)")
outSheet.write("AB12", "=(J12*1)+(K12*0.9)+(L12*0.8)+(M12*0.7)+(N12*0.6)+(O12*0.5)+(P12*0.4)+(Q12*0.3)+(R12*0.2)-(S12*0.5)-(T12*0.4)-(U12*0.3)-(V12*0.2)-(W12*0.1)")
outSheet.write("AC12", "=(S12*0.5)+(T12*0.4)+(U12*0.3)+(V12*0.2)+(W12*0.1)")
outSheet.write("AD12", "=(O12*0.5)+(P12*0.4)+(Q12*0.3)+(R12*0.2)")


outSheet.write("B13",t1_6)
outSheet.write("C13","Team")
outSheet.write("F13",r_2)
outSheet.write("G13","=SUM(G7:G11)")
outSheet.write("H13","=SUM(H7:H11)")
outSheet.write("I13","=SUM(I7:I11)")
outSheet.write("J13","=SUM(J7:J11)")
outSheet.write("K13","=SUM(K7:K11)")
outSheet.write("L13","=SUM(L7:L11)")
outSheet.write("M13","=SUM(M7:M11)")
outSheet.write("N13","=SUM(N7:N11)")
outSheet.write("O13","=SUM(O7:O11)")
outSheet.write("P13","=SUM(P7:P11)")
outSheet.write("Q13","=SUM(Q7:Q11)")
outSheet.write("R13","=SUM(R7:R11)")
outSheet.write("S13","=SUM(S7:S11)")
outSheet.write("T13","=SUM(T7:T11)")
outSheet.write("U13","=SUM(U7:U11)")
outSheet.write("V13","=SUM(V7:V11)")
outSheet.write("W13","=SUM(W7:W11)")
outSheet.write("X13","=(G13+I13)/(IF(H13=0,1,H13))")
outSheet.write("Y13","=AVERAGE(Y7:Y11)")
outSheet.write("Z13", "=(J13*1)+(K13*0.9)+(L13*0.8)+(M13*0.7)+(N13*0.6)+(O13*0.5)+(P13*0.4)+(Q13*0.3)+(R13*0.2)")
outSheet.write("AA13", "=(J13*1)+(K13*0.9)+(L13*0.8)+(M13*0.7)+(N13*0.6)")
outSheet.write("AB13", "=(J13*1)+(K13*0.9)+(L13*0.8)+(M13*0.7)+(N13*0.6)+(O13*0.5)+(P13*0.4)+(Q13*0.3)+(R13*0.2)-(S13*0.5)-(T13*0.4)-(U13*0.3)-(V13*0.2)-(W13*0.1)")
outSheet.write("AC13", "=(S13*0.5)+(T13*0.4)+(U13*0.3)+(V13*0.2)+(W13*0.1)")
outSheet.write("AD13", "=(O13*0.5)+(P13*0.4)+(Q13*0.3)+(R13*0.2)")

# #Writing Game Info Input
outSheet.write("B2",t1)
outSheet.write("C2",p1)
outSheet.write("D2",c1)
outSheet.write("J2",count1)
outSheet.write("K2",count2)
outSheet.write("L2",count3)
outSheet.write("M2",count4)
outSheet.write("N2",count5)
outSheet.write("O2",count6)
outSheet.write("P2",count7)
outSheet.write("Q2",count8)
outSheet.write("R2",count9)
outSheet.write("S2",count10)
outSheet.write("T2",count11)
outSheet.write("U2",count12)
outSheet.write("V2",count13)
outSheet.write("W2",count14)
#player2
outSheet.write("B3",t1_2)
outSheet.write("C3",p1_2)
outSheet.write("D3",c1_2)
outSheet.write("J3",count1_2)
outSheet.write("K3",count2_2)
outSheet.write("L3",count3_2)
outSheet.write("M3",count4_2)
outSheet.write("N3",count5_2)
outSheet.write("O3",count6_2)
outSheet.write("P3",count7_2)
outSheet.write("Q3",count8_2)
outSheet.write("R3",count9_2)
outSheet.write("S3",count10_2)
outSheet.write("T3",count11_2)
outSheet.write("U3",count12_2)
outSheet.write("V3",count13_2)
outSheet.write("W3",count14_2)
#player3
outSheet.write("B4",t1_3)
outSheet.write("C4",p1_3)
outSheet.write("D4",c1_3)
outSheet.write("J4",count1_3)
outSheet.write("K4",count2_3)
outSheet.write("L4",count3_3)
outSheet.write("M4",count4_3)
outSheet.write("N4",count5_3)
outSheet.write("O4",count6_3)
outSheet.write("P4",count7_3)
outSheet.write("Q4",count8_3)
outSheet.write("R4",count9_3)
outSheet.write("S4",count10_3)
outSheet.write("T4",count11_3)
outSheet.write("U4",count12_3)
outSheet.write("V4",count13_3)
outSheet.write("W4",count14_3)
#player4
outSheet.write("B5",t1_4)
outSheet.write("C5",p1_4)
outSheet.write("D5",c1_4)
outSheet.write("J5",count1_4)
outSheet.write("K5",count2_4)
outSheet.write("L5",count3_4)
outSheet.write("M5",count4_4)
outSheet.write("N5",count5_4)
outSheet.write("O5",count6_4)
outSheet.write("P5",count7_4)
outSheet.write("Q5",count8_4)
outSheet.write("R5",count9_4)
outSheet.write("S5",count10_4)
outSheet.write("T5",count11_4)
outSheet.write("U5",count12_4)
outSheet.write("V5",count13_4)
outSheet.write("W5",count14_4)
#player5
outSheet.write("B6",t1_5)
outSheet.write("C6",p1_5)
outSheet.write("D6",c1_5)
outSheet.write("J6",count1_5)
outSheet.write("K6",count2_5)
outSheet.write("L6",count3_5)
outSheet.write("M6",count4_5)
outSheet.write("N6",count5_5)
outSheet.write("O6",count6_5)
outSheet.write("P6",count7_5)
outSheet.write("Q6",count8_5)
outSheet.write("R6",count9_5)
outSheet.write("S6",count10_5)
outSheet.write("T6",count11_5)
outSheet.write("U6",count12_5)
outSheet.write("V6",count13_5)
outSheet.write("W6",count14_5)
#player6
outSheet.write("B7",t1_6)
outSheet.write("C7",p1_6)
outSheet.write("D7",c1_6)
outSheet.write("J7",count1_6)
outSheet.write("K7",count2_6)
outSheet.write("L7",count3_6)
outSheet.write("M7",count4_6)
outSheet.write("N7",count5_6)
outSheet.write("O7",count6_6)
outSheet.write("P7",count7_6)
outSheet.write("Q7",count8_6)
outSheet.write("R7",count9_6)
outSheet.write("S7",count10_6)
outSheet.write("T7",count11_6)
outSheet.write("U7",count12_6)
outSheet.write("V7",count13_6)
outSheet.write("W7",count14_6)
#player7
outSheet.write("B8",t1_7)
outSheet.write("C8",p1_7)
outSheet.write("D8",c1_7)
outSheet.write("J8",count1_7)
outSheet.write("K8",count2_7)
outSheet.write("L8",count3_7)
outSheet.write("M8",count4_7)
outSheet.write("N8",count5_7)
outSheet.write("O8",count6_7)
outSheet.write("P8",count7_7)
outSheet.write("Q8",count8_7)
outSheet.write("R8",count9_7)
outSheet.write("S8",count10_7)
outSheet.write("T8",count11_7)
outSheet.write("U8",count12_7)
outSheet.write("V8",count13_7)
outSheet.write("W8",count14_7)
#player8
outSheet.write("B9",t1_8)
outSheet.write("C9",p1_8)
outSheet.write("D9",c1_8)
outSheet.write("J9",count1_8)
outSheet.write("K9",count2_8)
outSheet.write("L9",count3_8)
outSheet.write("M9",count4_8)
outSheet.write("N9",count5_8)
outSheet.write("O9",count6_8)
outSheet.write("P9",count7_8)
outSheet.write("Q9",count8_8)
outSheet.write("R9",count9_8)
outSheet.write("S9",count10_8)
outSheet.write("T9",count11_8)
outSheet.write("U9",count12_8)
outSheet.write("V9",count13_8)
outSheet.write("W9",count14_8)
#player9
outSheet.write("B10",t1_9)
outSheet.write("C10",p1_9)
outSheet.write("D10",c1_9)
outSheet.write("J10",count1_9)
outSheet.write("K10",count2_9)
outSheet.write("L10",count3_9)
outSheet.write("M10",count4_9)
outSheet.write("N10",count5_9)
outSheet.write("O10",count6_9)
outSheet.write("P10",count7_9)
outSheet.write("Q10",count8_9)
outSheet.write("R10",count9_9)
outSheet.write("S10",count10_9)
outSheet.write("T10",count11_9)
outSheet.write("U10",count12_9)
outSheet.write("V10",count13_9)
outSheet.write("W10",count14_9)
#player10
outSheet.write("B11",t1_10)
outSheet.write("C11",p1_10)
outSheet.write("D11",c1_10)
outSheet.write("J11",count1_10)
outSheet.write("K11",count2_10)
outSheet.write("L11",count3_10)
outSheet.write("M11",count4_10)
outSheet.write("N11",count5_10)
outSheet.write("O11",count6_10)
outSheet.write("P11",count7_10)
outSheet.write("Q11",count8_10)
outSheet.write("R11",count9_10)
outSheet.write("S11",count10_10)
outSheet.write("T11",count11_10)
outSheet.write("U11",count12_10)
outSheet.write("V11",count13_10)
outSheet.write("W11",count14_10)

#Forumlas
#K
outSheet.write_formula("G2", "=SUM(J2:N2)")
outSheet.write_formula("G3", "=SUM(J3:N3)")
outSheet.write_formula("G4", "=SUM(J4:N4)")
outSheet.write_formula("G5", "=SUM(J5:N5)")
outSheet.write_formula("G6", "=SUM(J6:N6)")
outSheet.write_formula("G7", "=SUM(J7:N7)")
outSheet.write_formula("G8", "=SUM(J8:N8)")
outSheet.write_formula("G9", "=SUM(J9:N9)")
outSheet.write_formula("G10", "=SUM(J10:N10)")
outSheet.write_formula("G11", "=SUM(J11:N11)")

#A
outSheet.write_formula("I2", "=SUM(O2:R2)")
outSheet.write_formula("I3", "=SUM(O3:R3)")
outSheet.write_formula("I4", "=SUM(O4:R4)")
outSheet.write_formula("I5", "=SUM(O5:R5)")
outSheet.write_formula("I6", "=SUM(O6:R6)")
outSheet.write_formula("I7", "=SUM(O7:R7)")
outSheet.write_formula("I8", "=SUM(O8:R8)")
outSheet.write_formula("I9", "=SUM(O9:R9)")
outSheet.write_formula("I10", "=SUM(O10:R10)")
outSheet.write_formula("I11", "=SUM(O11:R11)")

#D
outSheet.write_formula("H2", "=SUM(S2:W2)")
outSheet.write_formula("H3", "=SUM(S3:W3)")
outSheet.write_formula("H4", "=SUM(S4:W4)")
outSheet.write_formula("H5", "=SUM(S5:W5)")
outSheet.write_formula("H6", "=SUM(S6:W6)")
outSheet.write_formula("H7", "=SUM(S7:W7)")
outSheet.write_formula("H8", "=SUM(S8:W8)")
outSheet.write_formula("H9", "=SUM(S9:W9)")
outSheet.write_formula("H10", "=SUM(S10:W10)")
outSheet.write_formula("H11", "=SUM(S11:W11)")

#KDA
outSheet.write_formula("X2", "=(G2+I2)/(IF(H2=0,1,H2))")
outSheet.write_formula("X3", "=(G3+I3)/(IF(H3=0,1,H3))")
outSheet.write_formula("X4", "=(G4+I4)/(IF(H4=0,1,H4))")
outSheet.write_formula("X5", "=(G5+I5)/(IF(H5=0,1,H5))")
outSheet.write_formula("X6", "=(G6+I6)/(IF(H6=0,1,H6))")
outSheet.write_formula("X7", "=(G7+I7)/(IF(H7=0,1,H7))")
outSheet.write_formula("X8", "=(G8+I8)/(IF(H8=0,1,H8))")
outSheet.write_formula("X9", "=(G9+I9)/(IF(H9=0,1,H9))")
outSheet.write_formula("X10", "=(G10+I10)/(IF(H10=0,1,H10))")
outSheet.write_formula("X11", "=(G11+I11)/(IF(H11=0,1,H11))")

#KP
outSheet.write_formula("Y2", "=(G2+I2)/(IF(G12=0,1,G12))")
outSheet.write_formula("Y3", "=(G3+I3)/(IF(G12=0,1,G12))")
outSheet.write_formula("Y4", "=(G4+I4)/(IF(G12=0,1,G12))")
outSheet.write_formula("Y5", "=(G5+I5)/(IF(G12=0,1,G12))")
outSheet.write_formula("Y6", "=(G6+I6)/(IF(G12=0,1,G12))")

outSheet.write_formula("Y7", "=(G7+I7)/(IF(G13=0,1,G13))")
outSheet.write_formula("Y8", "=(G8+I8)/(IF(G13=0,1,G13))")
outSheet.write_formula("Y9", "=(G9+I9)/(IF(G13=0,1,G13))")
outSheet.write_formula("Y10", "=(G10+I10)/(IF(G13=0,1,G13))")
outSheet.write_formula("Y11", "=(G11+I11)/(IF(G13=0,1,G13))")

#KAV
outSheet.write_formula("Z2", "=(J2*1)+(K2*0.9)+(L2*0.8)+(M2*0.7)+(N2*0.6)+(O2*0.5)+(P2*0.4)+(Q2*0.3)+(R2*0.2)")
outSheet.write_formula("Z3", "=(J3*1)+(K3*0.9)+(L3*0.8)+(M3*0.7)+(N3*0.6)+(O3*0.5)+(P3*0.4)+(Q3*0.3)+(R3*0.2)")
outSheet.write_formula("Z4", "=(J4*1)+(K4*0.9)+(L4*0.8)+(M4*0.7)+(N4*0.6)+(O4*0.5)+(P4*0.4)+(Q4*0.3)+(R4*0.2)")
outSheet.write_formula("Z5", "=(J5*1)+(K5*0.9)+(L5*0.8)+(M5*0.7)+(N5*0.6)+(O5*0.5)+(P5*0.4)+(Q5*0.3)+(R5*0.2)")
outSheet.write_formula("Z6", "=(J6*1)+(K6*0.9)+(L6*0.8)+(M6*0.7)+(N6*0.6)+(O6*0.5)+(P6*0.4)+(Q6*0.3)+(R6*0.2)")
outSheet.write_formula("Z7", "=(J7*1)+(K7*0.9)+(L7*0.8)+(M7*0.7)+(N7*0.6)+(O7*0.5)+(P7*0.4)+(Q7*0.3)+(R7*0.2)")
outSheet.write_formula("Z8", "=(J8*1)+(K8*0.9)+(L8*0.8)+(M8*0.7)+(N8*0.6)+(O8*0.5)+(P8*0.4)+(Q8*0.3)+(R8*0.2)")
outSheet.write_formula("Z9", "=(J9*1)+(K9*0.9)+(L9*0.8)+(M9*0.7)+(N9*0.6)+(O9*0.5)+(P9*0.4)+(Q9*0.3)+(R9*0.2)")
outSheet.write_formula("Z10", "=(J10*1)+(K10*0.9)+(L10*0.8)+(M10*0.7)+(N10*0.6)+(O10*0.5)+(P10*0.4)+(Q10*0.3)+(R10*0.2)")
outSheet.write_formula("Z11", "=(J11*1)+(K11*0.9)+(L11*0.8)+(M11*0.7)+(N11*0.6)+(O11*0.5)+(P11*0.4)+(Q11*0.3)+(R11*0.2)")

#KV
outSheet.write_formula("AA2", "=(J2*1)+(K2*0.9)+(L2*0.8)+(M2*0.7)+(N2*0.6)")
outSheet.write_formula("AA3", "=(J3*1)+(K3*0.9)+(L3*0.8)+(M3*0.7)+(N3*0.6)")
outSheet.write_formula("AA4", "=(J4*1)+(K4*0.9)+(L4*0.8)+(M4*0.7)+(N4*0.6)")
outSheet.write_formula("AA5", "=(J5*1)+(K5*0.9)+(L5*0.8)+(M5*0.7)+(N5*0.6)")
outSheet.write_formula("AA6", "=(J6*1)+(K6*0.9)+(L6*0.8)+(M6*0.7)+(N6*0.6)")
outSheet.write_formula("AA7", "=(J7*1)+(K7*0.9)+(L7*0.8)+(M7*0.7)+(N7*0.6)")
outSheet.write_formula("AA8", "=(J8*1)+(K8*0.9)+(L8*0.8)+(M8*0.7)+(N8*0.6)")
outSheet.write_formula("AA9", "=(J9*1)+(K9*0.9)+(L9*0.8)+(M9*0.7)+(N9*0.6)")
outSheet.write_formula("AA10", "=(J10*1)+(K10*0.9)+(L10*0.8)+(M10*0.7)+(N10*0.6)")
outSheet.write_formula("AA11", "=(J11*1)+(K11*0.9)+(L11*0.8)+(M11*0.7)+(N11*0.6)")

#DV
outSheet.write_formula("AC2", "=(S2*1)+(T2*0.9)+(U2*0.8)+(V2*0.7)+(W2*0.6)")
outSheet.write_formula("AC3", "=(S3*1)+(T3*0.9)+(U3*0.8)+(V3*0.7)+(W3*0.6)")
outSheet.write_formula("AC4", "=(S4*1)+(T4*0.9)+(U4*0.8)+(V4*0.7)+(W4*0.6)")
outSheet.write_formula("AC5", "=(S5*1)+(T5*0.9)+(U5*0.8)+(V5*0.7)+(W5*0.6)")
outSheet.write_formula("AC6", "=(S6*1)+(T6*0.9)+(U6*0.8)+(V6*0.7)+(W6*0.6)")
outSheet.write_formula("AC7", "=(S7*1)+(T7*0.9)+(U7*0.8)+(V7*0.7)+(W7*0.6)")
outSheet.write_formula("AC8", "=(S8*1)+(T8*0.9)+(U8*0.8)+(V8*0.7)+(W8*0.6)")
outSheet.write_formula("AC9", "=(S9*1)+(T9*0.9)+(U9*0.8)+(V9*0.7)+(W9*0.6)")
outSheet.write_formula("AC10", "=(S10*1)+(T10*0.9)+(U10*0.8)+(V10*0.7)+(W10*0.6)")
outSheet.write_formula("AC11", "=(S11*1)+(T11*0.9)+(U11*0.8)+(V11*0.7)+(W11*0.6)")

#AV

outSheet.write_formula("AD2", "=(O2*0.5)+(P2*0.4)+(Q2*0.3)+(R2*0.2)")
outSheet.write_formula("AD3", "=(O3*0.5)+(P3*0.4)+(Q3*0.3)+(R3*0.2)")
outSheet.write_formula("AD4", "=(O4*0.5)+(P4*0.4)+(Q4*0.3)+(R4*0.2)")
outSheet.write_formula("AD5", "=(O5*0.5)+(P5*0.4)+(Q5*0.3)+(R5*0.2)")
outSheet.write_formula("AD6", "=(O6*0.5)+(P6*0.4)+(Q6*0.3)+(R6*0.2)")
outSheet.write_formula("AD7", "=(O7*0.5)+(P7*0.4)+(Q7*0.3)+(R7*0.2)")
outSheet.write_formula("AD8", "=(O8*0.5)+(P8*0.4)+(Q8*0.3)+(R8*0.2)")
outSheet.write_formula("AD9", "=(O9*0.5)+(P9*0.4)+(Q9*0.3)+(R9*0.2)")
outSheet.write_formula("AD10", "=(O10*0.5)+(P10*0.4)+(Q10*0.3)+(R10*0.2)")
outSheet.write_formula("AD11", "=(O11*0.5)+(P11*0.4)+(Q11*0.3)+(R11*0.2)")

#KV-DV
outSheet.write_formula("AB2", "=(J2*1)+(K2*0.9)+(L2*0.8)+(M2*0.7)+(N2*0.6)+(O2*0.5)+(P2*0.4)+(Q2*0.3)+(R2*0.2)-(S2*1)-(T2*0.9)-(U2*0.8)-(V2*0.7)-(W2*0.6)")
outSheet.write_formula("AB3", "=(J3*1)+(K3*0.9)+(L3*0.8)+(M3*0.7)+(N3*0.6)+(O3*0.5)+(P3*0.4)+(Q3*0.3)+(R3*0.2)-(S3*1)-(T3*0.9)-(U3*0.8)-(V3*0.7)-(W3*0.6)")
outSheet.write_formula("AB4", "=(J4*1)+(K4*0.9)+(L4*0.8)+(M4*0.7)+(N4*0.6)+(O4*0.5)+(P4*0.4)+(Q4*0.3)+(R4*0.2)-(S4*1)-(T4*0.9)-(U4*0.8)-(V4*0.7)-(W4*0.6)")
outSheet.write_formula("AB5", "=(J5*1)+(K5*0.9)+(L5*0.8)+(M5*0.7)+(N5*0.6)+(O5*0.5)+(P5*0.4)+(Q5*0.3)+(R5*0.2)-(S5*1)-(T5*0.9)-(U5*0.8)-(V5*0.7)-(W5*0.6)")
outSheet.write_formula("AB6", "=(J6*1)+(K6*0.9)+(L6*0.8)+(M6*0.7)+(N6*0.6)+(O6*0.5)+(P6*0.4)+(Q6*0.3)+(R6*0.2)-(S6*1)-(T6*0.9)-(U6*0.8)-(V6*0.7)-(W6*0.6)")
outSheet.write_formula("AB7", "=(J7*1)+(K7*0.9)+(L7*0.8)+(M7*0.7)+(N7*0.6)+(O7*0.5)+(P7*0.4)+(Q7*0.3)+(R7*0.2)-(S7*1)-(T7*0.9)-(U7*0.8)-(V7*0.7)-(W7*0.6)")
outSheet.write_formula("AB8", "=(J8*1)+(K8*0.9)+(L8*0.8)+(M8*0.7)+(N8*0.6)+(O8*0.5)+(P8*0.4)+(Q8*0.3)+(R8*0.2)-(S8*1)-(T8*0.9)-(U8*0.8)-(V8*0.7)-(W8*0.6)")
outSheet.write_formula("AB9", "=(J9*1)+(K9*0.9)+(L9*0.8)+(M9*0.7)+(N9*0.6)+(O9*0.5)+(P9*0.4)+(Q9*0.3)+(R9*0.2)-(S9*1)-(T9*0.9)-(U9*0.8)-(V9*0.7)-(W9*0.6)")
outSheet.write_formula("AB10", "=(J10*1)+(K10*0.9)+(L10*0.8)+(M10*0.7)+(N10*0.6)+(O10*0.5)+(P10*0.4)+(Q10*0.3)+(R10*0.2)-(S10*1)-(T10*0.9)-(U10*0.8)-(V10*0.7)-(W10*0.6)")
outSheet.write_formula("AB11", "=(J11*1)+(K11*0.9)+(L11*0.8)+(M11*0.7)+(N11*0.6)+(O11*0.5)+(P11*0.4)+(Q11*0.3)+(R11*0.2)-(S11*1)-(T11*0.9)-(U11*0.8)-(V11*0.7)-(W11*0.6)")

outWorkbook.close()