import openpyxl
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import math
import csv

from Botswana import ExecuteBotswana
from HongKong import ExecuteHongKong
from Hungary import ExecuteHungary
from Italy import ExecuteItaly
from Malaysia import ExecuteMalaysia
from Mexico import ExecuteMexico
from Namibia import ExecuteNamibia
from Nigeria import ExecuteNigeria
from Philippines import ExecutePhilippines
from Singapore import ExecuteSingapore
from SouthAfrica import ExecuteSouthAfrica
from Taiwan import ExecuteTaiwan
from Ukraine import ExecuteUkraine
from US import ExecuteUS
from Ghana_Access import ExecuteGhanaAccess
from Ghana_Other import ExecuteGhanaOther



def main():

  window = tk.Tk()


  frame1 = tk.Frame(master=window, width=570, height=100,bg="#63b4ff")
  frame1.pack(fill=tk.BOTH, side=tk.TOP, expand=True)

  fileLabel = tk.Label(master=frame1, font=("Arial", 14), text="Enter the file to be cleaned", width=25, bg="#63b4ff")
  fileLabel.place(x=145, y=5)

  fileEntry = tk.Entry(master=frame1, width=60)
  fileEntry.place(x=100, y=30)

  buttonPickFile = tk.Button(master=frame1, text="Pick File", width=15)
  buttonPickFile.place(x=225, y=52)
  def filepicker(event):
    filePicker = askopenfilename()
    fileEntry.insert(0, filePicker)
  buttonPickFile.bind("<Button-1>", filepicker)
  


  frame = tk.Frame(master=window, width=570, height=290, bg="#005bb0")
  frame.pack(fill=tk.BOTH, side=tk.BOTTOM, expand=True)

  buttonBots = tk.Button(master=frame, text="Botswana", width=20)
  buttonBots.place(x=10, y=10)
  buttonHong = tk.Button(master=frame, text="Hong Kong", width=20)
  buttonHong.place(x=210, y=10)
  buttonHung = tk.Button(master=frame, text="Hungary", width=20)
  buttonHung.place(x=410, y=10)
  buttonItal = tk.Button(master=frame, text="Italy", width=20)
  buttonItal.place(x=10, y=70)
  buttonMala = tk.Button(master=frame, text="Malaysia", width=20)
  buttonMala.place(x=210, y=70)
  buttonMexi = tk.Button(master=frame, text="Mexico", width=20)
  buttonMexi.place(x=410, y=70)
  buttonNami = tk.Button(master=frame, text="Namibia", width=20)
  buttonNami.place(x=10, y=130)
  buttonNige = tk.Button(master=frame, text="Nigeria", width=20)
  buttonNige.place(x=210, y=130)
  buttonPhil = tk.Button(master=frame, text="Philippines", width=20)
  buttonPhil.place(x=410, y=130)
  buttonSing = tk.Button(master=frame, text="Singapore", width=20)
  buttonSing.place(x=10, y=190)
  buttonSoAf = tk.Button(master=frame, text="South Africa", width=20)
  buttonSoAf.place(x=210, y=190)
  buttonTaiw = tk.Button(master=frame, text="Taiwan", width=20)
  buttonTaiw.place(x=410, y=190)
  buttonUkra = tk.Button(master=frame, text="Ukraine", width=20)
  buttonUkra.place(x=10, y=250)
  buttonUS = tk.Button(master=frame, text="USA", width=20)
  buttonUS.place(x=210, y=250)
  buttonGhanaAccess = tk.Button(master=frame, text="Ghana Access", width=20)
  buttonGhanaAccess.place(x=410, y=260)
  buttonGhanaOther = tk.Button(master=frame, text="Ghana Other", width=20)
  buttonGhanaOther.place(x=410, y=230)
  
  def botswana(event):
    fileName = fileEntry.get()
    ExecuteBotswana(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonBots.bind("<Button-1>", botswana)


  def hongkong(event):
    fileName = fileEntry.get()
    ExecuteHongKong(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonHong.bind("<Button-1>", hongkong)


  def hungary(event):
    fileName = fileEntry.get()
    ExecuteHungary(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonHung.bind("<Button-1>", hungary)


  def italy(event):
    fileName = fileEntry.get()
    ExecuteItaly(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonItal.bind("<Button-1>", italy)


  def malaysia(event):
    fileName = fileEntry.get()
    ExecuteMalaysia(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonMala.bind("<Button-1>", malaysia)


  def mexico(event):
    fileName = fileEntry.get()
    ExecuteMexico(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonMexi.bind("<Button-1>", mexico)


  def namibia(event):
    fileName = fileEntry.get()
    ExecuteNamibia(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonNami.bind("<Button-1>", namibia)


  def nigeria(event):
    fileName = fileEntry.get()
    ExecuteNigeria(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonNige.bind("<Button-1>", nigeria)


  def philippines(event):
    fileName = fileEntry.get()
    ExecutePhilippines(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonPhil.bind("<Button-1>", philippines)


  def singapore(event):
    fileName = fileEntry.get()
    ExecuteSingapore(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonSing.bind("<Button-1>", singapore)


  def southafrica(event):
    fileName = fileEntry.get()
    ExecuteSouthAfrica(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonSoAf.bind("<Button-1>", southafrica)


  def taiwan(event):
    fileName = fileEntry.get()
    ExecuteTaiwan(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonTaiw.bind("<Button-1>", taiwan)


  def ukraine(event):
    fileName = fileEntry.get()
    ExecuteUkraine(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonUkra.bind("<Button-1>", ukraine)


  def us(event):
    fileName = fileEntry.get()
    ExecuteUS(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonUS.bind("<Button-1>", us)


  def ghanaAccess(event):
    fileName = fileEntry.get()
    ExecuteGhanaAccess(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonGhanaAccess.bind("<Button-1>", ghanaAccess)


  def ghanaOther(event):
    fileName = fileEntry.get()
    ExecuteGhanaOther(fileName)
    fileEntry.delete(0, tk.END)
    popupmsg()
  buttonGhanaOther.bind("<Button-1>", ghanaOther)




  def popupmsg():
    popup = tk.Tk()
    popup.wm_title("!")
    label = ttk.Label(popup, text="Process Complete")
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()


  window.mainloop()
main()
