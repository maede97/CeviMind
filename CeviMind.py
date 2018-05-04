#!/usr/bin/python3
# -*- coding: utf-8 -*-

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkf
import sqlite3 as sql
import os, string, shutil
import docx2txt
import PyPDF2

class CeviMind:
	def __init__(self, master):
		self.version = "1.1.3"
		self.master = master
		self.master.iconbitmap("icon.ico")
		self.filesFolder = "data"
		if(not os.path.exists("database.txt")):
			open("database.txt",'a').close()
			fileName = tkf.askdirectory(title="Wo sollen die Daten gelesen werden?",mustexist=True)
			if(fileName == None or fileName == ""): exit()
			with open("database.txt","w") as writefile:
				writefile.write(fileName)
			self.filesFolder = fileName
		else:
			self.filesFolder=open("database.txt","r").read()
		tk.Grid.columnconfigure(self.master, 0, weight=1)

		if(not os.path.exists(self.filesFolder)):
			os.system("mkdir "+self.filesFolder)
		self.informationText = "Informationen:\n\
		Hinzufügen: Füge neue Einträge mit dem Knopf \"Hinzufügen\" hinzu.\n\
		Suchen: Suche mit dem Feld oben links, starte die Suche mit Enter oder einem Klick auf \"Suchen\".\n\
		Speichern: Speichere Inhalte in deinen Ordnern mit einem Doppelklick auf den Eintrag.\n\
		Bearbeiten: Ändere und lösche Einträge, indem du einen Rechtsklick auf sie machst."
		# db connection
		self.db = sql.connect(self.filesFolder+"/data2.db")
		self.cursor = self.db.cursor()
		self.executeDB("CREATE TABLE IF NOT EXISTS files\
			(id INTEGER PRIMARY KEY, location TEXT, mainCat TEXT, \
			comments TEXT, \
			created DATETIME DEFAULT (datetime('now','localtime')));")
		#other vars
		self.newFileSave = ""
		self.font = ("Arial",12)
		self.error_color = "red"

		# GUI bauen
		self.master.title("CeviMind "+self.version)
		topFrame = tk.Frame(self.master,padx=15,pady=15)
		topFrame.grid(row=0)
		self.searchBox = tk.Entry(topFrame, font=self.font)
		self.searchBox.grid(row=0, column=0)
		self.searchBox.bind("<Return>", self.search)
		tk.Button(topFrame, text="Suchen", command=self.search).grid(row=0, column=2)

		self.checkVar = tk.IntVar()
		self.checkBox = tk.Checkbutton(topFrame, text="Auch in Dateien suchen",variable=self.checkVar)
		self.checkBox.grid(row=0,column=3)

		self.filesCount = tk.StringVar()
		#tk.Label(topFrame, width=20).grid(row=0, column=3)#placeholder
		tk.Label(topFrame, textvar = self.filesCount, font=self.font).grid(row=0, column=4)
		tk.Label(topFrame, width=10).grid(row=0, column=5)#placeholder
		tk.Button(topFrame, text="Neues Dokument", command=self.add).grid(row=0,column=6)

		self.popupMenu = tk.Menu(self.master, tearoff=0)
		self.popupMenu.add_command(label="Speichern",command=self.onDoubleClick)
		self.popupMenu.add_command(label="Bearbeiten", command=self.editSelected)
		self.popupMenu.add_command(label="Löschen",command=self.deleteSeleteced)

		self.treeView = ttk.Treeview(self.master, columns=("Bemerkungen","Datum"), height=15)
		self.treeView.grid(row=1,column=0,columnspan=6)
		self.treeView.heading("#0",text="Hauptkategorie/Datei")
		self.treeView.heading("Bemerkungen",text="Bemerkungen")
		self.treeView.heading("Datum",text="Datum")
		self.treeView.bind("<Double-1>",self.onDoubleClick)
		self.treeView.bind("<Button-3>",self.popup)

		tk.Label(self.master, text=self.informationText, font=self.font, justify=tk.LEFT,pady=15).grid(row=2, columnspan=6)

		self.updateTree()
	def popup(self, event):
		iid = self.treeView.identify_row(event.y)
		if(iid):
			if(self.treeView.item(iid)["values"]==""): return
			self.treeView.selection_set(iid)
			self.popupMenu.tk_popup(event.x_root, event.y_root)
		else:
			pass
	def editCommit(self):
		stop = False
		if(self.editMainCat.get() == ""):
			self.editMainCat.config(relief="sunken", bg=self.error_color)
			self.newFrame.wm_attributes("-topmost",1)
			self.newFrame.focus_force()
			return
		if(self.executeDB("UPDATE files SET mainCat=?, comments=? WHERE location=?;",(self.editMainCat.get(), self.editComments.get(1.0, tk.END),self.treeView.item(self.treeView.selection()[0])["text"]))):
			self.editFrame.destroy()
			self.updateTree()
	def editSelected(self):
		data = self.treeView.item(self.treeView.selection()[0])
		parent = self.treeView.parent(self.treeView.selection()[0])
		self.editFrame = tk.Toplevel(self.master, pady=15, padx=15)
		self.editFrame.resizable(0,0)
		self.editFrame.focus_force()
		self.editFrame.title("Dokument bearbeiten")
		tk.Label(self.editFrame, text="Dokument Bearbeiten:", font=(self.font[0], self.font[1], "bold")).grid(row=0, columnspan=3)
		tk.Label(self.editFrame, text="Speicherort", font=self.font).grid(row=1, column=0)
		tk.Label(self.editFrame, text=data["text"]).grid(row=1, column=1, columnspan=2)
		tk.Label(self.editFrame, text="*Hauptkategorie:", font=self.font).grid(row=2, column=0)
		self.editMainCat = tk.Entry(self.editFrame, font=self.font, width=25)
		self.editMainCat.insert(0,parent)
		self.editMainCat.grid(row=2, column=1, columnspan=2)
		tk.Label(self.editFrame, text="Bemerkungen:", font=self.font, width=25).grid(row=5,column=0)
		self.editComments = tk.Text(self.editFrame, font=self.font, width=25, height=3)
		self.editComments.insert(tk.END, data["values"][0])
		self.editComments.grid(row=5, column=1, columnspan=2)
		tk.Button(self.editFrame, text="Speichern", font=self.font, width=60, command=self.editCommit).grid(row=8, column=0, columnspan=3)
	def deleteSeleteced(self):
		data = self.treeView.item(self.treeView.selection()[0])
		try:
			os.remove(self.filesFolder+"/"+data["text"])
		except(Exception) as e:
			self.writeErrorMessage(e,"deleteSeleteced")
		self.executeDB("DELETE FROM files WHERE location=\'"+data["text"]+"\'")
		self.updateTree()
	def onDoubleClick(self, event=None):
		item = self.treeView.selection()[0]
		data = self.treeView.item(item)
		if(data["values"] == ""): return
		ext = data["text"].split(".")
		ext = ext[len(ext)-1]
		saveFile = tkf.asksaveasfilename(defaultextension="."+ext,filetypes=[(ext,"."+ext)], title="Speichern unter...",initialdir=os.path.expanduser('~'),initialfile=data["text"])
		if(saveFile ==() or saveFile == ''): return
		shutil.copy(self.filesFolder+"/"+data["text"],saveFile)
	def updateTree(self):
		if(self.checkVar.get()):
			return self.search()
		self.treeView.delete(*self.treeView.get_children())
		searchString = self.searchBox.get()
		if(searchString != ""):
			searchString = "%"+searchString+"%"
			data = self.fetchDB("SELECT * FROM files \
				WHERE location LIKE ? OR mainCat LIKE ? \
				OR comments LIKE ? OR created LIKE ? ORDER BY mainCat ASC;",
				(searchString, searchString, searchString, searchString, searchString, searchString))
		else: data = self.fetchDB("SELECT * FROM files ORDER BY mainCat ASC;")
		for entry in data:
			if(entry[2] not in self.treeView.get_children()):
				self.treeView.insert("","end",entry[2],text=entry[2],open=False)
			if(entry[3] == None): a = ""
			else: a = entry[3]
			self.treeView.insert(entry[2], "end", text=entry[1], values=(a.replace("\n",""),entry[4]))
		self.updateFilesCount()
	def get_all_children(self, item=""):
	    children = self.treeView.get_children(item)
	    for child in children:
	        children += self.get_all_children(child)
	    return children
	def updateFilesCount(self):
		self.filesCount.set("Anzahl Dateien: "+str(len(self.get_all_children())-len(self.treeView.get_children()))+"/"+str(self.fetchDB("SELECT count(*) FROM files;")[0][0]))
	def executeDB(self, command, values=None):
		try:
			if(values == None):
				self.cursor.execute(command)
			else:
				self.cursor.execute(command, values)
			self.db.commit()
			return True
		except(Exception) as e:
			self.writeErrorMessage(e,"executeDB")
			return False
	def fetchDB(self, command, values=None):
		try:
			if(values==None):
				self.cursor.execute(command)
			else:
				self.cursor.execute(command, values)

			return self.cursor.fetchall()
		except(Exception) as e:
			self.writeErrorMessage(e,"fetchDB")
	def addNew(self):
		stop = False
		if(self.newFileSave == ""):
			self.newFileButton.config(relief="sunken", bg=self.error_color)
			stop = True
		if(self.newMainCat.get() == ""):
			self.newMainCat.config(relief="sunken", bg=self.error_color)
			stop = True
		if(stop):
			self.newFrame.wm_attributes("-topmost",1)
			self.newFrame.focus_force()
			return
		shutil.copy(self.newFileSave,self.filesFolder+"/"+self.newFileName.get())
		if(self.executeDB("INSERT INTO files(location, mainCat, comments) VALUES (?,?,?)",(self.newFileName.get(), self.newMainCat.get(), self.newComments.get(1.0, tk.END)))):
			self.newFrame.destroy()
			self.updateTree()
	def askFile(self):
		defaultbg = self.master.cget('bg')
		self.newFileButton.config(relief="flat", bg=defaultbg)
		data = tkf.askopenfile(initialdir = "/",title = "Datei auswählen",filetypes = (("Dokumente","*.pdf *.docx *.doc *.xls *.xlsx *.pptx *.txt *.png *.jpg *.jpeg"),("all files","*.*")))
		if(data == None):
			self.newFileButton.config(relief="sunken", bg=self.error_color)
			self.newFrame.wm_attributes("-topmost",1)
			self.newFrame.focus_force()
			return
		name = data.name.split("/")
		self.newFileSave = data.name
		self.newFileName.set(name[len(name)-1])
		if(self.fetchDB("SELECT count(*) FROM files WHERE location LIKE '"+self.newFileName.get()+"'")[0][0]>0):
			self.newFileButton.config(relief="sunken", bg=self.error_color)
			self.newFrame.wm_attributes("-topmost",1)
			self.newFileName.set("Bitte umbenennen")
			return
		self.newFrame.wm_attributes("-topmost",1)
		self.newFrame.focus_force()
	def add(self):
		self.newFileName = tk.StringVar()
		self.newFileName.set("Auswählen")
		self.newFrame = tk.Toplevel(self.master, pady=15, padx=15)
		self.newFrame.resizable(0,0)
		self.newFrame.focus_force()
		self.newFrame.title("Neues Dokument hinzufügen")
		tk.Label(self.newFrame, text="Neues Dokument hinzufügen", font=(self.font[0], self.font[1], "bold")).grid(row=0, columnspan=3)
		tk.Label(self.newFrame, text="*Speicherort", font=self.font).grid(row=1, column=0)
		self.newFileButton = tk.Button(self.newFrame, textvar=self.newFileName, font=self.font, command=self.askFile, width=25)
		self.newFileButton.grid(row=1, column=1, columnspan=2)
		tk.Label(self.newFrame, text="*Hauptkategorie:", font=self.font).grid(row=2, column=0)
		self.newMainCat = tk.Entry(self.newFrame, font=self.font, width=25)
		self.newMainCat.grid(row=2, column=1, columnspan=2)
		tk.Label(self.newFrame, text="Bemerkungen:", font=self.font, width=25).grid(row=5,column=0)
		self.newComments = tk.Text(self.newFrame, font=self.font, width=25, height=3)
		self.newComments.grid(row=5, column=1, columnspan=2)
		tk.Button(self.newFrame, text="Hinzufügen", font=self.font, width=60, command=self.addNew).grid(row=8, column=0, columnspan=3)
	def fileContains(self, filename, searchString):
		try:
			#data = textract.process(self.filesFolder+"/"+filename,encoding="utf-8")
			ext = filename.split(".")
			ext = ext[len(ext)-1]
			if(ext.lower() == "docx"):
				data = docx2txt.process(self.filesFolder+"/"+filename)
				return searchString.lower() in data.lower()
			elif(ext.lower() == "txt"):
				return searchString.lower() in open(self.filesFolder+"/"+filename).read().lower()
			elif(ext.lower() == "pdf"):
				return searchString.lower() in self.pdf_to_text(self.filesFolder+"/"+filename).lower()
			else:
				return False
		except(Exception) as e:
			self.writeErrorMessage(e,"fileContains")
			return False
	def writeErrorMessage(self,err,f):
		message = str(err)
		#tkm.showerror("Error at "+f,message)
		with open("error.log","a") as wr:
			wr.writelines(f+": "+message+"\n")
	def pdf_to_text(self,_pdf_file_path):
		try:
			pdf_content = PyPDF2.PdfFileReader(_pdf_file_path)
			text_extracted = ""
			for x in range(0, pdf_content.getNumPages()):
				pdf_text = ""
				pdf_text = pdf_text + pdf_content.getPage(x).extractText()
				text_extracted = text_extracted + "".join(i for i in pdf_text)
			return text_extracted
		except(Exception) as e:
			self.writeErrorMessage(e,"readPDF")
			return ""
	def search(self, event=None):
		if(self.checkVar.get()):
			self.treeView.delete(*self.treeView.get_children())
			searchString = self.searchBox.get()
			data = self.fetchDB("SELECT * FROM files ORDER BY mainCat ASC;")
			for entry in data:
				if(searchString in entry or self.fileContains(entry[1], searchString)):
					if(entry[2] not in self.treeView.get_children()):
						self.treeView.insert("","end",entry[2],text=entry[2],open=False)
					if(entry[3] == None): a = ""
					else: a = entry[3]
					self.treeView.insert(entry[2], "end", text=entry[1], values=(a.replace("\n",""),entry[4]))
				else:
					pass
			self.updateFilesCount()
		else:
			self.updateTree()

if(__name__=='__main__'):
	root = tk.Tk()
	Mind = CeviMind(root)
	root.mainloop()
