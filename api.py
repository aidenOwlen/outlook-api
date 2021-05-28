# -*- coding: utf-8 -*-
import win32com.client
import os
from autoreply import * 
from PyQt4.QtCore import * 
from PyQt4.QtGui import * 
from PyQt4.QtCore import QThread 
from PyQt4 import QtCore,QtGui 
from PyQt4.QtGui import QApplication
global outlook


if not os.path.exists("automsg.txt"):
    

    f = open("automsg.txt","w")
    f.close()

def fetchMsgs(num=6):
    global listo
    listo = []
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(num)
    messages = inbox.Items
    message = messages.GetFirst()
    body = message.body
    subject = message.subject
    sender = message.Sender 
    sender_address = message.Sender.Address
    try:
        date = str(message.CreationTime)
    except:
        date = "none"
    listo.append([subject,body,sender,sender_address,date])



    i = 0
    while message and i < 200:
        try:
            message = messages.GetNext()
        except:
            pass
        try:
            body = message.body
        except:
            body = "None"
        try:
            subject = message.subject
        except:
            subject = "None"
        try:
            sender = message.Sender 
        except:
            sender = "None"
        try:
            sender_address = message.Sender.Address
        except:
            sender_address = "None"
        try:
            date = str(message.CreationTime)
        except:
            date = "none"
        listo.append([subject,body,sender,sender_address,date])
        i+= 1
  
             
fetchMsgs()
class myInterface(QtGui.QMainWindow,Ui_MainWindow):
    def retranslateUi(self,MainWindow):
         super(__class__,self).retranslateUi(MainWindow)
         self.view.clicked.connect(self.View)
         self.save.clicked.connect(self.Save)
         self.reply.clicked.connect(self.Reply)
         self.send.clicked.connect(self.Send)
         self.Attacj.clicked.connect(self.Attach)
         f = open("automsg.txt","r")
         r = f.read()
         f.close()
         self.automsg.setText(r)
         self.replymsg.setText(r)
         self.replymsg.hide()
         self.send.hide()
         self.subject.hide()
         self.sub.hide()
         self.Attacj.hide()
         header = self.tree.horizontalHeader()
         header.setResizeMode(0, QtGui.QHeaderView.Stretch)
         header.setResizeMode(1, QtGui.QHeaderView.Stretch)
         header.setResizeMode(2, QtGui.QHeaderView.Stretch)
         header.setResizeMode(3, QtGui.QHeaderView.Stretch)
         self.results = listo

         if len(self.results) >= 1:
                self.tree.setRowCount(0)
         for row_number,row_data in enumerate(self.results):
                self.tree.insertRow(row_number)
                for column_number,data in enumerate(row_data):
                    self.tree.setItem(row_number,column_number,QTableWidgetItem(str(data)))

         self.inbox1.clicked.connect(self.Inbox1)
         self.inbox2.clicked.connect(self.Inbox2)

        
    def Attach(self):
        self.filetoshare = QFileDialog.getOpenFileName()
        try:
            Msg.Attachments.Add(self.filetoshare)
        except Exception as K:
            print(K)
  
    def Inbox1(self):
        self.tree.setRowCount(0)
        fetchMsgs(6)
        self.results = listo

        if len(self.results) >= 1:
                self.tree.setRowCount(0)
        for row_number,row_data in enumerate(self.results):
                self.tree.insertRow(row_number)
                for column_number,data in enumerate(row_data):
                    self.tree.setItem(row_number,column_number,QTableWidgetItem(str(data)))

    def Inbox2(self):
        self.tree.setRowCount(0)
        fetchMsgs(5)
        self.results = listo

        if len(self.results) >= 1:
                self.tree.setRowCount(0)
        for row_number,row_data in enumerate(self.results):
                self.tree.insertRow(row_number)
                for column_number,data in enumerate(row_data):
                    self.tree.setItem(row_number,column_number,QTableWidgetItem(str(data)))
    def fillTree(self):
        print("hllo")
        

    def Send(self):
        self.subjectt = self.subject.text()
        Msg.Subject = self.subjectt
        Msg.Body = self.replymsg.toPlainText()
        Msg.Send()
        self.msg = QMessageBox()
        self.msg.setInformativeText("Message Sent ")
        self.msg.setIcon(QMessageBox.Information)

        self.msg.show()
        self.msg.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.msg.exec_()
        self.sub.hide()
        self.subject.hide()
        self.Attacj.hide()
        self.replymsg.hide()
    def Reply(self):
        global Msg

        o = win32com.client.Dispatch("Outlook.Application")

        
        row = self.tree.currentRow()
        self.senderr = self.tree.item(row,2).text()
        self.senderrAdress = self.tree.item(row,3).text()
        Msg = o.CreateItem(0)
        Msg.To = self.senderrAdress

        self.send.show()
        self.replymsg.show()
        self.sub.show()
        self.Attacj.show()
        self.subject.show()
        f = open("automsg.txt","r")
        r = f.read()
        f.close()
        try:
            r = r.replace("{sender}",self.senderr)
        except:
            pass
        self.replymsg.clear()
        self.replymsg.setText(r)
        
    def Save(self):
        self.textmsg = self.automsg.toPlainText()
        f = open("automsg.txt","w")
        f.write(self.textmsg)
        f.close()

    def View(self):
        row2 = self.tree.currentRow()
        self.subject2 = self.tree.item(row2,0).text()
        self.body2 = self.tree.item(row2,1).text()
        self.date2 = self.tree.item(row2,4).text()
        self.senderr2 = self.tree.item(row2,2).text()
        self.senderrAdress2 = self.tree.item(row2,3).text()

        self.msg2 = QMessageBox()
        self.msg2.setWindowTitle("Message details")
        self.msg2.setIcon(QMessageBox.Information)
        self.msg2.setText(self.subject2)
        self.msg2.setInformativeText("Sender : {}\nSender Address : {}\nDate : {}".format(self.senderr2,self.senderrAdress2,self.date2))
        self.msg2.setDetailedText(self.body2)
        self.msg2.show()

if __name__ == "__main__":
    import sys 
    app = QtGui.QApplication(sys.argv)
   
    
    MainWindow = QtGui.QMainWindow()
    ui = myInterface() 
    ui.setupUi(MainWindow)
    MainWindow.show()
    myInterface().fillTree()
    sys.exit(app.exec_())