from PyQt5 import QtWidgets, uic
import sys
import json 

attr_list=["Agent","Company","Source2","Market","Source1","Origin","Payment",
"picked_name","remark","RT_name","RT_address","RT_transaction",
"PT_name","PT_address","PT_transaction"]

class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__() # Call the inherited classes __init__ method
        uic.loadUi('res_simulator.ui', self) # Load the .ui file


        #buttons
        self.Newbtn.clicked.connect(self.new_profile)
        self.Savebtn.clicked.connect(self.save_profile)
        self.Deletebtn.clicked.connect(self.del_profile)
        self.Closebtn.clicked.connect(self.close_win)

        self.read_data_from_json()
        self.domestic_list.itemClicked.connect(self.agent_picked)
        self.oversea_list.itemClicked.connect(self.agent_picked)

        self.show() # Show the GUI

    #read agent list from json file
    def read_data_from_json(self):
        self.domestic_list.clear()
        self.oversea_list.clear()
        with open("test.json","r",encoding="utf-8") as f:
            global data
            data = json.load(f)
            for i in data:
                if data[i]["Market"] in ["OTAD","DIR","AGTD"]:
                    self.domestic_list.addItem(i)
                else:
                    self.oversea_list.addItem(i)

    #button function
    def new_profile(self):
        #right top
        self.Agent.setText("")
        self.Company.setText("")
        self.Source2.setText("")
 
        #middle
        self.Market.setText("")
        self.Source.setText("")
        self.Origin.setText("")
        self.Payment.setText("")
        self.picked_name.setText("")
        self.remark.setText("")

        #Routing
        self.RT_name.setText("")
        self.RT_address.setText("")
        self.RT_transaction.setText("")

        self.PT_name.setText("")
        self.PT_address.setText("")
        self.PT_transaction.setText("")

    def save_profile(self):
        try:
            with open("test.json","w+",encoding="utf-8") as f:

                agent_name=self.picked_name.text()

                if agent_name not in data and agent_name!="":
                    data[agent_name]={}
                    
                #middle
                data[agent_name]["Market"]=self.Market.text()
                data[agent_name]["Source"]=self.Source.text()
                data[agent_name]["Origin"]=self.Origin.text()
                data[agent_name]["Remark"]=self.remark.text()

                #right top
                data[agent_name]["Agent"]=self.Agent.text()
                data[agent_name]["Company"]=self.Company.text()
                data[agent_name]["Source2"]=self.Source2.text()

                #Routing
                data[agent_name]["Routing_RT"]=self.RT_name.text()
                data[agent_name]["Address_RT"]=self.RT_address.text()
                data[agent_name]["Transactions_RT"]=self.RT_transaction.text()

                data[agent_name]["Routing_PT"]=self.PT_name.text()
                data[agent_name]["Address_PT"]=self.PT_address.text()
                data[agent_name]["Transactions_PT"]=self.PT_transaction.text()

                json.dump(data,f,sort_keys=True,indent = 4)

            self.statusbar.showMessage(f"{agent_name} を保存しました",5000)
            self.new_profile()
            self.read_data_from_json()
        except:
            self.statusbar.showMessage("保存できませんでした",5000)
        print("save button is pressed")

    def del_profile(self,item):
        agent_name=self.picked_name.text()
        if agent_name !="":
            dialog = msgbox()
            dialog.msg.setText(f"{agent_name}\nを削除しますか?")
            retValue = dialog.exec_()
            if retValue == QtWidgets.QDialog.Accepted:
                with open("test.json","w",encoding="utf-8") as f:
                    del data[agent_name]
                    self.statusbar.showMessage(f"{agent_name} を削除しました",5000)
                    json.dump(data,f,sort_keys=True,indent = 4)
                self.new_profile()
                self.read_data_from_json()
                print("delete button is pressed")

    def close_win(self):
        exit()
        print("close button is pressed")
    
    #show the data of picked agent from the list
    def agent_picked(self,item):

        picked_name = item.text()

        with open("test.json","r",encoding="utf-8") as f:
                data = json.load(f)
                if picked_name in data:
                    try:
                        self.picked_name.setText(picked_name)
                        self.picked_name.setCursorPosition(0)
                        #right top
                        self.Agent.setText(data[picked_name]["Agent"])
                        self.Agent.setCursorPosition(0)
                        self.Company.setText(data[picked_name]["Company"])
                        self.Company.setCursorPosition(0)
                        self.Source2.setText(data[picked_name]["Source2"])
                        self.Source2.setCursorPosition(0)

                        #middle
                        self.Market.setText(data[picked_name]["Market"])
                        self.Market.setCursorPosition(0)
                        self.Source.setText(data[picked_name]["Source"])
                        self.Source.setCursorPosition(0)
                        self.Origin.setText(data[picked_name]["Origin"])
                        self.Origin.setCursorPosition(0)
                        self.Payment.setText("CA / DB")
                        self.remark.setText(data[picked_name]["Remark"])

                        #Routing
                        self.RT_name.setText(data[picked_name]["Routing_RT"])
                        self.RT_name.setCursorPosition(0)
                        self.RT_address.setText(data[picked_name]["Address_RT"])
                        self.RT_address.setCursorPosition(0)
                        self.RT_transaction.setText(data[picked_name]["Transactions_RT"])
                        self.RT_transaction.setCursorPosition(0)

                        self.PT_name.setText(data[picked_name]["Routing_PT"])
                        self.PT_name.setCursorPosition(0)
                        self.PT_address.setText(data[picked_name]["Address_PT"])
                        self.PT_address.setCursorPosition(0)
                        self.PT_transaction.setText(data[picked_name]["Transactions_PT"])
                        self.PT_transaction.setCursorPosition(0)

                    except:
                        print("something wrong")
        print(picked_name)

class msgbox(QtWidgets.QDialog):
    def __init__(self):
        super(msgbox,self).__init__() # Call the inherited classes __init__ method
        uic.loadUi('msgbox.ui', self) # Load the .ui file

        #buttons
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

app = QtWidgets.QApplication(sys.argv) # Create an instance of QtWidgets.QApplication
window = Ui() # Create an instance of our class
app.exec_() # Start the application