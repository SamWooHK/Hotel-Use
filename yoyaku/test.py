import json

#read data from json file

with open("test.json","r",encoding="utf-8") as f:
    global data
    data = json.load(f)
    # for i in range(len(data)):
    #     # print(data[i]["Name"])

    #     if "OTAD" in data[i]["Market"]:
    #         print(data[i])
    for i in data:
        print(i)

# # #change data
# with open("test.json","w+",encoding="utf-8") as f:
#     data[0]["Name"]="test"
#     data[1]["Name"]="test1"
#     data[2]["Name"]="test2"
#     json.dump(data,f,indent =4)

