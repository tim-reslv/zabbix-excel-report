#coding=utf-8
import requests,json,csv,codecs,datetime,time
import xlsxwriter

ApiUrl = 'https://example.zabbix.com/zabbix/api_jsonrpc.php'
header = {"Content-Type":"application/json"}
user="Admin"
password="password"

Title1=['Hostname','IP','CPU Load avg5','Memory Util (%)','Space Util (/ %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
Title2=['Hostname','IP','CPU Util (%)','Memory Util (%)','Space Util (C: %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
#x=(datetime.datetime.now()-datetime.timedelta(minutes=180)).strftime("%Y-%m-%d %H:%M:%S")
x=(datetime.datetime.now()-datetime.timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
y=(datetime.datetime.now()).strftime("%Y-%m-%d %H:%M:%S")

def gettoken():
    data = {"jsonrpc": "2.0",
                "method": "user.login",
                "params": {
                    "user": user,
                    "password": password
                },
                "id": 1,
                "auth": None
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    return json.loads(auth.content)['result']

def timestamp(x,y):
    p=time.strptime(x,"%Y-%m-%d %H:%M:%S")
    starttime = str(int(time.mktime(p)))
    q=time.strptime(y,"%Y-%m-%d %H:%M:%S")
    endtime = str(int(time.mktime(q)))

    print("start: ", x, " end: ", y)
#    print(y)
    return starttime,endtime

def logout(auth):
    data={
            "jsonrpc": "2.0",
            "method": "user.logout",
            "params": [],
            "id": 1,
            "auth": auth
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    return json.loads(auth.content)


def get_hosts(hostids,auth):
    data ={
            "jsonrpc": "2.0",
             "method": "host.get",
             "params": {
             "output": [ "name"],
             "hostids": hostids,
             "filter":{
                 "status": "0"
             },
             "selectInterfaces": [
                        "ip",
                        "interfaceid"
                    ],
           },
            "auth": auth,  # theauth id is what auth script returns, remeber it is string
            "id": 1
        }
    gethost=requests.post(url=ApiUrl,headers=header,json=data)
#    result=json.loads(gethost.content)["result"]
#    print(result)
    return json.loads(gethost.content)["result"]

def get_linux_hosts(auth):
    data ={
              "jsonrpc": "2.0",
              "method": "template.get",
              "params": {
                  "output": [
                      "templateid",
                      "host"
                  ],
                  "filter": {
                      "host": [
                          "Template OS Linux by Zabbix agent active",
                      ]
                  },
                  "selectHosts": [],
              },
            "auth": auth,  # theauth id is what auth script returns, remeber it is string
            "id": 1
        }
    gethost=requests.post(url=ApiUrl,headers=header,json=data)
    hostids=[]
    linux_hosts=json.loads(gethost.content)["result"]
    for hostlist in linux_hosts:
        for hostid in hostlist['hosts']:
            hostids.append(hostid['hostid'])
    return hostids

def get_windows_hosts(auth):
    data ={
              "jsonrpc": "2.0",
              "method": "template.get",
              "params": {
                  "output": [
                      "templateid",
                      "host"
                  ],
                  "filter": {
                      "host": [
                          "Template OS Windows by Zabbix agent active",
                      ]
                  },
                  "selectHosts": [],
              },
            "auth": auth,  # theauth id is what auth script returns, remeber it is string
            "id": 1
        }
    gethost=requests.post(url=ApiUrl,headers=header,json=data)
    hostids=[]
    linux_hosts=json.loads(gethost.content)["result"]
    for hostlist in linux_hosts:
        for hostid in hostlist['hosts']:
            hostids.append(hostid['hostid'])
    return hostids

def get_linux_host_hist(hostid,hostname,hostip,auth,timestamp):
    host=[]
    item1=[]
    item2=[]
    dic1={}
    for j in ['system.cpu.load[all,avg5]','vm.memory.size[pavailable]','vfs.fs.size[/,pused]','net.if.in[\"e*\"]','net.if.out[\"e*\"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"

                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "auth":auth,
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        hisdata={
            "jsonrpc":"2.0",
            "method":"history.get",
            "params":{
                "output":"extend",
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "history":0,
                "sortfield": "clock",
                "sortorder": "DESC",
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "auth": auth,
            "id":1
            }
        get_host_hist=requests.post(url=ApiUrl,headers=header,json=hisdata)
        hist=json.loads(get_host_hist.content)['result']
        item1.append(hist)

    for j in ['system.cpu.load[all,avg5]','vm.memory.size[pavailable]','vfs.fs.size[/,pused]','net.if.in[\"e*\"]','net.if.out[\"e*\"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "auth":auth,
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        trendata={
            "jsonrpc":"2.0",
            "method":"trend.get",
            "params":{
                "output": [
                    "itemid",
                    "value_max",
                    "value_avg"
                ],
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "auth": auth,
            "id":1
            }
        gettrend=requests.post(url=ApiUrl,headers=header,json=trendata)
        trend=json.loads(gettrend.content)['result']
        item2.append(trend)
    dic1['Hostname']=hostname
    dic1['IP']=hostip
    try:
        dic1['CPU Load avg5']=item2[0][0]['value_avg']
    except IndexError:
        dic1['CPU Load avg5']=0
    try:
        dic1['Memory Utilization']=item2[1][0]['value_avg']
    except IndexError:
        dic1['Memory Utilization']=0
    try:
        dic1['Space Utilization']=item2[2][0]['value_avg']
    except IndexError:
        dic1['Space Utilization']=0
    try:
        dic1['Traffic In']=item2[3][0]['value_avg']
    except IndexError:
        dic1['Traffic In']=0
    try:
        dic1['Traffic Out']=item2[4][0]['value_avg']
    except IndexError:
        dic1['Traffic Out']=0
    dic1['Start Time']=x
    dic1['End Time']=y
    host.append(dic1)
    if not host:
        print("no record")
    else:
        return host

def get_windows_host_hist(hostid,hostname,hostip,auth,timestamp):
    host=[]
    item1=[]
    item2=[]
    dic1={}
    for j in ['system.cpu.util','vm.memory.util','vfs.fs.size[C:,pused]','net.if.in["Amazon Elastic Network Adapter"]','net.if.out["Amazon Elastic Network Adapter"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"

                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "auth":auth,
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        hisdata={
            "jsonrpc":"2.0",
            "method":"history.get",
            "params":{
                "output":"extend",
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "history":0,
                "sortfield": "clock",
                "sortorder": "DESC",
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "auth": auth,
            "id":1
            }
        get_host_hist=requests.post(url=ApiUrl,headers=header,json=hisdata)
        hist=json.loads(get_host_hist.content)['result']
        item1.append(hist)

    for j in ['system.cpu.util','vm.memory.util','vfs.fs.size[C:,pused]','net.if.in["Amazon Elastic Network Adapter"]','net.if.out["Amazon Elastic Network Adapter"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"

                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "auth":auth,
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        trendata={
            "jsonrpc":"2.0",
            "method":"trend.get",
            "params":{
                "output": [
                    "itemid",
                    "value_max",
                    "value_avg"
                ],
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "auth": auth,
            "id":1
            }
        gettrend=requests.post(url=ApiUrl,headers=header,json=trendata)
        trend=json.loads(gettrend.content)['result']
        item2.append(trend)

    dic1['Hostname']=hostname
    dic1['IP']=hostip
    try:
        dic1['CPU Load avg5']=item2[0][0]['value_avg']
    except IndexError:
        dic1['CPU Load avg5']=0
    try:
        dic1['Memory Utilization']=item2[1][0]['value_avg']
    except IndexError:
        dic1['Memory Utilization']=0
    try:
        dic1['Space Utilization']=item2[2][0]['value_avg']
    except IndexError:
        dic1['Space Utilization']=0
    try:
        dic1['Traffic In']=item2[3][0]['value_avg']
    except IndexError:
        dic1['Traffic In']=0
    try:
        dic1['Traffic Out']=item2[4][0]['value_avg']
    except IndexError:
        dic1['Traffic Out']=0
    dic1['Start Time']=x
    dic1['End Time']=y
    host.append(dic1)
    if not host:
        print("no record")
    else:
        return host

def createreport():
        fname = 'Zabbix_Report_monthly.xlsx'
        workbook = xlsxwriter.Workbook(fname)
        # Title format
        format_title = workbook.add_format()
        format_title.set_border(1)  # border
        format_title.set_bg_color('#1ac6c0')  # background color
        format_title.set_align('center')
        format_title.set_bold()
        format_title.set_valign('vcenter')
        format_title.set_font_size(12)
        # history data format
        format_value = workbook.add_format()
        format_value.set_border(1)
        format_value.set_align('center')
        format_value.set_valign('vcenter')
        format_value.set_font_size(12)

        format_percentage = workbook.add_format()
        format_percentage.set_border(1)
        format_percentage.set_align('center')
        format_percentage.set_valign('vcenter')
        format_percentage.set_font_size(12)
        format_percentage.set_num_format('0.00')

        # create Linux workbook
        worksheet1 = workbook.add_worksheet("Linux Servers")
        # set colum width
        worksheet1.set_column('A:A', 51)
        worksheet1.set_column('B:G', 15)
        worksheet1.set_column('H:I', 20)
        # set row high
        worksheet1.set_default_row(25)
        # Freeze title
        worksheet1.freeze_panes(1, 1)
        # write title
        i = 0
        for title in Title1:
            worksheet1.write(0, i, title, format_title)
            i += 1
        # write host history data
        j = 1
        hosts=get_hosts(get_linux_hosts(token),token)
        for host in hosts:
            hostid=host['hostid']
            hostname=host['name']
            hostip=host['interfaces'][0]['ip']
            row=get_linux_host_hist(hostid,hostname,hostip,token,timestamp)
#            print(row)

            worksheet1.write(j, 0, row[0]['Hostname'], format_value)
            worksheet1.write(j, 1, row[0]['IP'], format_value)
            worksheet1.write_number(j, 2, float(row[0]['CPU Load avg5']), format_percentage)
            worksheet1.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
            worksheet1.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
            worksheet1.write_number(j, 5, int(row[0]['Traffic In']), format_value)
            worksheet1.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
            worksheet1.write(j, 7, row[0]['Start Time'], format_value)
            worksheet1.write(j, 8, row[0]['End Time'], format_value)
            j += 1

        # create Windows workbook
        worksheet2 = workbook.add_worksheet("Windows Servers")
        # set colum width
        worksheet2.set_column('A:A', 51)
        worksheet2.set_column('B:G', 15)
        worksheet2.set_column('H:I', 20)
        # set row high
        worksheet2.set_default_row(25)
        # Freeze title
        worksheet2.freeze_panes(1, 1)
        # write title
        i = 0
        for title in Title2:
            worksheet2.write(0, i, title, format_title)
            i += 1
        # write host history data
        j = 1
        hosts=get_hosts(get_windows_hosts(token),token)
        for host in hosts:
            hostid=host['hostid']
            hostname=host['name']
            hostip=host['interfaces'][0]['ip']
            row=get_windows_host_hist(hostid,hostname,hostip,token,timestamp)
#            print(row)

            worksheet2.write(j, 0, row[0]['Hostname'], format_value)
            worksheet2.write(j, 1, row[0]['IP'], format_value)
            worksheet2.write_number(j, 2, float(row[0]['CPU Load avg5']), format_percentage)
            worksheet2.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
            worksheet2.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
            worksheet2.write_number(j, 5, int(row[0]['Traffic In']), format_value)
            worksheet2.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
            worksheet2.write(j, 7, row[0]['Start Time'], format_value)
            worksheet2.write(j, 8, row[0]['End Time'], format_value)
            j += 1

        workbook.close()

token=gettoken()

timestamp=timestamp(x,y)

createreport()

logout(token)