# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            time.sleep(0.2)
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')   
            time.sleep(0.2) 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次') 
            time.sleep(0.2)   
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/directReports',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/me/profile',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/me/insights/trending',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/directReports',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/users',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/me/planner/tasks',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/groups',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/me/findRooms',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/contacts',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/recent',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/schemaExtensions',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/onenote/notebooks',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/sites/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/sites/root/drives',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/sites/root/lists',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/v1.0/applications?$count=true',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/applicationTemplates',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
        if req.get(r'https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorgroups',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
            time.sleep(0.2)
    except:
        print("pass")
        pass
for _ in range(3000):
    main()
    time.sleep(0.8)
