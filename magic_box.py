#!/usr/bin/env/ python
#encoding:utf-8
'''
@author: Kelly Chen
@license: Dell
@time: 2/7/2017 9:40 PM
@desc:
'''
#http://blog.163.com/yang_jianli/blog/static/16199000620138922841273/
#http://blog.csdn.net/frank_good/article/details/52934306
#http://www.liaoxuefeng.com/wiki/0014316089557264a6b348958f449949df42a6d3a2e542c000
#http://www.cnblogs.com/rollenholt/archive/2012/04/23/2466179.html
from win32com.client.gencache import EnsureDispatch as Dispatch
import os
from datetime import datetime

outlook=Dispatch("Outlook.Application")
mapi=outlook.GetNamespace("MAPI")
MAILNAME="GSQE_Data@Dell.com"
FOLDERNAME="Inbox"
SUBFOLDERNAME="Audit"
PROCEEDFOLDERNAME="proceed"
DESTPATH="C:\\Users\\Kelly_Chen1\\PycharmProjects\\magic\\backup\\"
#inbox=mapi.GetDefaultFolder(win32.constants.olFolderInbox)
#for attachment in item.Attachments:
class Olk():
    def __init__(self,outlook_object):
        self._obj=outlook_object

def main():
    while mapi.Folders[MAILNAME].Folders[FOLDERNAME].Items.Count>0:
        item=mapi.Folders[MAILNAME].Folders[FOLDERNAME].Items[1]
        try:
            for attachment in item.Attachments:
                if os.path.exists(DESTPATH+attachment.DisplayName):
                    tuplefile=os.path.splitext(attachment.DisplayName)
                    attachment.SaveAsFile(DESTPATH + tuplefile[0]+datetime.now().strftime('%Y-%m-%d %H-%M-%S') +tuplefile[1])
                else:
                    attachment.SaveAsFile(DESTPATH+attachment.DisplayName)
        except Exception as e:
            print("My Error: ", e )
        item.Move(mapi.Folders[MAILNAME].Folders[PROCEEDFOLDERNAME])
    print("Done", datetime.now())
if __name__ == '__main__':
     main()

