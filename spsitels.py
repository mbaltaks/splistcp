#!/usr/bin/env python
# full site list for SharePoint 2003
import sys
import subprocess

try:
    site_url = sys.argv[1]
except:
    print u'Provide a site url.'
    sys.exit()

spadmin = "C:\\Program Files\\Common Files\\Microsoft Shared\\web server extensions\\60\\BIN\\stsadm.exe"

def list_subsites(url):
    #print "New sub site:", url
    return_list = [""]
    list = subprocess.Popen([spadmin, "-o", "enumsubwebs", "-url", url], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    output, errors = list.communicate()
    split = output.split("<Subweb>")
    for s in split:
        if (s.find("http://") == 0):
            s = s.replace("</Subweb>", "")
            s = s.replace("</Subwebs>", "")
            s = s.rstrip(" \r\n")
            return_list.append(s)
            #print s
    return_list.remove("")
    for url in return_list:
        list = list_subsites(url)
        for l in list:
            return_list.append(l)
    return return_list


full_url_list = [""]

list = subprocess.Popen([spadmin, "-o", "enumsites", "-url", site_url], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
output, errors = list.communicate()
split = output.split("\"")
for s in split:
    if (s.find("http://") == 0):
        full_url_list.append(s)
full_url_list.remove("")

for url in full_url_list:
    sub_list = list_subsites(url)
    for sub in sub_list:
        full_url_list.append(sub)

for url in full_url_list:
    print url
