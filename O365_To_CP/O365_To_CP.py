import urllib
import requests
import json
import wget
import os
import uuid
import socket
import struct
import glob


# based on code from https://geektechstuff.com/2018/07/09/microsoft-office-365-endpoints-v2-python/ by geektechstuff
# original source from https://support.office.com/en-us/article/managing-office-365-endpoints-99cab9d4-ef59-4207-9f2b-3728eb46bf9a?ui=en-US&rs=en-US&ad=US#ID0EACAAA=4._Web_service
# Modified by DrkNite
# cidr to netmask by https://stackoverflow.com/users/2225682/falsetru from https://stackoverflow.com/questions/33750233/convert-cidr-to-subnet-mask-in-python


def cidr_to_netmask(cidr):
    network, net_bits = cidr.split('/')
    host_bits = 32 - int(net_bits)
    netmask = socket.inet_ntoa(struct.pack('!I', (1 << 32) - (1 << host_bits)))
    return network, netmask

# helper to call the webservice and parse the response
def webApiGet(methodName, instanceName, clientRequestId):
    ws = "https://endpoints.office.com"
    requestPath = ws + '/' + methodName + '/' + instanceName + '?clientRequestId=' + clientRequestId
    request = urllib.request.Request(requestPath)
    with urllib.request.urlopen(request) as response:
        return json.loads(response.read().decode())
    return()

def SortAndRemove(filein, fileout):
        lines_seen = set() # holds lines already seen
        outfile = open(fileout, "w")
        for line in open(filein, "r"):
            if line not in lines_seen: # not a duplicate
                outfile.write(line)
                lines_seen.add(line)
        outfile.close()

def writeDataOut(destFile,fileData):
    if os.path.exists(destFile):
        with open(destFile, 'a') as data_write:
            data_write.write(fileData)
    else:
        with open(destFile, 'w') as data_write:
            data_write.write(fileData)
    return()

def parseURL(urls1):
    outputstr = ""
    for url in urls1:
        urlsplit = url.split(".")
        x = len(urlsplit)
        i=1
        y=""
        DN1 = urlsplit[0]
        if DN1[:1] == "*":
            while i < x:
                y += "."+urlsplit[i]
                i += 1
            outputstr+=(y+" FQDN"+'\n')
        else:
            i=0
            while i < x:
                y += "."+urlsplit[i]
                i += 1
            outputstr+=(y+" non-FQDN"+'\n')
    return(outputstr)

# path where client ID and latest version number will be stored
datapath = './endpoints_clientid_latestversion.txt'

# URL List for for Firewall
url_list = './365_url_list.txt'
url_list_sorted = './365_url_list_sorted.txt'

# IPv4 List for Firewall
ip_hosts = './365_hosts_list.tmp'
ip_hosts_sorted = './365_hosts_list_sorted.txt'
ip_nets = './365_nets_list.tmp'
ip_nets_sorted = './365_nets_list_sorted.txt'

# Group creating script
o365_groups = './365_groups.txt'
net_groups = "./net_group.txt"

#Temp delete txt files for testing
r = [f for f in glob.glob("*.txt")]
for f in r:
    os.remove (f)


# fetch client ID and version if data exists; otherwise create new file
if os.path.exists(datapath):
    with open(datapath, 'r') as fin:
        clientRequestId = fin.readline().strip()
        latestVersion = fin.readline().strip()
else:
    clientRequestId = str(uuid.uuid4())
    latestVersion = '0000000000'
    with open(datapath, 'w') as fout:
        fout.write(clientRequestId + '\n' + latestVersion)

# call version method to check the latest version, and pull new data if version number is different
version = webApiGet('version', 'Worldwide', clientRequestId)
if version['latest'] > latestVersion:
    print('New version of Office 365 worldwide commercial service instance endpoints detected')

    # write the new version number to the data file
    with open(datapath, 'w') as fout:
        fout.write(clientRequestId + '\n' + version['latest'])

    # invoke endpoints method to get the new data
    endpointSets = webApiGet('endpoints', 'Worldwide', clientRequestId)

    # filter results for required, and transform these into tuples with port and category
    flatUrls = []
    for endpointSet in endpointSets:
        if endpointSet['required'] == True:
            iid = endpointSet['id']
            serviceArea = endpointSet['serviceArea']
            category = endpointSet['category']
            required = endpointSet['required']
            urls = endpointSet['urls'] if 'urls' in endpointSet else []
            tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
            udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
            flatUrls.extend([(iid, serviceArea, category, required, url, tcpPorts, udpPorts) for url in urls])
            #               url[0]   url[1]      url[2]     url[3] url[4]  url[5]   url[6]

    flatIps = []
    for endpointSet in endpointSets:
        if endpointSet['required'] == True:
            iid = endpointSet['id']
            serviceArea = endpointSet['serviceArea']
            ips = endpointSet['ips'] if 'ips' in endpointSet else []
            category = endpointSet['category']
            required = endpointSet['required']
            ip4s = [ip for ip in ips if '.' in ip]
            tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
            udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
            flatIps.extend([(iid, serviceArea, category, required, ip, tcpPorts, udpPorts) for ip in ip4s])
            #               ip[0]   ip[1]       ip[2]       ip[3]  ip[4]  ip[5]   ip[6]

    print('IPv4 IP hosts')
    HT1 = ""
    HT2 = ""
    for ip in flatIps:
        IP2 = cidr_to_netmask(ip[4])
        ipAdd = IP2[0]
        CIDR = IP2[1]
        serviceArea2 = ip[1]
        if CIDR == "255.255.255.255":
            HT1+=str('add network name "Microsoft_O365_H_'+ipAdd+'" ip-address "'+ipAdd+'\n')
 #           HT2+=str('add network name "Microsoft_O365_H_'+ipAdd+'" ip-address "'+ipAdd+'\n')
            if ip[1] == "Exchange":
                HT2+=str('set group name Microsoft_O365_Exchange members.add "Microsoft_O365_H_'+ipAdd+'"'+'\n')
            elif ip[1] == "SharePoint":
                HT2+=str('set group name Microsoft_O365_SharePoint members.add "Microsoft_O365_H_'+ipAdd+'"'+'\n')
            elif ip[1] == "Skype":
                HT2+=str('set group name Microsoft_O365_Skype members.add "Microsoft_O365_H_'+ipAdd+'"'+'\n')
            else:
                HT2+=str('set group name Microsoft_O365_Common members.add "Microsoft_O365_H_'+ipAdd+'"'+'\n')
    writeDataOut(ip_hosts,HT1)
    writeDataOut(net_groups,HT2)
    print('IPv4 hosts Done')

    SortAndRemove(ip_hosts,ip_hosts_sorted)


    print('IPv4 IP Address Ranges')
    IPR1 = ""
    IPR2 = "" # use for creating group file, create a def to parse and return string.
    for ip in flatIps:
        IP2 = cidr_to_netmask(ip[4])
        ipRange = IP2[0]
        CIDR = IP2[1]
        serviceArea2 = ip[1]
        if CIDR != "255.255.255.255":
           IPR1+=str('add network name "Microsoft_O365_N_'+ip[4]+'" subnet "'+ipRange+'" subnet-mask "'+CIDR+'"'+'\n')
           if ip[1] == "Exchange":
                IPR2+=str('set group name Microsoft_O365_Exchange members.add "Microsoft_O365_N_'+ip[4]+'"'+'\n')
           elif ip[1] == "SharePoint":
                IPR2+=str('set group name Microsoft_O365_SharePoint members.add "Microsoft_O365_N_'+ip[4]+'"'+'\n')
           elif ip[1] == "Skype":
                IPR2+=str('set group name Microsoft_O365_Skype members.add "Microsoft_O365_N_'+ip[4]+'"'+'\n')
           else:
                IPR2+=str('set group name Microsoft_O365_Common members.add "Microsoft_O365_N_'+ip[4]+'"'+'\n')
    writeDataOut(ip_nets,IPR1)
    writeDataOut(net_groups,IPR2)
    print('IPv4 ranges Done')

    SortAndRemove(ip_nets,ip_nets_sorted)

    #print('URLs List')
    #with open(url_list, 'w') as data_write:
    #    for url in flatUrls:
    #        URL2 = url[4]
    #        data_write.write(str(url[0])+' , '+url[1]+' , '+URL2+'\n')
    #print('URLs Done')

    URL2 = []
    for url in flatUrls:
        URL1 = str(url[4])
        URL2.append(URL1)
    urlstrList = parseURL(URL2)
    writeDataOut(url_list,urlstrList)

    SortAndRemove(url_list,url_list_sorted)

    with open(o365_groups, 'w') as data_write:
        data_write.write('add group name "Microsoft_O365_Common"'+'\n'+
                         'add group name "Microsoft_O365_Exchange"'+'\n'+
                         'add group name "Microsoft_O365_SharePoint"'+'\n'+
                         'add group name "Microsoft_O365_Skype"'+'\n'+'\n'+
                         'add group name "Microsoft_O365_DO_Common"'+'\n'+
                         'add group name "Microsoft_O365_DO_Exchange"'+'\n'+
                         'add group name "Microsoft_O365_DO_SharePoint"'+'\n'+
                         'add group name "Microsoft_O365_DO_Skype"'+'\n')

    #delete tmp files
    r = [f for f in glob.glob("*.tmp")]
    for f in r:
        os.remove (f)

else:
    print('Office 365 worldwide commercial service, no updates detected.')