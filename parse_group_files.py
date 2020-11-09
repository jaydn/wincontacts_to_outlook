import re 
import sys
import base64
import lxml
import os
from lxml import etree
import csv

def main():
    xtree = etree.parse(sys.argv[1]).getroot()
#    print(xtree)
#    print(etree.tostring(xtree, pretty_print=True))
    for email in xtree.findall(".//{http://schemas.microsoft.com/Contact/Extended/MSWABMAPI}PropTag0x66001102"):
#        print(etree.text_content(email))
#        print(etree.strip_tags(etree.dump(email)))
#        print(etree.dump(email))
#        print(email.text)
        group_content = base64.b64decode(email.text).decode('utf-16')
#        print(group_content)
        contacts = []
        for path in group_content.split('/PATH:"')[1:]:
            contact_file_path = path.split('"')[0]
            contact_file_path = contact_file_path.replace('\\', '/').replace("C:/Users/user/Contacts/", './Contacts/')
            print(contact_file_path)
            try:
                ctree = etree.parse(contact_file_path).getroot()
#                print(etree.tostring(ctree, pretty_print=True))
#                print(ctree)
#                print(ctrs)
#{http://schemas.microsoft.com/Contact}contact
#                for n in ctree.findall(".//{http://schemas.microsoft.com/Contact}c:EmailAdd:
                addrz = [x.text for x in ctree.findall(".//{http://schemas.microsoft.com/Contact}Address")]
                namez = [x.text for x in ctree.findall(".//{http://schemas.microsoft.com/Contact}FormattedName")]

                contacts.append(list(zip(addrz, namez))[0])

                os.unlink(contact_file_path)
            except Exception as e:
                print(str(e))
        print(contacts)
        with open("outputs/"+os.path.basename(sys.argv[1])+'.csv', mode='a') as outhandle:
            csvout =  csv.writer(outhandle)
            for contact in contacts:
                csvout.writerow(contact)

#            print("outputs/"+os.path.basename(sys.argv[1])+'.csv')

    
        
#        print(etree.dump(email))
#    for name in xtree.iter():
#        print(dir(name.tag))
#        print(name)
#        print(name.tag)


if __name__ == '__main__':
    main()
