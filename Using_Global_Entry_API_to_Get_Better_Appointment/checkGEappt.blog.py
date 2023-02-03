#!/usr/bin/python
#from https://kevinisms.fason.org
#written by Jeremy Fason

import requests
import time, os
import re

url = "https://ttp.cbp.dhs.gov/schedulerapi/slots?orderBy=soonest&limit=1&locationId={}&minimum=1"
# add your email
email = 'user@domain.com'
# add your city and city code
LOCATION_IDS = {
    'Site1': 1111,
    #'Site2': 2222
}

for city, id in LOCATION_IDS.items():
        u = url.format(id)
        ua = {'User-Agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36"}
        appointments = requests.get(u, headers=ua).json()
        if appointments:
          results = appointments[0]['startTimestamp']
          print(results)
          if not os.path.isfile("~/checkGE.disable"):
              if re.match('2019-(10|11|12)|2020-(01|02)-',results):
                os.mknod("~/checkGE.disable")
                o = "{}: Found an appointment at {}!\n".format(city, appointments[0]['startTimestamp'])
                f = open("/tmp/checkGE.appt", "w+")
                f.write(o)
                sendEmail = 'cat /tmp/checkGE.appt|mailx -s "GlobalEntry appt available" '+ email
                os.system(sendEmail)
                os.system('date >> ~/checkGE.history')
                os.system('cat ~/checkGE.appt >> ~/checkGE.history')
                os.system('cat /dev/null > ~/checkGE.appt')
        else:
            print("{}: No appointments available".format(city))
        #time.sleep(5)
