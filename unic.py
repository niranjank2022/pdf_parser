import re
t = """
sello
A, helo
The student wants to emphasize how Romney wool  differs from Merino and  Rambouillet wool.  Which choice most 
effectively uses relevant information from the notes to accomplish this goal? 
A.  Romney wool is just one of the many kinds of wools, each originating from a different breed of sheep. 
B.  Sheep wool varies from breed to breed, so Romney wool will  be different than other kinds of wool.
Abra"""

t = t.replace('\n', '###')
# print(t)
s = re.findall("A\\. (.*?)###", t)
print(s)
# print(s.replace("$$$", '\n'))
