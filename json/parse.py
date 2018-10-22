#!/usr/bin/python
import json
import pprint

file = open("json.example","r")
#print(file.read())

jsondata = json.loads(file.read())
# print(jsondata)

# output to file example
# output = open("json.example2", "w")
# pprint.pprint(jsondata, output)

# left_lines = jsondata["left_lines"]
# print(left_lines)
# easier than above

domains = ""
# print("\ngetting all domains:")
for item in jsondata['result']['top']:
    domains += item.get("domain") + ","

# print(domains)

urls = ""
# print("\ngetting all URLS:")
for item in jsondata['result']['top']:
    urls += item.get("url") + ","

# print(urls)

print(jsondata['left_lines'])
print(jsondata['result']['results'])
# print(jsondata['result']['top'][0]['domain'] + ", " + jsondata['result']['top'][0]['url'])
print("found_domains:" + domains)
print("found_urls:" + urls)
print()