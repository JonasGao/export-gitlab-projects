import http.client
import json
import re
from openpyxl import Workbook

HOST = '<Your gitlab host>'
TOKEN = '<Your gitlab access token>'


def req(url):
    conn = http.client.HTTPSConnection(HOST)
    payload = ''
    headers = {
        'Authorization': 'Bearer ' + TOKEN
    }
    conn.request("GET", url, payload, headers)
    res = conn.getresponse()
    data = res.read()
    str_value = data.decode("utf-8")
    json_obj = json.loads(str_value)
    link = res.headers['Link']
    if not link:
        return json_obj, None
    m = re.search('^<(.+)>;', link)
    found = m.group(1)
    return json_obj, found


def write_project(projects, ws):
    for proj in projects:
        ws.append([proj['id'], proj['name'], proj['description'], proj['http_url_to_repo']])


def main():
    wb = Workbook()
    ws = wb.active
    ws.append(['id', 'name', 'description', 'http_url_to_repo'])
    url = "/api/v4/projects?pagination=keyset&per_page=50&order_by=id&sort=asc"
    json_obj, next_url = req(url)
    write_project(json_obj, ws)
    while next_url:
        json_obj, next_url = req(next_url[21:])
        write_project(json_obj, ws)
    wb.save('/home/demo/Documents/projects.xlsx')


if __name__ == '__main__':
    main()
