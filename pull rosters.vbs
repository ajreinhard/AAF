Set fso = CreateObject("Scripting.FileSystemObject")
file_path = "C:\Users\Owner\Documents\GitHub\AAF\"

dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
dim teams(8)
teams(1) = "arizona-hotshots"
teams(2) = "atlanta-legends"
teams(3) = "birmingham-iron"
teams(4) = "memphis-express"
teams(5) = "orlando-apollos"
teams(6) = "salt-lake-stallions"
teams(7) = "san-antonio-commanders"
teams(8) = "san-diego-fleet"

for tm = 1 to 8
team_roster = "https://aaf.com/" & teams(tm) & "/roster/"

xHttp.Open "GET", team_roster, False
xHttp.Send

with bStrm
    .type = 1
    .open
    .write xHttp.responseBody
    .savetofile file_path & "\rosters\" & teams(tm) & ".txt", 2
    .close
end with

next


Set bStrm = Nothing
Set xHttp = Nothing
Set fso = Nothing
Set fl = Nothing



msgbox "Done"
wscript.quit