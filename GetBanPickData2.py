import urllib3
import xmltodict
import xlsxwriter
import re

def main():
    pageid = input("Please enter Esportspedia Page ID: ")
    weekcolumn = input("Please enter the name of week column: ")
    teamname = input("Please enter the name of team: ")
    url = 'http://lol.esportspedia.com/w/api.php?action=query&pageids='+pageid+'&prop=revisions&rvprop=content&format=xml'
    http = urllib3.PoolManager()
    file = http.request('GET',url)
    data = file.data
    data = xmltodict.parse(data)
    data = data['api']['query']['pages']['page']['revisions']['rev']['#text']
    data="".join(data.split())
    start = "{{PicksAndBans/SectionButton|name="+weekcolumn;
    end = "{{BlockBox|end}}";
    startPos = data.find(start);
    endPos = data.find(end,startPos);
    data=data[startPos:endPos];
    start = "{{PicksAndBans|";
    end = "}}";
    startPos = 0;
    endPos = 0;
    matchstartPos=[];
    matchendPos=[];
    search=True;
    while search :
        startPos = data.find(start,endPos);
        if startPos == -1:
            search=False;
        else:
            endPos = data.find(end,startPos);
            matchstartPos.append(startPos+len(start));
            matchendPos.append(endPos);
    numberofmatch=len(matchstartPos);
    match=[];
    for i in range(numberofmatch):
        match.append(data[matchstartPos[i]:matchendPos[i]]);
    blueteam=[];
    redteam=[];
    teamside=[];
    blueban1=[];
    blueban2=[];
    blueban3=[];
    redban1=[];
    redban2=[];
    redban3=[];
    bluepick1=[];
    bluepick2=[];
    bluepick3=[];
    bluepick4=[];
    bluepick5=[];
    bluepick1role=[];
    bluepick2role=[];
    bluepick3role=[];
    bluepick4role=[];
    bluepick5role=[];
    redpick1=[];
    redpick2=[];
    redpick3=[];
    redpick4=[];
    redpick5=[];
    redpick1role=[];
    redpick2role=[];
    redpick3role=[];
    redpick4role=[];
    redpick5role=[];
    for i in range(numberofmatch):
        start = "team1=";
        end = "|";
        startPos = 0;
        endPos = 0;
        startPos = match[i].find(start)+len(start);
        endPos = match[i].find(end,startPos);
        testteamname=match[i][startPos:endPos];
        testteamname=testteamname.upper();
        if testteamname == teamname:
            blueteam.append(match[i][startPos:endPos])
            teamside.append(1);
            start = "team2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redteam.append(match[i][startPos:endPos])
            start = "blueban1=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            blueban1.append(match[i][startPos:endPos])
            start = "blueban2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            blueban2.append(match[i][startPos:endPos])
            start = "blueban3=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            blueban3.append(match[i][startPos:endPos])
            start = "redban1=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redban1.append(match[i][startPos:endPos])
            start = "redban2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redban2.append(match[i][startPos:endPos])
            start = "redban3=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redban3.append(match[i][startPos:endPos])
            start = "bluepick1=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick1.append(match[i][startPos:endPos])
            start = "bluepick2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick2.append(match[i][startPos:endPos])
            start = "bluepick3=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick3.append(match[i][startPos:endPos])
            start = "bluepick4=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick4.append(match[i][startPos:endPos])
            start = "bluepick5=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick5.append(match[i][startPos:endPos])
            start = "redpick1=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick1.append(match[i][startPos:endPos])
            start = "redpick2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick2.append(match[i][startPos:endPos])
            start = "redpick3=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick3.append(match[i][startPos:endPos])
            start = "redpick4=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick4.append(match[i][startPos:endPos])
            start = "redpick5=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick5.append(match[i][startPos:endPos])
            start = "bluepick1role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick1role.append(match[i][startPos:endPos])
            start = "bluepick2role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick2role.append(match[i][startPos:endPos])
            start = "bluepick3role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick3role.append(match[i][startPos:endPos])
            start = "bluepick4role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick4role.append(match[i][startPos:endPos])
            start = "bluepick5role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            bluepick5role.append(match[i][startPos:endPos])
            start = "redpick1role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick1role.append(match[i][startPos:endPos])
            start = "redpick2role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick2role.append(match[i][startPos:endPos])
            start = "redpick3role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick3role.append(match[i][startPos:endPos])
            start = "redpick4role=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            redpick4role.append(match[i][startPos:endPos])
            start = "redpick5role=";
            startPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = len(match[i]);
            redpick5role.append(match[i][startPos:endPos])
        else:
            start = "team2=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            testteamname=match[i][startPos:endPos];
            testteamname=testteamname.upper();
            if testteamname == teamname:
                teamside.append(0);
                start = "team1=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                blueteam.append(match[i][startPos:endPos])
                start = "team2=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redteam.append(match[i][startPos:endPos])
                start = "blueban1=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                blueban1.append(match[i][startPos:endPos])
                start = "blueban2=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                blueban2.append(match[i][startPos:endPos])
                start = "blueban3=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                blueban3.append(match[i][startPos:endPos])
                start = "redban1=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redban1.append(match[i][startPos:endPos])
                start = "redban2=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redban2.append(match[i][startPos:endPos])
                start = "redban3=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redban3.append(match[i][startPos:endPos])
                start = "bluepick1=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick1.append(match[i][startPos:endPos])
                start = "bluepick2=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick2.append(match[i][startPos:endPos])
                start = "bluepick3=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick3.append(match[i][startPos:endPos])
                start = "bluepick4=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick4.append(match[i][startPos:endPos])
                start = "bluepick5=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick5.append(match[i][startPos:endPos])
                start = "redpick1=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick1.append(match[i][startPos:endPos])
                start = "redpick2=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick2.append(match[i][startPos:endPos])
                start = "redpick3=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick3.append(match[i][startPos:endPos])
                start = "redpick4=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick4.append(match[i][startPos:endPos])
                start = "redpick5=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick5.append(match[i][startPos:endPos])
                start = "bluepick1role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick1role.append(match[i][startPos:endPos])
                start = "bluepick2role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick2role.append(match[i][startPos:endPos])
                start = "bluepick3role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick3role.append(match[i][startPos:endPos])
                start = "bluepick4role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick4role.append(match[i][startPos:endPos])
                start = "bluepick5role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                bluepick5role.append(match[i][startPos:endPos])
                start = "redpick1role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick1role.append(match[i][startPos:endPos])
                start = "redpick2role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick2role.append(match[i][startPos:endPos])
                start = "redpick3role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick3role.append(match[i][startPos:endPos])
                start = "redpick4role=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                redpick4role.append(match[i][startPos:endPos])
                start = "redpick5role=";
                startPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = len(match[i]);
                redpick5role.append(match[i][startPos:endPos])
        
    workbook = xlsxwriter.Workbook('test.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0;
    col = 0;
    teamnameformat = workbook.add_format({'bold': True, 'bg_color': '#B6D7A8'})
    enemyteamnameformat = workbook.add_format({'bold': True})
    banformat = workbook.add_format({'bold': True, 'bg_color': '#D9D2E9'})
    teampickformat = workbook.add_format({'bold': True, 'bg_color': '#B6D7A8'})
    enemyteampickformat = workbook.add_format({'bold': True, 'bg_color': '#EA9999'})
    teamtopformat = workbook.add_format({'bold': True, 'bg_color': '#9FC5E8'})
    teamjgformat = workbook.add_format({'bold': True, 'bg_color': '#B4A7D6'})
    teammidformat = workbook.add_format({'bold': True, 'bg_color': '#B6D7A8'})
    teamadformat = workbook.add_format({'bold': True, 'bg_color': '#EA9999'})
    teamsupformat = workbook.add_format({'bold': True, 'bg_color': '#FFE599'})
    enemyteamroleformat = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9'})
    numberofmatch=len(blueteam);
    for i in range(numberofmatch):
        blueteam[i]=re.sub('[^A-Za-z0-9]','',blueteam[i])
        redteam[i]=re.sub('[^A-Za-z0-9]', '',redteam[i])
        blueban1[i]=re.sub('[^A-Za-z0-9]', '',blueban1[i])
        redban1[i]=re.sub('[^A-Za-z0-9]', '',redban1[i])
        blueban2[i]=re.sub('[^A-Za-z0-9]', '',blueban2[i])
        redban2[i]=re.sub('[^A-Za-z0-9]', '',redban2[i])
        blueban3[i]=re.sub('[^A-Za-z0-9]', '',blueban3[i])
        redban3[i]=re.sub('[^A-Za-z0-9]', '',redban3[i])
        bluepick1[i]=re.sub('[^A-Za-z0-9]', '',bluepick1[i])
        bluepick2[i]=re.sub('[^A-Za-z0-9]', '',bluepick2[i])
        bluepick3[i]=re.sub('[^A-Za-z0-9]', '',bluepick3[i])
        bluepick4[i]=re.sub('[^A-Za-z0-9]', '',bluepick4[i])
        bluepick5[i]=re.sub('[^A-Za-z0-9]', '',bluepick5[i])
        redpick1[i]=re.sub('[^A-Za-z0-9]', '',redpick1[i])
        redpick2[i]=re.sub('[^A-Za-z0-9]', '',redpick2[i])
        redpick3[i]=re.sub('[^A-Za-z0-9]', '',redpick3[i])
        redpick4[i]=re.sub('[^A-Za-z0-9]', '',redpick4[i])
        redpick5[i]=re.sub('[^A-Za-z0-9]', '',redpick5[i])
        bluepick1role[i]=re.sub('[^A-Za-z0-9]', '',bluepick1role[i])
        bluepick2role[i]=re.sub('[^A-Za-z0-9]', '',bluepick2role[i])
        bluepick3role[i]=re.sub('[^A-Za-z0-9]', '',bluepick3role[i])
        bluepick4role[i]=re.sub('[^A-Za-z0-9]', '',bluepick4role[i])
        bluepick5role[i]=re.sub('[^A-Za-z0-9]', '',bluepick5role[i])
        redpick1role[i]=re.sub('[^A-Za-z0-9]', '',redpick1role[i])
        redpick2role[i]=re.sub('[^A-Za-z0-9]', '',redpick2role[i])
        redpick3role[i]=re.sub('[^A-Za-z0-9]', '',redpick3role[i])
        redpick4role[i]=re.sub('[^A-Za-z0-9]', '',redpick4role[i])
        redpick5role[i]=re.sub('[^A-Za-z0-9]', '',redpick5role[i])
        blueteam[i]=blueteam[i].upper()
        redteam[i]=redteam[i].upper()
        blueban1[i]=blueban1[i].lower()
        redban1[i]=redban1[i].lower()
        blueban2[i]=blueban2[i].lower()
        redban2[i]=redban2[i].lower()
        blueban3[i]=blueban3[i].lower()
        redban3[i]=redban3[i].lower()
        bluepick1[i]=bluepick1[i].lower()
        bluepick2[i]=bluepick2[i].lower()
        bluepick3[i]=bluepick3[i].lower()
        bluepick4[i]=bluepick4[i].lower()
        bluepick5[i]=bluepick5[i].lower()
        redpick1[i]=redpick1[i].lower()
        redpick2[i]=redpick2[i].lower()
        redpick3[i]=redpick3[i].lower()
        redpick4[i]=redpick4[i].lower()
        redpick5[i]=redpick5[i].lower()
        blueban1[i]=blueban1[i].capitalize()
        redban1[i]=redban1[i].capitalize()
        blueban2[i]=blueban2[i].capitalize()
        redban2[i]=redban2[i].capitalize()
        blueban3[i]=blueban3[i].capitalize()
        redban3[i]=redban3[i].capitalize()
        bluepick1[i]=bluepick1[i].capitalize()
        bluepick2[i]=bluepick2[i].capitalize()
        bluepick3[i]=bluepick3[i].capitalize()
        bluepick4[i]=bluepick4[i].capitalize()
        bluepick5[i]=bluepick5[i].capitalize()
        redpick1[i]=redpick1[i].capitalize()
        redpick2[i]=redpick2[i].capitalize()
        redpick3[i]=redpick3[i].capitalize()
        redpick4[i]=redpick4[i].capitalize()
        redpick5[i]=redpick5[i].capitalize()
        bluepick1role[i]=bluepick1role[i].capitalize()
        bluepick2role[i]=bluepick2role[i].capitalize()
        bluepick3role[i]=bluepick3role[i].capitalize()
        bluepick4role[i]=bluepick4role[i].capitalize()
        bluepick5role[i]=bluepick5role[i].capitalize()
        redpick1role[i]=redpick1role[i].capitalize()
        redpick2role[i]=redpick2role[i].capitalize()
        redpick3role[i]=redpick3role[i].capitalize()
        redpick4role[i]=redpick4role[i].capitalize()
        redpick5role[i]=redpick5role[i].capitalize()
        if teamside[i]:
            worksheet.write(row,col,blueteam[i],teamnameformat);
            worksheet.write(row,col+1,redteam[i],enemyteamnameformat);
            row += 1;
            worksheet.write(row,col,blueban1[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban1[i],banformat);
            row += 1;
            worksheet.write(row,col,blueban2[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban2[i],banformat);
            row += 1;
            worksheet.write(row,col,blueban3[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban3[i],banformat);
            row += 1;
            worksheet.write(row,col,bluepick1[i],teampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick1[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick2[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,bluepick2[i],teampickformat);
            row += 1;
            worksheet.write(row,col,bluepick3[i],teampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick3[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick4[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,bluepick4[i],teampickformat);
            row += 1;
            worksheet.write(row,col,bluepick5[i],teampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick5[i],enemyteampickformat);
            row += 3;
            if bluepick1role[i][0] == 'T':
                worksheet.write(row,col,bluepick1role[i][0],teamtopformat);
            elif bluepick1role[i][0] == 'J':
                worksheet.write(row,col,bluepick1role[i][0],teamjgformat);
            elif bluepick1role[i][0] == 'M':
                worksheet.write(row,col,bluepick1role[i][0],teammidformat);
            elif bluepick1role[i][0] == 'A':
                worksheet.write(row,col,bluepick1role[i][0],teamadformat);
            elif bluepick1role[i][0] == 'S':
                worksheet.write(row,col,bluepick1role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick1role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick2role[i][0],enemyteamroleformat);
            row += 1;
            if bluepick2role[i][0] == 'T':
                worksheet.write(row,col,bluepick2role[i][0],teamtopformat);
            elif bluepick2role[i][0] == 'J':
                worksheet.write(row,col,bluepick2role[i][0],teamjgformat);
            elif bluepick2role[i][0] == 'M':
                worksheet.write(row,col,bluepick2role[i][0],teammidformat);
            elif bluepick2role[i][0] == 'A':
                worksheet.write(row,col,bluepick2role[i][0],teamadformat);
            elif bluepick2role[i][0] == 'S':
                worksheet.write(row,col,bluepick1role[i][0],teamsupformat);
            row += 1;
            if bluepick3role[i][0] == 'T':
                worksheet.write(row,col,bluepick3role[i][0],teamtopformat);
            elif bluepick3role[i][0] == 'J':
                worksheet.write(row,col,bluepick3role[i][0],teamjgformat);
            elif bluepick3role[i][0] == 'M':
                worksheet.write(row,col,bluepick3role[i][0],teammidformat);
            elif bluepick3role[i][0] == 'A':
                worksheet.write(row,col,bluepick3role[i][0],teamadformat);
            elif bluepick3role[i][0] == 'S':
                worksheet.write(row,col,bluepick3role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick3role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick4role[i][0],enemyteamroleformat);
            row += 1;
            if bluepick4role[i][0] == 'T':
                worksheet.write(row,col,bluepick4role[i][0],teamtopformat);
            elif bluepick4role[i][0] == 'J':
                worksheet.write(row,col,bluepick4role[i][0],teamjgformat);
            elif bluepick4role[i][0] == 'M':
                worksheet.write(row,col,bluepick4role[i][0],teammidformat);
            elif bluepick4role[i][0] == 'A':
                worksheet.write(row,col,bluepick4role[i][0],teamadformat);
            elif bluepick4role[i][0] == 'S':
                worksheet.write(row,col,bluepick4role[i][0],teamsupformat);
            row += 1;
            if bluepick5role[i][0] == 'T':
                worksheet.write(row,col,bluepick5role[i][0],teamtopformat);
            elif bluepick5role[i][0] == 'J':
                worksheet.write(row,col,bluepick5role[i][0],teamjgformat);
            elif bluepick5role[i][0] == 'M':
                worksheet.write(row,col,bluepick5role[i][0],teammidformat);
            elif bluepick5role[i][0] == 'A':
                worksheet.write(row,col,bluepick5role[i][0],teamadformat);
            elif bluepick5role[i][0] == 'S':
                worksheet.write(row,col,bluepick5role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick5role[i][0],enemyteamroleformat);
            row += 1;
        else:
            worksheet.write(row,col,blueteam[i],enemyteamnameformat);
            worksheet.write(row,col+1,redteam[i],teamnameformat);
            row += 1;
            worksheet.write(row,col,blueban1[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban1[i],banformat);
            row += 1;
            worksheet.write(row,col,blueban2[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban2[i],banformat);
            row += 1;
            worksheet.write(row,col,blueban3[i],banformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redban3[i],banformat);
            row += 1;
            worksheet.write(row,col,bluepick1[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick1[i],teampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick2[i],teampickformat);
            row += 1;
            worksheet.write(row,col,bluepick2[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,bluepick3[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick3[i],teampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick4[i],teampickformat);
            row += 1;
            worksheet.write(row,col,bluepick4[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,bluepick5[i],enemyteampickformat);
            row += 1;
            worksheet.write(row,col,None);
            worksheet.write(row,col+1,redpick5[i],teampickformat);
            row += 3;
            worksheet.write(row,col,bluepick1role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,None);
            if redpick1role[i][0] == 'T':
                worksheet.write(row,col+1,redpick1role[i][0],teamtopformat);
            elif redpick1role[i][0] == 'J':
                worksheet.write(row,col+1,redpick1role[i][0],teamjgformat);
            elif redpick1role[i][0] == 'M':
                worksheet.write(row,col+1,redpick1role[i][0],teammidformat);
            elif redpick1role[i][0] == 'A':
                worksheet.write(row,col+1,redpick1role[i][0],teamadformat);
            elif redpick1role[i][0] == 'S':
                worksheet.write(row,col+1,redpick1role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,None);
            if redpick2role[i][0] == 'T':
                worksheet.write(row,col+1,redpick2role[i][0],teamtopformat);
            elif redpick2role[i][0] == 'J':
                worksheet.write(row,col+1,redpick2role[i][0],teamjgformat);
            elif redpick2role[i][0] == 'M':
                worksheet.write(row,col+1,redpick2role[i][0],teammidformat);
            elif redpick2role[i][0] == 'A':
                worksheet.write(row,col+1,redpick2role[i][0],teamadformat);
            elif redpick2role[i][0] == 'S':
                worksheet.write(row,col+1,redpick2role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,bluepick2role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,bluepick3role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,None);
            if redpick3role[i][0] == 'T':
                worksheet.write(row,col+1,redpick3role[i][0],teamtopformat);
            elif redpick3role[i][0] == 'J':
                worksheet.write(row,col+1,redpick3role[i][0],teamjgformat);
            elif redpick3role[i][0] == 'M':
                worksheet.write(row,col+1,redpick3role[i][0],teammidformat);
            elif redpick3role[i][0] == 'A':
                worksheet.write(row,col+1,redpick3role[i][0],teamadformat);
            elif redpick3role[i][0] == 'S':
                worksheet.write(row,col+1,redpick3role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,None);
            if redpick4role[i][0] == 'T':
                worksheet.write(row,col+1,redpick4role[i][0],teamtopformat);
            elif redpick4role[i][0] == 'J':
                worksheet.write(row,col+1,redpick4role[i][0],teamjgformat);
            elif redpick4role[i][0] == 'M':
                worksheet.write(row,col+1,redpick4role[i][0],teammidformat);
            elif redpick4role[i][0] == 'A':
                worksheet.write(row,col+1,redpick4role[i][0],teamadformat);
            elif redpick4role[i][0] == 'S':
                worksheet.write(row,col+1,redpick4role[i][0],teamsupformat);
            row += 1;
            worksheet.write(row,col,bluepick4role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,bluepick5role[i][0],enemyteamroleformat);
            row += 1;
            worksheet.write(row,col,None);
            if redpick5role[i][0] == 'T':
                worksheet.write(row,col+1,redpick5role[i][0],teamtopformat);
            elif redpick5role[i][0] == 'J':
                worksheet.write(row,col+1,redpick5role[i][0],teamjgformat);
            elif redpick5role[i][0] == 'M':
                worksheet.write(row,col+1,redpick5role[i][0],teammidformat);
            elif redpick5role[i][0] == 'A':
                worksheet.write(row,col+1,redpick5role[i][0],teamadformat);
            elif redpick5role[i][0] == 'S':
                worksheet.write(row,col+1,redpick5role[i][0],teamsupformat);
            row += 1;
        
    workbook.close()
    
if __name__=="__main__":
    main()
