import urllib3
import xmltodict
import xlsxwriter
import re
import championData;
from collections import Counter

def writepickwithselectedrole(workbook,worksheet,i,row,col,pickrole1,pickrole2,pickrole3,pickrole4,pickrole5,pick1,pick2,pick3,pick4,pick5,role,teamside,toptenchampion):
    blueteamformat = workbook.add_format({'bold': True, 'font_color': '#0000FF'})
    redteamformat = workbook.add_format({'bold': True, 'font_color': '#FF0000'})
    blue1stformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#00FFFF'})
    blue2ndformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#FFFF00'})
    blue3rdformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#A64D79'})
    blue4thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#00FF00'})
    blue5thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#FF9900'})
    blue6thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#674EA7'})
    blue7thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#FF4FD9'})
    blue8thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#274E13'})
    blue9thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#E06666'})
    blue10thformat = workbook.add_format({'bold': True, 'font_color': '#0000FF','bg_color': '#4A86E8'})
    red1stformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#00FFFF'})
    red2ndformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#FFFF00'})
    red3rdformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#A64D79'})
    red4thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#00FF00'})
    red5thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#FF9900'})
    red6thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#674EA7'})
    red7thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#FF4FD9'})
    red8thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#274E13'})
    red9thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#E06666'})
    red10thformat = workbook.add_format({'bold': True, 'font_color': '#FF0000','bg_color': '#4A86E8'})
    if teamside == 'blue':
        if pickrole1[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick1[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick1[i],blueteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick1[i],blue1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick1[i],blue2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick1[i],blue3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick1[i],blue4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick1[i],blue5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick1[i],blue6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick1[i],blue7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick1[i],blue8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick1[i],blue9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick1[i],blue10thformat);
        elif pickrole2[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick2[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick2[i],blueteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick2[i],blue1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick2[i],blue2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick2[i],blue3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick2[i],blue4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick2[i],blue5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick2[i],blue6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick2[i],blue7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick2[i],blue8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick2[i],blue9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick2[i],blue10thformat);
        elif pickrole3[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick3[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick3[i],blueteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick3[i],blue1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick3[i],blue2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick3[i],blue3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick3[i],blue4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick3[i],blue5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick3[i],blue6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick3[i],blue7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick3[i],blue8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick3[i],blue9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick3[i],blue10thformat);
        elif pickrole4[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick4[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick4[i],blueteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick4[i],blue1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick4[i],blue2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick4[i],blue3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick4[i],blue4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick4[i],blue5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick4[i],blue6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick4[i],blue7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick4[i],blue8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick4[i],blue9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick4[i],blue10thformat);
        elif pickrole5[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick5[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick5[i],blueteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick5[i],blue1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick5[i],blue2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick5[i],blue3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick5[i],blue4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick5[i],blue5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick5[i],blue6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick5[i],blue7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick5[i],blue8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick5[i],blue9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick5[i],blue10thformat);
    elif teamside == 'red':
        if pickrole1[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick1[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick1[i],redteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick1[i],red1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick1[i],red2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick1[i],red3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick1[i],red4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick1[i],red5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick1[i],red6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick1[i],red7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick1[i],red8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick1[i],red9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick1[i],red10thformat);
        elif pickrole2[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick2[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick2[i],redteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick2[i],red1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick2[i],red2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick2[i],red3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick2[i],red4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick2[i],red5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick2[i],red6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick2[i],red7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick2[i],red8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick2[i],red9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick2[i],red10thformat);
        elif pickrole3[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick3[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick3[i],redteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick3[i],red1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick3[i],red2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick3[i],red3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick3[i],red4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick3[i],red5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick3[i],red6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick3[i],red7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick3[i],red8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick3[i],red9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick3[i],red10thformat);
        elif pickrole4[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick4[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick4[i],redteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick4[i],red1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick4[i],red2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick4[i],red3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick4[i],red4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick4[i],red5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick4[i],red6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick4[i],red7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick4[i],red8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick4[i],red9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick4[i],red10thformat);
        elif pickrole5[i][0] == role:
            istopten=0;
            for j in range(10):
                if toptenchampion[j] == pick5[i]:
                    istopten=j+1;
            if istopten == 0:
                worksheet.write(row,col,pick5[i],redteamformat);
            elif istopten == 1:
                worksheet.write(row,col,pick5[i],red1stformat);
            elif istopten == 2:
                worksheet.write(row,col,pick5[i],red2ndformat);
            elif istopten == 3:
                worksheet.write(row,col,pick5[i],red3rdformat);
            elif istopten == 4:
                worksheet.write(row,col,pick5[i],red4thformat);
            elif istopten == 5:
                worksheet.write(row,col,pick5[i],red5thformat);
            elif istopten == 6:
                worksheet.write(row,col,pick5[i],red6thformat);
            elif istopten == 7:
                worksheet.write(row,col,pick5[i],red7thformat);
            elif istopten == 8:
                worksheet.write(row,col,pick5[i],red8thformat);
            elif istopten == 9:
                worksheet.write(row,col,pick5[i],red9thformat);
            elif istopten == 10:
                worksheet.write(row,col,pick5[i],red10thformat);

def calculatetoptenpick(numberofmatch,teamside,blueban1,blueban2,blueban3,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,redban1,redban2,redban3,redpick1,redpick2,redpick3,redpick4,redpick5):
    totalpick=[];
    teamban=[];
    teampick=[];
    enemyban=[];
    enemypick=[];
    for i in range(numberofmatch):
        totalpick.append(blueban1[i]);
        totalpick.append(blueban2[i]);
        totalpick.append(blueban3[i]);
        totalpick.append(bluepick1[i]);
        totalpick.append(bluepick2[i]);
        totalpick.append(bluepick3[i]);
        totalpick.append(bluepick4[i]);
        totalpick.append(bluepick5[i]);
        totalpick.append(redban1[i]);
        totalpick.append(redban2[i]);
        totalpick.append(redban3[i]);
        totalpick.append(redpick1[i]);
        totalpick.append(redpick2[i]);
        totalpick.append(redpick3[i]);
        totalpick.append(redpick4[i]);
        totalpick.append(redpick5[i]);
        if teamside[i]:
            teamban.append(blueban1[i]);
            teamban.append(blueban2[i]);
            teamban.append(blueban3[i]);
            teampick.append(bluepick1[i]);
            teampick.append(bluepick2[i]);
            teampick.append(bluepick3[i]);
            teampick.append(bluepick4[i]);
            teampick.append(bluepick5[i]);
            enemyban.append(redban1[i]);
            enemyban.append(redban2[i]);
            enemyban.append(redban3[i]);
            enemypick.append(redpick1[i]);
            enemypick.append(redpick2[i]);
            enemypick.append(redpick3[i]);
            enemypick.append(redpick4[i]);
            enemypick.append(redpick5[i]);
        else:
            teamban.append(redban1[i]);
            teamban.append(redban2[i]);
            teamban.append(redban3[i]);
            teampick.append(redpick1[i]);
            teampick.append(redpick2[i]);
            teampick.append(redpick3[i]);
            teampick.append(redpick4[i]);
            teampick.append(redpick5[i]);
            enemyban.append(blueban1[i]);
            enemyban.append(blueban2[i]);
            enemyban.append(blueban3[i]);
            enemypick.append(bluepick1[i]);
            enemypick.append(bluepick2[i]);
            enemypick.append(bluepick3[i]);
            enemypick.append(bluepick4[i]);
            enemypick.append(bluepick5[i]);
    c=Counter(totalpick).most_common(10);
    d=Counter(teamban);
    e=Counter(teampick);
    f=Counter(enemyban);
    g=Counter(enemypick);
    toptenchampion=[];
    total=[];
    numberteamban=[];
    numberteampick=[];
    numberenemyban=[];
    numberenemypick=[];
    for idx, val in enumerate(c):
        toptenchampion.append(val[0]);
        numberteamban.append(d[val[0]]);
        numberteampick.append(e[val[0]]);
        numberenemyban.append(f[val[0]]);
        numberenemypick.append(g[val[0]]);
        total.append(val[1]);
    return toptenchampion,total,numberteampick,numberenemypick,numberteamban,numberenemyban;
   
def main():
    pageid = input("Please enter Esportspedia Page ID: ")
    teamname = input("Please enter the name of team: ")
    url = 'http://lol.esportspedia.com/w/api.php?action=query&pageids='+pageid+'&prop=revisions&rvprop=content&format=xml'
    http = urllib3.PoolManager()
    file = http.request('GET',url)
    data = file.data
    data = xmltodict.parse(data)
    data = data['api']['query']['pages']['page']['revisions']['rev']['#text']
    data="".join(data.split())
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
    winner=[];
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
            start = "winner=";
            end = "|";
            startPos = 0;
            endPos = 0;
            startPos = match[i].find(start)+len(start);
            endPos = match[i].find(end,startPos);
            winner.append(match[i][startPos:endPos])
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
                start = "winner=";
                end = "|";
                startPos = 0;
                endPos = 0;
                startPos = match[i].find(start)+len(start);
                endPos = match[i].find(end,startPos);
                winner.append(match[i][startPos:endPos])
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
        
    workbook = xlsxwriter.Workbook('banpicktable.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0;
    col = 0;
    weekwinformat = workbook.add_format({'bold': True, 'bg_color': '#00FF00'})
    weekloseformat = workbook.add_format({'bold': True, 'bg_color': '#FF0000'})
    blueteamformat = workbook.add_format({'bold': True, 'font_color': '#0000FF'})
    redteamformat = workbook.add_format({'bold': True, 'font_color': '#FF0000'})
    numberofmatch=len(blueteam);
    championTonameMap = championData.nameMap;
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
        blueban1[i]=championTonameMap[blueban1[i]];
        redban1[i]=championTonameMap[redban1[i]];
        blueban2[i]=championTonameMap[blueban2[i]];
        redban2[i]=championTonameMap[redban2[i]];
        blueban3[i]=championTonameMap[blueban3[i]];
        redban3[i]=championTonameMap[redban3[i]];
        bluepick1[i]=championTonameMap[bluepick1[i]];
        bluepick2[i]=championTonameMap[bluepick2[i]];
        bluepick3[i]=championTonameMap[bluepick3[i]];
        bluepick4[i]=championTonameMap[bluepick4[i]];
        bluepick5[i]=championTonameMap[bluepick5[i]];
        redpick1[i]=championTonameMap[redpick1[i]];
        redpick2[i]=championTonameMap[redpick2[i]];
        redpick3[i]=championTonameMap[redpick3[i]];
        redpick4[i]=championTonameMap[redpick4[i]];
        redpick5[i]=championTonameMap[redpick5[i]];
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
    toptenchampion,total,numberteampick,numberenemypick,numberteamban,numberenemyban=calculatetoptenpick(numberofmatch,teamside,blueban1,blueban2,blueban3,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,redban1,redban2,redban3,redpick1,redpick2,redpick3,redpick4,redpick5);
    for i in range(numberofmatch):
        if teamside[i]:
            if winner[i] == '1':
                worksheet.merge_range(row,col,row,col+1,i+1,weekwinformat);
            elif  winner[i] == '2':
                worksheet.merge_range(row,col,row,col+1,i+1,weekloseformat);
            col += 2;
            worksheet.write(row,col,blueteam[i],blueteamformat);
            col += 1;
            worksheet.write(row,col,blueban1[i],blueteamformat);
            col += 2;
            worksheet.write(row,col,blueban2[i],blueteamformat);
            col += 1;
            worksheet.write(row,col,blueban3[i],blueteamformat);
            col += 3;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'T','blue',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'J','blue',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'M','blue',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'A','blue',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'S','blue',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'T','red',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'J','red',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'M','red',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'A','red',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'S','red',toptenchampion);
            col += 3;
            worksheet.write(row,col,redban3[i],redteamformat);
            col += 1;
            worksheet.write(row,col,redban2[i],redteamformat);
            col += 2;
            worksheet.write(row,col,redban1[i],redteamformat);
            col += 1;
            worksheet.write(row,col,redteam[i],redteamformat);
            col += 1;
            if winner[i] == '1':
                worksheet.merge_range(row,col,row,col+1,i+1,weekloseformat);
            elif  winner[i] == '2':
                worksheet.merge_range(row,col,row,col+1,i+1,weekwinformat);
            row += 1;
            col = 0;
        else:
            if winner[i] == '1':
                worksheet.merge_range(row,col,row,col+1,i+1,weekloseformat);
            elif  winner[i] == '2':
                worksheet.merge_range(row,col,row,col+1,i+1,weekwinformat);
            col += 2;
            worksheet.write(row,col,redteam[i],redteamformat);
            col += 1;
            worksheet.write(row,col,redban1[i],redteamformat);
            col += 2;
            worksheet.write(row,col,redban2[i],redteamformat);
            col += 1;
            worksheet.write(row,col,redban3[i],redteamformat);
            col += 3;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'T','red',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'J','red',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'M','red',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'A','red',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,redpick1role,redpick2role,redpick3role,redpick4role,redpick5role,redpick1,redpick2,redpick3,redpick4,redpick5,'S','red',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'T','blue',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'J','blue',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'M','blue',toptenchampion);
            col += 1;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'A','blue',toptenchampion);
            col += 2;
            writepickwithselectedrole(workbook,worksheet,i,row,col,bluepick1role,bluepick2role,bluepick3role,bluepick4role,bluepick5role,bluepick1,bluepick2,bluepick3,bluepick4,bluepick5,'S','blue',toptenchampion);
            col += 3;
            worksheet.write(row,col,blueban3[i],blueteamformat);
            col += 1;
            worksheet.write(row,col,blueban2[i],blueteamformat);
            col += 2;
            worksheet.write(row,col,blueban1[i],blueteamformat);
            col += 1;
            worksheet.write(row,col,blueteam[i],blueteamformat);
            col += 1;
            if winner[i] == '1':
                worksheet.merge_range(row,col,row,col+1,i+1,weekwinformat);
            elif  winner[i] == '2':
                worksheet.merge_range(row,col,row,col+1,i+1,weekloseformat);
            row += 1;
            col = 0;
    row +=3;
    for i in range(10):
        worksheet.write(row,col,toptenchampion[i]);
        worksheet.write(row,col+1,total[i]);
        worksheet.write(row,col+2,numberteampick[i]);
        worksheet.write(row,col+3,numberenemypick[i]);
        worksheet.write(row,col+4,numberteamban[i]);
        worksheet.write(row,col+5,numberenemyban[i]);
        row += 1;
    workbook.close()
    
if __name__=="__main__":
    main()
