from collections import OrderedDict, deque
import json
from PIL import ImageFont, Image
import re

import xlsxwriter as xw


WB_PATH = 'hi3guide.xlsm'
VALKYRIE_DATA_PATH = 'json/valkyrie.json'

class Main():
    def __init__(self):
        self.wb = None
        self.ws = None
        self.currCellR = 0
        self.currCellC = 1
        self.defaultCellWidth = 3.5
    
        self.data = None
        self.currValkData = None
        
        ## Use font Calibri, fontsize 11
        self.font = ImageFont.truetype('calibri.ttf', size=16)
        ## First row height is 15 pts, the default in Excel
        
        
    def run(self, initAsXlsm=True):
        self.loadData()
        self.createWorkbook(initAsXlsm)
        self.initializeWorkbook()
        self.writeGuide()
        
        self.terminate()
    
    def loadData(self):
        global data
        with open(VALKYRIE_DATA_PATH, 'r', encoding='utf-8') as jsonFile:
            self.data = json.load(jsonFile, object_hook=OrderedDict)
    
        data = self.data 
        
    def createWorkbook(self, initAsXlsm):
        if initAsXlsm:
            self.wb = xw.Workbook(WB_PATH)
            
            ## add dummy vba project so that xs will save it as xlsm instead of xls
            self.wb.add_vba_project('./vbaProject.bin')
            
        else:
            self.wb = xw.Workbook(WB_PATH[:-1] + 'x')
        
    
    def initializeWorkbook(self):
        colorG8 = '#E34234'
        colorGW = '#1B4D3E'
        titleFontSize = 20
        infoFontSize = 16
        
        self.formatTopTitle = self.wb.add_format({
                'align': 'center_across',
                'font_size': infoFontSize})
        self.formatTitle = self.wb.add_format({
                'align': 'center_across',
                'font_size': titleFontSize,
                'bold': True,
                'valign': 'vcenter'})
        self.formatSection = self.wb.add_format({
                'align': 'center',
                'font_size': infoFontSize})
        self.formatSubsection = self.wb.add_format({
                'align': 'center',
                'font_size': infoFontSize})
        self.formatInfo = self.wb.add_format({
                #'align': 'center_across',
                'text_wrap': True,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatStat = self.wb.add_format({
                'align': 'center_across',
                'text_wrap': True,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatInfoG8 = self.wb.add_format({
                'color': colorG8,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatInfoGW = self.wb.add_format({
                'color': colorGW,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatInfoTitle = self.wb.add_format({
                'bold': True,
                'align': 'justify',
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatEquipment = self.wb.add_format({
                'align': 'center_across',
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatEquipmentGW = self.wb.add_format({
                'align': 'center_across',
                'text_wrap': True,
                'valign': 'vcenter',
                'color': colorGW,
                'font_size': infoFontSize})
        self.formatEquipmentG8 = self.wb.add_format({
                'align': 'center_across',
                'text_wrap': True,
                'valign': 'vcenter',
                'color': colorG8,
                'font_size': infoFontSize})
        self.formatEquipmentInfo = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter'})
        self.formatSkillType = self.wb.add_format({
                'text_wrap': True,
                'bold': True,
                'font_size': infoFontSize})
        self.formatPotentialHeader = self.wb.add_format({
                'align': 'center_across',
                'text_wrap': True,
                'bold': True,
                'font_size': infoFontSize})
        self.formatPotentialRank = self.wb.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'bold': True,
                'font_size': infoFontSize})
        self.formatPotentialStar = self.formatPotentialRank
        self.formatPotentialSkills = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter'})
        self.formatPotentialSkillsDescGW = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter',
                'color': colorGW})
        self.formatPotentialSkillsDescG8 = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter',
                'color': colorG8})
        self.formatSynergyName = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter',
                'align': 'center',
                'bold': True
                })
        self.formatSynergyG8 = self.wb.add_format({
                'text_wrap': True,
                'color': colorG8,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatSynergyGW = self.wb.add_format({
                'text_wrap': True,
                'color': colorGW,
                'valign': 'vcenter',
                'font_size': infoFontSize})
        self.formatSynergy = self.wb.add_format({
                'align': 'center_across',
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter'})
        self.formatSynergyDesc = self.wb.add_format({
                'text_wrap': True,
                'font_size': infoFontSize,
                'valign': 'vcenter'})
        
    def writeGuide(self):
        for valk in self.data:
            self.currCellR = 0
            self.currCellC = 1
            self.rowHeights = [15]
            self.ws = self.wb.add_worksheet()
            self.currValkData = self.data[valk]
            self.initializeWorksheet()
            self.writeCharGuide()
            self.fillTypeColor()
            
    def initializeWorksheet(self):
        self.ws.set_column(0, 25, self.defaultCellWidth)
        self.ws.hide_gridlines(2)
        self.ws.name = self.currValkData['name'] + ' ' + self.currValkData['char'] 
        
    def writeCharGuide(self):
        self.writeName()
        self.addEmptyRow()
        self.addValkAvatar()
        self.writeScore()
        self.addEmptyRow()
        self.writeStrengths()
        self.addEmptyRow()
        self.writeWeaknesses()
        self.addEmptyRow()
        self.writePremiumLoadout()
        self.addEmptyRow()
        self.writeDiscountLoadout()
        self.addEmptyRow()
        self.writePotential()
        self.addEmptyRow()
        self.writeSynergy()
        
    def addEmptyRow(self):
        self.nextRowWrite('', self.formatInfo)
        
    def writeName(self):
        self.nextRowWrite(self.currValkData['name'] + ' ' + self.currValkData['char'], 
                          self.formatTitle, minRowHeight=68)
        
    def addValkAvatar(self):
        valkName = self.currValkData['name']
        valkName = re.sub(r'\W+', '', valkName)
        
        IMG_DIR = 'img/valkyrie/' + valkName + '.png'
        self.addImage(1, 2, 2, 6, IMG_DIR)
        
    def writeScore(self):
        self.nextRowWrite("Score", self.formatTitle)
        self.nextRowWrite(("Damage", 
                                (self.formatInfoG8, self.currValkData["score-damage-g8"]), 
                                "Support", 
                                (self.formatInfoG8, self.currValkData["score-support-g8"])),
                                (self.formatStat,) * 4, (6,) * 4)
        self.nextRowWrite(("Interrupt", (self.formatInfoG8, self.currValkData["score-interrupt-g8"]), 
                                "Difficulty", (self.formatInfoG8, self.currValkData["score-difficulty-g8"])),
                                (self.formatStat,) * 4, (6,) * 4)     
        self.nextRowWrite(("Co-op", 
                                (self.formatInfoG8, self.currValkData["score-coop-g8"],
                                 self.formatInfo, "\n",
                                 self.formatInfoGW, self.currValkData["score-coop-gw"]),
                                "Infinite Abyss", 
                                (self.formatInfoG8, self.currValkData["score-abyss-g8"],
                                 self.formatInfo, "\n",
                                 self.formatInfoGW, self.currValkData["score-abyss-gw"]),
                                 "Memorial Arena", 
                                (self.formatInfoG8, self.currValkData["score-arena-g8"],
                                 self.formatInfo, "\n",
                                 self.formatInfoGW, self.currValkData["score-arena-gw"])),
                                (self.formatStat,) * 6, (4,) * 6)                     
    
    def writeStrengths(self):
        self.nextRowWrite("Strengths", self.formatTitle)
        for strength in self.currValkData['strengths']:
            self.addTitledDesc(strength)
            
    def writeWeaknesses(self):
        self.nextRowWrite("Weaknesses", self.formatTitle)
        for weakness in self.currValkData['weaknesses']:
            self.addTitledDesc(weakness)
            
    def writePremiumLoadout(self):
        self.nextRowWrite("Recommended Premium Sets", self.formatTitle)
        for loadout in self.currValkData['loadouts-premium']:
            self.addLoadout(loadout)
            self.addEmptyRow()
        
    def writeDiscountLoadout(self):
        self.nextRowWrite("Recommended Discount Sets", self.formatTitle)
        for loadout in self.currValkData['loadouts-discount']:
            self.addLoadout(loadout)
            self.addEmptyRow()
    
    def writePotential(self):
        self.nextRowWrite('Potential', self.formatTitle)
        self.addPotentialHeader()
        for rank in self.currValkData['potential']:
            self.addPotentialRank(rank, self.currValkData['potential'][rank])
            
    def writeSynergy(self):
        valkDir = 'img/valkyrie/'
        self.nextRowWrite('Team Synergy',self.formatTitle)
        for teammate in self.currValkData['team-synergy']:
            richString = self.teamSynergyToRichString(teammate)
            self.nextRowWrite(richString, 
                              (self.formatInfo, self.formatSynergy, self.formatSynergyDesc), 
                              (3, 6, 15), merged=True, minRowHeight=67.5)
            
            valkName = teammate['valk-name']
            valkName = re.sub(r'[^\w+\(\)]', '', valkName)
            self.addImage(self.currCellR, 1, self.currCellR, 3, 
                          valkDir + valkName + '_c.png')
            
    def fillTypeColor(self):
        colors = ['#c8f8fd', 'orange', 'purple']
        color = colors[self.currValkData['type']]
        cell_format = self.wb.add_format()
        cell_format.set_bg_color(color)
        for i in range(self.currCellR + 2):
            if i == 0 or i == self.currCellR + 1:
                for j in range(26):
                    self.ws.write(i, j, '', cell_format)
            else:
                self.ws.write(i, 0, '', cell_format)
                self.ws.write(i, 25, '', cell_format)
             
            
    def teamSynergyToRichString(self, teamData):
        richString = ('',)
        richString += ((self.formatSynergyName, teamData['valk-name']),)
        
        descString = tuple()
        
        if 'desc-gw' in teamData:
            descString += (self.formatSynergyGW, 
                           teamData['desc-gw'])
        if 'desc-gw' in teamData and 'desc-g8' in teamData:
            descString += (self.formatSynergyGW, '\n',)
        if 'desc-g8' in teamData:
            descString += (self.formatSynergyG8, 
                           teamData['desc-g8'])
        richString += (descString,)
        return richString
        
        
    
    def addPotentialRank(self, rank, rankData):
        isFirst = True
        
        skillPriority = self.skillPriorityToRichString(rankData)
        for skill in rankData['skills']:
            skillStr = self.getSkillStr(skill)
            if isFirst:
                self.nextRowWrite((rank, skillPriority, skillStr), 
                                  (self.formatPotentialRank, 
                                   self.formatPotentialStar, 
                                   self.formatPotentialSkills),  
                                  (4, 4, 16), merged=(False, False, True))
                rowStart = self.currCellR
                isFirst = False
            else:
                self.nextRowWrite(('', skillStr), 
                                  (self.formatPotentialSkills, ) * 2, 
                                  (8, 16), merged=(False, True))
        else:
            self.nextRowWrite(('', rankData['rank-desc-gw']), 
                                  (self.formatPotentialSkillsDescGW, ) * 2, 
                                  (8, 16), merged=(False, True))
            rowEnd = self.currCellR
            
        self.ws.merge_range(rowStart, 1, rowEnd, 4, '')
        self.ws.merge_range(rowStart, 5, rowEnd, 8, '')
        
        rankDir = 'img/rank/'
        self.addImage(rowStart, 2, rowEnd, 3, rankDir + rank + '.png', center=False)
        
    def addPotentialHeader(self):
        self.nextRowWrite(('Rank', 'Priority', 'Description'), 
                          (self.formatPotentialHeader,) * 3, 
                          (4, 4, 16))
                
            
    def getSkillStr(self, skillJson):
        typeList = ['Leader', 'Passive', 'Evasion', 'Basic', 'Ultimate', 'Special']
        return ('⟦ ' + typeList[skillJson['skill-type']] + ' ⟧   ' + skillJson['skill-name'])
    
    def skillPriorityToRichString(self, rankData):
        ## temporily use equipment string format
        richString = tuple()
        fullStar, emptyStar = '★', '☆'
        if 'priority-gw' in rankData:
            richString += (self.formatEquipmentGW, 
                           (fullStar * rankData['priority-gw'] + 
                           emptyStar * (3 - rankData['priority-gw'])))
        if 'priority-gw' in rankData and 'priority-g8' in rankData:
            richString += (('\n'),)
        if 'priority-g8' in rankData:
            richString += (self.formatEquipmentG8, 
                           (fullStar * rankData['priority-g8'] + 
                           emptyStar * (3 - rankData['priority-g8'])))
        return richString            
            
    def addLoadout(self, loadout):
        loadoutGrade = self.loadoutScoreToRichString(loadout)
        richDescStr, numLines = self.getLoadoutDesc(loadout)
        
        self.nextRowWrite((loadoutGrade, 'Description'), 
                          (self.formatEquipment, ) * 2,
                          (8, 16))
        
        weaponDir = 'img/weapon/'
        stigmataDir = 'img/stigmata/'
        descRowStart = self.currCellR + 1
        rowHeight = 80
        
        if type(loadout['weapon']) is str:
            weapons = (loadout['weapon'],)
        else:
            weapons = loadout['weapon']
            
        for weapon in weapons:
            self.nextRowWrite(('', weapon),
                              (self.formatEquipment,) * 3,
                              (4, 4), minRowHeight=rowHeight)
            weapon = re.sub(r'[^\w+\(\)]', '', weapon)
            self.addImage(self.currCellR, 1, self.currCellR, 4,
                              weaponDir + weapon + '.png')
                
        if type(loadout['stigT']) is str:
            stigTs = (loadout['stigT'],)
        else:
            stigTs = loadout['stigT']

        for stigT in stigTs:
            self.nextRowWrite(('', stigT),
                               (self.formatEquipment,) * 3,
                               (4, 4), minRowHeight=rowHeight)
            stigT = re.sub(r'[^\w+\(\)]', '', stigT)
            self.addImage(self.currCellR, 1, self.currCellR, 4, 
                          stigmataDir + stigT + '.png')
            
        if type(loadout['stigM']) is str:
            stigMs = (loadout['stigM'],)
        else:
            stigMs = loadout['stigM']

        for stigM in stigMs:
            self.nextRowWrite(('', stigM),
                               (self.formatEquipment,) * 3,
                               (4, 4), minRowHeight=rowHeight)
            stigM = re.sub(r'[^\w+\(\)]', '', stigM)
            self.addImage(self.currCellR, 1, self.currCellR, 4, 
                          stigmataDir + stigM + '.png')
            
        if type(loadout['stigB']) is str:
            stigBs = (loadout['stigB'],)
        else:
            stigBs = loadout['stigB']

        for stigB in stigBs:
            self.nextRowWrite(('', stigB),
                               (self.formatEquipment,) * 3,
                               (4, 4), minRowHeight=rowHeight)
            stigB = re.sub(r'[^\w+\(\)]', '', stigB)
            self.addImage(self.currCellR, 1, self.currCellR, 4, 
                          stigmataDir + stigB + '.png')
        
        descRowEnd = self.currCellR
        self.ws.merge_range(descRowStart, 9, descRowEnd, 24, '')
        self.ws.write_rich_string(descRowStart, 9, *richDescStr, 
                                  self.formatEquipmentInfo)
        
    def getLoadoutDesc(self, loadout):
        lineCount = 0
        richString = tuple()
        if 'desc-gw' in loadout:
            richString += (self.formatInfoGW, loadout['desc-gw'])
        if 'desc-gw' in loadout and 'desc-g8' in loadout:
            richString += ('\n',)
        if 'desc-g8' in loadout:
            richString += (self.formatInfoG8, loadout['desc-g8'])
        return richString, lineCount
        
                
    def loadoutScoreToRichString(self, loadout):
        richString = tuple()
        fullStar, emptyStar = '★', '☆'
        if 'rating-gw' in loadout:
            richString += (self.formatEquipmentGW, 
                           (fullStar * loadout['rating-gw'] + 
                           emptyStar * (loadout['rating-gw-max'] - loadout['rating-gw'])))
        if 'rating-gw' in loadout and 'rating-g8' in loadout:
            richString += (self.formatEquipment, '\n')
        if 'rating-g8' in loadout:
            richString += (self.formatEquipmentG8, "❤︎")
        return richString            
        
            
    def addTitledDesc(self, info):
        '''
            info: a dict-like object containing title and desc
        '''
        richString = (self.formatInfoTitle, info['title'])
        
        for key in info:
            src = key.split('-')[-1]
            if src == 'gw':
                style = self.formatInfoGW
            elif src == 'g8':
                style = self.formatInfoG8
            else:
                continue
        
            richString += ('\n', 
                           style, info[key])
            self.nextRowWrite(richString, self.formatInfo, merged=True)
            #self.ws.set_row(self.currCellR, (len(lines) + 1) * 25)
            
    def addImage(self, row1, col1, row2, col2, imgDir, autoscale=True, center=True, vcenter=True):
        # pixel to point ratio is 1.333 for height and 8.403 for width. May
        # depend on font.
        rangeHeight = sum(self.rowHeights[row1:row2+1]) * 1.333
        rangeWidth = (col2 - col1 + 1) * self.defaultCellWidth * 8.403
        
        img = Image.open(imgDir)
        imgWidth, imgHeight = img.size
        
        heightRatio = rangeHeight / imgHeight
        widthRatio = rangeWidth / imgWidth
        
        if autoscale:
            scale = min(heightRatio, widthRatio)
        else:
            scale = 1
            
        if center:
            xOffset = (rangeWidth - imgWidth * scale) / 2
        else:
            xOffset = 0
        
        if vcenter:
            yOffset = (rangeHeight - imgHeight * scale) / 2
        else:
            yOffset = 0
            
        options = {
                    'x_offset':    xOffset,
                    'y_offset':    yOffset,
                    'x_scale':     scale,
                    'y_scale':     scale,
                    'positioning': 1
                    }
        self.ws.insert_image(row1, col1, imgDir,
                             options)
        
        
        
    
    def nextRowWrite(self, strings, styles, spaces=(24,), merged=False, minRowHeight=None):
        '''
        Writes the value(s) of the next row. Can take in raw string, a tuple 
        of strings, rich strings in the form of a tuple, or a tuple of 
        strings and rich strings.
        
        The type of style must be consistent with that of strings, e.g. both 
        must be tuples of the same length (for tuples of raw strings) or 
        type of string is str and type of styles is xw.format.Format.
        
        The row being written will then be automatically resized according to 
        the number of lines taken. 
        '''
        
        def countNewLines(item):
            if type(item) is tuple:
                return sum(map(countNewLines, item))
            elif type(item) is xw.format.Format:
                return 0
            
            return item.count('\n')
        
        self.currCellR += 1
        self.currCellC = 1
        maxNumLines = 0
        maxRowHeight = 0
        
        if type(strings) == str or self.isRichString(strings):
            strings = (strings,)
            styles = (styles, )
            
        if merged is True or merged is False:
            merged = (merged, ) * len(strings)
            
        for i, string in enumerate(strings):
            if styles[i] is None:
                fontSize = 16
            else:
                fontSize = styles[i].font_size
                
            if merged[i]:
                self.ws.merge_range(self.currCellR, self.currCellC, 
                                    self.currCellR, self.currCellC + spaces[i] - 1,
                                    '')
            if self.isRichString(string):
                splittedRichStr = self.splitRichStringForWrap(string, spaces[i])
                numLines = countNewLines(splittedRichStr) + 1
                
                if styles[i] is not None:
                    self.ws.write_rich_string(self.currCellR, self.currCellC, *splittedRichStr, styles[i])
                else:
                    self.ws.write_rich_string(self.currCellR, self.currCellC, *splittedRichStr)
                    
            else:
                splittedStr = self.splitSimpleStringForWrap(string, spaces[i])
                numLines = countNewLines(splittedStr) + 1
                if styles[i] is not None:
                    self.ws.write(self.currCellR, self.currCellC, splittedStr, styles[i])
                else:
                    self.ws.write(self.currCellR, self.currCellC, splittedStr)
            
            if not merged[i] and styles[i] is not None:
                for j in range(1, spaces[i]):
                    self.ws.write(self.currCellR, self.currCellC + j, '', styles[i]) 
                    
            
            maxNumLines = max(maxNumLines, numLines)
            maxRowHeight = max(maxRowHeight, numLines * 25 * fontSize / 16)
            
            if minRowHeight is not None:
                maxRowHeight = max(minRowHeight, maxRowHeight) 
            
            self.currCellC += spaces[i]
            
        self.ws.set_row(self.currCellR, maxRowHeight)
        self.rowHeights.append(maxRowHeight)
                           
                
    def isRichString(self, richString):
        if type(richString) is not tuple:
            return False
        
        for ele in richString:
            if isinstance(ele, xw.format.Format):
                return True
        
        return False
    
    def splitStringForWrap(self, longString, space):
        #strings = longString.split('\n')
        strings = [longString]
        newStrings = []
        for string in strings:
            lines = self.wordWrap(string, space)
            newStrings.extend(lines)
        
        return newStrings
    
    def splitSimpleStringForWrap(self, longString, space):
        return '\n'.join(self.splitStringForWrap(longString, space))
    
    def splitRichStringForWrap(self, richString, space):
        if len(richString) == 0:
            return tuple()
        
        ele, *richString = richString
        if isinstance(ele, xw.format.Format):
            string, *richString = richString
        else:
            string = ele
            
        newString = '\n'.join(self.splitStringForWrap(string, space))
        
        if isinstance(ele, xw.format.Format):
            richLines = (ele, newString)
        
        else:
            richLines = (newString, )
        
        return richLines + self.splitRichStringForWrap(richString, space)
    
    def wordWrap(self, string, space=24):
        assert string != None
        
        words = deque(string.split(' '))
        lineToAdd = ''
        lines = []
        
        while len(words) > 0:
            hasNewLine = False
            word = words.popleft()
            if '\n' in word:
                hasNewLine = True
                word, *remainder = word.split('\n')
                words.extendleft(remainder[::-1])
                
            totalLength = self.font.getsize(lineToAdd + word)[0]
            assert type(totalLength) is int
            if totalLength > 530 * space / 24:
                lines.append(lineToAdd)
                lineToAdd = word + ' '
            else:
                lineToAdd += word + ' '
            if hasNewLine:
                lines.append(lineToAdd.strip())
                lineToAdd = ''
        lines.append(lineToAdd.strip())
        
        return lines
        
    
    def terminate(self):
        self.wb.close()
            

if __name__ == '__main__':
    main = Main()
    main.run()