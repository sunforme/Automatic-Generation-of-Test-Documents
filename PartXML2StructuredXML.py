from shutil import copyfile
from xml.dom.minidom import parse
import xml.dom.minidom
import numpy as np
import os
def file_name(path):  # 传入存储的list
    L = []
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if not os.path.isdir(file_path):
            if os.path.splitext(file_path)[1] == '.xml':
                L.append(file_path)
    return L
def BuildCellNode(Cell, cell_node):
    Ranking = Cell.getAttribute('Ranking')
    cell_node.setAttribute('Ranking', Ranking)
    Paragraph = Cell.getAttribute('Paragraph')
    cell_node.setAttribute('Paragraph', Paragraph)
    Type = Cell.getAttribute('Type')
    cell_node.setAttribute('Type', Type)
    Font_list1 = Cell.getAttribute("Fonts")
    cell_node.setAttribute('Fonts', Font_list1)
    FontSize_list1 = Cell.getAttribute("FontSizes")
    cell_node.setAttribute('FontSizes', FontSize_list1)
    isBold1 = Cell.getAttribute("Bold")
    cell_node.setAttribute('Bold', isBold1)
    isItalic1 = Cell.getAttribute("Italic")
    cell_node.setAttribute('Italic', isItalic1)
    FontColor1 = Cell.getAttribute('FontColor')
    cell_node.setAttribute('FontColor', FontColor1)
    ColumnSpan = Cell.getAttribute('ColumnSpan')
    cell_node.setAttribute('ColumnSpan', ColumnSpan)
    ColumnSize = Cell.getAttribute('ColumnSize')
    cell_node.setAttribute('ColumnSize', ColumnSize)
    RowSpan = Cell.getAttribute('RowSpan')
    cell_node.setAttribute('RowSpan', RowSpan)
    RowSize = Cell.getAttribute('RowSize')
    cell_node.setAttribute('RowSize', RowSize)
    return cell_node
def BuildTree(cellnode, Cell):
    PreCells = Cell.childNodes
    Cell_s = []
    for PreCell in PreCells:
        if PreCell.nodeName == 'Cell':
            Cell_s.append(PreCell)
    if len(Cell_s) != 0:
        for Cell_ in Cell_s:
            cell_node = dom.createElement('Cell')
            cell_node = BuildCellNode(Cell_, cell_node)
            cell_node = BuildTree(cell_node, Cell_)
            cellnode.appendChild(cell_node)
    return cellnode
def getChild(cell):
    cellss = []
    if cell.hasChildNodes():
        childnodes = cell.childNodes
        for childnode in childnodes:
            cellss.append(childnode)
            cells1 = getChild(childnode)
            if len(cells1) != 0:
                for cell1 in cells1:
                    cellss.append(cell1)
    return cellss
def getTextHead(textData):
    ChineseNum = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖',
                  '拾', '第']
    symbol = ['-', '.', '(', ')', '（', '）', '、', '·', '，', ',']
    foreignNum = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
                  'I', 'II', 'III','IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII',
                  'Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ', 'Ⅺ', 'Ⅻ',
                  'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x',
                  'ⅰ', 'ⅱ', 'ⅲ', 'ⅳ', 'ⅴ', 'ⅵ', 'ⅶ', 'ⅷ', 'ⅸ', 'ⅹ', 'α', 'β', 'γ', 'δ', 'ε']
    serial = ''
    for char in textData:
        if (ord(char) in (97, 122)) or (
                ord(char) in (
        65, 90)) or char.isspace() or char.isdigit() or char in ChineseNum or char in symbol or char in foreignNum:
            serial = serial + char
        else:
            break
    CharList = []
    for char1 in serial:
        if not (char1.isspace() or char1 in symbol):
            CharList.append(char1)
    return CharList
def getBestPreHead(table_node1, precellList, currentCell): #当前单元格占全列时，获取前边占全列单元格中最符合当前单元格父节点的单元格
    bestPreHead = table_node1
    textData1 = currentCell.getAttribute("Paragraph")
    currentCellHead = getTextHead(textData1)
    FontList1 = currentCell.getAttribute("Fonts")
    FontSizeList1 = currentCell.getAttribute("FontSizes")
    isBold1 = currentCell.getAttribute("Bold")
    isItalic1 = currentCell.getAttribute("Italic")
    fontcolors1 = currentCell.getAttribute('FontColor')
    flag1 = False
    flag2 = False
    flag3 = False
    for precell in precellList:
        textData = precell.getAttribute("Paragraph")
        precellHead = getTextHead(textData)
        if len(currentCellHead) != 0 and len(currentCellHead) == len(precellHead) + 1:
            equal = True
            for charNum in range(len(currentCellHead)-1):
                if currentCellHead[charNum] != precellHead[charNum]:
                    equal = False
            if equal:
                childNodes = getChild(table_node1)
                for childNode in childNodes:
                    if childNode.getAttribute('Ranking') == precell.getAttribute("Ranking"):
                        bestPreHead = childNode
                        flag1 = True
    if not flag1:
        for precell in precellList:
            textData = precell.getAttribute("Paragraph")
            precellHead = getTextHead(textData)
            if len(currentCellHead) != 0 and len(currentCellHead) == len(precellHead):
                equal = True
                for charNum in range(len(currentCellHead) - 1):
                    if currentCellHead[charNum] != precellHead[charNum]:
                        equal = False
                if equal:
                    childNodes = getChild(table_node1)
                    for childNode in childNodes:
                        if childNode.getAttribute('Ranking') == precell.getAttribute("Ranking"):
                            bestPreHead = childNode.parentNode
                            flag2 = True
    if (not flag2) and (not flag1) and len(currentCellHead) == 0 and currentCell.getAttribute("ColumnSize") == table_node1.getAttribute("列数"):
        for precell in precellList:
            FontList = precell.getAttribute("Fonts")
            FontSizeList = precell.getAttribute("FontSizes")
            isBold = precell.getAttribute("Bold")
            isItalic = precell.getAttribute("Italic")
            fontcolors = precell.getAttribute('FontColor')
            if FontList==FontList1 and FontSizeList == FontSizeList1 and isBold ==isBold1 and isItalic ==isItalic1 and fontcolors == fontcolors1:
                childNodes = getChild(table_node1)
                for childNode in childNodes:
                    if childNode.getAttribute('Ranking') == precell.getAttribute("Ranking"):
                        if childNode.parentNode.hasAttribute("Paragraph"):
                            parentText = childNode.parentNode.getAttribute("Paragraph")
                            if len(parentText) != 0:
                                parentTextHead = getTextHead(parentText)
                                if len(parentTextHead) != 0:
                                    bestPreHead = childNode.parentNode
                                    flag3 = True
                        elif childNode.parentNode.nodeName == 'Table':
                            bestPreHead = childNode.parentNode
                            flag3 = True
    if (not flag2) and (not flag1) and (not flag3):
        precell = precellList[-1]
        childNodes = getChild(table_node1)
        for childNode in childNodes:
            if childNode.getAttribute('Ranking') == precell.getAttribute("Ranking"):
                bestPreHead = childNode
    return bestPreHead
XML_path = 'D:\\PythonWorkspace\\TableRecognition\\TableIdentification\\Data\\PredictPartStructuredXML'
XML_files = file_name(XML_path)
for XML_file in XML_files:
    print(XML_file)
    XMLName = os.path.basename(XML_file)
    # 使用minidom解析器打开 XML 文档
    DOMTree = xml.dom.minidom.parse(XML_file)
    AllTables = DOMTree.documentElement
    tables = AllTables.getElementsByTagName('Table')
    # 1.创建DOM树对象
    dom = xml.dom.minidom.Document()
    # 2.创建根节点。每次都要用DOM对象来创建任何节点。
    root_node = dom.createElement('Tables')
    # 3.用DOM对象添加根节点
    dom.appendChild(root_node)
    for table in tables:
        TableRowSize = table.getAttribute("行数")
        TableColSize = table.getAttribute("列数")
        table_node = dom.createElement('Table')
        root_node.appendChild(table_node)
        table_node.setAttribute("列数", TableColSize)
        table_node.setAttribute("行数", TableRowSize)
        PreHeadList = []
        parts = table.getElementsByTagName('Part')
        for part in parts:
            PreCells = part.childNodes
            Cells = []
            for PreCell in PreCells:
                if PreCell.nodeName == 'Cell':
                    Cells.append(PreCell)
            for Cell in Cells:
                if len(PreHeadList) != 0:
                    precell_list = table_node
                    Ranking = int(Cell.getAttribute("Ranking"))
                    UpIsParent = False
                    for PreHead in PreHeadList:
                        if int(PreHead.getAttribute("Ranking")) == Ranking - 1:
                            childNodes = getChild(table_node)
                            for childNode in childNodes:
                                if childNode.getAttribute('Ranking') == PreHead.getAttribute("Ranking"):
                                    precell_list = childNode
                                    UpIsParent = True
                    if not UpIsParent:  # 如果前一个单元格不占全列，但前面的单元格还有占全列的，并且是表头，判断当前单元格的父单元格是哪个
                        precell_list = getBestPreHead(table_node, PreHeadList, Cell)
                    cell_node = dom.createElement('Cell')
                    cell_node = BuildCellNode(Cell, cell_node)
                    cell_node = BuildTree(cell_node, Cell)
                    precell_list.appendChild(cell_node)
                    if Cell.getAttribute("ColumnSize") == TableColSize and Cell.getAttribute("Type") == 'indication':
                        PreHeadList.append(Cell)
                else:
                    cell_node = dom.createElement('Cell')
                    cell_node = BuildCellNode(Cell, cell_node)
                    cell_node = BuildTree(cell_node, Cell)
                    table_node.appendChild(cell_node)
                    if Cell.getAttribute("ColumnSize") == TableColSize and Cell.getAttribute("Type") == 'indication':
                        PreHeadList.append(Cell)
    StructuredXMLPath = 'D:\\PythonWorkspace\\TableRecognition\\TableIdentification\\Data\\PredictStructuredXML\\' + XMLName
    try:
        with open(StructuredXMLPath, 'w', encoding='UTF-8') as fh:
            dom.writexml(fh, indent='', addindent='\t', newl='\n', encoding='UTF-8')
            print('整体逻辑结构XML OK!')
    except Exception as err:
        print('错误信息：{0}'.format(err))
