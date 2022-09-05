from zipfile import ZipFile 
from lxml import etree
import re
import tempfile
import requests
import os
import shutil
from PIL import Image

def pretty_print(n):
	print(etree.tostring(n,pretty_print=True, encoding='utf-8').decode())
					
def fix_listing(tmp_dir): # Замена таблицы стилей для списков
	f = open(os.path.join(tmp_dir,"word/numbering.xml"), "wb")
	s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:pStyle w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:smallCaps w:val="false"/><w:caps w:val="false"/><w:dstrike w:val="false"/><w:strike w:val="false"/><w:vertAlign w:val="baseline"/><w:position w:val="0"/><w:sz w:val="26"/><w:sz w:val="26"/><w:spacing w:val="0"/><w:i w:val="false"/><w:u w:val="none"/><w:b/><w:kern w:val="0"/><w:effect w:val="none"/><w:iCs w:val="false"/><w:bCs w:val="false"/><w:em w:val="none"/><w:vanish w:val="false"/><w:rFonts w:cs="Times New Roman"/><w:color w:val="000000"/></w:rPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:pStyle w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%3."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="4899" w:hanging="504"/></w:pPr><w:rPr><w:smallCaps w:val="false"/><w:caps w:val="false"/><w:dstrike w:val="false"/><w:strike w:val="false"/><w:vertAlign w:val="baseline"/><w:position w:val="0"/><w:sz w:val="26"/><w:sz w:val="26"/><w:spacing w:val="0"/><w:i w:val="false"/><w:u w:val="none"/><w:b w:val="false"/><w:kern w:val="0"/><w:effect w:val="none"/><w:iCs w:val="false"/><w:bCs w:val="false"/><w:em w:val="none"/><w:vanish w:val="false"/><w:rFonts w:cs="Times New Roman"/><w:color w:val="000000"/></w:rPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="0"/></w:pPr></w:lvl></w:abstractNum>'
	for i in range (2, 50):
		s += '<w:abstractNum w:abstractNumId="'+str(i)+'"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="227"/></w:tabs><w:ind w:left="227" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="454"/></w:tabs><w:ind w:left="454" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="680"/></w:tabs><w:ind w:left="680" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="907"/></w:tabs><w:ind w:left="907" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1134"/></w:tabs><w:ind w:left="1134" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1361"/></w:tabs><w:ind w:left="1361" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1587"/></w:tabs><w:ind w:left="1587" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1814"/></w:tabs><w:ind w:left="1814" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="2041"/></w:tabs><w:ind w:left="2041" w:hanging="227"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
	for i in range (51, 100):
		s += '<w:abstractNum w:abstractNumId="'+str(i)+'"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="397"/></w:tabs><w:ind w:left="754" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="794"/></w:tabs><w:ind w:left="1151" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1191"/></w:tabs><w:ind w:left="1548" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%4."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1588"/></w:tabs><w:ind w:left="1945" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%5."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1985"/></w:tabs><w:ind w:left="2342" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%6."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="2381"/></w:tabs><w:ind w:left="2738" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="2778"/></w:tabs><w:ind w:left="3135" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%8."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="3175"/></w:tabs><w:ind w:left="3532" w:hanging="397"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%9."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="3572"/></w:tabs><w:ind w:left="3929" w:hanging="397"/></w:pPr></w:lvl></w:abstractNum>'
	for i in range (1, 100):
		s += '<w:num w:numId="'+str(i)+'"><w:abstractNumId w:val="'+str(i)+'"/></w:num>'
	s += '</w:numbering>'
	f.write(s.encode())
	f.close()
		
def get_t_of_p(namespace, p):
	text = ''
	for t in p.findall('.//'+namespace+'t'):
		if t.text != None:
			text += t.text
	return text
		
def get_schemes(root, namespace):
	refs = 0
	refname = []
	for t in root.findall('.//'+namespace+'t'): # Найти список схем	
		if t.text == "Схемы":
			p = t.getparent().getparent() #P
			while p.getnext() != None: # До конца документа
				p = p.getnext() #P
				if p.tag == namespace+'p':
					for x in p.findall('.//'+namespace+'t'): #T
						if x.text != None: # Название таблицы
							if len(x.text) > 1:
								k = x.getparent()
								refs += 1 #Вставка назначения ссылки
								n = etree.Element(namespace+"bookmarkStart")
								n.attrib[namespace+'id']=str(refs)
								n.attrib[namespace+'name']="Ref_R"+str(refs)
								k.addprevious(n)			
								
								n = etree.Element(namespace+"bookmarkEnd")
								n.attrib[namespace+'id']=str(refs)
								k.addnext(n)
								refname.append(x.text)
	
	for tc in root.findall('.//'+namespace+'tc'): # Вставить ссылки
		if get_t_of_p(namespace, tc) == "Конфигурация":
			tc = tc.getparent().find(namespace+'tc[2]') #Найти ячейку с названием схемы
			if get_t_of_p(namespace, tc) in refname:
				r = tc.findall('.//'+namespace+'r')
				
				r1 = etree.Element(namespace+"r")
				rpr = etree.SubElement(r1, namespace+"rPr")
				l1 = etree.SubElement(r1, namespace+"fldChar")
				l1.attrib[namespace+"fldCharType"]="begin"
				
				r2 = etree.Element(namespace+"r")
				rpr = etree.SubElement(r2, namespace+"rPr")
				l2 = etree.SubElement(r2, namespace+"instrText")
				l2.text = " REF Ref_R"+str(refname.index(get_t_of_p(namespace, tc))+1)+" \h "
				
				r3 = etree.Element(namespace+"r")
				rpr = etree.SubElement(r3, namespace+"rPr")
				l3 = etree.SubElement(r3, namespace+"fldChar")
				l3.attrib[namespace+"fldCharType"]="separate"
				
				r4 = etree.Element(namespace+"r")
				rpr = etree.SubElement(r4, namespace+"rPr")
				l4 = etree.SubElement(r4, namespace+"fldChar")
				l4.attrib[namespace+"fldCharType"]="end"
				
				r[0].addprevious(r1)
				r[0].addprevious(r2)
				r[0].addprevious(r3)
				r[-1].addnext(r4)
						
def table_styler(root, namespace):
	# pagewidth = int(root.find('.//'+namespace+'pgSz').attrib.get(namespace+'w')) - int(root.find('.//'+namespace+'pgMar').attrib.get(namespace+'left')) - int(root.find('.//'+namespace+'pgMar').attrib.get(namespace+'right'))
	tags = '<table.*?>|</table>|<thead>|</thead>|<tbody>|</tbody>|<tfoot>|</tfoot>|<tr>|</tr>|<th>|</th>|<td>|</td>'
	table_builder = False
	table_inserter = False
	table_cell = ''
	table_content = []
	table_layer = 2
	table_row = 0
	p_kills = []
	extra_before = ''

	for p in root.findall('.//'+namespace+'p'): # Найти все тексты				
		tag = re.findall(tags, get_t_of_p(namespace, p))
		cuts = re.split(tags, get_t_of_p(namespace, p))
		if len(tag) > 0:
			p_kills.append(p)		
			for i in range(len(tag)):
				if "<table" in tag[i] and table_builder == False:
					table_builder = True
					extra_before = cuts[i]
				elif tag[i] == '<thead>' and table_builder == True:
					table_layer = 1
				elif tag[i] == '<tbody>' and table_builder == True:
					table_layer = 2
				elif tag[i] == '<tfoot>' and table_builder == True:
					table_layer = 3
				elif (tag[i] == '</thead>' or tag[i] == '</tbody>' or tag[i] == '</tfoot>') and table_builder == True:
					table_layer = 2
					
				elif tag[i] == '<tr>' and table_builder == True:
					table_content.append([table_layer, []])
				elif tag[i] == '</tr>' and table_builder == True:
					table_row += 1
					
				elif tag[i] == '<td>' and table_builder == True and table_inserter == False:
					table_inserter = True
				elif tag[i] == '<th>' and table_builder == True and table_inserter == False:
					table_inserter = True
				elif tag[i] == '</td>' and table_builder == True and table_inserter == True:
					table_inserter = False
					table_content[table_row][1].append(("td", table_cell))
					table_cell = ''
				elif tag[i] == '</th>' and table_builder == True and table_inserter == True:
					table_inserter = False
					table_content[table_row][1].append(("th", table_cell))
					table_cell = ''
					
				if table_builder == True and table_inserter == True:
					table_cell += cuts[i+1]
				
				if tag[i] == '</table>' and table_builder == True and table_inserter == False:
					if extra_before != '':
						p1 = etree.Element(namespace+'p')
						r = etree.SubElement(p1, namespace+'r')
						t = etree.SubElement(r, namespace+"t")
						t.text = extra_before
						p.addprevious(p1)
					table_content.sort(key = lambda x: x[0])
					tbl = etree.Element(namespace+"tbl")
					# tblPr = etree.SubElement(tbl, namespace+"tblPr")
					# tblW = etree.SubElement(tblPr, namespace+"tblW")
					# tblW.attrib[namespace+"w"]=str(pagewidth)
					# tblGrid = etree.SubElement(tbl, namespace+"tblGrid")
					for row in table_content:
						tr = etree.SubElement(tbl, namespace+"tr")
						for cell in row[1]:
							tc = etree.SubElement(tr, namespace+"tc")
							tcPr = etree.SubElement(tc, namespace+"tcPr")
							tcBorders = etree.SubElement(tcPr, namespace+"tcBorders")
							for d in ["top", "left", "bottom", "right"]:
								border = etree.SubElement(tcBorders, namespace+d)
								border.attrib[namespace+"val"]="single"
								border.attrib[namespace+"sz"]="2"
							pt = etree.SubElement(tc, namespace+"p")
							if cell[0] == "th":
								pPr = etree.SubElement(pt, namespace+"pPr")
								align = etree.SubElement(pPr, namespace+"jc")
								align.attrib[namespace+"val"] = "center"									
							rt = etree.SubElement(pt, namespace+"r")
							# if cell[0] == "th":
								# rPr = etree.SubElement(rt, namespace+"rPr")
								# b = etree.SubElement(rPr, namespace+"b")
							t = etree.SubElement(rt, namespace+"t")
							t.text = cell[1]
							if cell[0] == "th":
								t.text = "<b>"+t.text+"</b>"
					p.addprevious(tbl)
					if len(cuts) > i:
						if cuts[i+1] != '':
							p2 = etree.Element(namespace+'p')
							r = etree.SubElement(p2, namespace+'r')
							t = etree.SubElement(r, namespace+"t")
							t.text = cuts[i+1]
							p.addprevious(p2)
					table_builder = False
					table_row = 0
					table_content.clear()
	for p in p_kills:
		if p.getparent() != None:
			p.getparent().remove(p)
		else:
			root.remove(p)

def list_styler(root, namespace):
	tags = '<ol>|</ol>|<ul>|</ul>|<li>|</li>'
	p_kills = []
	list_builder = False
	list_inserter = False
	ordered_list = False
	list_items = []
	extra_before = ''
	li = 1

	for p in root.findall('.//'+namespace+'p'): # Найти все тексты				
		tag = re.findall(tags, get_t_of_p(namespace, p))
		cuts = re.split(tags, get_t_of_p(namespace, p))
		if len(tag) > 0:
			p_kills.append(p)			
			for i in range(len(tag)):
				if tag[i] == '<ol>' and list_builder == False:
					list_builder = True
					ordered_list = True
					li += 1
					extra_before = cuts[i]
				elif tag[i] == '<ul>' and list_builder == False:
					list_builder = True
					ordered_list = False
					li += 1
					extra_before = cuts[i]
				elif tag[i] == '<li>' and list_builder == True and list_inserter == False:
					list_inserter = True
				elif tag[i] == '</li>' and list_builder == True and list_inserter == True:
					list_inserter = False
					
				if list_builder == True and list_inserter == True:
					list_items.append(cuts[i+1])
					
				if (tag[i] == '</ol>' or tag[i] == '</ul>') and list_builder == True and list_inserter == False:
					if extra_before != '':
						p1 = etree.Element(namespace+'p')
						r = etree.SubElement(p1, namespace+'r')
						t = etree.SubElement(r, namespace+"t")
						t.text = extra_before
						p.addprevious(p1)
					for item in list_items:
						pl = etree.Element(namespace+'p')
						pPr = etree.SubElement(pl, namespace+"pPr")
						numPr = etree.SubElement(pPr, namespace+"numPr")
						rPr = etree.SubElement(pPr, namespace+"rPr")
						ilvl = etree.SubElement(numPr, namespace+"ilvl")
						ilvl.attrib[namespace+"val"]="0"
						numId = etree.SubElement(numPr, namespace+"numId")
						if ordered_list == False:
							numId.attrib[namespace+"val"]=str(li)
						else:
							numId.attrib[namespace+"val"]=str(li+50)
						r = etree.SubElement(pl, namespace+'r')
						t = etree.SubElement(r, namespace+"t")
						t.text = item
						p.addprevious(pl)
					if len(cuts) > i:
						if cuts[i+1] != '':
							p2 = etree.Element(namespace+'p')
							r = etree.SubElement(p2, namespace+'r')
							t = etree.SubElement(r, namespace+"t")
							t.text = cuts[i+1]
							p.addprevious(p2)
					list_builder = False
					ordered_list = False
					list_items.clear()	
	for p in p_kills:
		if p.getparent() != None:
			p.getparent().remove(p)
		else:
			root.remove(p)
	
def image_styler(root, namespace, tmp_dir, files):
	tags = '<img.*?>'
	for p in root.findall('.//'+namespace+'p'): # Найти все тексты				
		tag = re.findall(tags, get_t_of_p(namespace, p))
		if len(tag) > 0:
			t = p.find('.//'+namespace+'t')
			t.text = t.text.replace(tag[0], '')
			link = tag[0].split('src="')[1].split('"')[0]
			err = False
			media = tmp_dir+'/word/media/'
			if not os.path.exists(media):
				os.mkdir(media)
			if "http" in link:
				try:
					request = requests.get(link, timeout=2)
				except:
					print("Can't get image by URL", link)
					err = True
				else:
					link = link.split('/')[-1]
					with open(link, 'wb') as f:
						f.write(request.content)
					shutil.move(link, media)
			else:
				try:
					shutil.copy(link, media)  
				except e:
					print("Image saving error", e)
					err = True
			if err == True:
				tag, width, height = (link, 300, 50)
			else:
				files.append('word/media/'+link)
				image = Image.open(media+link)
				width, height = (image.size[0], image.size[1])
				image.close()
				tag = "rId2"
				if link.split(".")[1] == "jpg":
					type = "image/jpeg"
				else:
					type = "image/png"
				with open(os.path.join(tmp_dir,"word/_rels/document.xml.rels"), "rb+") as f:
					tree = etree.parse(f)
					root = tree.getroot()
					tag = "rId"+str(len(root.getchildren())+1)
					root.append(etree.Element("Relationship", Id=tag, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", Target="media/"+link))
					f.seek(0)
					tree.write(f, encoding='utf-8', xml_declaration=False)
				with open(os.path.join(tmp_dir,"[Content_Types].xml"), "rb+") as f:
					tree = etree.parse(f)
					root = tree.getroot()
					root.append(etree.Element("Override", PartName="/word/media/"+link, ContentType=type))
					f.seek(0)
					tree.write(f, encoding='utf-8', xml_declaration=False)
			etree.register_namespace('a',"http://schemas.openxmlformats.org/drawingml/2006/main")
			etree.register_namespace('pic',"http://schemas.openxmlformats.org/drawingml/2006/picture")
			wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
			nsa = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
			nspic = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
			rns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
			pi = etree.Element(namespace+"p")
			p.addnext(pi)
			r = etree.SubElement(pi, namespace+"r")
			rPr = etree.SubElement(r, namespace+"rPr")
			drawing = etree.SubElement(r, namespace+"drawing")
			inline = etree.SubElement(drawing, wp+"inline", distT="0", distB="0", distL="0", distR="0") #
			extent = etree.SubElement(inline,  wp+"extent", cx=str(width*10000), cy=str(height*10000))
			# effectExtent = etree.SubElement(inline, wp+"effectExtent", l="0", t="0", r="0", b="0") #
			docPr = etree.SubElement(inline, wp+"docPr", id="3", name=link, descr="") #
			# cNvGraphicFramePr = etree.SubElement(inline, wp+"cNvGraphicFramePr") #
			# graphicFrameLocks = etree.SubElement(cNvGraphicFramePr, nsa+"graphicFrameLocks") #
			# graphicFrameLocks.attrib["noChangeAspect"]="1" #
			graphic = etree.SubElement(inline, nsa+"graphic")
			graphicData = etree.SubElement(graphic, nsa+"graphicData")
			graphicData.attrib["uri"]="http://schemas.openxmlformats.org/drawingml/2006/picture"
			pic = etree.SubElement(graphicData, nspic+"pic")
			nvPicPr = etree.SubElement(pic, nspic+"nvPicPr") #
			cNvPr = etree.SubElement(nvPicPr, nspic+"cNvPr", id="3", name=link, descr="") #
			cNvPicPr = etree.SubElement(nvPicPr, nspic+"cNvPicPr") #
			# picLocks = etree.SubElement(cNvPicPr, nsa+"picLocks", noChangeAspect="1") #
			blipFill = etree.SubElement(pic, nspic+"blipFill")
			blip = etree.SubElement(blipFill, nsa+"blip")
			blip.attrib[rns+"embed"]=tag
			stretch = etree.SubElement(blipFill, nsa+"stretch") #
			# fillRect = etree.SubElement(stretch, nsa+"fillRect") #
			spPr = etree.SubElement(pic, nspic+"spPr", bwMode="auto") #
			xfrm = etree.SubElement(spPr, nsa+"xfrm") #
			# off = etree.SubElement(xfrm, nsa+"off", x="0", y="0") #
			ext = etree.SubElement(xfrm, nsa+"ext", cx=str(width*10000), cy=str(height*10000)) #
			prstGeom = etree.SubElement(spPr, nsa+"prstGeom", prst="rect") #
			# avLst = etree.SubElement(prstGeom, nsa+"avLst") #
			
def parag_styler(root, namespace):
	tags = '<p.*?>|</p>|<div.*?>|</div>|<font.*?>|</font>|<pre>|</pre>'
	p_kills = []
	change_font = False
	face = ''
	size = ''
	color = ''
	for p in root.findall('.//'+namespace+'p'): # Найти все тексты	
		tag = ['']
		tag.extend(re.findall(tags, get_t_of_p(namespace, p)))
		cuts = re.split(tags, get_t_of_p(namespace, p))	
		if len(tag) > 1:
			p_kills.append(p)	
			for i in range(len(tag)):
				if cuts[i] != '':
					pp = etree.Element(namespace+'p')	
					p.addnext(pp)
					p = pp	
				if tag[i] == '<p>' or '<p ' in tag[i]: #p
					if "align=" in tag[i]:
						val = tag[i].split('align="')[1].split('"')[0]
						alignment_flag = True
						alignment = val
						pPr = etree.SubElement(p, namespace+"pPr")
						align = etree.SubElement(pPr, namespace+"jc")
						align.attrib[namespace+"val"] = alignment
				elif tag[i] == '<div>' or '<div ' in tag[i]: #div
					if "style=" in tag[i]:
						if "text-align:" in tag[i]:
							val = tag[i].split('text-align:')[1].split(';')[0].replace(" ", "")
							alignment_flag = True
							alignment = val
							pPr = etree.SubElement(p, namespace+"pPr")
							align = etree.SubElement(pPr, namespace+"jc")
							align.attrib[namespace+"val"] = alignment
					if "align=" in tag[i]:
						val = tag[i].split('align="')[1].split('"')[0]
						alignment_flag = True
						alignment = val
						pPr = etree.SubElement(p, namespace+"pPr")
						align = etree.SubElement(pPr, namespace+"jc")
						align.attrib[namespace+"val"] = alignment
				elif tag[i] == '<pre>': #pre
					pass
				elif tag[i] == '<font>' or '<font ' in tag[i]: #font
					face = ''
					size = ''
					color = ''
					change_font = True
					if "face=" in tag[i]:
						face = tag[i].split('face="')[1].split('"')[0]
					if "size=" in tag[i]:
						size = tag[i].split('size="')[1].split('"')[0]
					if "color=" in tag[i]:
						color = tag[i].split('color="')[1].split('"')[0]
				elif tag[i] == '</div>' or tag[i] == '</pre>' or tag[i] == '</font>':
					alignment_flag = False
					alignment = "left"
					change_font = False
				if cuts[i] != '':
					r = etree.SubElement(p, namespace+'r')						
					rPr = etree.SubElement(r, namespace+"rPr")
					if change_font == True:
						if face != '':				
							rFonts = etree.SubElement(rPr, namespace+"rFonts")
							rFonts.attrib[namespace+"cs"] = face
						if size != '':
							sz = etree.SubElement(rPr, namespace+"sz")
							sz.attrib[namespace+"val"] = str(24*int(size)/3)
						if color != '':
							col = etree.SubElement(rPr, namespace+"color")
							col.attrib[namespace+"val"] = color
					t = etree.SubElement(r, namespace+"t")
					t.text = cuts[i]
	for p in p_kills:
		if p.getparent() != None:
			p.getparent().remove(p)
		else:
			root.remove(p)

def other_styler(root, namespace):
	xmlns = "{http://www.w3.org/XML/1998/namespace}"
	b_mode = False
	i_mode = False
	u_mode = False
	br_mode = False
	tags = '<b>|</b>|<i>|</i>|<u>|</u>|<br>'
	for p in root.findall('.//'+namespace+'p'): # Найти все тексты		
		for t in p.findall('.//'+namespace+'t'):
			if t.text != None:
				if "&nbsp;" in t.text:  # Замена пробела
					t.text = t.text.replace("&nbsp;", "\xA0")		
		tag = ['']
		tag.extend(re.findall(tags, get_t_of_p(namespace, p)))
		cuts = re.split(tags, get_t_of_p(namespace, p))
		if len(tag) > 1:
			for r in p.findall('.//'+namespace+'r'):
				p.remove(r)	
			for i in range(len(tag)):
				if tag[i] == '<b>': #b
					b_mode = True
				elif tag[i] == '</b>':
					b_mode = False
				elif tag[i] == '<i>': #i
					i_mode = True
				elif tag[i] == '</i>':
					i_mode = False
				elif tag[i] == '<u>': #u
					u_mode = True
				elif tag[i] == '</u>':
					u_mode = False
				elif tag[i] == '<br>': #br
					br_mode = True
						
				if cuts[i] != '':
					r = etree.SubElement(p, namespace+'r')						
					rpr = etree.SubElement(r, namespace+"rPr")
					if b_mode == True:
						b = etree.SubElement(rpr, namespace+"b")
					if i_mode == True:
						b = etree.SubElement(rpr, namespace+"i")
					if u_mode == True:	
						b = etree.SubElement(rpr, namespace+"u")
						b.attrib[namespace+"val"]="single"
					if br_mode == True:	
						b = etree.SubElement(r, namespace+"br")
						br_mode = False
					t = etree.SubElement(r, namespace+"t")
					t.text = cuts[i]
					t.attrib[xmlns+"space"]="preserve"
					# pretty_print(pp)
			# p.addprevious(pp)

if __name__ == "__main__":
	file = "Test.docx"
	with ZipFile(file, 'a') as zip:
		tmp_dir = tempfile.mkdtemp()
		zip.extractall(tmp_dir)
		files = zip.namelist()
		fix_listing(tmp_dir)
		print(tmp_dir)
		with open(os.path.join(tmp_dir,"word/document.xml"), "rb+") as myfile:
			tree = etree.parse(myfile)
			root = tree.getroot()
			namespace=root.tag.split('document')[0]
			get_schemes(root, namespace) #приоритет 1
			table_styler(root, namespace) #приоритет 2
			list_styler(root, namespace) #приоритет 3
			image_styler(root, namespace, tmp_dir, files) #приоритет 4
			parag_styler(root, namespace) #приоритет 5
			other_styler(root, namespace) #приоритет 6
			myfile.seek(0)
			myfile.truncate(0)
			tree.write(myfile, encoding='utf-8', xml_declaration=True, standalone=True)
			tree.write("document.xml", encoding='utf-8', xml_declaration=True, standalone=True) #debug
		with ZipFile("Test_Marked.docx", "w") as docx:
		# with ZipFile(file, "w") as docx:
			for filename in files:
				docx.write(os.path.join(tmp_dir, filename), filename)
	shutil.rmtree(tmp_dir)