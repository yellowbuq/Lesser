#!/usr/bin/env python3

version = '3.2'

from PIL import Image#https://pypi.org/project/image/
import os


#docx2txt Source: https://pypi.org/project/docx2txt/
import re
import xml.etree.ElementTree as ET
import zipfile
import os

nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


def xml2text(xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    root = ET.fromstring(xml)
    for child in root.iter():
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
    return text


def process(docx, img_dir=None):
    text = u''

    # unzip the docx in memory
    zipf = zipfile.ZipFile(docx)
    filelist = zipf.namelist()

    # get header text
    # there can be 3 header files in the zip
    header_xmls = 'word/header[0-9]*.xml'
    for fname in filelist:
        if re.match(header_xmls, fname):
            text += xml2text(zipf.read(fname))

    # get main text
    doc_xml = 'word/document.xml'
    text += xml2text(zipf.read(doc_xml))

    # get footer text
    # there can be 3 footer files in the zip
    footer_xmls = 'word/footer[0-9]*.xml'
    for fname in filelist:
        if re.match(footer_xmls, fname):
            text += xml2text(zipf.read(fname))

    if img_dir is not None:
        # extract images
        for fname in filelist:
            _, extension = os.path.splitext(fname)
            if extension in [".jpg", ".jpeg", ".png", ".bmp"]:
                dst_fname = os.path.join(img_dir, os.path.basename(fname))
                with open(dst_fname, "wb") as dst_f:
                    dst_f.write(zipf.read(fname))

    zipf.close()
    return text.strip()



def felieton(filenamedocx):

	tytu?? = ''
	tekst = []#tytu?? i tekst
	Linki = []#liniki do grafik
	Podpis = []#podpis i komentarz
	autor = 'Czytelnik SGI Lesser'#Autor domy??lny gdy nie ma podpisu
	grafika = 0
	zawarto???? = []
	zako??czenia = ['KONIEC', 'CDN.']

	try:

		for i,line in enumerate(process(filenamedocx,'r').strip().split('\n')):
			if i == 0:
				tytu?? += line+'\n'#Na pocz??tku wydzielam tytu??
				continue
			if not line:
				continue
			if grafika and len(line) > 8 and line != 'Czytelnik SGI Lesser':#Tworzenie link??w gdy line nie jest podpisem
				if line.find('https://') >= 0:#https
					index = line[line.find('https://')+8:].find('/')
					line = line[line.find('https://'):]
					Linki.append(f'<li><a href="https://{line[8:]}">{line[8:index+8]}</a></li>\n')
					continue
				elif line.find('http://') >= 0:#http
					index = line[line.find('http://')+7:].find('/')
					line = line[line.find('http://'):]
					Linki.append(f'<li><a href="http://{line[7:]}">{line[7:index+7]}</a></li>\n')
					continue
				Linki.append('</ul>\n')#Koniec listy
				Podpis.append(autor+'\n')#Podpis
				if line.find('Komentarz:') != -1:
					Podpis.append(line[line.find('Komentarz:')+10:])#Komentarz bez "Komenatrz:"
				break
			if line.strip() == 'Grafika:':#Start grafika
				grafika = 1
				Linki.append('<p>Grafika:</p>\n<ul>\n')
			else:
				if grafika and len(line) > 3:
					autor = line#Podpis autora
					continue

				if line in zako??czenia:#Specialne zako??czenia
					tekst.append(f'<p class="ct">{line}</p>\n')
				#Linki w tek??cie
				if line.find('https://') == 0:
					index = line[line.find('https://')+8:].find('/')
					tekst.append(f'<a href="https://{line[8:]}">{line[8:index+8]}</a>\n')
					continue
				if line.find('http://') == 0:
					index = line[line.find('https://')+7:].find('/')
					tekst.append(f'<a href="http://{line[7:]}">{line[7:index+7]}</a>\n')
					continue

				if len(line) > 8 and len(line) < 20 and not grafika:
					tekst.append(f'<b>{line}</b>\n')#Podpis w Tekstach Czytelnika i podtytu??y

				else:#Tre???? artyku??u
					tekst.append(f'<p>{line}</p>\n')


		#Rozmieszczenie grafik w tekscie
		################################

		#Indexy akapit??w
		indeksy_akapit??w = []#Na jakim indeksie zaczynaj?? si?? akapity
		d??ugo???? = 0#D??ugo??????tekstu
		licznik_akapit??w = 0#Ile jest akapit??w
		kt??ra_grafika = 0#Przed kt??r?? grafik?? jest okre??lona liczba akapit??w
		for x in range(len(tekst)):
			d??ugo???? += len(tekst[x])#D??ugo???? tekstu w akapicie "x"
			indeksy_akapit??w.append(d??ugo????)#Dodanie do listy indeks??w

		#Odleg??o???? od siebie jako ??rednia
		Tekst = ''#Ko??cowy tekst
		index_grafik = []#Lista przechowuj??ca indexy gdzie powinny znajdowa?? sie grafiki
		tekst = ''.join(tekst)#Z????czenie listy tre??ci artyku??u na string
		d??ugo???? = len(tekst)#D??ugo???? artyku??u
		co_ile = d??ugo????//(ile_jest_grafik)#Co ile grafika

		[index_grafik.append(co_ile*x) for x in range(1, ile_jest_grafik)]#Tworzenie listy kt??ra przechowuje kolejne indeksy grafik

		#Uwzgl??dnienie wysoko??ci grafik
		for i in range( len(wysoko????_obraz??w)-1):
			if wysoko????_obraz??w[i] > 300:#Przesuni??cie obrazu o "przesuni??cie" gdy wysoko???? jest wi??ksza od 300px
				przesuni??cie = round((abs(300 - wysoko????_obraz??w[i]) / 21)*86)
				if wysoko????_obraz??w[i] > 400: index_grafik[0] -= przesuni??cie#Gdy obraz jest wy??szy ni?? 400px to przesu?? te?? pierwszy obraz
				index_grafik[i+1] += przesuni??cie

		#Uwzgl??dnienie odst??p??w akapit??w
		for x in range(len(tekst)):
			if x in indeksy_akapit??w:
				licznik_akapit??w += 1
			if x in index_grafik:
				index_grafik[kt??ra_grafika] -= 86*licznik_akapit??w#86 znak??w ma najd??u??szy wiersz
				licznik_akapit??w = 0
				kt??ra_grafika += 1


		#Specjalnie znaki
		#################

		for i,litera in enumerate(tekst):#index, litera
			if i in index_grafik: Tekst += '<<<GRAFIKA>>>' #Je??eli index litery jest taki sam jak index grafiki
			if chr(i) == 8222: Tekst += i+'<i>' # ???
			if chr(i) == 8221: Tekst += '</i>'+i # ???
			else: Tekst += litera

		zawarto????.append(tytu??)
		zawarto????.append(''.join(Tekst))
		zawarto????.append(''.join(Linki))
		zawarto????.append(''.join(Podpis))
		open(filenamedocx[:-5]+'.txt','w').write(''.join(zawarto????))
		print('?? Plik',filenamedocx[:-4]+'txt','poprawnie stworzony')

	except TabError:
		print('?? B????d pliku',filenamedocx,'- mo??liwe z??e rozszerzenie')
		print(TabError)
		return 0


def poezja(filenamedocx):

	Tytu?? = ''
	Tekst = []#tytu?? i tekst
	Linki = []#liniki do grafik
	Podpis = []#podpis i komentarz
	grafika = 0
	enter = 0
	po = False#Gdy jest "po" akapicie wers??w
	zawarto???? = []

	try:

		for i,line in enumerate(process(filenamedocx,'r').strip().split('\n')):#args.filenamedocx

			if i == 0:#Wykonuje sie tylko raz
				Tytu?? += line+'\n'
				Tekst.append(f'<br /><br /><p class="ct"><strong>{line}</strong></p><br />\n\n')#Tytu?? wiersza
				po = True
				continue
			if not line:#Gdy jest enter
				enter += 1
				continue
			if grafika and len(line) > 8 and line != 'Czytelnik SGI Lesser':#Tworzenie link??w gdy line nie jest podpisem
				if line.find('https://') >= 0:#https
					index = line[line.find('https://')+8:].find('/')
					line = line[line.find('https://'):]
					Linki.append(f'<li><a href="https://{line[8:]}">{line[8:index+8]}</a></li>\n')
					continue
				elif line.find('http://') >= 0:#http
					index = line[line.find('http://')+7:].find('/')
					line = line[line.find('http://'):]
					Linki.append(f'<li><a href="http://{line[7:]}">{line[7:index+7]}</a></li>\n')
					continue
				Linki.append('</ul>\n')#Koniec listy
				Podpis.append(autor+'\n')#Podpis
				if line.find('Komentarz:') != -1:
					Podpis.append(line[line.find('Komentarz:')+10:])#Komentarz bez "Komenatrz:"
				break
			if line.strip() == 'Grafika:':#Grafika start
				grafika = 1
				Linki.append('</p>\n\n<p>Grafika:</p>\n<ul>\n')
			else:
				if grafika:#Podpis
					try:
						autor = line#Podpis
					except:
						autor = 'Czytelnik SGI Lesser\n'
				else:
					if po:#Gdy jest po tytule wiersza
						Tekst.append(f'<p class="ct">{line}')#Akapit wers??w
						enter = 0
						po = False#Akapit wers??w stop
						continue
					if enter > 3:#Gdy s?? 4 entery to jest odst??p mi??dzy wierszami(dla bezpiecze??twa wi??ksze od 3)
						Tekst.append(f'</p>\n\n<br /><br /><p class="ct"><strong>{line}</strong></p><br />\n\n')#Tytu?? wiersza
						enter = 0
						po = True#Akapit wers??w start
						continue

					else:
						if enter == 3: Tekst.append('<br />\n')#Gdy wyliczono 3 entery to jest odst??p
						Tekst.append(f'<br />\n{line}')#Wersy
						enter = 0#Brak enter??w


		zawarto????.append(Tytu??)
		zawarto????.append(''.join(Tekst))
		zawarto????.append(''.join(Linki))
		zawarto????.append(''.join(Podpis))
		open(filenamedocx[:-5]+'.txt','w').write(''.join(zawarto????))
		print('?? Plik',filenamedocx[:-4]+'txt','poprawnie stworzony')

	except TabError:
		print('?? B????d pliku',filenamedocx,'- mo??liwe z??e rozszerzenie')#Spotkany b????d z rozszerzeniem ".rtf"
		print(TabError)
		return 0


def zdj??cie(filenameimg):#Tworzenie obrazu z odpowiednimi wymiarami

	img = Image.open(filenameimg)#Objekt obrazu
	szeroko????,wysoko???? = img.size[0],img.size[1]#szeroko????, wysoko????

	if img.size[0] == 300 or img.size[0] == 280:#Gdy obraz ma odpowiednie wymiary
		print('\t?? Plik',filenameimg+' ma odpowiednie wymiary')
		wysoko????_obraz??w[int(filenameimg[-5])-1] = img.size[1]
		return 0

	if szeroko???? < wysoko????:#Gdy wysoko???? jest wi??ksza od szeroko??ci
		nowa_szeroko???? = 280
	if szeroko???? > wysoko????:#Gdy wysoko???? jest wi??ksza od szeroko??ci
		nowa_szeroko???? = 300
	if szeroko???? == wysoko????:#Gdy wysoko???? i szeroko???? s?? r??wne
		nowa_szeroko???? = 300

	nowa_wysoko???? = int(nowa_szeroko???? * wysoko???? / szeroko????)#Nowa wysoko???? z proporcji

	nowy_img = img.resize((nowa_szeroko????, nowa_wysoko????), Image.ANTIALIAS)#Nowy obraz
	wysoko????_obraz??w[int(filenameimg[-5])-1] = nowa_wysoko????#Dodanie wysoko??ci

	try:
		if filenameimg[-3:] == 'jpg' or filenameimg[-3:] == 'jpeg':#Gdy stary obraz jest rozszerzenia jpg, jpeg
			nowy_img.save(filenameimg)
			print('?? Plik',filenameimg+' poprawnie stworzony')
			return
		else:
			try:
				nowy_img.save(filenameimg[:-3]+'jpg')#Pr??ba zapisania jako jpg
			except:
				nowy_img.save(filenameimg)#Nadpisanie istniej??cego
				print('?? Plik',filenameimg[:-3]+filenameimg[-3:]+' poprawnie stworzony')
				return
			print('?? Plik',filenameimg[:-3]+'jpg'+' poprawnie stworzony')
			return

	except:
		print('?? B????d pliku',filenameimg)
		return



print('\n  |                                        ')
print('  |        _ \    __|    __|    _ \    __| ')
print('  |        __/  \__ \  \__ \    __/   |    ')
print(' _____|  \___|  ____/  ____/  \___|  _|    ')
print(f'                                         -{version}-\n')

jestdocx = False#Czy wykryto plik docx
ile_jest_grafik = 0#Liczba grafik
wysoko????_obraz??w = {}#S??ownik {"numer_obrazu":wysoko????}
nazwy_plik??w = []

[nazwy_plik??w.append(nazwa_pliku) for nazwa_pliku in os.listdir()]#Nazwy plik??w zawartych w folderze

for filename in nazwy_plik??w:#Przeszukiwanie nazw plik??w obraz??w i docx

	if (filename[-3:] == 'jpg' or filename[-3:] == 'png' or filename[-4:] == 'jpeg'):#Wykryto obraz
		ile_jest_grafik += 1

		if filename[-4:] == 'jpeg':
			newfilename = filename.replace('.jpeg', '.jpg')
			os.rename(os.path.join(filename), os.path.join(newfilename))#Zmiana rozszerzenia zdj??cia z jpeg na jpg
			filename = newfilename#Nowa nazwa pliku

		if filename[-5].isdigit(): zdj??cie(filename)#Zmiana obrazu gdy nie jest g????wnym

	if filename[-4:] == 'docx':
		filenamedocx = filename
		jestdocx = True

if ile_jest_grafik > 1 and jestdocx:
	felieton(filenamedocx)
if ile_jest_grafik == 1 and jestdocx:#Poezje maj?? tylko jedno zdj??cie
	poezja(filenamedocx)

if not jestdocx:
	print('?? Nie znaleziono pliku .docx\n')
if ile_jest_grafik == 0:
	print('?? Nie znaleziono grafik\n')

exit()