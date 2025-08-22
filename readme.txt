--Features:
1. Extracts Texts and Images from pptx/pdf/xlsx or image files
2. create a new copy of the file and import the extracted content (After you modify them)
3. Translates these files with LLM(beta)

--Usage:
1. install python and run setup.bat(for pip install dependencies from requirements.txt)
2. launch MoonTranslator from AppRoot

--To enable cuda for faster OCR: 
pip uninstall easyocr, torchvision
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu126
pip install easyocr

--Warning!!!: 
1. Using less capable models may results in more API calls
2. traslating from or to an unsupported language will cause infinite calls
3. configure "LLM_maxTries" to protect your wallet 

--available options for OCR_Language
Abaza	abq
Adyghe	ady
Afrikaans	af
Angika	ang
Arabic	ar
Assamese	as
Avar	ava
Azerbaijani	az
Belarusian	be
Bulgarian	bg
Bihari	bh
Bhojpuri	bho
Bengali	bn
Bosnian	bs
Simplified Chinese	ch_sim
Traditional Chinese	ch_tra
Chechen	che
Czech	cs
Welsh	cy
Danish	da
Dargwa	dar
German	de
English	en
Spanish	es
Estonian	et
Persian (Farsi)	fa
French	fr
Irish	ga
Goan Konkani	gom
Hindi	hi
Croatian	hr
Hungarian	hu
Indonesian	id
Ingush	inh
Icelandic	is
Italian	it
Japanese	ja
Kabardian	kbd
Kannada	kn
Korean	ko
Kurdish	ku
Latin	la
Lak	lbe
Lezghian	lez
Lithuanian	lt
Latvian	lv
Magahi	mah
Maithili	mai
Maori	mi
Mongolian	mn
Marathi	mr
Malay	ms
Maltese	mt
Nepali	ne
Newari	new
Dutch	nl
Norwegian	no
Occitan	oc
Pali	pi
Polish	pl
Portuguese	pt
Romanian	ro
Russian	ru
Serbian (cyrillic)	rs_cyrillic
Serbian (latin)	rs_latin
Nagpuri	sck
Slovak	sk
Slovenian	sl
Albanian	sq
Swedish	sv
Swahili	sw
Tamil	ta
Tabassaran	tab
Telugu	te
Thai	th
Tajik	tjk
Tagalog	tl
Turkish	tr
Uyghur	ug
Ukranian	uk
Urdu	ur
Uzbek	uz
Vietnamese	vi