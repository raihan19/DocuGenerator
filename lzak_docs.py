import os
import sys
import docx
import functions as f

weaNo = 1  # WTG number

if len(sys.argv) == 3:
    folderIn = str(sys.argv[1])
    folderOut = str(sys.argv[2])
else:
    folderIn = os.getcwd()
    folderOut = str("{}/documents/".format(os.getcwd()))

data = f.get_data_from_json(folderIn)
atlasList = f.atlas_list(data)
English = True

# loop through all site
for refName in atlasList:
    print("Working on {}".format(refName))
    bookmarks = []  # array to collect all the informations for the bookmarks

    # bookmark 1 [Bookmark name: titel_atlasname]
    bookmarks.append({
        "bookmark": "titel_atlasname",
        "type": "string",
        "value": str(refName),
        "bold": True,
        "italic": True
        })

    # bookmark 2 [Bookmark name: ort]
    valList = f.wbReader(folderIn, data, refName, weaNo)
    bookmarks.append({
        "bookmark": "ort",
        "type": "string",
        "value": valList[0],
        "bold": True,
        "italic": True
        })

    # bookmark 3 [Bookmark name: wea]
    '''Read from json file'''
    '''We are subtracting 1 because index starts from 0'''
    bookmarks.append({
        "bookmark": "wea",
        "type": "string",
        "value": str(data['WEAs'][weaNo-1]['Anlage']),
        "bold": True,
        "italic": True
        })

    # bookmark 4 [Bookmark name: anlage]
    '''Read from json file'''
    bookmarks.append({
        "bookmark": "anlage",
        "type": "string",
        "value": str(data['WEAs'][weaNo-1]['Typ']),
        "bold": True,
        "italic": True
        })

    # bookmark 5 [Bookmark name: nh]
    '''Read from json file'''
    bookmarks.append({
        "bookmark": "nh",
        "type": "string",
        "value": str(data['WEAs'][weaNo-1]['NH']) + 'm',
        "bold": True,
        "italic": True
        })

    # bookmark 6 and 7 [Bookmark name: evon & ebis]
    d_string = f.get_date(folderIn, refName, data, English, weaNo)
    bookmarks.append({
        "bookmark": "evon",
        "type": "string",
        "value": d_string[0]
    })

    bookmarks.append({
        "bookmark": "ebis",
        "type": "string",
        "value": d_string[1]
    })

    # bookmark 8 [Source: a text, Bookmark name: zeitschritt]
    dmTrack = ''
    if English == True:
        if data['WEAs'][weaNo - 1]['timestep'] == 'monthly':
            dmTrack = 'monthly values'
        elif data['WEAs'][weaNo - 1]['timestep'] == 'daily':
            dmTrack = 'daily values'
        else:
            dmTrack = ''
    if English == False:
        if data['WEAs'][weaNo - 1]['timestep'] == 'daily':
            dmTrack = 'Tageswerte'
        elif data['WEAs'][weaNo - 1]['timestep'] == 'monthly':
            dmTrack = 'Monatswerte'
        else:
            dmTrack = ''

    bookmarks.append({
        "bookmark": "zeitschritt",
        "type": "string",
        "value": dmTrack
    })

    # bookmark 9 [Source: wea1.csv, Bookmark name: fehlwerte]
    bookmarks.append({
        "bookmark": "fehlwerte",
        "type": "string",
        "value": str(f.count_missing_values(folderIn, refName, data, weaNo))
    })

    # bookmark 10 [Source: a text, Bookmark name: verfuegbarkeitsangaben] ***modify it later***
    bookmarks.append({
        "bookmark": "verfuegbarkeitsangaben",
        "type": "string",
        "value": d_string[2]
    })

    # bookmark 11 [Source: Config.json, Bookmark name: leistungskennlinie]
    if English == True:
        powerCurveVar = 'Power curve: ' + str(data['WEAs'][weaNo-1]['power_curve'][:-4])
    if English == False:
        powerCurveVar = "Leistungskurve: " + str(data['WEAs'][weaNo-1]['power_curve'][:-4])
    bookmarks.append({
        "bookmark": "leistungskennlinie",
        "type": "string",
        "value": powerCurveVar
    })

    # bookmark 12 [Source: Config.json, Bookmark name: auswahlzelle_titel]
    if English == True:
        selectionVar = 'Selection of the ' + str(refName) + '-Index-cell'
    if English == False:
        selectionVar = 'Auswahl der ' + str(refName) + '-Index-Zelle'

    bookmarks.append({
        "bookmark": "auswahlzelle_titel",
        "type": "string",
        "value": selectionVar,
        "underline": True
    })

    # bookmark 13 [Source: wea1-cor.csv, Bookmark name: auswahlzelle]
    sentence1 = f.get_coordinates(folderIn, refName, weaNo, data, English)
    bookmarks.append({
        "bookmark": "auswahlzelle",
        "type": "string",
        "value": sentence1
    })

    # bookmark 14 and 15 [Source: regression_eng_1.png & ertrag_messzeitraum_eng_1.png, Bookmark name: regression & messertrag]
    bookmarks.append({
        "bookmark": "regression",
        "type": "picture",
        "value": "{}/{}/regression_{}.png".format(folderIn, refName, weaNo)
    })

    bookmarks.append({
        "bookmark": "messertrag",
        "type": "picture",
        "value": "{}/{}/ertrag_messzeitraum_{}.png".format(folderIn, refName, weaNo)
    })

    # bookmark 16 and 17 [Source: Config.json, Bookmark name: atlas3 & atlas4]
    bookmarks.append({
        "bookmark": "atlas3",
        "type": "string",
        "value": str(refName),
        "italic": True
    })
    bookmarks.append({
        "bookmark": "atlas4",
        "type": "string",
        "value": str(refName),
        "italic": True
    })

    # bookmark 18 and 19 [Source: Config.json, Bookmark name: referenzzeitraum4 & referenzzeitraum3]
    bookmarks.append({
        "bookmark": "referenzzeitraum4",
        "type": "string",
        "value": str(f.refprd_list(data, refName)),
        "underline": True
    })
    bookmarks.append({
        "bookmark": "referenzzeitraum3",
        "type": "string",
        "value": " " + str(f.refprd_list(data, refName)),
    })

    # bookmark 20 [Source: Config.json, Bookmark name: atlas5]
    bookmarks.append({
        "bookmark": "atlas5",
        "type": "string",
        "value": str(refName)
    })

    # bookmark 21 [Source: langzeitertrag_eng_1.png, Bookmark name: langzeitetrag]
    bookmarks.append({
        "bookmark": "langzeitetrag",
        "type": "picture",
        "value": "{}/{}/langzeitertrag_{}.png".format(folderIn, refName, weaNo)
    })

    # bookmark 22 [Source: config json, Bookmark name: referenzzeitraum5]
    bookmarks.append({
        "bookmark": "referenzzeitraum5",
        "type": "string",
        "value": str(f.refprd_list(data, refName)),
        "italic": True
    })

    # bookmark 23 [Source: Config.json, Bookmark name: atlas6]
    bookmarks.append({
        "bookmark": "atlas6",
        "type": "string",
        "value": str(refName),
        "italic": True
    })

    # bookmark 24 [Source: wea1.xls, Bookmark name: ertragsindex]
    bookmarks.append({
        "bookmark": "ertragsindex",
        "type": "string",
        "value": valList[1]
    })

    # bookmark 25 [Source: weas_D-3km.E5_k2.txt, Bookmark name: jahresenergieertrag]
    bookmarks.append({
        "bookmark": "jahresenergieertrag",
        "type": "string",
        "value": valList[2]
    })

    template_file = str('{}/templates/template_german.docx'.format(os.path.dirname(os.path.abspath(__file__))))

    if English == True:
        template_file = str('{}/templates/template_english.docx'.format(os.path.dirname(os.path.abspath(__file__))))

    doc = docx.Document(template_file)

    f.write_bookmarks(doc, bookmarks)

    try:
        if not os.path.isdir(folderOut):
            os.makedirs("{}".format(folderOut))
        filename = ''
        if English == True:
            filename = str("{}/{}_WEA{}_eng.docx".format(folderOut, f.atlasName(data)[refName], weaNo))
        if English == False:
            filename = str("{}/{}_WEA{}.docx".format(folderOut, f.atlasName(data)[refName], weaNo))
        print('Creating file: {}'.format(filename))
        doc.save(filename)

    except OSError as e:
        print(e)

