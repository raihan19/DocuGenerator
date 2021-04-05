import json
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import pandas as pd

def get_data_from_json(folder): # load the config.json file
    with open('{}/config.json'.format(folder)) as filePath:
        return json.load(filePath)

# defined function to get all site name
def atlas_list(data):
    allAtlas = []
    for atlas in data['Referenzen']:
        if (atlas['type'] == 'Atlas') or (atlas['type'] == 'file'):
            allAtlas.append(atlas['name'])
    sortedAllatlas = sorted(allAtlas)
    return sortedAllatlas

# format the name of the sites
#defined function to format the name of the atlas
def atlasName(data):
    newS = {}
    for s in atlas_list(data):
        tempS = []
        for char in s:
            if char == '-':
                continue
            if char == '.':
                tempS.append('_')
                continue
            tempS.append(char)
        newS[s] = ''.join(tempS)
    return newS

def get_ref_data(all_refs,refname):
    for ref in all_refs:
        if ref["name"] == refname:
            return ref

def get_first_common_date(all_refs):
    first_dates = {}
    for ref in all_refs:
        first_date = -1
        for dateV in sorted(ref["data"].keys()):
            if (first_date > dateV) or (first_date == -1):
                first_date = dateV
        first_dates[ref["name"]] = first_date
    common_first_date = -1
    for date in first_dates.values():
        if date > common_first_date:
            common_first_date = date
    return common_first_date

def get_ref_data_first_date(all_refs, refName, first_date):
    results = {}
    for ref in all_refs:
        if ref["name"] == refName:
            for dateV in sorted(ref["data"].keys()):
                if dateV >= first_date:
                    results[dateV] = ref["data"][dateV]
            return results

def get_all_date_val(date_val, limit):
    all_date_val = []
    for key in date_val:
        if key < limit:
            continue
        all_date_val.append(date_val[key].strftime('%b/%y'))
    return all_date_val

def get_new_all_date_val(date_val, limit):
    all_date_val = get_all_date_val(date_val, limit)

    new_all_date_val = []
    temp_new_all_date_val = []

    for j in range(len(all_date_val)):
        if j % 24 == 0:
            new_all_date_val.append(all_date_val[j])
        else:
            new_all_date_val.append('')

    for j in range(len(all_date_val)):
        if j % 6 == 0 and (j % 12 == 0 or j % 24 == 0):
            temp_new_all_date_val.append(all_date_val[j])
        else:
            temp_new_all_date_val.append('')

    mashedUplist = []
    mashedUplist.append(temp_new_all_date_val)
    mashedUplist.append(new_all_date_val)

    return mashedUplist

def figure_1(val_row_list, atlasList, date_val, title_name, limit, figNo):
    plt.rcParams["font.weight"] = "bold"
    plt.rcParams["axes.labelweight"] = "bold"

    fig_1, ax_1 = plt.subplots()
    fig_1.set_canvas(plt.gcf().canvas)
    width = 3

    for i in range(len(atlasList)):
        all_date_val = get_all_date_val(date_val, limit)

        while len(all_date_val) > len(val_row_list[i]):
            all_date_val.pop()

        df_new = pd.DataFrame({'x': all_date_val, '{}'.format(atlasList[i]): np.array(val_row_list[i])})
        ax_1.plot('x', '{}'.format(atlasList[i]), data=df_new, linewidth=width)

    x = np.array(get_new_all_date_val(date_val, limit)[0])
    ax_1.set_xticks(x)
    ax_1.set_xticklabels(get_new_all_date_val(date_val, limit)[1])

    ax_1.set_title('{}'.format(title_name), fontsize=14, weight="bold")
    ax_1.yaxis.set_major_formatter(mtick.PercentFormatter())

    ax_1.grid()
    # uncomment to set the limits
    # ax_1.set_ylim(ymin=0)
    # ax_1.set_xlim(xmin=0)

    plt.legend(bbox_to_anchor=(0.5, -0.15), loc='lower center', ncol=len(atlasList))
    fig_1.set_size_inches(10, 7)
    plt.savefig('Fig{}.png'.format(figNo), dpi=300, bbox_inches='tight')

def figure_2(atlasList, new_date_val, val_row_list):
    plt.rcParams["font.weight"] = "bold"
    plt.rcParams["axes.labelweight"] = "bold"

    fig_1, ax_1 = plt.subplots()
    fig_1.set_canvas(plt.gcf().canvas)
    width = 3

    maxCountEntry = 0
    for i in range(len(atlasList)):
        all_date_val = []
        countEntry = 0
        for key in new_date_val:
            if key < 2:
                continue
            all_date_val.append(new_date_val[key])
            countEntry += 1

        while len(all_date_val) > len(val_row_list[i]):
            all_date_val.pop()
            countEntry -= 1

        df_new = pd.DataFrame({'x': all_date_val, '{}'.format(atlasList[i]): np.array(val_row_list[i])})
        ax_1.plot('x', '{}'.format(atlasList[i]), data=df_new, linewidth=width)
        z = np.polyfit(all_date_val, val_row_list[i], 1)
        p = np.poly1d(z)
        ax_1.plot(all_date_val, p(all_date_val), "--", label='{}'.format(atlasList[i]), linewidth=width)
        if maxCountEntry < countEntry:
            maxCountEntry = countEntry
            plt.gca().set_xticks(df_new["x"].unique())

    ax_1.set_title('Jahresmittelwerte', fontsize=14, weight="bold")
    ax_1.yaxis.set_major_formatter(mtick.PercentFormatter())

    ax_1.grid()
    # uncomment to set the limits
    # ax_1.set_ylim(ymin=0)
    ax_1.set_xlim(xmin=new_date_val[2], xmax=new_date_val[len(new_date_val)-1])

    plt.legend(bbox_to_anchor=(0.5, -0.15), loc='lower center', ncol=len(atlasList))
    fig_1.set_size_inches(15, 7)
    plt.savefig('Fig4.png', dpi=300, bbox_inches='tight')


def deviation_list(countMax, val_row_list):
    diff_list = []
    for i in range(countMax):
        maxV = max(val_row_list[i])
        minV = min(val_row_list[i])
        if maxV == minV:
            diff_list.append(maxV)
        else:
            diff_list.append(maxV - minV)

    return diff_list

def column_wise_value(data, countMax, columnPlot, sh, limit):
    '''create a list of list where each list contains entry from each column'''
    val_row_list = []
    for i in range(len(atlas_list(data))):
        val_row = []
        for j in range(countMax):
            if sh.cell(j + limit, columnPlot + 1).value == None:
                continue
            val_row.append(sh.cell(j + limit, columnPlot + 1).value)
        val_row_list.append(val_row)
        columnPlot += 1
    return val_row_list