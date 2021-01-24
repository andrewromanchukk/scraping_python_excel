from bs4 import BeautifulSoup
import requests as req
import xlrd
import pandas as pd
import xlsxwriter

link = "https://habr.com/ru/top/page"


def make_hyperlink(text, link):
    return '=HYPERLINK("%s","%s")'%(text,link.format(link))

def goParse(link):
    titles, links, times, hyperlinks = [], [], [], []
    counter = 1
    while(counter<=3):
        print(counter)
        res = req.get(link + str(counter))
        html = BeautifulSoup(res.text, 'lxml')
        times += html.find_all('span', class_='post__time')
        links_a = html.find_all('a', class_='post__title_link')
        page = html.find_all('a', id='next_page')

        for a in links_a:
            # titles.append(a.text) 
            # links.append(a['href']) 
            hyperlinks.append(make_hyperlink(a['href'], a.text))
 
        if page == None:
            break
        counter += 1

    for i, time in enumerate(times):
        times[i] = time.text

    df = pd.DataFrame()
    df['Time'] = times
    # df['Titles'] = titles
    # df['Links'] = links
    df['HyperLinks'] = hyperlinks

    writer = pd.ExcelWriter('habr.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Аркуш1', index=False)

    writer.sheets['Аркуш1'].set_column('A:A', 20)
    writer.sheets['Аркуш1'].set_column('B:B', 100)
    # writer.sheets['Аркуш1'].set_column('C:C', 50)

    writer.save()

goParse(link)