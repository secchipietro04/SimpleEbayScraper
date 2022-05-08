import requests
import re
from urlextract import URLExtract
from bs4 import BeautifulSoup
import xlsxwriter
import io
import urllib.request
import urllib
from PIL import Image

request = "server"
numberPerPage = [60, 120, 240][2]
maxPrice = 600
minPrice = 0
country = 3  # Eu


def getElementByIdT(identif, text):
    page = BeautifulSoup(text)
    res = page.find(id=identif)
    return res


def getElementsByClassT(className, typeElement, text):
    page = BeautifulSoup(text)
    res = page.find_all(typeElement, className)
    return res


def getContentByClassT(className, typeElement, idx, text):
    res = getElementsByClassT(className, typeElement, text)

    if len(res) > 0:
        return re.findall(">(.*)<", str(res[idx]))[0]
    else:
        return ""


def getContentByIdT(identif, text):
    res = getElementByIdT(identif, text)
    if len(res) > 0:
        return re.findall(">(.*)<", str(res))[0]
    else:
        return ""


class Collection:
    def __init__(self):
        self.items = set()
        self.a = 0

    def add(self, setUrls):
        for i in setUrls:
            item = ItemL(i)
            print("preso item:  " + str(self.a))
            self.items.add(item.cleanPage())
            self.a += 1


class ItemL:
    def __init__(self, url):
        self.url = url
        self.page = self.getPage()
        self.price = self.getPrice()
        self.imageUrl = self.getImageUrl()
        self.title = self.getTitle()

    def getElementById(self, identif):
        page = BeautifulSoup(self.page)
        res = page.find(id=identif)
        return res

    def getElementsByClass(self, className, typeElement):
        page = BeautifulSoup(self.page)
        res = page.find_all(typeElement, className)
        return res

    def getContentByClass(self, className, typeElement, idx):
        res = self.getElementsByClass(className, typeElement)
        if (res is not None):
            return re.findall(">(.*)<", str(res[idx]))[0]
        else:
            return ""

    def getContentById(self, identif):
        res = self.getElementById(identif)
        if (res is not None):
            return re.findall(">(.*)<", str(res))[0]
        else:
            return ""

    def getPage(self):
        x = requests.get(self.url)
        return x.text

    def cleanPage(self):
        self.page = None
        return self

    def getInformations(self):
        self.price = self.getPrice()
        self.imageUrl = self.getImageUrl()
        self.title = self.getTitle()

    def getPrice(self):
        out = ""
        out = self.getContentById("prcIsum")
        self.auction = False
        if (out == ""):
            out = self.getContentById("prcIsum_bidPrice")
            self.auction = True
        try:
            out = re.findall("EUR (.*)", out)[0]
            out = out.replace(",", ".")
            return float(str(out))
        except:
            return -1

    def getImageUrl(self):
        try:
            res = self.getElementById("icImg")
            return res.attrs["src"]
        except:
            return "https://www.salonlfc.com/wp-content/uploads/2018/01/image-not-found-1-scaled-1150x647.png"

    def getTitle(self):
        res = self.getElementById("LeftSummaryPanel")
        try:
            res = getElementsByClassT("ux-textspans ux-textspans--BOLD", "span", str(res))[0]
            return re.findall(">(.*)<", str(res))[0]
        except:
            return ""


def cleanResults(linksP):
    out = set()
    for i in linksP:
        if "hash=item" in i:
            out.add(i)
    return out


def getListingLinks():
    query = "https://www.ebay.de/sch/i.html?_from=R40&_nkw=" + request + "&_sacat=0&_sop=10&_udhi=" + str(
        maxPrice) + "&rt=nc&LH_PrefLoc=" + str(country) + "&_ipg=" + str(numberPerPage)
    allItems = set()
    for i in range(1, 22):
        x = requests.get(query + "&_pgn=" + str(i))
        text = x.text
        extractor = URLExtract()
        allItems.update(tuple(extractor.find_urls(text)))
        print("presi link:  " + str(i))
    return allItems


if __name__ == '__main__':
    workbook = xlsxwriter.Workbook('Listing.xlsx')
    worksheet = workbook.add_worksheet()
    links = cleanResults(getListingLinks())
    items = Collection()

    items.add(links)
    row = 1

    head = ["link", "title", "price", "auction", "image"]
    width = 140.0
    height = 182.0
    for col_num, data in enumerate(head):
        worksheet.write(0, col_num, data)
    for i in items.items:
        col = 0
        worksheet.write(row, col, i.url)
        col += 1
        worksheet.write(row, col, i.title)
        col += 1
        worksheet.write(row, col, i.price)
        col += 1
        worksheet.write(row, col, str(i.auction))
        col += 1

        image_data = io.BytesIO(urllib.request.urlopen(i.imageUrl).read())
        image = Image.open(image_data).convert("RGBA")
        scale_x = width / image.size[0]
        scale_y = width / image.size[1]

        worksheet.set_row_pixels(row=row, height=height)
        worksheet.set_column_pixels(first_col=col, last_col=col, width=width)
        worksheet.insert_image(row, col, "",
                               {'image_data': image_data, 'object_position': 1, 'x_scale': scale_x, 'y_scale': scale_y})
        row += 1

    workbook.close()
