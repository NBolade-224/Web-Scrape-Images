from requests_html import AsyncHTMLSession, HTMLSession
import openpyxl, os
import urllib.request


ExcelFile = openpyxl.load_workbook("C:\\Users\\nickb\\Videos\\Commercial Testing\MHStar SKU test.xlsx")
ExcelSheet = ExcelFile.active
SKUList = []
SKUwebpage = {}
for cell in ExcelSheet['A']:
    SKUList.append(cell.value)

destfolder = "C:\\Users\\nickb\\Videos\\Commercial Testing\\AOSOM IMAGE\\"

asession  = AsyncHTMLSession()
session  = HTMLSession()

urls1 = [
"https://www.aosom.co.uk/category/home~420/",
"https://www.aosom.co.uk/category/office~199/",
"https://www.aosom.co.uk/category/sports-leisure~161/",
"https://www.aosom.co.uk/category/health-beauty~208/",
"https://www.aosom.co.uk/category/toys-games~230/",
"https://www.aosom.co.uk/category/baby-products~250/",
"https://www.aosom.co.uk/patio-lawn-garden/garden-shades-c789.html",
"https://www.aosom.co.uk/patio-lawn-garden/garden-buildings-c798.html",
"https://www.aosom.co.uk/patio-lawn-garden/garden-tools-c800.html",
"https://www.aosom.co.uk/patio-lawn-garden/barbecues-c748.html",]

urls2 = [
"https://www.aosom.co.uk/patio-lawn-garden/garden-decor-c801.html",
"https://www.aosom.co.uk/patio-lawn-garden/garden-planters-stands-c799.html",
"https://www.aosom.co.uk/category/fire-pits~132/",
"https://www.aosom.co.uk/category/kitchen-equipment~443/",
"https://www.aosom.co.uk/category/diy~215/",
"https://www.aosom.co.uk/category/garden-furniture-accessories~131/",
"https://www.aosom.co.uk/category/pet-supplies~184/",
"https://www.aosom.co.uk/category/sofa-lounges~434/",
"https://www.aosom.co.uk/category/home-furniture-all~423/",
"https://www.aosom.co.uk/category/storage-cleaning-solutions~425/",
"https://www.aosom.co.uk/category/shoe-bench~626/",
"https://www.aosom.co.uk/category/bathroom~421/"]

async def Scraper1(vurl):
    m = 1
    for x in range(0,100):
        url = vurl+"?column=0&page=%d&psort=0" % int(m)
        page = await asession.get(url)
        products = page.html.find('.display-good-item')
        if products == []:
            break
        for x in products:
            if x.attrs.get('sellersku') in SKUList:
                print("https://www.aosom.co.uk"+x.attrs.get('href'))
                SKUwebpage[x.attrs.get('sellersku')] = "https://www.aosom.co.uk"+x.attrs.get('href')
                print()
        m += 1


asession.run(*[lambda vurl=vurl: Scraper1(vurl) for vurl in urls1])
asession.run(*[lambda vurl=vurl: Scraper1(vurl) for vurl in urls2])

for skus in SKUwebpage:
    page = session.get(SKUwebpage.get(skus))
    images = page.html.find('.cloudzoom-gallery')
    try:
        os.mkdir(destfolder+'%s' % skus) 
    except:
        pass
    for y,x in enumerate(images):
        badurl = x.attrs.get('data-src')
        fixedurl = badurl.replace("/thumbnail/100/n6/","/100/")
        urllib.request.urlretrieve(fixedurl, destfolder+'%s\\%s.jpg' % (skus,y+1))
        print(fixedurl)
        Row = SKUList.index(skus) + 1
        print(x.attrs.get('data-src'))
        ExcelSheet['B%d' % int(Row)] = "FOUND"
        print()


ExcelFile.save("C:\\Users\\nickb\\Videos\\Commercial Testing\\MHStar SKU test.xlsx")


