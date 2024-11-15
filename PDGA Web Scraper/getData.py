import requests
from lxml import etree
from io import StringIO


def getStd(rtgs):

    mean = sum(rtgs) / len(rtgs)
    variance = sum([((x - mean) ** 2) for x in rtgs]) / len(rtgs)
    res = variance**0.5
    return res


def getRatingsInfo(p):

    parser = etree.HTMLParser()

    r = requests.get("https://www.pdga.com/player/" + p + "/details")
    html = r.content.decode("utf-8")

    tree = etree.parse(StringIO(html), parser=parser)
    root = tree.getroot()

    year = "2024"

    try:
        tourneys = root[1][2][1][0][0][0][0][2][0][0][0][0][0][8][1][-1][0][1]

        rtgs = []

        for child in tourneys:
            if child[7].text == "No":
                continue
            # if child[2].text[-4:] == year:
            rtgs.append(
                (
                    int(child[6].text),
                    child[0][0].text,
                    child[2].text,
                    child[3].text,
                    child[4].text,
                    child[1].text,
                )
            )

        return rtgs
    except:
        return []


def getRatings(root):

    tourneys = root[1][2][1][0][0][0][0][2][0][0][0][0][0][8][1][-1][0][1]
    rtgs = []

    for child in tourneys:
        if child[7].text == "No":
            continue
        rtgs.append(int(child[6].text))

    return rtgs


def getData(p):

    parser = etree.HTMLParser()

    r = requests.get("https://www.pdga.com/player/" + p + "/details")
    html = r.content.decode("utf-8")

    tree = etree.parse(StringIO(html), parser=parser)
    root = tree.getroot()

    r = root[1][2][1][0][0][0][0][2][0][0][0][0][0][6][1][0]
    for c in r:
        if str(c.attrib) == "{'class': 'current-rating'}":
            r = c
            break

    rtg = int(r[0].tail)
    rtgs = getRatings(root)

    std = getStd(rtgs)

    return rtg, std
