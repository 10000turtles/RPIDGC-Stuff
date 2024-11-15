import getData
import xlsxwriter as xw
import time
from datetime import datetime, timedelta, date
import math
import statistics as stat


def get_date(date_str, event, round):
    mon_to_num = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }

    if event == "Paint Bureau Pro League Season 4 Front 9 ":
        end_date = date_str[date_str.index("to") + 3 :]
        end_date_spl = end_date.split("-")
        end_date_obj = date(
            day=int(end_date_spl[0]),
            month=mon_to_num[end_date_spl[1]],
            year=int(end_date_spl[2]),
        )

        week_delta = timedelta(weeks=1)

        final_date_obj = end_date_obj - week_delta * (9 - round)
        return_str = final_date_obj.strftime("%m/%d/%Y")

        start_day = date(2024, 1, 1)
        date_diff = abs(final_date_obj - start_day)
        return_frac = date_diff.days / 365

        return return_str, return_frac

    if event == "Paint Bureau Pro League Season 4 Back 9":
        end_date = date_str[date_str.index("to") + 3 :]
        end_date_spl = end_date.split("-")
        end_date_obj = date(
            day=int(end_date_spl[0]),
            month=mon_to_num[end_date_spl[1]],
            year=int(end_date_spl[2]),
        )

        week_delta = timedelta(weeks=1)

        final_date_obj = end_date_obj - week_delta * (9 - round)
        return_str = final_date_obj.strftime("%m/%d/%Y")

        start_day = date(2024, 1, 1)
        date_diff = abs(final_date_obj - start_day)
        return_frac = date_diff.days / 365

        return return_str, return_frac

    try:
        end_date = date_str[date_str.index("to") + 3 :]
    except:
        end_date = date_str

    end_date_spl = end_date.split("-")
    end_date_obj = date(
        day=int(end_date_spl[0]),
        month=mon_to_num[end_date_spl[1]],
        year=int(end_date_spl[2]),
    )

    return_str = end_date_obj.strftime("%m/%d/%Y")

    start_day = date(2024, 1, 1)
    date_diff = abs(end_date_obj - start_day)
    return_frac = date_diff.days / 365

    return return_str, return_frac


discapers = [
    ("Justin Mucelli", 215058),
    ("Meghan Mercier", 148113),
    ("Brian Bickersmith", 126043),
    ("Kyle Denman", 218407),
    ("Kyle Hirsch", 168321),
    ("Corey Cook", 110582),
    ("Ryan Travis", 148343),
    ("Jaimen Hume", 68908),
    ("Tucker Kozloski", 148801),
    ("Ethan Hadders", 166444),
    ("Troy Vassari", 161405),
    ("David Koch", 141999),
    ("Adam Nelson", 111953),
    ("Timothy Jiardini", 18910),
    ("Chad Larson", 20532),
    ("Daniel Eignor", 57613),
    ("Jason Gorsage", 194672),
    ("Todd Kurnat", 34565),
    ("Danny White", 35189),
    ("Greg Kurtz", 27735),
    ("Adam Shumaker", 183726),
    ("Peter Hodge", 268026),
    ("Tyler Gannon", 213970),
    ("Jeremy Kasprick", 120271),
    ("Ryan Yaddow", 115179),
    ("Anthony Bohanske", 162533),
    ("Roderick Perry", 210489),
    ("Austin Swartout", 191747),
    ("Troy Moran", 250957),
    ("Joshua Le", 276847),
    ("Alex Potts", 111106),
    ("Brandon Medina", 148805),
    ("Dan Gibbons", 176709),
    ("Marcia Focht", 66680),
    ("Terry Hudson", 199571),
    ("Steve Marshall", 185089),
    ("Evan Parsley", 194630),
    ("Peter Cornish", 156639),
    ("Kenji Cline", 37696),
    ("Jeff Darling", 125270),
    ("Carletta Darling", 191328),
    ("Kevin T. Kroencke", 209045),
    ("Mark Ungerman", 240095),
    ("Rhyan Lorenc", 209288),
    ("Ross O’Toole", 112821),
    ("David Herodes", 131502),
    ("Ben Pickering", 254381),
    ("Nicholas Bohmer", 161895),
    ("Jared Borelli", 279380),
    ("Christopher Evanchuk", 190841),
    ("Trevor Cline", 269172),
    ("Noah Fry", 175683),
    ("Brian Monahan", 195823),
    ("Sam Acevedo", 208874),
    ("Michael Merlini", 238974),
    ("Branden Cline", 253454),
    ("Kyle Walsh", 283885),
    ("Frank DeGregorio", 198910),
    ("Todd Lorenc", 209285),
    ("Derek Styczynski", 145428),
    ("Erich Struna", 235692),
    ("Nick Silvano", 184886),
    ("Joey Herrington", 223146),
    ("Eamon Foster", 272651),
    ("John Hafner", 146525),
    ("Justin Frey", 124224),
    ("Joe Jaskolka", 53716),
    ("Randy Bemis", 106525),
    ("Joseph Caron", 210706),
    ("Justin Park", 139848),
    ("Eric Bossenbroek", 242938),
    ("Chandler Harvey", 236395),
    ("Josh Manko", 238277),
    ("Dennis Daley", 246046),
    ("Logan Akins", 279217),
    # ("Michael P. Filicky", 40823),
    ("Peter Fitzgerald", 237135),
    ("Doug Iannon", 172034),
    ("Robert Gerstman", 217674),
    ("Kathleen Bemis", 120308),
    ("Rob Hendricks", 176954),
    ("Nick Warren", 145261),
    ("William Stone", 238341),
    ("Thomas Hutchinson", 171912),
    ("Dell Potts", 225398),
    ("Mindy Potts", 112622),
    ("William Harrison", 147961),
    ("Niel Hall", 238273),
    ("Eric Conine", 244582),
    ("Nicholas Esposito", 253951),
    ("Bart Welch", 76909),
    ("Brian Neary", 279717),
    ("Cody Conine", 280282),
    ("Zachary Jordan", 252739),
    ("George Henry", 269324),
    ("Mark Bryan", 146901),
    ("Sean Dollard", 239450),
    ("Justin Scott", 281748),
    ("Padriac Higgins", 275508),
    ("Shaina Desapio", 269323),
    ("Megan Bursell", 272570),
    ("Melissa Calabria", 198003),
    ("Aaron Holtby", 179962),
    ("Abi Bleier", 258009),
    ("Adam Gill", 170856),
    ("Adrian VanHall", 143524),
    ("Aiden Glennon", 146030),
    ("Alden Slack", 157289),
    ("Alex Helenek", 106577),
    ("Alexis White", 246012),
    ("Andrew Barger", 122203),
    ("Andrew Gulak", 211528),
    ("Andrew Tighe", 62529),
    ("Andy Ziemens", 150957),
    ("April King-Hampel", 91391),
    ("Austin Dudla", 200334),
    ("Ben Hayko", 139443),
    ("Beth Spaulding", 147521),
    ("Bobby Hallum", 212441),
    ("Brendan Sisk", 124857),
    ("Brendan Woods", 140935),
    ("Brent Irving", 165202),
    ("Brett Delamater", 56208),
    ("Brock Degraw", 163833),
    ("Carly Spaulding", 146017),
    ("Charles Carpenter III", 153703),
    ("Chris Dahl", 107474),
    ("Chris DelBianco", 83008),
    ("Christina Nardi", 156558),
    ("Cody Meddis", 86190),
    ("Cole Mediratta", 247764),
    ("Crysta Kovach", 258108),
    ("Dan Eignor", 57613),
    ("Danny Partin", 81246),
    ("Daryl Menton", 138175),
    ("Dave Chaiken", 203276),
    ("Dave Herodes", 131502),
    ("Dave Hudson", 98094),
    ("Dave Macisaac", 149644),
    ("Dave Moore", 59659),
    ("Derek Fostyk", 101301),
    ("Devin Declerk", 94965),
    ("Dylan Courtney", 56434),
    ("Dylan Johanson", 113276),
    ("Earl Steenburg", 57658),
    ("Evan Lloyd", 235704),
    ("Gene Gierka", 40396),
    ("Harrison Lehmann", 60755),
    ("Ian Daughton", 265220),
    ("JJ Knapp", 240246),
    ("Jacob Najac", 40016),
    ("Jacqueline Kirkpatrick", 114349),
    ("Jake Vandelinde", 146056),
    ("James Black", 201511),
    ("James Scow", 103199),
    ("Jamie Borge", 185528),
    ("Jamin Totino", 100661),
    ("Jasan LaSasso", 65363),
    ("Jeff Wiechowski", 11653),
    ("Jeremy Bledsoe", 64217),
    ("Jeremy Milyon", 146069),
    ("Jeremy Whitaker", 143465),
    ("Jessica LaSasso", 70789),
    ("Jibreel Frawan", 147404),
    ("Jim Fulmer", 119539),
    ("Jim Seymour", 246863),
    ("Jirapat Khamgaseam", 233397),
    ("Joe Caron", 210706),
    ("John Tamburino", 163070),
    ("Jon Stephan", 139811),
    ("Jon Vermilyea", 187838),
    ("Joseph Kowalik", 148421),
    ("Josh Weinstock", 43843),
    ("Josh Winn", 58801),
    ("Julia Marger", 182826),
    ("Juliet Barney", 177946),
    ("Justin Lamarche", 149677),
    ("Justin Lor", 230615),
    ("Justin Meddis", 105528),
    ("Kaitlyn Clay", 125178),
    ("Kaley Russel", 127316),
    ("Karly Calzada", 129561),
    ("Kat Bemis", 120308),
    ("Kathryn Chiacchia", 65398),
    ("Kevin Nguyen", 179013),
    ("Kristin McDonald", 160273),
    ("Kurt Osterlitz", 104553),
    ("Kyle Moore", 203291),
    ("Levi Vadnais", 260796),
    ("Liam Doyle", 72954),
    ("Lisa Lundquest", 141307),
    ("Marc Ox", 54859),
    ("Mark Dami", 180373),
    ("Mark Hay", 112400),
    ("Matt Culley", 75193),
    ("Matt Ellis", 226444),
    ("Matt Gamache", 94319),
    ("Matt Sharp", 156321),
    ("Matthew Starzyk", 149893),
    ("Michael Young", 88401),
    ("Mike LaRue", 109016),
    ("Mike Lazzaro", 62457),
    ("Mike Tehan", 71915),
    ("Mike Winn", 144354),
    ("Mike Zanchelli", 73252),
    ("Miranda Carpenter", 168999),
    ("Nick Esposito", 79390),
    ("Nick Terralavoro", 91398),
    ("Parker Cerone", 153437),
    ("Patrick Shanley", 243506),
    ("Peter Lunstead", 48755),
    ("Rob Gertsman", 217674),
    ("Ryan Carpenter", 172044),
    ("Ryan Clair", 62363),
    ("Ryan Jeskie", 192838),
    ("Ryan Kendrick", 56439),
    ("Ryan Martin", 195939),
    ("Ryan Nelson", 77652),
    ("Sam Richter", 213251),
    ("Savanna Burke", 147531),
    ("Scott Mosher", 196128),
    ("Seth Thomas", 193002),
    ("Shane Delameter", 117902),
    ("Shane Osterlitz", 253293),
    ("Sparky Spauling", 117638),
    ("Thomas Rascona", 132140),
    ("Tim Defranco", 53270),
    ("Tim Goyette", 194215),
    ("Tim Martino", 241992),
    ("Todd Martin", 117962),
    ("Tony Malikowski", 54361),
    ("Travis Bushore", 58542),
    ("Travis West", 192553),
    ("Tucker Middleton", 58742),
    ("Tyler Calzada", 119183),
    # ("William Bach", 31885),
    ("Willy Harrison", 147961),
    ("Zac McDonald", 160274),
    ("Zach Beaudet", 31608),
    ("Zach Hatch", 189127),
    ("Jon Meschutt", 219104),
    ("Bill Archer", 86097),
    ("EJ Bilodeau", 238267),
    ("Kevin Jacobs", 45608),
    # ("Jason Murak", 137521),
    ("Zak Joyce", 191000),
    ("Cory Manning", 245732),
    ("Justin D'Aust", 121854),
    ("Nathan Prentice", 122064),
    ("Charles Gordon Jr", 157278),
    ("Benjamin Hayko", 139443),
    ("Eric Moreira", 231709),
    ("Michael Kaiserian", 191197),
    ("James Squier", 215094),
    ("Alexander Trainor", 209712),
    ("Kelli Terpening", 189583),
    ("Julia Mae Marger", 182826),
    ("Brian Betit", 66454),
    ("Jacob Driskill", 189284),
    ("Kevin Relyea", 169004),
    ("Jay O'Leary", 216907),
    ("Pat Keenan", 72514),
    ("Christy Betit", 69552),
    # ("Ron Tyre", 16266),
    ("Robert Virginia", 216663),
    ("Jason Matts", 196109),
    ("Jim Seymour", 238070),
    ("Christopher Yeager", 78732),
    ("Ryan Phillips", 127782),
    ("Ivan Potocnik", 269829),
    ("Todd Everleth", 166917),
    ("Andrew Disalvi", 195268),
    ("Connor Eckhardt", 244210),
    ("Luke Debritz", 232216),
    ("Jacob Dunbar", 225669),
    ("Tyler Reynolds", 121025),
    ("Tony Baxter", 187360),
    ("Anthony McQuiston", 283612),
    ("Greg Lai", 277403),
    ("Amanda Dunbar", 225671),
    ("Danielle Wooddell", 233154),
    ("Lucas Kotleski", 239875),
    ("Justin Hickok", 233361),
    ("Nathan Farley", 177810),
    ("Devin Lamke", 175858),
    ("Nolan Savage", 281360),
    ("Justin Buskey", 169801),
    ("Roy Lancaster", 156066),
    ("Adam Dahl", 235584),
    ("Christopher Wilson", 270576),
    ("Ryan Lucht", 283616),
    ("Joe Ponessa", 250583),
    ("Victor Disalvi", 241280),
    ("Kat Backman", 255848),
    ("Adam Hauck", 62634),
    ("Adam Selmon", 118039),
    ("Alex Peet", 274725),
    ("Alson Peterson", 85724),
    ("Amber Stout", 158859),
    ("Andrew Chamberlaine", 196900),
    ("Andrew DiSalvi", 195268),
    ("Andy Kloss", 106608),
    ("Ben Houck", 151901),
    ("Bill Bowen", 85219),
    ("Billie Riley", 146748),
    ("Blaine Mendrysa", 205616),
    ("Bob Kulchuck", 54607),
    ("Bradley Martin", 269580),
    ("Brian J Neary", 279717),
    ("Bryan Elsworth", 146887),
    ("Caleb Hellinger-Rock", 270772),
    ("Cameron Rivers", 100423),
    ("Chris Evanchuck", 190841),
    ("Chris Heminway", 218930),
    ("Chris Williams", 161096),
    ("Chris Wilson", 19014),
    ("Dan Blake", 158366),
    ("Dan McCrea", 279970),
    ("Daniel Pupke", 256677),
    ("Daniel Smith", 133820),
    ("Danielle Woddell", 233154),
    ("Darren Mckinney", 196657),
    ("Dave Swain", 202045),
    ("Eric London", 162158),
    ("Francisco Caamano", 141704),
    ("Frank Degregorio", 198910),
    ("Gertie Czarnecki", 226431),
    ("Gil Flanders", 190796),
    ("Glen Oldrich", 232810),
    ("Hoffman Bob", 43262),
    ("Hoffman Greg", 43162),
    ("Jack Bradley", 51744),
    ("Jay Halpin", 159189),
    ("Jeff Anderson", 285891),
    ("Jeff Daniels", 158633),
    ("Jeff Tehan", 71927),
    ("Jeff Wikstrom", 163823),
    ("Jeff Zipkin", 30996),
    ("Joe Gaspardi", 53052),
    ("Joe Kowalik", 148421),
    ("John Furman", 184372),
    ("Jordan Johns", 178787),
    ("Josh Sack", 239268),
    ("Justin Hickock", 233361),
    ("Justin Winter", 273671),
    ("Laurie Borge", 192398),
    ("Mark Stryker", 53554),
    ("Matt Wall", 88112),
    ("Matt Woodell", 273093),
    ("Matthew Knauf", 113247),
    ("Mike (Nassua) Warner", 166313),
    ("Mike Schwartz", 77987),
    ("Mike (Highland) Warner", 166845),
    ("Nick Cardone", 221517),
    ("Nick D'Amato", 84217),
    ("Rob Samson", 272039),
    ("Robert Immel", 197589),
    ("Scott Radford", 125952),
    ("David Van Scoy", 146016),
    ("Sean Partee", 161205),
    ("Shari Fish", 232802),
    ("Shawn Weber", 118508),
    ("Steve Banatoski", 66403),
    ("Tim Golden", 80142),
    ("Tom Rascona", 132140),
    ("Travis Nutting", 123633),
    ("Tyler Holloway", 201807),
    ("Will Borge", 192397),
    ("Will Mehls", 35820),
    ("Dunccino", 187533),
]

players = [("Mooch", 215058)]

discaptians = [
    ("Kyle Hirsch", 168321),
    ("Corey Cook", 110582),
    ("Justin Mucelli", 215058),
    ("Meghan Mercier", 148113),
    ("Brian Bickersmith", 126043),
    ("Ethan Hadders", 166444),
    ("David Koch", 141999),
    ("Adam Nelson", 111953),
    ("Timothy Jiardini", 18910),
    ("Jason Gorsage", 194672),
    ("Anthony Bohanske", 162533),
    ("Harrison Lehmann", 60755),
    ("Josh Winn", 58801),
    ("Matt Culley", 75193),
    ("Parker Cerone", 153437),
    ("Thomas Rascona", 132140),
    ("Adam Selmon", 118039),
    ("Greg Kurtz", 27735),
    ("Jeff Wiechowski", 11653),
    ("Chris Bolton", 39910),
    ("Kenji Cline", 37696),
    ("Troy Vassari", 161405),
    ("Kaitlyn Clay", 125178),
    ("Julia Mae Marger", 182826),
    ("Marcia Focht", 66680),
    ("Jaimen Hume", 68908),
    ("Jasan LaSasso", 65363),
    ("Jessica LaSasso", 70789),
]
devils_grove = [
    ("Dan Racaniello", 106240),
    ("Scott Walsh", 18991),
    ("Chase Short", 113386),
    ("Aidan Bailey", 251218),
    ("David Janco", 166039),
    ("JT Haggett", 94436),
    ("Ryan Sullivan", 149101),
    ("Ashton McLaughlin", 101161),
    ("Phil Denoncourt", 73774),
    ("Aaron Wilmot", 59811),
    ("Brad Russell", 163016),
    ("Matt Burnette", 88230),
    ("Ervin Brown", 96490),
    ("Keith Eaton", 85097),
    ("Alic Shorey", 92130),
    ("Chris Brunick", 96013),
    ("Gavin Joannides", 101085),
    ("Jared Girardin", 94559),
    ("Jason Dix", 96147),
    ("Jon Ross", 91185),
    ("Miles Knight", 125494),
    ("Nathan Albert", 259574),
    ("Allison Hagget", 91335),
    ("Dominique Ross", 87869),
    ("Nicole Russell", 88131),
]

all_round_ratings = []  # tuples of (name, rating, event, date, division,round,tier)
mpo_round_ratings = []
fpo_round_ratings = []
mp40_round_ratings = []
mp50_round_ratings = []
mp60_round_ratings = []
fp40_round_ratings = []
fp50_round_ratings = []
fp60_round_ratings = []
ma1_round_ratings = []
fa1_round_ratings = []
ma40_round_ratings = []
ma50_round_ratings = []
ma60_round_ratings = []
fa40_round_ratings = []
fa50_round_ratings = []
fa60_round_ratings = []
ma2_round_ratings = []
ma3_round_ratings = []
fa2_round_ratings = []
fa3_round_ratings = []


for player in players:
    print(str(player[1]))
    rating_data = getData.getRatingsInfo(str(player[1]))

    for data in rating_data:

        if data[4] != "Semis" and data[4] != "Finals":
            date_mod, date_frac = get_date(data[2], data[1], int(data[4]))

        else:
            date_mod, date_frac = get_date(data[2], data[1], 1)

        all_round_ratings.append(
            (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
        )
        if data[3] == "MPO":
            mpo_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FPO":
            fpo_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MP40":
            mp40_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MP50":
            mp50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MP60":
            mp60_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FP40":
            fp40_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FP50":
            fp50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FP60":
            fp60_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA1":
            ma1_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA1":
            fa1_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA40":
            ma40_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA50":
            ma50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA60":
            ma50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA40":
            fa40_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA50":
            fa50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA60":
            fa50_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA2":
            ma2_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "MA3":
            ma3_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA2":
            fa2_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        if data[3] == "FA3":
            fa3_round_ratings.append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
    time.sleep(1)


all_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
mpo_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fpo_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
mp40_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
mp50_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
mp60_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fp40_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fp50_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fp60_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma1_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa1_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma40_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma50_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma60_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa40_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa50_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa60_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma2_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
ma3_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa2_round_ratings.sort(key=lambda tup: tup[1], reverse=True)
fa3_round_ratings.sort(key=lambda tup: tup[1], reverse=True)

concat_round_ratings = [
    mpo_round_ratings,
    fpo_round_ratings,
    mp40_round_ratings,
    mp50_round_ratings,
    mp60_round_ratings,
    fp40_round_ratings,
    fp50_round_ratings,
    fp60_round_ratings,
    ma1_round_ratings,
    fa1_round_ratings,
    ma40_round_ratings,
    ma50_round_ratings,
    ma60_round_ratings,
    fa40_round_ratings,
    fa50_round_ratings,
    fa60_round_ratings,
    ma2_round_ratings,
    ma3_round_ratings,
    fa2_round_ratings,
    fa3_round_ratings,
]

group_tags = [
    "MPO",
    "FPO",
    "MP40",
    "MP50",
    "MP60",
    "FP40",
    "FP50",
    "FP60",
    "MA1",
    "FA1",
    "MA40",
    "MA50",
    "MA60",
    "FA40",
    "FA50",
    "FA60",
    "MA2",
    "MA3",
    "FA2",
    "FA3",
]


def team_analysis(worksheet, players):

    players_round_ratings = []

    for player in players:
        print(str(player[1]))
        rating_data = getData.getRatingsInfo(str(player[1]))

        players_round_ratings.append([])

        for data in rating_data:
            if data[4] != "Semis" and data[4] != "Finals":
                date_mod, date_frac = get_date(data[2], data[1], int(data[4]))

            else:
                date_mod, date_frac = get_date(data[2], data[1], 1)

            players_round_ratings[-1].append(
                (player[0], data[0], data[1], date_mod, data[3], data[5], date_frac)
            )
        time.sleep(1)
    c_o = 15
    data_columns = 5

    for pn in range(0, len(players), 1):

        worksheet.write(2, c_o + data_columns * pn, "Name")
        worksheet.write(2, c_o + data_columns * pn + 1, players[pn][0])
        worksheet.write(2, c_o + data_columns * pn + 2, "PDGA #")
        worksheet.write(2, c_o + data_columns * pn + 3, players[pn][1])

        count = 0
        for round in players_round_ratings[pn]:
            worksheet.write(4 + count, c_o + data_columns * pn, round[1])
            worksheet.write(4 + count, c_o + data_columns * pn + 1, round[2])
            worksheet.write(4 + count, c_o + data_columns * pn + 2, round[3])
            worksheet.write(4 + count, c_o + data_columns * pn + 3, round[5])
            worksheet.write(4 + count, c_o + data_columns * pn + 4, round[6])
            count = count + 1

        worksheet.add_table(
            3,
            c_o + data_columns * pn,
            3 + len(players_round_ratings[pn]),
            c_o + data_columns * pn + 4,
        )

        worksheet.write(3, c_o + data_columns * pn, "Rating")
        worksheet.write(3, c_o + data_columns * pn + 1, "Event")
        worksheet.write(3, c_o + data_columns * pn + 2, "Date")
        worksheet.write(3, c_o + data_columns * pn + 3, "Tier")
        worksheet.write(3, c_o + data_columns * pn + 4, "Date Frac")

        rnds = [round[1] for round in players_round_ratings[pn]]
        if len(rnds) == 0:
            rnds = [0, 0]
        elif len(rnds) == 1:
            rnds = [rnds[0], rnds[0]]

        avg_round_rating = stat.mean(rnds)
        pdga_rating = 0
        median_round_rating = stat.median(rnds)
        std_dev = stat.stdev(rnds)
        min_rating = min(rnds)
        max_rating = max(rnds)
        rating_range = max_rating - min_rating

        a_rnds = [round[1] for round in players_round_ratings[pn] if round[5] == "A"]
        b_rnds = [round[1] for round in players_round_ratings[pn] if round[5] == "B"]
        c_rnds = [round[1] for round in players_round_ratings[pn] if round[5] == "C"]

        if len(a_rnds) == 0:
            a_rnds = [0, 0]
        elif len(a_rnds) == 1:
            a_rnds = [a_rnds[0], a_rnds[0]]
        if len(b_rnds) == 0:
            b_rnds = [0, 0]
        elif len(b_rnds) == 1:
            b_rnds = [b_rnds[0], b_rnds[0]]
        if len(c_rnds) == 0:
            c_rnds = [0, 0]
        elif len(c_rnds) == 1:
            c_rnds = [c_rnds[0], c_rnds[0]]

        a_tier_avg = stat.mean(a_rnds)
        b_tier_avg = stat.mean(b_rnds)
        c_tier_avg = stat.mean(c_rnds)

        worksheet.write(3 + pn, 2, players[pn][0])
        worksheet.write(3 + pn, 3, players[pn][1])
        worksheet.write(3 + pn, 4, avg_round_rating)
        worksheet.write(3 + pn, 5, median_round_rating)
        worksheet.write(3 + pn, 6, pdga_rating)
        worksheet.write(3 + pn, 7, std_dev)
        worksheet.write(3 + pn, 8, min_rating)
        worksheet.write(3 + pn, 9, max_rating)
        worksheet.write(3 + pn, 10, rating_range)
        worksheet.write(3 + pn, 11, a_tier_avg)
        worksheet.write(3 + pn, 12, b_tier_avg)
        worksheet.write(3 + pn, 13, c_tier_avg)
    worksheet.add_table(2, 2, 2 + len(players), 13)

    worksheet.write(2, 2, "Name")
    worksheet.write(2, 3, "PDGA Number")
    worksheet.write(2, 4, "Avg Round Rating")
    worksheet.write(2, 5, "Median Round Rating")
    worksheet.write(2, 6, "PDGA Rating (approx)")
    worksheet.write(2, 7, "Standard Deviation")
    worksheet.write(2, 8, "Min")
    worksheet.write(2, 9, "Max")
    worksheet.write(2, 10, "Range")
    worksheet.write(2, 11, "A Tier Avg")
    worksheet.write(2, 12, "B Tier Avg")
    worksheet.write(2, 13, "C Tier Avg")
    return


# tuples of (0: name, 1: rating, 2: event, 3: date, 4: division, 5: tier, 6: date_frac)

wb = xw.Workbook("discap_data.xlsx")
ws1 = wb.add_worksheet("All of Discap")
ws2 = wb.add_worksheet("By Division")
ws3 = wb.add_worksheet("My Stats")
ws4 = wb.add_worksheet("Discaptians")
ws5 = wb.add_worksheet("Devils Grove")
ws6 = wb.add_worksheet("Discapers")

center = wb.add_format()
center.set_align("center")

collumn_offset = 5

row_counter = 3

ws1.merge_range(1, 2, 1, 5, "All Discapers", center)

ws1.write(2, 2, "Player")
ws1.write(2, 3, "Rating")
ws1.write(2, 4, "Event")
ws1.write(2, 5, "Date")

for round in all_round_ratings:
    ws1.write(row_counter, 2, round[0])
    ws1.write(row_counter, 3, round[1])
    ws1.write(row_counter, 4, round[2])
    ws1.write(row_counter, 5, round[3])
    row_counter = row_counter + 1

group_counter = 0
for group in concat_round_ratings:
    ws2.merge_range(
        1,
        2 + group_counter * collumn_offset,
        1,
        5 + group_counter * collumn_offset,
        group_tags[group_counter],
        center,
    )

    ws2.write(2, 2 + group_counter * collumn_offset, "Player")
    ws2.write(2, 3 + group_counter * collumn_offset, "Rating")
    ws2.write(2, 4 + group_counter * collumn_offset, "Event")
    ws2.write(2, 5 + group_counter * collumn_offset, "Date")

    row_counter = 3

    for round in group:
        ws2.write(row_counter, 2 + group_counter * collumn_offset, round[0])
        ws2.write(row_counter, 3 + group_counter * collumn_offset, round[1])
        ws2.write(row_counter, 4 + group_counter * collumn_offset, round[2])
        ws2.write(row_counter, 5 + group_counter * collumn_offset, round[3])
        row_counter = row_counter + 1

    group_counter = group_counter + 1

team_analysis(ws4, discaptians)
team_analysis(ws5, devils_grove)
team_analysis(ws6, discapers)


wb.close()
