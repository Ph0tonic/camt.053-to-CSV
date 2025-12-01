# Takes a CAMT053 XML as input and converts it to a
# comma separated values (CSV) file - e.g. for further analysis in Excel
# Tested with ISO-20022 / CAMT.053 Files of Swiss Migrosbank
# (see namespace definition below for file format version details)
# Tested with python 3.9.4
# CAMT.053 Specs
# https://www.swift.com/swift_resource/35371
# https://www.swift.com/search?keywords=camt.053&search-origin=onsite_search
# https://www.ebics.de/de/datenformate

# 28. March 2023 Joerg Kummer

import PySimpleGUI as sg
import xml.etree.ElementTree as ET

# CAMT.053 Namespace is used as the default namespace (i.e. empty string in code)
# more namespaces can be included
ns = {"": "urn:iso:std:iso:20022:tech:xsd:camt.053.001.04"}

# Keeping the CSV extension and comma as a separator means
# the file can be opened in Windows/Excel by double-clicking it in Windows
# Explorer and import dialogs can be skipped

infile = ""
infile = sg.popup_get_file(
    "Please select the file in CAMT.053 format", default_path=infile
)
if not infile:
    exit()

outfile = infile + ".csv"
print(f"Converting to output CSV File at\n{outfile}")
# separator in outfile
p = ","

# load and parse the file
tree = ET.parse(infile)
# get the root node
root = tree.getroot()

# To allows double-clicking the CSV in Windows Explorer,
# - comma should be the separator and
# - encoding should be windows-1252
f = open(outfile, "w", encoding="windows-1252")


# Safe Access to Element.text
# preventing python errors if the Element was not found etc
def sa(elm, st):
    s = elm.find(st, ns)
    if s is None:
        return "-"
    else:
        return s.text


# Safe Access to Element.attr at path st with name n
def sat(elm, st, n):
    s = elm.find(st, ns)
    if s is None:
        return "-"
    else:
        if n in s.attrib:
            return s.attrib[n]
        else:
            return "-"


def pr(elm, st):
    f.write(sa(elm, st))
    f.write(p)


# print headers
s = root.find("./BkToCstmrStmt/Stmt", ns)
f.write(
    f"Kontoauszug\nKonto{p}{sa(s, './Acct/Id/IBAN')}\nWährung{p}{sa(s, './Acct/Ccy')}\nvon{p}{sa(s, './FrToDt/FrDtTm')}\nbis{p}{sa(s, './FrToDt/ToDtTm')}\nerstellt am{p}{sa(s, './CreDtTm')}\n"
)
f.write(
    f"Booking Date{p}Valuta Date{p}Reversed{p}Status{p}Additional Info{p}Additional Tx Info{p}Number of Transactions in Booking{p}Amount{p}Currency{p}Credit/Debitor{p}Debitor{p}Creditor{p}Reference\n"
)

# Then iterate through the Ntry nodes - the bookings
for entry in root.findall(".//Ntry", ns):
    # A booking (entry) can consist of multiple transactions (tx)
    # Each transaction can involve multiple creditors, debitors
    # One row per transaction is appended to the CSV file, i.e. we
    # iterate through all transactions
    for tx in entry.findall("NtryDtls/TxDtls", ns):
        pr(entry, "BookgDt/Dt")
        pr(entry, "ValDt/Dt")
        pr(entry, "RvslInd")
        pr(entry, "Sts")
        # adding double quotes for strings, which may contain a comma
        a = '"' + sa(entry, "AddtlNtryInf") + '",'
        f.write(a)
        a = '"' + sa(tx, "AddtlTxInf") + '",'
        f.write(a)

        pr(entry, "NtryDtls/Btch/NbOfTxs")
        pr(tx, "Amt")
        f.write(sat(tx, "Amt", "Ccy"))
        f.write(p)
        pr(tx, "CdtDbtInd")
        a = '"' + sa(tx, "RltdPties/Dbtr/Nm") + '",'
        f.write(a)
        a = '"' + sa(tx, "RltdPties/Cdtr/Nm") + '",'
        f.write(a)
        a = '"' + sa(tx, "RmtInf/Ustrd") + '",'
        f.write(a)
        f.write("\n")

f.close()
