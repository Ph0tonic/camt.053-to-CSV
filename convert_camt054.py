# Takes a CAMT054 XML as input and converts it to a
# comma separated values (CSV) file - e.g. for further analysis in Excel
# Tested with ISO-20022 / CAMT.054 Files
# (see namespace definition below for file format version details)
# Tested with python 3.9.4
# CAMT.054 Specs
# https://www.swift.com/swift_resource/35371
# https://www.swift.com/search?keywords=camt.054&search-origin=onsite_search
# https://www.ebics.de/de/datenformate

# Based on convert.py for CAMT.053
# Adapted for CAMT.054 (Debit/Credit Notification)
# 2 February 2026

import xml.etree.ElementTree as ET
import sys

# CAMT.054 Namespace is used as the default namespace (i.e. empty string in code)
# more namespaces can be included
# Note: CAMT.054 uses a different namespace than CAMT.053
# Support for multiple versions (04, 08, etc.)
ns = {"": "urn:iso:std:iso:20022:tech:xsd:camt.054.001.08"}

# Keeping the CSV extension and comma as a separator means
# the file can be opened in Windows/Excel by double-clicking it in Windows
# Explorer and import dialogs can be skipped

infile = ""
try:
    infile = sys.argv[1]
except:
    import PySimpleGUI as sg

    infile = sg.PopupGetFile(
        "Please select the file in CAMT.054 format", default_path=infile
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
f = open(outfile, "w", encoding="utf-8")


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
# CAMT.054 uses BkToCstmrDbtCdtNtfctn instead of BkToCstmrStmt
s = root.find("./BkToCstmrDbtCdtNtfctn/Ntfctn", ns)
if s is not None:
    f.write(
        f"Notification de Débit/Crédit\nCompte{p}{sa(s, './Acct/Id/IBAN')}\nDevise{p}{sa(s, './Acct/Ccy')}\nDe{p}{sa(s, './FrToDt/FrDtTm')}\nÀ{p}{sa(s, './FrToDt/ToDtTm')}\nCréé le{p}{sa(s, './CreDtTm')}\n"
    )
else:
    f.write("Notification de Débit/Crédit\n")

f.write(
    f"Booking Date{p}Valuta Date{p}Reversed{p}Status{p}Additional Info{p}AcctSvcrRef{p}InstrId{p}Number of Transactions in Booking{p}Amount{p}Currency{p}Credit/Debit{p}Debitor{p}Creditor{p}Reference{p}Remittance Info{p}Notification ID\n"
)

# Then iterate through the Ntry nodes - the notifications
# In CAMT.054, the structure is BkToCstmrDbtCdtNtfctn/Ntfctn/Ntry
# Using .//Ntry to find all Ntry elements regardless of exact path
for entry in root.findall(".//Ntry", ns):
    # A notification entry can consist of multiple transactions (tx)
    # Each transaction can involve multiple creditors, debitors
    # One row per transaction is appended to the CSV file, i.e. we
    # iterate through all transactions

    # Get notification ID if available
    ntfctn = entry.find("../../", ns)
    ntfctn_id = sa(ntfctn, "Id") if ntfctn is not None else "-"

    for tx in entry.findall("NtryDtls/TxDtls", ns):
        pr(entry, "BookgDt/Dt")
        pr(entry, "ValDt/Dt")
        pr(entry, "RvslInd")
        pr(entry, "Sts/Cd")
        # adding double quotes for strings, which may contain a comma
        a = '"' + sa(entry, "AddtlNtryInf") + '",'
        f.write(a)
        # AcctSvcrRef and InstrId from transaction references
        a = '"' + sa(tx, "Refs/AcctSvcrRef") + '",'
        f.write(a)
        a = '"' + sa(tx, "Refs/InstrId") + '",'
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
        # Extract reference - try different paths
        ref = sa(tx, "RmtInf/Ustrd")
        if ref == "-":
            ref = sa(tx, "RmtInf/Strd/CdtrRefInf/Ref")
        a = '"' + ref + '",'
        f.write(a)
        # Extract remittance info (additional details)
        rmt_info = sa(tx, "RmtInf/Strd/AddtlRmtInf")
        if rmt_info == "-":
            rmt_info = sa(tx, "RmtInf/Ustrd")
        a = '"' + rmt_info + '",'
        f.write(a)
        a = '"' + ntfctn_id + '",'
        f.write(a)
        f.write("\n")

f.close()

print(f"Conversion completed successfully!")
