import win32com.client
import re
import csv

DL_NAME = 'DL-Jama-Champions'
outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
namespace = outApp.GetNamespace("MAPI")
address_lists = namespace.AddressLists

PR_COUNTRY = "http://schemas.microsoft.com/mapi/proptag/0x3A26001F"
PR_LOCALITY = "http://schemas.microsoft.com/mapi/proptag/0x3A27001F"

processed_members = set()

# Funcție pentru a extrage divizia și subdiviziunea și eliminarea parantezei dacă există
def extract_division_and_subdivision(name):
    match = re.search(r'\((\S+)\s+(\S+)', name)
    if match:
        division = match.group(1).replace(")", "")  # Elimină paranteza închisă
        subdivision = match.group(2).replace(")", "")  # Asigură că nu există paranteză în subdivizie
        return division, subdivision
    return "Unknown", "Unknown"

def get_members_from_dl(dl, writer):
    for member in dl.Members:
        name = member.Name
        smtp_address = None
        country = "Unknown"
        city = "Unknown"
        division, subdivision = extract_division_and_subdivision(name)

        try:
            if member.AddressEntryUserType == win32com.client.constants.olExchangeDistributionListAddressEntry:
                sub_dl = member.GetExchangeDistributionList()
                if sub_dl is not None:
                    writer.writerow([country, city, division, subdivision, f"Subgroup: {name}", ""])
                    get_members_from_dl(sub_dl, writer)
            else:
                if member.AddressEntryUserType == win32com.client.constants.olExchangeUserAddressEntry:
                    smtp_address = member.GetExchangeUser().PrimarySmtpAddress
                else:
                    smtp_address = member.Address

                if smtp_address in processed_members:
                    continue
                else:
                    processed_members.add(smtp_address)

                property_accessor = member.PropertyAccessor
                try:
                    country = property_accessor.GetProperty(PR_COUNTRY)
                except:
                    country = "Unknown"

                try:
                    city = property_accessor.GetProperty(PR_LOCALITY)
                except:
                    city = "Unknown"

            # Elimină membrii dacă există mai multe "Unknown"-uri (mai mult de unul)
            if [country, city, division, subdivision].count("Unknown") > 1:
                continue

            if smtp_address:
                writer.writerow([country, city, division, subdivision, name, smtp_address])
            else:
                writer.writerow([country, city, division, subdivision, name, "Unknown"])

        except Exception as e:
            print(f"Error processing member: {name}, {e}")
            writer.writerow([country, city, division, subdivision, name, "Error"])

for address_list in address_lists:
    if address_list.Name == "Global Address List":
        gal = address_list.AddressEntries
        dl = gal.Item(DL_NAME)
        
        if dl.AddressEntryUserType == win32com.client.constants.olExchangeDistributionListAddressEntry:
            print(f'Membrii din lista de distribuție {DL_NAME}:')
            
            with open('champions.csv', mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['Country', 'City', 'Division', 'Subdivision', 'Name', 'Email Address'])
                get_members_from_dl(dl, writer)

        break

print("Fișierul champions.csv a fost creat.")