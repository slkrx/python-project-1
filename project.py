import random
from datetime import timedelta, date
from faker import Faker
import xml.etree.ElementTree as ET
from pathlib import Path
import xlsxwriter
from zipfile import ZipFile
import glob
import json
from sys import exit
from census_faker_io import get_dropdown_input, get_positive_integer, get_boolean

with open("state_codes.json") as file:
    states = json.loads(file.read())
owner_types = { "Business Entity": "1", "Qualified Plan": "4", "Trust": "2" }
with open("naic_states.txt") as file:
    naic_states = file.read().splitlines()
ns = {'ns': 'http://ACORD.org/Standards/Life/2'}
ET.register_namespace('', ns['ns'])

def get_options():
    global states, owner_types
    options = dict()
    options["plan_type"] = get_dropdown_input("Plan type: ", ["COLI", "CSIO"])
    options["fp_fn"] = input("FP first name: ")
    options["fp_ln"] = input("FP last name: ")
    options["count"] = get_positive_integer("Count: ")
    options["owner_name"] = input("Owner name: ")
    options["owner_state"] = get_dropdown_input("Owner state: ", states.keys())
    options["owner_type"] = get_dropdown_input("Owner type: ", owner_types.keys())
    options["replacements"] = get_boolean("Replacements (y/n): ")
    options["product_type"] = get_dropdown_input("Product type: ", ["AIUL", "AVUL"])
    options["include_address"] = get_boolean("Include address (y/n): ")
    options["has_pdf"] = get_boolean("Has PDF (y/n): ")
    return options

def generate_person(opts):
    global states, naic_states

    # feitelbergStates = [ "TX", "FL", "NY", "SC", "CT", "RI", "GA", "CA", ]
    # residenceStates = [ "TX", "FL", "SC", "PA", "MD", "OH", "DC" ]
    residenceStates = [ "PA" ]
    fake = Faker()
    person = dict()
    person["Gender (M/F)"] = random.choice(["M", "F"])
    if person["Gender (M/F)"] == "M":
        person["First Name"] = fake.first_name_male()
        person["Middle Name"] = fake.first_name_male()
    else:
        person["First Name"] = fake.first_name_female()
        person["Middle Name"] = fake.first_name_female()
    person["Last Name"] = fake.last_name()
    person["Date of Birth (mm/dd/yyyy)"] = fake.date_of_birth()
    person["Date of Hire (mm/dd/yyyy)"] = person["Date of Birth (mm/dd/yyyy)"] + timedelta(days = (20 + random.randint(0, 20)) * 365 + random.randint(0, 12) * 30 + random.randint(0, 365))
    person["SSN"] = f"{random.randint(100,999)}-{random.randint(10,99)}-{random.randint(1000,9999)}"
    person["Email"] = "samuel.rechsteiner@alphafmc.com"
    person["Owner Type"] = "Insured"
    person["Owner Name"] = f"{person['First Name']} {person['Last Name']}"
    person["Salary"] = round(30000 + 100000 * random.random(), 2)
    person["Job Title"] = fake.job()
    person["Work Street Address"] = fake.street_address()
    person["Work Street Address Line 2"] = f"Suite {random.randint(100, 999)}"
    person["Work City"] = fake.city()
    if opts["plan_type"] == "COLI":
        person["Work Site State"] = fake.state_abbr()
        if (person["Work Site State"] in ["DC", "AS", "FM", "GU", "MH", "MP", "PW", "PR", "VI"]):
            person["Work Site State"] = "PA"
    else:
        person["Work Site State"] = random.choice(residenceStates)
    person["Residence State"] = person["Work Site State"]
    person["Work Zip Code"] = f"{random.randint(10000, 99999)}"
    person["Tobacco Use (Yes/No)"] = "No"
    person["US Citizenship (Yes/No)"] = "Yes"
    if opts["replacements"]:
        person["Amount of Insurance"] = round(30000 + 100000 * random.random(), 2)
        person["Insurance Company"] = random.choice(["Penn Mutual", "Liberty"])
        person["Product Type"] = random.choice(["AIUL", "AWL"])
        person["Inforce GI/SI or Fully UW"] = random.choice(["GI", "SI", "Fully UW"])
        person["Intend to Replace (Yes/No)"] = "Yes"
        person["Same Owner (Yes/No)"] = "Yes"
        person["PML or Other"] = "PML" if person["Insurance Company"] == "Penn Mutual" else "Other"
        person["Policy Number"] = f"{random.randint(900000, 999999)}"
        person["Policy 1035 Exchange (Yes/No)"] = "No"
        person["1035 Exchange Has Loan (Yes/No)"] = "No"
        if opts["owner_state"] == "CA":
            person["Comparison to Current Coverage - Improvement"] = "Fewer benefits, but lower premiums"
            person["Comparison to Current Coverage - Improvement (specify)"] = ""
        if opts["owner_state"] == "FL":
            person["Comparison Requested (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "MA":
            person["Yield Indices Requested (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "DE":
            person["DE Summary Requested (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "GA":
            person["GA Summary Requested (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "OK":
            person["Notify Present Insurer (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "WY":
            person["Insurance Company Home Office Location - City"] = fake.city()
            person["Insurance Company Home Office Location - State"] = random.choice([s for s in list(states.keys()) if s != "Other"])
            person["Surrender Charge of Policy Being Replaced %"] = f"{random.randint(0,100)}"
            person["Surrender Charge of Policy Being Replaced $"] = f"{random.randint(10000,50000)}"
        if opts["owner_state"] == "WA":
            person["Reduced Benefits or Increased Premiums (Yes/No)"] = "Yes"
            person["New Penalties or Charges (Yes/No)"] = "Yes"
            person["Existing Penalties or Charges (Yes/No)"] = "Yes"
            person["Tax Consequences (Yes/No)"] = "Yes"
            person["Interest Earnings Considered (Yes/No)"] = "Yes"
            person["Minimum Amounts Required (Yes/No)"] = "Yes"
            person["Other Materially Adverse Effects (Yes/No)"] = "Yes"
            person["Reduced Benefits or Increased Premiums Explanation"] = "I'm ok with reduced benefits" if person["Reduced Benefits or Increased Premiums (Yes/No)"] == "Yes" else ""
            person["New Penalties or Charges Explanation"] = "The new penalties are required" if person["New Penalties or Charges (Yes/No)"] == "Yes" else ""
            person["Existing Penalties or Charges Explanation"] = "The previous penalties are no longer necessary" if person["Existing Penalties or Charges (Yes/No)"] == "Yes" else ""
            person["Tax Consequences Explanation"] = "The participant is in a new tax bracket" if person["Tax Consequences (Yes/No)"] == "Yes" else ""
            person["Interest Earnings Considered Explanation"] = "The participant has considered the interest rate" if person["Interest Earnings Considered (Yes/No)"] == "Yes" else ""
            person["Minimum Amounts Required Explanation"] = "A minimum of x dollars is required" if person["Minimum Amounts Required (Yes/No)"] else ""
            person["Other Materially Adverse Effects Explanation"] = "There will be capital degradation" if person["Other Materially Adverse Effects (Yes/No)"] else ""
        if opts["owner_state"] == "AR":
            person["Life Insurance Product Name"] = random.choice(["Accumulation Indexed Universal Life", "Accumulation Variable Universal Life", "Accumulation Whole Life"])
            person["Current Premium"] = f"${random.randint(10000,50000):,}"
            person["Surrender Value"] = f"${random.randint(10000,50000):,}"
            person["Current Interest Rate"] = f"{random.randint(1,100)}"
            person["Guarantee Period"] = f"{random.randint(1, 12)}"
            person["Guaranteed Minimum Interest Rate"] = f"{random.randint(1,100)}"
            person["Surrender Charge Period"] = f"{random.randint(1, 12)}"
            person["Surrender Charge Percentage Per Year"] = f"{random.randint(1,100)}"
            person["Years Remaining"] = f"{random.randint(1, 12)}"
            person["Free Withdrawals (Yes/No)"] = random.choice(["Yes", "No"])
            person["Percentage of Free Withdrawals"] = f"{random.randint(1,100)}"
            person["Significant Provisions (Yes/No)"] = random.choice(["Yes", "No"])
            person["Significant Provisions Details"] = "The provisions will be significant" if person["Significant Provisions (Yes/No)"] == "Yes" else ""
            person["Significant Provisions - New Policy (Yes/No)"] = random.choice(["Yes", "No"])
            person["Significant Provisions - New Policy Details"] = "The provisions will be significant" if person["Significant Provisions - New Policy (Yes/No)"] == "Yes" else ""
            person["Death Benefit Amount"] = f"${random.randint(10000,50000):,}"
        if opts["owner_state"] in naic_states:
            person["Reason for Replacement"] = "Reason for replacement"
            person["Sales Material - Consumer Brochure (Yes/No)"] = random.choice(["Yes", "No"])
            person["Sales Material - Consumer Kit (Yes/No)"] = random.choice(["Yes", "No"])
            person["Sales Material - Other (Yes/No)"] = random.choice(["Yes", "No"])
            person["Sales Material - Other (specify)"] = "Other sales material specification" if person["Sales Material - Other (Yes/No)"] == "Yes" else ""
            person["Financing (Yes/No)"] = random.choice(["Yes", "No"])
        if opts["owner_state"] == "IN":
            person["Date of Issue (mm/dd/yyyy)"] = fake.date_of_birth()
            person["Type of Optional Benefits"] = "Type of optional benefits"

    if opts["include_address"]:
        person["Residence Street Address"] = fake.building_number() + " " + fake.street_name() + " " + fake.street_suffix()
        person["Residence Street Address Line 2"] = random.choice(["", random.choice(["Apt.", "Suite", "Unit"]) + " " + fake.building_number()])
        person["Residence City"] = fake.city()
        person["Residence Zip Code"] = f"{random.randint(10000, 99999)}"
    return person

def write_census(people):
    census_path = Path("output/census.xlsx")
    census_path.parent.mkdir(parents=True, exist_ok=True)
    with xlsxwriter.Workbook(census_path) as workbook:
        worksheet = workbook.add_worksheet()
        date_format = workbook.add_format({"num_format": 'mm/dd/yyyy'})
        for col, header in enumerate(people[0].keys()):
            worksheet.write(0, col, header)
        row = 1
        for person in people:
            for col, key in enumerate(person.keys()):
                if isinstance(person[key], date):
                    worksheet.write_datetime(row, col, person[key], date_format)
                else:
                    worksheet.write(row, col, person[key])
            row += 1

def write_individual_illustrations(people, opts):
    global ns, states
    root = ET.parse(f"./{opts["product_type"].lower()}-individual.xml")
    olife = root.find("ns:TXLifeRequest/ns:OLifE", ns)
    policy = olife.find("ns:Holding/ns:Policy", ns)
    life = policy.find("ns:Life", ns)
    person_node = olife.find("ns:Party[@id='Insured_1']/ns:Person", ns)
    gender = person_node.find("ns:Gender", ns)

    coverages = root.findall(".//ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:Life/ns:Coverage", ns)
    base_cov = life.find(".//ns:Coverage/ns:IndicatorCode[@tc='1']/..", ns)
    root.find("ns:TXLifeRequest/ns:OLifE/ns:Party[@id='Producer_1']/ns:Person/ns:FirstName", ns).text = opts["fp_fn"]
    annualPremiumAmountNode = ET.SubElement(base_cov, "TotAnnualPremAmt")
    totalAnnualPremiumAmount = 0

    for i, person in enumerate(people):
        person_node.find("ns:FirstName", ns).text = person["First Name"]
        person_node.find("ns:LastName", ns).text = person["Last Name"]
        gender.set("tc", "1" if person["Gender (M/F)"] == "M" else "2")
        gender.text = person["Gender (M/F)"]
        person_node.find("ns:BirthDate", ns).text = str(person["Date of Birth (mm/dd/yyyy)"])
        root.find("ns:TXLifeRequest/ns:OLifE/ns:Party[@id='Producer_1']/ns:Person/ns:FirstName", ns).text = opts["fp_fn"]
        root.find("ns:TXLifeRequest/ns:OLifE/ns:Party[@id='Producer_1']/ns:Person/ns:LastName", ns).text = opts["fp_ln"]
        annualPremiumAmount = round(random.random() * 10000.0 + 5000.0, 2)
        totalAnnualPremiumAmount += annualPremiumAmount
        annualPremiumAmountNode.text = f"{annualPremiumAmount}"
        root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:OLifEExtension/ns:PolicyExtension/ns:PolicyUWAmount", ns).text = f"{round(500000 + random.random() * 1000000, 2)}"
        if opts["plan_type"] == "COLI":
            app_jurisdiction = root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:ApplicationInfo/ns:ApplicationJurisdiction", ns)
            app_jurisdiction.text = opts["owner_state"]
            app_jurisdiction.set("tc", str(states[opts["owner_state"]]))
        if opts["has_pdf"]:
            parent = root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:Life", ns)
            child = ET.SubElement(parent, "PremiumDepositFundAmt")
            child.text = f"{random.randint(10000, 99999)}"
        if opts["replacements"]:
            ancestor = root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:OLifEExtension/ns:PolicyExtension/ns:StateReplacementInfo", ns)
            if opts["owner_state"] == "WA":
                ancestor.find("ns:WA_After", ns).text = "X"
                ancestor.find("ns:WA_HowLong", ns).text = "N/A"
                ancestor.find("ns:WA_HowMuch", ns).text = "N/A"
                ancestor.find("ns:WA_Rate", ns).text = "5.30%"
                ancestor.find("ns:WA_Limit", ns).text = "90-95% CV"
                ancestor.find("ns:WA_SurrChg", ns).text = "decreasing over 10 years."
                ancestor.find("ns:WA_DeathBen", ns).text = "level, starting at $500,000"
            if opts["owner_state"] == "WY":
                ancestor.find("ns:WY_SurrCharges", ns).text = "100,100,98,95,89,78,65,51,34"
            if opts["owner_state"] == "AR":
                ancestor.find("ns:AR_Insurer", ns).text = random.choice(["Liberty", "Penn Insurance and Annuity"])
                ancestor.find("ns:AR_ProductType", ns).text = random.choice(["Indexed Life Insurance", "Variable Life Insurance"])
                ancestor.find("ns:AR_ProductName", ns).text = random.choice(["Accumulation VUL", "Accumulation IUL"])
                ancestor.find("ns:AR_ProposedPremium", ns).text = f"${random.randint(1,9)},{random.randint(100,999)}.{random.randint(0,99)} Ann"
                ancestor.find("ns:AR_AnnualConsideration", ns).text = f"${random.randint(1,9)},{random.randint(100,999)}.{random.randint(0,99)}  Annually"
                ancestor.find("ns:AR_CurrentCV", ns).text = f"${random.randint(1,9)},{random.randint(100,999)} end of year 1"
                ancestor.find("ns:AR_CurrentSV", ns).text = f"${random.randint(1,9)},{random.randint(100,999)} end of year 1"
                ancestor.find("ns:AR_Face", ns).text = f"${random.randint(1,9)}00,000"
                ancestor.find("ns:AR_CurrInt", ns).text = "N/A"
                ancestor.find("ns:AR_GuarMinAccum", ns).text = f"{random.randint(1,9)}.00%"
                ancestor.find("ns:AR_SurrChgPeriod", ns).text = f"{random.randint(1,9)}  / 100% decr. {random.randint(0,100)}% per yr."
                ancestor.find("ns:AR_FreeWithdrawals", ns).text = "N/A"
        ttfInd = root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:Life/ns:LifeUSA/ns:Exchange1035Ind", ns)
        ttfInd.text = "False"
        ttfInd.set("tc", "0")
        root.write(f"output/life{i + 1}.xml")
    return (totalAnnualPremiumAmount, coverages)

def write_composite_illustration(totalAnnualPremiumAmount, coverages, opts, people):
    global ns

    composite_root = ET.parse("composite-template.xml")
    cov_parent = composite_root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:Life", ns)
    for cov in coverages:
        cov_parent.append(cov)
    base_cov = cov_parent.find("ns:Coverage", ns)
    ET.SubElement(base_cov, "TotAnnualPremAmt").text = f"{totalAnnualPremiumAmount}"
    ET.SubElement(base_cov, "NumLives").text = f"{opts["count"]}"
    if opts["plan_type"] == "COLI":
        owner = composite_root.find("ns:TXLifeRequest/ns:OLifE/ns:Party[@id='Owner_1']", ns) 
        owner.find("ns:FullName", ns).text = opts["owner_name"]
        otype = owner.find("ns:OLifEExtension/ns:PartyExtension/ns:OwnerType", ns)
        otype.text = opts["owner_type"]
        otype.set("tc", owner_types[opts["owner_type"]])
    plan_type = composite_root.find("ns:TXLifeRequest/ns:OLifE/ns:Holding/ns:Policy/ns:GroupPolicy/ns:OLifEExtension/ns:GroupPlanType", ns)
    plan_type.text = opts["plan_type"]
    plan_type.set("tc", "1" if opts["plan_type"] == "COLI" else "2")
    elt = composite_root.findall("ns:TXLifeRequest/ns:OLifE", ns)[0]
    if opts["plan_type"] == "CSIO":
        planTypeElt = elt.findall("ns:Holding/ns:Policy/ns:GroupPolicy/ns:OLifEExtension/ns:GroupPlanType", ns)[0]
        planTypeElt.text = "CSIO"
        planTypeElt.set("tc", "2")
        ownerElt = elt.findall("ns:Party[@id='Owner_1']", ns)[0]
        elt.remove(ownerElt)
    for i, person in enumerate(people):
        subelt = ET.SubElement(elt, "Party", { "id": f"Insured_{i + 1}" })
        subelt = ET.SubElement(subelt, "Person")
        ET.SubElement(subelt, "FirstName").text = person["First Name"]
        ET.SubElement(subelt, "LastName").text = person["Last Name"]
        ET.SubElement(subelt, "Gender", { "tc": "1" if person["Gender (M/F)"] == "M" else "2" }).text = person["Gender (M/F)"]
        ET.SubElement(subelt, "BirthDate").text = str(person["Date of Birth (mm/dd/yyyy)"])
        if opts["plan_type"] == "CSIO":
            ET.SubElement(subelt, "ResidenceState").text = person["Residence State"]
            subelt = ET.SubElement(elt, "Party", { "id": f"Owner_{i + 1}" })
            subelt = ET.SubElement(subelt, "OLifEExtension")
            subelt = ET.SubElement(subelt, "PartyExtension")
            ET.SubElement(subelt, "OwnerType", { "tc": "3" }).text = "Insured"
    composite_root.write(f"output/composite.xml")

def zip_illustrations():
    illustrations = glob.glob("./output/*.xml")
    zip_path = Path("output/Archive.zip")
    with ZipFile(zip_path, "w") as zip_file:
        for illustration in illustrations:
            zip_file.write(illustration)

def main():
    opts = get_options()
    people = [generate_person(opts) for _ in range(opts["count"])]
    write_census(people)
    totalAnnualPremiumAmount, coverages = write_individual_illustrations(people, opts)
    write_composite_illustration(totalAnnualPremiumAmount, coverages, opts, people)
    zip_illustrations()

if __name__ == "__main__":
    main()
