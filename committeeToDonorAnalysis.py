import PyPDF2
import re
import pandas as pd
import spacy
import xlsxwriter


# Open the PDF file
pdf_path = 'boardmembers.pdf'
pdf_file = open(pdf_path, 'rb')
nlp = spacy.load('en_core_web_sm')
workbook = xlsxwriter.Workbook('last_name_of_donor_and_committee_member_match.xlsx')
worksheet = workbook.add_worksheet()
#write column headers to worksheet
worksheet.write(0, 0, "Committee Member")
worksheet.write(0, 1, "Donor")
worksheet.write(0, 2, "Amount")
worksheet.write(0, 3, "Campaign")

# Create a PDF reader object
pdf_reader = PyPDF2.PdfReader(pdf_file)

# Initialize variables
names_and_nominators = {}
current_name = None
extracted_names = []


"""
    Primary way we determine if the line has the name of a committee member
"""
def isPositionHeader(line):
    position_titles = ["District", "Position", "Non-Voting"]
    for pos in position_titles:
        if pos in line:
            return True
    return False

def isVacantPosition(line):
    if "VACANT" in line:
        return True
    return False

"""
    Purpose: Ignore any paragraphs or text that include the keywords we use to identify nomination positions
    
    How it works: if there are more than 5 distinct words in the line, ignore it bc it's likely a sentence
    
    Opportunity for improvement: Use spaCy to identify if the word is a name entity instead of using a weak length indicator
"""
def isDescription(line):
    if len(line.split()) > 8:
        return True
    return False

def isSuffix(word):
    suffixList = ["jr", "Jr.", "Jr", "JR", "II", "ii", "III", "iii", "SR", "Sr", "Sr.", "sr", "PH.D"]
    if word in suffixList:
        return True
    return False

"""
    Positioning is different between the nomination titles, so this handles that edge case
"""
def extractNameFromLine(line):
    if "Non-Voting" in line:
        return line.split()[1:]
    else:
        return line.split()[2:]

# Iterate over each page of the PDF
print("Reading from committee membership document")
for page in pdf_reader.pages:
    # Extract the text from the current page
    lines = page.extract_text().split("\n")
    for line in lines:
        if isPositionHeader(line) and not isVacantPosition(line):
            if isDescription(line):
                continue
            name = " ".join(extractNameFromLine(line))
            extracted_names.append(name)

# Close the PDF file
pdf_file.close()

"""
Still not working.

Example: Demetris Sampson donated to Zarin Gracey and is on the DFW - DALLAS FORT WORTH INTERNATIONAL AIRPORT BOARD
"""

# Load the Excel file
print("Opening campaign donor worksheet")
excel_path = '2023 Dallas Campaign Donors.xlsx'
df = pd.read_excel(excel_path)
# Iterate over each donation from excel
last_name_match_row = 1
print("Comparing names of committee members to campaign donors")
for committee_member in extracted_names:
    member_last_name = committee_member.split()[-2] if isSuffix(committee_member.split()[-1]) else committee_member.split()[-1].lower()

    # Search for matching name of a nominee
    for index, row in df.iterrows():
        donor_last_name = str(row['Donor']).split(",")[0].lower()  # Convert the cell value to string
        if donor_last_name == member_last_name:
            worksheet.write(last_name_match_row, 0, committee_member)
            worksheet.write(last_name_match_row, 1, str(row['Donor']))
            worksheet.write(last_name_match_row, 2, str(row['Amount']))
            worksheet.write(last_name_match_row, 3, str(row['Candidate']))
            last_name_match_row = last_name_match_row + 1

workbook.close()
