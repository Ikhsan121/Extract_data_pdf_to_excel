import re
import pdfplumber
from create_excel import create_xlsx

# Function to extract data from a given page of the PDF
def extract_page_data(page_lines):
    page_data = {}
    emails = []
    complete_text = ' '.join(page_lines)

    # Extract Notice Address
    notice_address = extract_notice_address(complete_text)
    print(notice_address)
    if notice_address:
        page_data["notice_address"] = notice_address
        address_components = extract_address_components(notice_address)
        print("address component: ", address_components)
        page_data.update(address_components)

    # Extract other page-specific data
    for line in page_lines:
        extract_strata_plan_and_lot(line, page_data)
        extract_lot_number(line, page_data)
        extract_unit_no(line, page_data)
        extract_levy_entitlement(line, page_data)
        extract_owner_info(line, page_data)
        extract_levy_address(line, page_data)
        extract_purchase_and_entry_dates(line, page_data)
        extract_emails(line, emails, page_data)
        extract_tenant_info(line, page_data)
        extract_lease_info(line, page_data)
    owner_email = ", ".join(emails)
    page_data["owner_emails"] = owner_email
    return page_data

# Helper function to extract the notice address using regex
def extract_notice_address(text):
    pattern = r"Notice Address\s+(.*?)(?=\s*(Owner Email|Levy Address|$))"
    match = re.search(pattern, text)
    if match:
        return re.sub(r"Notice Address | Date of entry \d{2}/\d{2}/\d{4}", "", match.group())
    return None

# Helper function to extract address components (street, suburb, state, postcode)
def extract_address_components(notice_address):
    # Define a regex pattern to extract street address, suburb, state (optional), and postcode
    address_pattern = r"(?P<street_address>(?:\d{1,2}/?\d*\s)?(?:[A-Za-z0-9\s&/-]+(?:\s(?:Road|Rd|Street|Ave|Avenue|Drive|Blvd|Lane|Court|Cres|St|Terrace|PO Box|Box))?))\s(?P<suburb>[A-Za-z\s]+)\s(?P<state>[A-Za-z]{2,3})?\s(?P<postcode>\d{4})"

    # Apply the regex pattern to extract address components
    match = re.search(address_pattern, notice_address)

    # Default values for missing components
    street_address = None
    suburb = None
    state = None
    postcode = None

    # If a match is found, assign the matched components
    if match:
        street_address = match.group('street_address') if match.group('street_address') else None
        suburb = match.group('suburb') if match.group('suburb') else None
        state = match.group('state') if match.group('state') else "Not Available"  # Handling missing state
        postcode = match.group('postcode') if match.group('postcode') else None

    # If state is missing, manually check for that and adjust the regex accordingly
    if not state and len(notice_address.split()) > 2:
        print("state tidak ada")
        # If the state is missing, we still try to capture street, suburb, and postcode
        missing_state_pattern = r"(?P<street_address>(?:\d{1,2}/?\d*\s)?(?:[A-Za-z0-9\s&/-]+(?:\s(?:Road|Rd|Street|Ave|Avenue|Drive|Blvd|Lane|Court|Cres|St|Terrace|PO Box|Box))?))\s(?P<suburb>[A-Za-z\s]+)\s(?P<postcode>\d{4})"
        match_missing_state = re.search(missing_state_pattern, notice_address)
        if match_missing_state:
            street_address = match_missing_state.group('street_address')
            suburb = match_missing_state.group('suburb')
            postcode = match_missing_state.group('postcode')
            state = ""  # Explicitly handle the missing state here

    # Return the extracted components or default value if no match is found
    return {
        "notice_address_street_address": street_address if street_address else "Not Available",
        "notice_address_suburb": suburb if suburb else "Not Available",
        "notice_address_state": state,
        "notice_address_postcode": postcode if postcode else "Not Available"
    }


# Helper function to extract Strata Plan and Lot number
def extract_strata_plan_and_lot(line, page_data):
    if "Strata Plan" in line:
        parts = line.split("Strata Plan")[1].strip().split(" ")
        page_data["strata_plan"] = parts[0]
        page_data['lot_street_name'] = line.split("44264773546 75")[1].strip()

# Helper function to extract Unit Number
def extract_unit_no(line, page_data):
    if "Unit no.:" in line:
        page_data["unit_no"] = line.split(" ")[6]

def extract_lot_number(line, page_data):
    if "Lot:" in line:
        page_data["lot"] = line.split(" ")[1].strip()

# Helper function to extract Levy Entitlement
def extract_levy_entitlement(line, page_data):
    try:
        levy_entitlement = re.search(r'\d{3} / \d{1,3}(?:,\d{3})*\.\d{2}', line).group()
        page_data["levy_entitlement"] = levy_entitlement
    except AttributeError:
        pass

# Helper function to extract Owner Name and Contact Information
def extract_owner_info(line, page_data):
    if "Owner Name" in line:
        page_data["owner_name"] = line.split("Owner Name")[1].strip()
    if "Contact Number" in line:
        page_data["contact_number"] = line.split("Contact Number")[1].strip()

# Helper function to extract Levy Address
def extract_levy_address(line, page_data):
    if "Levy Address" in line:
        page_data["levy_address"] = line.split("Levy Address")[1].strip()

# Helper function to extract Date of Purchase and Date of Entry
def extract_purchase_and_entry_dates(line, page_data):
    if "Date of purchase" in line:
        page_data["date_purchase"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()
    if "Date of entry" in line:
        page_data["date_entry"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()

# Helper function to extract emails
def extract_emails(line, emails, page_data):
    email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    found_emails = re.findall(email_pattern, line)
    for email in found_emails:
        if email not in emails:
            emails.append(email)

# Helper function to extract Tenant and Agent Information
def extract_tenant_info(line, page_data):
    if "Tenant name" in line:
        page_data["tenant_name"] = line.split("Tenant name")[1].strip()
    if "Tenant Contact" in line:
        page_data["tenant_contact"] = line.split("Tenant Contact")[1].strip()
    if "Agent Name" in line:
        page_data["agent_name"] = line.split("Agent Name")[1].strip()
    if "Vacant" in line:
        page_data["vacant"] = line.split("Vacant")[1].strip()

# Helper function to extract Lease Start and End Dates
def extract_lease_info(line, page_data):
    if "Lease Start Date" in line:
        page_data["lease_start_date"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()
    if "Lease End Date" in line:
        page_data["lease_end_date"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()
    if "Move in Date" in line:
        page_data["move_in_date"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()
    if "Review Date" in line:
        page_data["review_date"] = re.search(r'\d{2}/\d{2}/\d{4}', line).group()

# Load the PDF and extract data
def extract_pdf_data(pdf_path):
    pages_data = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                pages_data[page_num] = extract_page_data(lines)
    return pages_data

if __name__=="__main__":
    # Main execution
    pdf_path = "./source/Source Data.pdf"
    pdf_data = extract_pdf_data(pdf_path)
    create_xlsx(pdf_data)
    print(pdf_data)



