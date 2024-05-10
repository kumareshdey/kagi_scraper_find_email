import re

def extract_emails(texts, name, number):
    email_list = []
    # Convert name and number to lowercase for case-insensitive matching
    name = name.lower()
    # Construct a case-insensitive regular expression pattern to match email addresses
    email_pattern = r"\b[A-Za-z0-9._%+-]+@(?!.*\.edu)[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    
    for text in texts:
        text = text.lower()
        if name in text and number in text:
            emails = re.findall(email_pattern, text)
            if emails:
                # Add the email address to the list
                email_list.extend(emails)
        # Return a list of unique email addresses
    return list(set(email_list))

texts = ["Jul 5, 2022 Email. Davis@saadiangroup.ed. Davis Saadian | CA DRE #01921748. Carolwood Partners | DRE# 02200006. All information is deemed reliable but not"]
name = "DAVIS SAADIAN"
number = "01921748"

print(extract_emails(texts, name, number))
