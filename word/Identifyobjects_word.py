from docx import Document

# Predefined file path
file_path = r'C:\Users\e057495\OneDrive - Mastercard\Desktop\Dummy.docx'

# Open the Word document
document = Document(file_path)

# List to store the attachments
attachments = []

# Iterate over the document parts
for part in document.part.rels.values():
    if "embed" in part.reltype:
        attachment_name = part.target_ref
        if attachment_name not in attachments:
            attachments.append(attachment_name)

# Print the list of attachments
for attachment in attachments:
    print(attachment)
