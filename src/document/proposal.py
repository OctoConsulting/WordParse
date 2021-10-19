from document.document_set import DocumentSet


class Proposal(DocumentSet):
    def __init__(self):
        pass

    def sort(self):
        # Sort documents chronologically (if we have dates) else by filename/type
        # Example: Amendments and Q&A come later
        # Can use RFP-specific logic here
        pass
