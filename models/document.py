from PySide6.QtCore import QDate

class Document:
    def __init__(self, number="", name="", counterparty="", 
                 start_date=QDate.currentDate(), end_date=None,
                 description="", attachments=None):
        self.number = number
        self.name = name
        self.counterparty = counterparty
        self.start_date = start_date
        self.end_date = end_date
        self.description = description
        self.attachments = attachments or []

    def is_document_expiring(self, days_threshold=30):
        if not self.end_date:
            return False
        return self.end_date.daysTo(QDate.currentDate()) <= days_threshold

    def to_dict(self):
        return {
            "number": self.number,
            "name": self.name,
            "counterparty": self.counterparty,
            "start_date": self.start_date.toString("dd.MM.yyyy"),
            "end_date": self.end_date.toString("dd.MM.yyyy") if self.end_date else "",
            "description": self.description,
            "attachments": ", ".join(self.attachments)
        }