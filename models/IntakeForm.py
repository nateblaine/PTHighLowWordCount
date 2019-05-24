class IntakeForm:
    def __init__(self, record_id, chief_complaint):
        self.record_id = record_id;
        self.chief_complaint = chief_complaint;

    def __str__(self):
        return ('Record ID: ' + str(self.record_id) + ' Chief Complaint: ' + str(self.chief_complaint))
