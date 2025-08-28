

class AbstractReportGenerator:
    def __init__(self, excel_path: str, output_path: str):
        self.excel_path = excel_path
        self.output_path = output_path
    
    def generate_report(self):
        raise NotImplementedError


class ExcelReportGenerator(AbstractReportGenerator):
    def __init__(self, excel_path: str, output_path: str):
        self.excel_path = excel_path
        self.output_path = output_path
    
    def generate_report(self):
        pass

        