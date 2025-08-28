from datetime import datetime
from collections import defaultdict
import calendar
from dateutil.rrule import rrule, MONTHLY
import openpyxl
import pandas as pd
from data_importer import FirestoreDataImporter

class ExcelReportGenerator:
    """
    Generates an Excel attendance report by fetching data from Firestore.
    The logic is a Python port of the Dart ReportManager.
    """
    def __init__(self, db_client):
        if not db_client:
            raise ValueError("A valid Firestore database client is required.")
        self.db = db_client


    def _get_months_between(self, start_date: datetime, end_date: datetime) -> dict[str, datetime]:
        """Generates a dictionary of months between two dates."""
        months = {}
        # rrule generates dates for each month between the start and end
        for dt in rrule(MONTHLY, dtstart=start_date, until=end_date):
            month_key = dt.strftime("%B-%Y") # Format as 'MONTH_NAME-YYYY'
            months[month_key] = dt
        return months

    def _get_days_for_month(self, month_date: datetime) -> list[datetime]:
        """Gets all days for a specific month."""
        _, num_days = calendar.monthrange(month_date.year, month_date.month)
        return [datetime(month_date.year, month_date.month, day) for day in range(1, num_days + 1)]

    def _generate_excel_with_pandas(self, df: pd.DataFrame, start_date: datetime, end_date: datetime, output_path: str):
        """
        Generates an Excel report from a DataFrame using pandas.
        This approach is more efficient for data manipulation and produces cleaner code.
        """
        months = self._get_months_between(start_date, end_date)
        if not months:
            raise ValueError("Date range does not cover any full months.")

        # Get unique student information
        student_info = df[['lrn', 'lastName', 'firstName', 'studentYear', 'studentSection']].drop_duplicates(subset='lrn')
        student_info = student_info.set_index('lrn')

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for month_key, month_date in months.items():
                month_df = df[(df['timestamp'].dt.year == month_date.year) & (df['timestamp'].dt.month == month_date.month)].copy()
                
                weekdays_in_month = [d for d in self._get_days_for_month(month_date) if d.weekday() < 5]
                weekday_days = [d.day for d in weekdays_in_month]

                if not weekdays_in_month:
                    pd.DataFrame(columns=["No weekdays in this month."]).to_excel(writer, sheet_name=month_key, index=False)
                    continue

                # Create the main structure of the report for this month with all students
                month_report = student_info.copy()

                # Add columns for each weekday, initialized to empty string
                for day in weekday_days:
                    month_report[day] = ''

                if not month_df.empty:
                    month_df['day'] = month_df['timestamp'].dt.day
                    # Get only absent records
                    absences = month_df[month_df['isAbsent'] == True]
                    # Create a pivot for absences
                    if not absences.empty:
                        absence_pivot = absences.pivot_table(index='lrn', columns='day', values='isAbsent', aggfunc='any')
                        
                        # Map absences ('A') to the report
                        for lrn, row in absence_pivot.iterrows():
                            if lrn in month_report.index:
                                for day, is_absent in row.items():
                                    if is_absent and day in weekday_days:
                                        month_report.loc[lrn, day] = 'A'

                # Calculate totals
                month_report['Total Absences'] = (month_report[weekday_days] == 'A').sum(axis=1)
                month_report['Total Present'] = len(weekday_days) - month_report['Total Absences']
                
                month_report.reset_index(inplace=True)
                
                # Reorder columns to match header
                month_report = month_report[['lrn', 'lastName', 'firstName', 'studentYear', 'studentSection'] + weekday_days + ['Total Absences', 'Total Present']]

                # Write the DataFrame data without the header
                month_report.to_excel(writer, sheet_name=month_key, index=False, header=False, startrow=2)

                # Now, write the headers manually using openpyxl
                sheet = writer.sheets[month_key]
                
                prefix_cols = ['LRN', 'Last Name', 'First Name', 'Year', 'Section']
                day_letters = [d.strftime('%a') for d in weekdays_in_month]
                day_numbers = [str(d.day) for d in weekdays_in_month]
                suffix_cols = ['Total Absences', 'Total Present']

                header_top = [''] * len(prefix_cols) + day_letters + [''] * len(suffix_cols)
                header_bottom = prefix_cols + day_numbers + suffix_cols

                # Write the top header row (day letters)
                for col_num, cell_value in enumerate(header_top, 1):
                    sheet.cell(row=1, column=col_num, value=cell_value)

                # Write the bottom header row (main headers)
                for col_num, cell_value in enumerate(header_bottom, 1):
                    sheet.cell(row=2, column=col_num, value=cell_value)

    def generate_report(self, start_date: datetime, end_date: datetime, output_path: str, section: str | None = None):
        """
        Main method to generate the complete Excel report.
        Orchestrates fetching data and writing the file.
        """
        try:
            print("Fetching student records...")
            records = FirestoreDataImporter(self.db).import_data(start_date, end_date, section)

            if not records:
                print("No records found for the given date range and section.")
                return

            # Convert records to a pandas DataFrame
            records_list = [doc.to_dict() for doc in records]
            df = pd.DataFrame(records_list)
            df['timestamp'] = pd.to_datetime(df['timestamp'])


            # The DataFrame can now be used for statistics, for example:
            stats_generator = StatisticsGenerator(df)
            stats_generator.generate_statistics()

            print("Generating Excel report using pandas...")
            self._generate_excel_with_pandas(df, start_date, end_date, output_path)
            print("Report generated successfully.")


        except Exception as e:
            print(f"Error generating report: {e}")
            raise



class StatisticsGenerator:
    def __init__(self, df: pd.DataFrame):
        if df.empty:
            raise ValueError("DataFrame cannot be empty.")
        self.df = df

    def generate_statistics(self):
        """
        Generates and prints basic attendance statistics.
        """
        print("--- Attendance Statistics ---")
        
        # Ensure 'isAbsent' column exists and is boolean
        if 'isAbsent' not in self.df.columns:
            print("'isAbsent' column not found. Cannot generate statistics.")
            return
            
        # Total records
        total_records = len(self.df)
        print(f"Total attendance records: {total_records}")

        # Absence/Presence count
        attendance_counts = self.df['isAbsent'].value_counts()
        absences = attendance_counts.get(True, 0)
        presents = attendance_counts.get(False, 0)
        
        print(f"Total absences: {absences}")
        print(f"Total presents: {presents}")

        # Absences by year level
        if 'studentYear' in self.df.columns:
            absences_by_year = self.df[self.df['isAbsent'] == True]['studentYear'].value_counts()
            print("\nAbsences by Year Level:")
            print(absences_by_year)

        # Absences by section
        if 'studentSection' in self.df.columns:
            absences_by_section = self.df[self.df['isAbsent'] == True]['studentSection'].value_counts()
            print("\nAbsences by Section:")
            print(absences_by_section)
            
        print("--- End of Statistics ---")
        

        