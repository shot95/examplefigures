import pandas as pd


class ExcelParser:
    def __init__(self, file_path):
        """
        Initializes the ExcelParser object with the path to the Excel file.

        :param file_path: str, path to the Excel file
        """
        self.file_path = file_path
        self.data = None
        self.sheets = None

    def load_excel(self):
        """
        Loads the Excel file and stores sheet names.
        """
        try:
            # Read the Excel file into a pandas ExcelFile object
            self.sheets = pd.ExcelFile(self.file_path)
            print(f"Sheets in the Excel file: {self.sheets.sheet_names}")
        except Exception as e:
            print(f"Error loading Excel file: {e}")

    def load_sheet(self, sheet_name):
        """
        Loads a specific sheet into a DataFrame.

        :param sheet_name: str, name of the sheet to load
        :return: DataFrame or None if the sheet is not found
        """
        if self.sheets:
            try:
                self.data = self.sheets.parse(sheet_name,index_col = 0)
                print(f"Loaded sheet: {sheet_name}")
                return self.data
            except ValueError:
                print(f"Sheet '{sheet_name}' not found.")
                return None
        else:
            print("Excel file not loaded. Call load_excel() first.")
            return None

    def get_data(self):
        """
        Returns the data from the last loaded sheet.

        :return: DataFrame or None if no sheet has been loaded
        """
        if self.data is not None:
            return self.data
        else:
            print("No sheet data available. Load a sheet first.")
            return None

    def show_sheet_info(self):
        """
        Displays information about the currently loaded sheet.
        """
        if self.data is not None:
            print(f"Data shape: {self.data.shape}")
            print(f"Columns: {self.data.columns}")
            print(f"First 5 rows:\n{self.data.head()}")
        else:
            print("No sheet loaded. Please load a sheet first.")

    def save_to_csv(self, output_file):
        """
        Saves the current DataFrame to a CSV file.

        :param output_file: str, path to the output CSV file
        """
        if self.data is not None:
            try:
                self.data.to_csv(output_file, index=False)
                print(f"Data saved to {output_file}")
            except Exception as e:
                print(f"Error saving to CSV: {e}")
        else:
            print("No data to save. Load a sheet first.")


# Example usage:
if __name__ == "__main__":
    # Create an instance of the ExcelParser
    parser = ExcelParser('testdata.xlsx')

    # Load the Excel file
    parser.load_excel()

    # Load a specific sheet
    data = parser.load_sheet('Tabelle1')

    # Show basic sheet info
    parser.show_sheet_info()

    # Save the data to a CSV file
    #parser.save_to_csv('output.csv')