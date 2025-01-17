import json
import pandas as pd
import os
from datetime import datetime
from typing import Dict, List, Optional

class PortfolioDataValidator:
    
    @staticmethod
    def validate_month_year(month_year: str) -> bool:
        try:
            datetime.strptime(month_year, '%B %Y')
            return True
        except ValueError:
            return False
    
    @staticmethod
    def validate_excel_structure(df: pd.DataFrame) -> bool:
        required_columns = {'Name', 'ISIN', 'Industry', 'Quantity', 'MarketValue', 'PercentNAV'}
        return all(col in df.columns for col in required_columns)

class PortfolioManager:
    def __init__(self, storage_file: str = 'portfolio_data.json'):
        
        self.storage_file = storage_file
        self.data = self._load_data()
        self.validator = PortfolioDataValidator()

    def _load_data(self) -> Dict:
        if os.path.exists(self.storage_file):
            try:
                with open(self.storage_file, 'r', encoding='utf-8') as file:
                    return json.load(file)
            except json.JSONDecodeError:
                print(f"Warning: Could not decode {self.storage_file}. Starting with empty data.")
                return {}
            except Exception as e:
                print(f"Warning: Error reading {self.storage_file}: {str(e)}. Starting with empty data.")
                return {}
        return {}

    def _save_data(self) -> bool:
        try:
            with open(self.storage_file, 'w', encoding='utf-8') as file:
                json.dump(self.data, file, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving data: {str(e)}")
            return False

    def process_excel_data(self, excel_file: str, month_year: str) -> bool:
       
        try:
            if not self.validator.validate_month_year(month_year):
                raise ValueError("Invalid month-year format. Use 'Month YYYY' format (e.g., 'January 2024')")

            if not os.path.exists(excel_file):
                raise FileNotFoundError(f"Excel file not found: {excel_file}")

            df = pd.read_excel(excel_file)
            
            portfolio_data = df.iloc[6:, [2, 3, 4, 5, 6, 7]].copy()
            portfolio_data.columns = ['Name', 'ISIN', 'Industry', 'Quantity', 'MarketValue', 'PercentNAV']
            
            if not self.validator.validate_excel_structure(portfolio_data):
                raise ValueError("Excel file structure is invalid")

            portfolio_data = portfolio_data.dropna(how='all')
            
            for col in ['Quantity', 'MarketValue', 'PercentNAV']:
                portfolio_data[col] = pd.to_numeric(portfolio_data[col], errors='coerce')

            month_data = []
            for _, row in portfolio_data.iterrows():
                if pd.notna(row['ISIN']) and str(row['ISIN']).strip():
                    entry = {
                        "MutualFundDetails": {
                            "Name": str(row['Name']).strip(),
                            "ISIN": str(row['ISIN']).strip(),
                            "Industry": str(row['Industry']).strip()
                        },
                        "MonthData": {
                            "Quantity": float(row['Quantity']) if pd.notna(row['Quantity']) else 0,
                            "MarketValueInLakhs": float(row['MarketValue']) if pd.notna(row['MarketValue']) else 0,
                            "%ToNAV": float(row['PercentNAV']) if pd.notna(row['PercentNAV']) else 0
                        }
                    }
                    month_data.append(entry)
            
            if not month_data:
                raise ValueError("No valid data found in Excel file")

            self.data[month_year] = month_data
            if self._save_data():
                print(f"\nSuccessfully processed data for {month_year}")
                return True
            return False
            
        except Exception as e:
            print(f"\nError processing Excel file: {str(e)}")
            return False

    def search_and_calculate_changes(self, fund_name: str, start_month: str, end_month: str) -> Optional[List[Dict]]:
        
        try:
            if not all(self.validator.validate_month_year(month) for month in [start_month, end_month]):
                raise ValueError("Invalid month format. Use 'Month YYYY' format")

            if start_month not in self.data or end_month not in self.data:
                raise KeyError("One or both months not found in the data")

            results = []
            start_data = {item['MutualFundDetails']['ISIN']: item for item in self.data[start_month]}
            end_data = {item['MutualFundDetails']['ISIN']: item for item in self.data[end_month]}

            fund_name = fund_name.lower()
            for isin, end_item in end_data.items():
                if fund_name in end_item['MutualFundDetails']['Name'].lower():
                    result = {
                        "FundDetails": end_item['MutualFundDetails'].copy(),
                        "Changes": {},
                        "Status": "Existing"
                    }

                    start_item = start_data.get(isin)
                    if not start_item:
                        result["Status"] = "New Addition"
                        results.append(result)
                        continue

                    for key in ['Quantity', 'MarketValueInLakhs', '%ToNAV']:
                        start_value = start_item['MonthData'][key]
                        end_value = end_item['MonthData'][key]
                        
                        result["Changes"][key] = {
                            "StartValue": start_value,
                            "EndValue": end_value,
                            "PercentageChange": (
                                ((end_value - start_value) / start_value * 100)
                                if start_value != 0 else None
                            )
                        }
                    
                    results.append(result)

            self._print_analysis_results(results, start_month, end_month)
            return results

        except Exception as e:
            print(f"\nError: {str(e)}")
            return None

    def _print_analysis_results(self, results: List[Dict], start_month: str, end_month: str) -> None:
        """Print formatted analysis results"""
        print(f"\nAnalysis Results ({start_month} to {end_month}):")
        print("=" * 80)

        for result in results:
            print(f"\nFund Name: {result['FundDetails']['Name']}")
            print(f"ISIN: {result['FundDetails']['ISIN']}")
            print(f"Industry: {result['FundDetails']['Industry']}")
            print(f"Status: {result['Status']}")
            
            if result['Status'] == "New Addition":
                print("This is a new fund addition - no change calculations available")
                print("-" * 40)
                continue

            print("\nChanges:")
            for metric, values in result['Changes'].items():
                print(f"{metric}:")
                print(f"  Start: {values['StartValue']:,.2f}")
                print(f"  End: {values['EndValue']:,.2f}")
                if values['PercentageChange'] is not None:
                    print(f"  Change: {values['PercentageChange']:+,.2f}%")
                else:
                    print("  Change: N/A (starting value was 0)")
            print("-" * 40)

def main():
    manager = PortfolioManager()
    
    while True:
        print("\nPortfolio Manager Menu:")
        print("=" * 50)
        print("1. Import Excel Data")
        print("2. Analyze Fund Performance")
        print("3. Exit")
        
        try:
            choice = input("\nEnter your choice (1-3): ").strip()
            
            if choice == '1':
                excel_file = input("\nEnter Excel file path: ").strip()
                month_year = input("Enter month and year (e.g., 'September 2024'): ").strip()
                
                if manager.process_excel_data(excel_file, month_year):
                    print("\nData imported successfully!")
                else:
                    print("\nFailed to import data.")
            
            elif choice == '2':
                if not manager.data:
                    print("\nNo data available. Please import data first.")
                    continue
                
                print("\nAvailable months:")
                for month in sorted(manager.data.keys()):
                    print(f"- {month}")
                
                fund_name = input("\nEnter fund name to search: ").strip()
                start_month = input("Enter start month (e.g., 'September 2024'): ").strip()
                end_month = input("Enter end month (e.g., 'October 2024'): ").strip()
                
                manager.search_and_calculate_changes(fund_name, start_month, end_month)
            
            elif choice == '3':
                print("\nThank you for using Portfolio Manager!")
                break
            
            else:
                print("\nInvalid choice. Please enter 1, 2, or 3.")
                
        except KeyboardInterrupt:
            print("\n\nProgram interrupted by user.")
            break
        except Exception as e:
            print(f"\nAn unexpected error occurred: {str(e)}")
            print("Please try again.")

if __name__ == "__main__":
    main()