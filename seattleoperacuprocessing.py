import json
import logging
import sys
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast, List, Dict
from dataclasses import dataclass
import re
from collections import defaultdict

import requests
import pandas as pd

import os

## Change to the directory where this script is located
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Environment variable loading
try:
    from dotenv import load_dotenv

    load_dotenv('.\.env', override=True)
except ImportError:
    print("Note: python-dotenv not installed. Using existing environment variables.")


class SeattleOperaTableConverter:
    """
    A class for converting Seattle Opera JSON data to Excel tables organized by year.
    """
    
    def __init__(self):
        self.fieldnames = ['SHOW', 'DATES', 'ROLE', 'ARTIST', 'OTHER', 'FILENAME']
    
    def load_json_data(self, file_path: str) -> Dict[str, Any]:
        """Load JSON data from file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
            raise
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON in file '{file_path}': {e}")
            raise

    def extract_table_data(self, json_data: Dict[str, Any], filename: str = '') -> List[Dict[str, str]]:
        """
        Extract table data from JSON structure.
        Returns a list of dictionaries with keys: SHOW, DATES, ROLE, ARTIST, OTHER, FILENAME
        """
        table_data = []
        
        # Navigate through the JSON structure
        if 'result' in json_data and 'contents' in json_data['result']:
            contents = json_data['result']['contents']
            
            for content in contents:
                if 'fields' in content:
                    fields = content['fields']
                    
                    # Extract show name
                    show = ""
                    if 'SHOW' in fields and 'valueString' in fields['SHOW']:
                        show = fields['SHOW']['valueString']
                    
                    # Extract dates
                    dates = ""
                    if 'DATES' in fields and 'valueString' in fields['DATES']:
                        dates = fields['DATES']['valueString']
                    elif 'DATE' in fields and 'valueDate' in fields['DATE']:
                        dates = fields['DATE']['valueDate']
                    
                    # Extract roles and artists
                    if 'ROLES' in fields and 'valueArray' in fields['ROLES']:
                        for role_entry in fields['ROLES']['valueArray']:
                            if 'valueObject' in role_entry:
                                role_obj = role_entry['valueObject']
                                
                                role = ""
                                artist = ""
                                other = ""
                                
                                if 'ROLE' in role_obj and 'valueString' in role_obj['ROLE']:
                                    role = role_obj['ROLE']['valueString']
                                
                                if 'ARTIST' in role_obj and 'valueString' in role_obj['ARTIST']:
                                    artist = role_obj['ARTIST']['valueString']
                                
                                if 'OTHER' in role_obj and 'valueString' in role_obj['OTHER']:
                                    other = role_obj['OTHER']['valueString']
                                
                                # Add row to table data
                                if role and artist:  # Only add if both role and artist exist
                                    table_data.append({
                                        'SHOW': show,
                                        'DATES': dates,
                                        'ROLE': role,
                                        'ARTIST': artist,
                                        'OTHER': other,
                                        'FILENAME': filename
                                    })
        
        return table_data

    def extract_year_from_dates(self, dates_str: str) -> str:
        """Extract year from dates string using regex."""
        if not dates_str:
            return "Unknown"
        
        # Look for 4-digit year patterns
        year_match = re.search(r'\b(19|20)\d{2}\b', dates_str)
        if year_match:
            return year_match.group(0)
        
        # If no 4-digit year found, try to extract from date ranges like "1980-81"
        range_match = re.search(r'\b(19|20)\d{2}[-–]\d{2}\b', dates_str)
        if range_match:
            # Extract the starting year from ranges like "1980-81"
            return range_match.group(0).split('-')[0].split('–')[0]
        
        return "Unknown"

    def organize_data_by_year(self, table_data: List[Dict[str, str]]) -> Dict[str, List[Dict[str, str]]]:
        """Organize table data by year extracted from DATES column."""
        data_by_year = defaultdict(list)
        
        for row in table_data:
            year = self.extract_year_from_dates(row['DATES'])
            data_by_year[year].append(row)
        
        return dict(data_by_year)

    def process_json_files(self, json_files: List[str]) -> tuple[Dict[str, List[Dict[str, str]]], List[str]]:
        """Process multiple JSON files and organize data by year. Returns data and list of successfully processed files."""
        all_data_by_year = defaultdict(list)
        processed_files = []
        
        for json_file in json_files:
            try:
                print(f"Processing: {json_file}")
                json_data = self.load_json_data(json_file)
                # Extract just the filename from the full path
                filename = Path(json_file).name
                table_data = self.extract_table_data(json_data, filename)
                
                if table_data:
                    # Organize this file's data by year
                    year_data = self.organize_data_by_year(table_data)
                    
                    # Merge with overall data
                    for year, data in year_data.items():
                        all_data_by_year[year].extend(data)
                    
                    processed_files.append(json_file)
                    print(f"  Extracted {len(table_data)} rows")
                else:
                    print(f"  No valid data found in {json_file}")
                    
            except Exception as e:
                print(f"  Error processing {json_file}: {e}")
                continue
        
        return dict(all_data_by_year), processed_files

    def save_to_excel_by_year(self, data_by_year: Dict[str, List[Dict[str, str]]], output_path: str):
        """Save data to Excel file with separate sheets for each year."""
        if not data_by_year or all(not data for data in data_by_year.values()):
            print("No data to save to Excel.")
            return
            
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                total_rows = 0
                sheets_created = 0
                
                # Sort years for consistent sheet order
                sorted_years = sorted(data_by_year.keys(), key=lambda x: x if x != "Unknown" else "0000")
                
                for year in sorted_years:
                    data = data_by_year[year]
                    if not data:
                        continue
                    
                    # Create DataFrame
                    df = pd.DataFrame(data)
                    
                    # Ensure columns are in the correct order
                    df = df[self.fieldnames]
                    
                    # Create sheet name (Excel sheet names have limitations)
                    sheet_name = f"Year_{year}" if year != "Unknown" else "Unknown_Year"
                    sheet_name = sheet_name[:31]  # Excel sheet name limit
                    
                    # Save to sheet
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    total_rows += len(data)
                    sheets_created += 1
                    print(f"  Sheet '{sheet_name}': {len(data)} rows")
                
                if sheets_created > 0:
                    # Create a summary sheet
                    self._create_summary_sheet(writer, data_by_year, sorted_years)
                    print(f"Data successfully saved to Excel: {output_path}")
                    print(f"Total rows: {total_rows}, Sheets: {sheets_created + 1} (including Summary)")
                else:
                    print("No valid data found to create Excel sheets.")
                
        except ImportError:
            print("Error: openpyxl library not found. Please install it with: pip install openpyxl")
            raise

    def _create_summary_sheet(self, writer, data_by_year: Dict[str, List[Dict[str, str]]], sorted_years: List[str]):
        """Create a summary sheet with statistics by year."""
        summary_data = []
        
        for year in sorted_years:
            data = data_by_year[year]
            if not data:
                continue
                
            # Get unique shows for this year
            shows = set(row['SHOW'] for row in data)
            
            summary_data.append({
                'Year': year,
                'Shows': ', '.join(sorted(shows)),
                'Total_Roles': len(data),
                'Unique_Shows': len(shows),
                'Sheet_Name': f"Year_{year}" if year != "Unknown" else "Unknown_Year"
            })
        
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            print(f"  Sheet 'Summary': Overview of all years")

    def move_processed_json_files(self, processed_files: List[str], base_folder: str = "./curesults"):
        """Move processed JSON files to a processed subfolder."""
        if not processed_files:
            return
        
        processed_folder = Path(base_folder) / "processed"
        processed_folder.mkdir(exist_ok=True)
        
        moved_count = 0
        for json_file in processed_files:
            try:
                source_path = Path(json_file)
                if source_path.exists():
                    # Create destination path with timestamp prefix to avoid conflicts
                    import datetime
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    destination_path = processed_folder / f"PROCESSED_{timestamp}_{source_path.name}"
                    
                    # Move the file
                    source_path.rename(destination_path)
                    moved_count += 1
                    print(f"  Moved {source_path.name} to processed folder")
            except Exception as e:
                print(f"  Error moving {json_file}: {e}")
        
        print(f"📁 Moved {moved_count} processed JSON files to {processed_folder}")

    def convert_curesults_to_excel(self, curesults_folder: str = "./curesults", output_file: str = "seattle_opera_data_by_year.xlsx"):
        """Convert all JSON files in curesults folder to Excel with sheets organized by year."""
        curesults_path = Path(curesults_folder)
        
        if not curesults_path.exists():
            print(f"Error: Curesults folder '{curesults_folder}' does not exist.")
            return
        
        # Find all JSON files
        json_files = list(curesults_path.glob("*.json"))
        
        if not json_files:
            print(f"No JSON files found in {curesults_folder}")
            return
        
        print(f"Found {len(json_files)} JSON files to process.")
        
        # Process files and organize by year
        data_by_year, processed_files = self.process_json_files([str(f) for f in json_files])
        
        if not data_by_year:
            print("No valid data extracted from any files.")
            return
        
        # Save to Excel with separate sheets by year
        self.save_to_excel_by_year(data_by_year, output_file)
        
        # Move processed JSON files to subfolder
        if processed_files:
            print("\n" + "="*50)
            print("Moving processed JSON files...")
            print("="*50)
            self.move_processed_json_files(processed_files, curesults_folder)
        
        return output_file


def main():
    settings = Settings(
        endpoint=os.getenv("AZURE_CONTENT_UNDERSTANDING_ENDPOINT"),
        api_version="2025-05-01-preview",
        # Either subscription_key or aad_token must be provided. Subscription Key is more prioritized.
        subscription_key=os.getenv("AZURE_CONTENT_UNDERSTANDING_SUBSCRIPTION_KEY"),
        aad_token="AZURE_CONTENT_UNDERSTANDING_AAD_TOKEN",
        # Insert the analyzer name.
        analyzer_id=os.getenv("AZURE_CONTENT_UNDERSTANDING_ANALYZER_ID"),
        # Insert the supported file types of the analyzer.
        # file_location="./playbills/1980-81-Manon-Lescaut-Artists-Page.jpg",
    )
    client = AzureContentUnderstandingClient(
        settings.endpoint,
        settings.api_version,
        subscription_key=settings.subscription_key,
        token_provider=settings.token_provider,
    )
    
    # Process all files in the playbills folder
    playbills_folder = Path("./playbills")
    files_to_process = [f for f in playbills_folder.glob("*") if f.is_file() and not f.name.startswith("DONE_")]

    for file_path in files_to_process:
        print(f"\nProcessing file: {file_path.name}")
        file_location = str(file_path)

        response = client.begin_analyze(settings.analyzer_id, file_location)
        result = client.poll_result(
            response,
            timeout_seconds=60 * 60,
            polling_interval_seconds=1,
        )

        try:
            result_filename = Path(file_location).stem + "_result.json"
            with open(f"./curesults/{result_filename}", "w", encoding="utf-8") as f:
                json.dump(result, f, indent=2)
            print(f"Result saved to {result_filename}")
            # Rename and move the original file to processed subfolder
            original_file = Path(file_location)
            processed_folder = original_file.parent / "processed"
            processed_folder.mkdir(exist_ok=True)
            renamed_file = processed_folder / f"DONE_{original_file.name}"
            original_file.rename(renamed_file)
            print(f"Renamed and moved {original_file.name} to {renamed_file.name}")
        except Exception as e:
            print(f"Error saving result to file: {e}")
    
    # After processing all files, convert the results to Excel table organized by year
    print("\n" + "="*60)
    print("Converting JSON results to Excel table organized by year...")
    print("="*60)
    
    converter = SeattleOperaTableConverter()
    excel_output = converter.convert_curesults_to_excel(
        curesults_folder="./curesults",
        output_file="seattle_opera_complete_by_year.xlsx"
    )
    
    if excel_output:
        print(f"\n✅ Excel conversion completed: {excel_output}")
        print("📊 Data is organized in separate sheets by year")
    else:
        print("\n❌ Excel conversion failed or no data found")



@dataclass(frozen=True, kw_only=True)
class Settings:
    endpoint: str
    api_version: str
    subscription_key: str | None = None
    aad_token: str | None = None
    analyzer_id: str
    # file_location: str

    def __post_init__(self):
        key_not_provided = (
            not self.subscription_key
            or self.subscription_key == "AZURE_CONTENT_UNDERSTANDING_SUBSCRIPTION_KEY"
        )
        token_not_provided = (
            not self.aad_token
            or self.aad_token == "AZURE_CONTENT_UNDERSTANDING_AAD_TOKEN"
        )
        if key_not_provided and token_not_provided:
            raise ValueError(
                "Either 'subscription_key' or 'aad_token' must be provided"
            )

    @property
    def token_provider(self) -> Callable[[], str] | None:
        aad_token = self.aad_token
        if aad_token is None:
            return None

        return lambda: aad_token


class AzureContentUnderstandingClient:
    def __init__(
        self,
        endpoint: str,
        api_version: str,
        subscription_key: str | None = None,
        token_provider: Callable[[], str] | None = None,
        x_ms_useragent: str = "cu-sample-code",
    ) -> None:
        if not subscription_key and token_provider is None:
            raise ValueError(
                "Either subscription key or token provider must be provided"
            )
        if not api_version:
            raise ValueError("API version must be provided")
        if not endpoint:
            raise ValueError("Endpoint must be provided")

        self._endpoint: str = endpoint.rstrip("/")
        self._api_version: str = api_version
        self._logger: logging.Logger = logging.getLogger(__name__)
        self._logger.setLevel(logging.INFO)
        self._headers: dict[str, str] = self._get_headers(
            subscription_key, token_provider and token_provider(), x_ms_useragent
        )

    def begin_analyze(self, analyzer_id: str, file_location: str):
        """
        Begins the analysis of a file or URL using the specified analyzer.

        Args:
            analyzer_id (str): The ID of the analyzer to use.
            file_location (str): The path to the file or the URL to analyze.

        Returns:
            Response: The response from the analysis request.

        Raises:
            ValueError: If the file location is not a valid path or URL.
            HTTPError: If the HTTP request returned an unsuccessful status code.
        """
        if Path(file_location).exists():
            with open(file_location, "rb") as file:
                data = file.read()
            headers = {"Content-Type": "application/octet-stream"}
        elif "https://" in file_location or "http://" in file_location:
            data = {"url": file_location}
            headers = {"Content-Type": "application/json"}
        else:
            raise ValueError("File location must be a valid path or URL.")

        headers.update(self._headers)
        if isinstance(data, dict):
            response = requests.post(
                url=self._get_analyze_url(
                    self._endpoint, self._api_version, analyzer_id
                ),
                headers=headers,
                json=data,
            )
        else:
            response = requests.post(
                url=self._get_analyze_url(
                    self._endpoint, self._api_version, analyzer_id
                ),
                headers=headers,
                data=data,
            )

        response.raise_for_status()
        self._logger.info(
            f"Analyzing file {file_location} with analyzer: {analyzer_id}"
        )
        return response

    def poll_result(
        self,
        response: requests.Response,
        timeout_seconds: int = 120,
        polling_interval_seconds: int = 2,
    ) -> dict[str, Any]:  # pyright: ignore[reportExplicitAny]
        """
        Polls the result of an asynchronous operation until it completes or times out.

        Args:
            response (Response): The initial response object containing the operation location.
            timeout_seconds (int, optional): The maximum number of seconds to wait for the operation to complete. Defaults to 120.
            polling_interval_seconds (int, optional): The number of seconds to wait between polling attempts. Defaults to 2.

        Raises:
            ValueError: If the operation location is not found in the response headers.
            TimeoutError: If the operation does not complete within the specified timeout.
            RuntimeError: If the operation fails.

        Returns:
            dict: The JSON response of the completed operation if it succeeds.
        """
        operation_location = response.headers.get("operation-location", "")
        if not operation_location:
            raise ValueError("Operation location not found in response headers.")

        headers = {"Content-Type": "application/json"}
        headers.update(self._headers)

        start_time = time.time()
        while True:
            elapsed_time = time.time() - start_time
            self._logger.info(
                "Waiting for service response", extra={"elapsed": elapsed_time}
            )
            if elapsed_time > timeout_seconds:
                raise TimeoutError(
                    f"Operation timed out after {timeout_seconds:.2f} seconds."
                )

            response = requests.get(operation_location, headers=self._headers)
            response.raise_for_status()
            result = cast(dict[str, str], response.json())
            status = result.get("status", "").lower()
            if status == "succeeded":
                self._logger.info(
                    f"Request result is ready after {elapsed_time:.2f} seconds."
                )
                return response.json()  # pyright: ignore[reportAny]
            elif status == "failed":
                self._logger.error(f"Request failed. Reason: {response.json()}")
                raise RuntimeError("Request failed.")
            else:
                self._logger.info(
                    f"Request {operation_location.split('/')[-1].split('?')[0]} in progress ..."
                )
            time.sleep(polling_interval_seconds)

    def _get_analyze_url(self, endpoint: str, api_version: str, analyzer_id: str):
        return f"{endpoint}/contentunderstanding/analyzers/{analyzer_id}:analyze?api-version={api_version}&stringEncoding=utf16"

    def _get_headers(
        self, subscription_key: str | None, api_token: str | None, x_ms_useragent: str
    ) -> dict[str, str]:
        """Returns the headers for the HTTP requests.
        Args:
            subscription_key (str): The subscription key for the service.
            api_token (str): The API token for the service.
            enable_face_identification (bool): A flag to enable face identification.
        Returns:
            dict: A dictionary containing the headers for the HTTP requests.
        """
        headers = (
            {"Ocp-Apim-Subscription-Key": subscription_key}
            if subscription_key
            else {"Authorization": f"Bearer {api_token}"}
        )
        headers["x-ms-useragent"] = x_ms_useragent
        return headers


if __name__ == "__main__":
    main()
