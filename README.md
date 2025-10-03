# Seattle Opera Data Processing Suite

## 🎭 Help Preserve Seattle's Cultural Legacy (#GIVE2025)

**Join us in digitizing decades of Seattle Opera's historic performance programs!**

Much of Seattle Opera's rich cultural history is currently locked away in scanned images that aren't searchable or accessible. This volunteer-driven digitization project transforms those historical treasures into structured, searchable data that will be available for generations to come.

### Why This Matters

- **Preserve Cultural Heritage**: Help safeguard Seattle's performing arts legacy
- **Enable Accessibility**: Make historical performance data searchable and AI-accessible  
- **Celebrate History**: Honor the artists, conductors, and productions that shaped our city's cultural landscape
- **Community Impact**: No experience required—contribute to your community's cultural preservation

### What We're Digitizing

Transform scanned playbill images into structured Excel files containing:

- **Performer names and roles** from iconic productions (1980s-present)
- **Production details** including dates, venues, and cast information
- **Historical context** that brings Seattle's opera history to life

---

## Technical Overview

A comprehensive Python toolkit for processing Seattle Opera playbill images and converting the extracted JSON data into structured Excel tables. The suite provides both standalone conversion tools and an integrated Azure Content Understanding workflow.

## Project Workflow

This project processes Seattle Opera historical data through two complementary approaches:

1. **Image Processing**: Uses Azure Content Understanding to extract text and structure from playbill images
2. **Data Conversion**: Converts the resulting JSON data into organized Excel spreadsheets with year-based sheets

## Project Structure

```text
seattleopera/
├── json_to_table_converter.py      # Standalone JSON→Excel/CSV converter
├── seattleoperacuprocessing.py     # Integrated Azure AI + conversion workflow
├── requirements.txt                # Python dependencies
├── .env                            # Azure credentials configuration (create from sample)
├── .env_sample                     # Sample environment file template
├── cuanalyer_template/             # Azure Content Understanding analyzer template
│   └── seattleopera.json          # Pre-configured analyzer for opera playbills
├── curesults/                      # JSON processing directory
│   └── processed/                  # Auto-managed processed files folder
├── playbills/                      # Input images directory
└── README.md                       # This documentation
```

## Main Scripts

- **`json_to_table_converter.py`** - Standalone conversion tool with command-line interface
- **`seattleoperacuprocessing.py`** - Complete workflow: image processing + automatic Excel conversion with file management
- **`cuanalyer_template/seattleopera.json`** - Azure Content Understanding analyzer template for opera playbills

## Quick Start

### Option 1: Full Workflow (Recommended)

```bash
# Process playbill images and convert to Excel (all-in-one)
python seattleoperacuprocessing.py
```

### Option 2: JSON-Only Conversion

```bash
# Convert JSON files in curesults/ folder to Excel
python json_to_table_converter.py curesults/

# Convert single JSON file to CSV
python json_to_table_converter.py curesults/processed/curesult.json

# Get help with all available options
python json_to_table_converter.py --help
```

## Integrated Processing Workflow

The `seattleoperacuprocessing.py` script provides a complete end-to-end workflow:

1. **Process playbill images** with Azure Content Understanding API
2. **Extract structured data** from images to JSON files in `curesults/` folder
3. **Automatically convert** JSON results to Excel organized by year
4. **Create separate sheets** for each performance year (1980, 1981, etc.)
5. **Move processed JSON files** to `processed/` subfolder with timestamps
6. **Generate consolidated Excel** with multiple sheets and summary statistics

### Excel Output Structure

When using the integrated processor, the generated Excel file contains:

- **Year_1980 sheet**: All performances from 1980
- **Year_1981 sheet**: All performances from 1981  
- **Year_1982 sheet**: All performances from 1982
- **Summary sheet**: Overview with show counts and statistics by year
- **Unknown_Year sheet**: Data where year couldn't be determined

## Usage Examples

### Basic Operations

```bash
# Convert single JSON file
python json_to_table_converter.py your_file.json

# Convert to Excel with custom output
python json_to_table_converter.py your_file.json -f excel -o output.xlsx
```

### Append Mode

```bash
# Append data to existing CSV file
python json_to_table_converter.py new_data.json -o existing_file.csv --append

# Append data to existing Excel file  
python json_to_table_converter.py new_data.json -o existing_file.xlsx -f excel --append
```

### Multiple File Processing

```bash
# Process all JSON files in a directory
python json_to_table_converter.py /path/to/json/directory

# Process with custom output file
python json_to_table_converter.py /path/to/json/directory -o combined_results.csv

# Process recursively (including subdirectories)
python json_to_table_converter.py /path/to/json/directory --recursive

# Process with custom file pattern
python json_to_table_converter.py /path/to/json/directory --pattern "opera_*.json"
```

### Real-World Workflow Examples

```bash
# Complete processing workflow: images → JSON → Excel
1. Place playbill images in playbills/ folder
2. Run: python seattleoperacuprocessing.py
3. Find Excel output with year-organized sheets
4. Processed JSON files automatically moved to curesults/processed/

# Convert existing JSON files to Excel by year
python json_to_table_converter.py curesults/processed/ -f excel -o seattle_opera_complete.xlsx

# Combine multiple directories into one file
python json_to_table_converter.py /path/to/first/directory -o master_file.csv
python json_to_table_converter.py /path/to/second/directory -o master_file.csv --append

# Process specific pattern recursively
python json_to_table_converter.py /path/to/base/directory --recursive --pattern "*_results.json"
```

## Table Structure

The converter creates tables with these columns:

- **SHOW**: Name of the opera/show
- **DATES**: Performance dates  
- **ROLE**: Role or position (conductor, character, etc.)
- **ARTIST**: Name of the performer/artist
- **OTHER**: Additional information (if available)
- **FILENAME**: Original JSON filename (for tracking data source)

## Features

### Standalone Converter (json_to_table_converter.py)

- ✅ Single file conversion (JSON → CSV/Excel)
- ✅ Batch processing of multiple JSON files
- ✅ Append mode for incremental data collection
- ✅ Recursive directory processing
- ✅ Custom file pattern matching
- ✅ Error handling and progress reporting
- ✅ Support for both CSV and Excel output formats

### Integrated Processor (seattleoperacuprocessing.py)


- ✅ Azure Content Understanding API integration
- ✅ Automatic processing of playbill images from `playbills/` folder
- ✅ Automatic Excel conversion organized by year
- ✅ Separate Excel sheets for each performance year (Year_1980, Year_1981, etc.)
- ✅ Summary sheet with statistics and show counts by year
- ✅ Year extraction from date strings (supports ranges like "1980-81")
- ✅ **Automatic file management**: Moves processed JSON files to `processed/` subfolder with timestamps
- ✅ **Incremental processing**: Skips already-processed files to avoid duplicates
- ✅ **Progress tracking**: Detailed logging and status reporting

## Requirements & Installation

### Prerequisites

- **Python 3.8+**
- **Azure AI account** (for integrated processor)
- **Required packages**: See `requirements.txt`

### Quick Setup (Recommended)

1. **Clone and Navigate to Project**:

   ```bash
   git clone <repository-url>
   cd seattleoperadigitization
   ```

2. **Create Virtual Environment**:

   ```bash
   # Create virtual environment
   python -m venv .venv
   
   # Activate virtual environment
   # On Windows (PowerShell):
   .\.venv\Scripts\Activate.ps1
   
   # On Windows (Command Prompt):
   .\.venv\Scripts\activate.bat
   
   # On macOS/Linux:
   source .venv/bin/activate
   ```

3. **Install Dependencies**:

   ```bash
   # Install all required packages from requirements.txt
   pip install -r requirements.txt
   ```

4. **Set up Azure credentials** (for integrated processor):

   ```bash
   # Copy the sample environment file
   cp .env_sample .env
   # Edit .env file with your actual Azure credentials
   ```

### Manual Installation (Alternative)

```bash
# Install required dependencies individually
pip install pandas openpyxl azure-ai-documentintelligence python-dotenv requests xlsxwriter

# Navigate to the tool directory
cd seattleoperadigitization

# Set up Azure credentials (copy from .env_sample)
cp .env_sample .env
# Edit .env file with your actual Azure credentials
```

### Azure Setup (Required for Integrated Processor)

1. **Create Azure Document Intelligence Resource**:
   - Go to [Azure Portal](https://portal.azure.com)
   - Create a new "Document Intelligence" resource
   - Copy the endpoint URL and subscription key

2. **Configure Environment Variables**:

   ```bash
   # Copy the sample environment file
   cp .env_sample .env
   
   # Edit .env file with your credentials:
   # AZURE_CONTENT_UNDERSTANDING_ENDPOINT=https://your-resource-name.cognitiveservices.azure.com/
   # AZURE_CONTENT_UNDERSTANDING_SUBSCRIPTION_KEY=your_32_character_key
   ```

3. **Set Up Custom Content Understanding Analyzer**:

   This project includes a pre-configured analyzer template specifically designed for Seattle Opera playbill processing. The template defines the exact data structure and extraction fields optimized for opera program data.

   **Using the Provided Template:**

   ```bash
   # The analyzer template is located at:
   cuanalyer_template/seattleopera.json
   ```

   **Template Structure:**
   - **SHOW**: Extracts the opera name from playbill images
   - **DATES**: Generates performance dates (supports year-only when days unclear)
   - **ROLES**: Array of cast and crew information with:
     - **ROLE**: Position (Conductor, Director, Character names, etc.)
     - **ARTIST**: Person's name assigned to the role
     - **OTHER**: Additional context (company, language, performance order)

   **To set up your own Content Understanding instance:**

   1. **Create Azure AI Foundry Resource**:
      - Go to [Azure Portal](https://portal.azure.com)
      - Create "Azure AI Foundry" resource
      - Note the endpoint and subscription key

   2. **Import the Analyzer Template**:
      - Access your Content Understanding Studio on [Foundry Portal](https://ai.azure.com)
      - Create new analyzer project
      - Import the provided `cuanalyer_template/seattleopera.json` template
      - Train the analyzer with sample playbill images
      - Deploy the trained model

   3. **Update Environment Configuration**:

      ```bash
      # Add to your .env file:
      AZURE_CONTENT_UNDERSTANDING_ANALYZER_ID=your_deployed_analyzer_id
      ```

   **Benefits of the Custom Template:**
   - Optimized field extraction for opera program data
   - Handles multiple cast members per role
   - Extracts crew positions (directors, designers, etc.)
   - Captures additional context information
   - Supports flexible date formats common in playbills

4. **Verify Setup**:

   ```bash
   # Test standalone converter (no Azure required)
   python json_to_table_converter.py --help
   
   # Test integrated processor (requires Azure setup)
   python seattleoperacuprocessing.py
   ```

## File Management & Organization

### Automatic File Processing

The integrated processor automatically manages files:

- **Input**: Playbill images go in `playbills/` folder
- **Intermediate**: JSON results saved to `curesults/` folder
- **Processed**: Completed JSON files moved to `curesults/processed/` with timestamps
- **Output**: Excel files generated in main directory with year-based sheets

### File Naming Convention

- **Processed JSON files**: `PROCESSED_YYYYMMDD_HHMMSS_originalname.json`
- **Excel output**: `seattle_opera_data_by_year_YYYYMMDD_HHMMSS.xlsx`

### Current Project Status

✅ **Fully Operational**: Both tools are production-ready  
✅ **7 Seattle Opera productions processed** (1980-1982 seasons)  
✅ **173 individual artist/role records** extracted and organized  
✅ **Automatic file management** prevents reprocessing  
✅ **Year-based organization** with separate Excel sheets  

## Error Handling

The suite provides comprehensive error handling:

- Skips files that can't be processed and continues with others
- Reports which files were processed successfully
- Shows total number of rows extracted from each file
- Displays sample data for verification
- Provides detailed progress information during batch processing
- Gracefully handles Azure API errors and network issues
- Validates JSON structure before processing

## Troubleshooting

### Common Issues

1. **Azure Authentication**: Ensure `.env` file has correct `AZURE_CONTENT_UNDERSTANDING_*` credentials
2. **File Permissions**: Check write access to output directories  
3. **JSON Format**: Validate JSON structure matches expected schema
4. **Dependencies**: Install all required packages with `pip install`

### Getting Help

```bash
# Get detailed help for standalone converter
python json_to_table_converter.py --help

# Check Azure configuration
python seattleoperacuprocessing.py
```

---

## About

**Created**: October 3, 2025  
**Purpose**: Seattle Opera Historical Data Processing Suite  
**Components**: Azure AI Document Intelligence + Excel Table Generation  
**Status**: Production Ready  
**Author**: Jan Goergen
