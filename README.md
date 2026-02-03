Connect Rate Dashboard Automation

An automated data pipeline designed to consolidate daily agent performance reports into a single, clean master file for Power BI visualization. This tool handles the "Heavy Lifting" of navigating complex folder structures, cleaning raw Excel exports, and preparing data for reporting.

🚀 Overview

This project automates the extraction and transformation of agent contact data (Connect Rates, Voicemails, and Handle Times). It replaces manual copy-pasting by walking through a directory, skipping header noise, and merging multi-sheet data into a structured Final.xlsx output.


🛠️ Technical Implementation

The script performs the following automated steps:

    Directory Traversal: Uses os.walk to find all .xlsx files within the raw data folder.

    Smart Extraction: Reads data starting from Row 8 to bypass report headers, specifically capturing columns C through AQ.

    Data Cleansing:

        Removes empty rows based on the Agent Name column.

        Filters out "Total" summary rows using case-insensitive regex.

        Appends the original filename as a "Source File" column for auditability.

    OneDrive Compatibility: Includes a retry mechanism (3-second delay) to handle file locking issues common with OneDrive/SharePoint syncing.

⚙️ Setup & Usage
Prerequisites

    Python 3.8+

    Pandas & Openpyxl libraries

Installation

    Clone this repository to your local machine.

    Install the required packages:
    Bash

    pip install -r requirements.txt

Running the Pipeline

    Place your raw agent reports into the data/raw/ folder.

    Execute the script:
    Bash

    python src/process_data.py

    Open Power BI and click Refresh to update the "Connect Rate & Voicemail Performance" dashboard.

📊 Dashboard Visuals

The output data is optimized for Power BI, supporting:

    Connect Rate % vs. Voicemail % comparisons.

    Daily performance trends by Agent Name or Team.

    Sentiment analysis based on message counts.
