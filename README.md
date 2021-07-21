# Birmingham Tech Talent Dashboard

This repository contains a collection of data cleaning and management scripts 
created to help support the [Magic City Data Collective's](https://sites.uab.edu/data-collective/) 
Tech Talent Gap Analysis Dashboard.

# Installation

To clone and run these scripts, you'll need one of
[Miniconda](https://docs.conda.io/en/latest/miniconda.html) or 
[Anaconda](https://www.anaconda.com/) installed on your computer.  From your
Anaconda command prompt:

```
# Clone this repository
git clone https://github.com/manage-this/mcdc-tech-talent.git

# Go into the repository
cd mcdc-tech-talent

# Create a new environment with all required packages
conda env create -f environment.yml
```

# Google API Key

A valid Google API credential file named `key.json` must be placed in the 
project's root folder. 

# Running the cleaning scripts

Run this project when any of the primary data sources have been refreshed with 
updated releases.  All data found in the appropriate Google Drive folders
will first be downloaded to your computer.  After all cleaning and aggregation
steps are complete, the resulting datasets will be uploaded back into their
correct locations in Google Sheets.

Go to this project's root directory from your Anaconda command prompt:

```
# Activate the virtual environment
conda activate tech_talent

# Run the main data processing script
python mcdc_tech_talent.py
```

# License