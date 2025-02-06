# Land Use Land Cover Ontology Instantiation
The aim of this repository is to store the code to instantiate automatically the **LULC ontology** from an excel file, or a folder containing excel files.

## Overview

The LULC Ontology is an ongoing project designed to assist the research community in conducting systematic reviews on Land Use and Land Cover (LULC) and related subjects. This ontology provides a structured representation of LULC information, facilitating data organization and knowledge extraction. The work is still under development and has not yet been officially published.

## Features

- **Automated Ontology Instantiation**: Converts data from an Excel file into an ontology in OWL format.
- **Metadata Enrichment**: Automatically retrieves metadata for research articles using DOIs.
- **Excel Template Guide**: Provides instructions on how to format the input Excel file.

## Repository Contents

### Documentation
- **LULC_ontology_template_guide.docx** is a detailed guide on how to fill the Excel template.

### Templates
- **LULC_Ontology_template.xlsm** is the empty Excel file needed to instantiate the ontology.
- **LULC_Ontology_example.xlsm** is an example Excel file.

### Ontology File
- **lulc_review.owl** is the ontology represented in the owl format.

### Python Scripts
- **owl_filler.py** is the python code that take as input the excel file and outputs the ontology instantiated with the articles in the Excel file.
- **metadata_enrichment.py** is a python code that allows to automatically complete in the Excel file the metadata of a paper from its DOI.

## Usage
### 1. Prepare the input Excel File
Complete the **LULC_Ontology_template** with the information from the paper you to describe. Follow the **template guide** to correctly format the Excel file.

### 2. Enrich Metadata (Optional)
To automatically complete metadata for articles using their DOI, run:

`python metadata_enrichment.py`

This script will:
- Extract DOIs from the Excel file.
- Retrieve metadata (such as title, authors, journal, and year) from online databases.
- Update the Excel file with the retrieved metadata.

### 3. Run the Ontology Instantiation
To automatically populate the ontology from the Excel file, execute the script:

`python owl_filler.py`

This script will:
- Load the base ontology from lulc_review.owl.
- Read the Excel file(s) from the specified directory.
- Instantiate the ontology using the data from the Excel file.
- Save the newly instantiated ontology as lulc_review_instantiated.owl.

## Dependencies
Ensure the following Python libraries are installed before running the scripts:

`pip install pandas owlready2 pylatexenc requests geopy html`