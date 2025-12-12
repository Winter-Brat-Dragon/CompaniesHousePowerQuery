# Companies House Power Query Scripts

> **Disclaimer:** The scripts in this repository were generated almost entirely by AI. Use at your own risk. 

## Overview

This repository contains **Power Query (M) scripts** designed to retrieve company data from the Companies House API. The scripts are able to generate:

- A table of company profiles 
- A list of officers for all companies on the Input Table
- A list of PSCs for all companies on the Input Table

The Input Table needs Company Numbers to be entered to generate the api URLs.

---

## Folder Structure

- `Functions/`
  - `GetCompanyProfile.m` – Fetch company profile from Companies House
  - `GetCompanyOfficers.m` – Fetch company officers data from Companies House
  - `GetCompanyPSCs.m` – Fetch PSC data from Companies House
  - `GetNestedField.m` – Helper function used in all queries

- `Queries/` 
  - `CompanyProfiles.m` – Generates a table of company profiles
  - `CompanyOfficers.m` – Generates a list of officers for all companies on the Input Table
  - `CompanyPSCs.m` – Generates a list of PSCs for all companies on the Input Table

- `ExampleInput.xlsx` – Example input table (Sorry DW)f

---

## Basic Usage

1. You'll need an API key from Companies House. 
   -This requires signing up for an account and generating a key
   -https://developer.company-information.service.gov.uk/
   
1. Prepare the input table in Excel  
   - The CompanyNumber field is mandatory
   - The API URLs will generate automatically
   - There is an optional CompanyName column for reference. Data entered here won't pull through to the output tables.

2. Add each function and query in the query editor in excel

3. Review the outputs and load them to new tables

---

Notes

- Not sure what you're doing? Neither was I!
- Try asking ChatGPT to get simple instructions on how to use these scripts in excel. It worked for me!
