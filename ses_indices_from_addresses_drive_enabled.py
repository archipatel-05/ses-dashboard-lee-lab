#!/usr/bin/env python3
"""
ses_indices_from_addresses_v4_refined.py

Refined address-to-index pipeline.

What this version changes
-------------------------
1. Uses project-relative sourcing instead of hard-coded machine-specific paths.
2. Auto-extracts reference ZIP archives found beside the program.
3. Adds ADI, DCI, and AHRF support.
4. Keeps the main output clean: one geography choice per index family, no tract+ZIP
   duplicate score sets in the main file.
5. Produces five output workbooks:
      - main
      - supplementary SVI
      - supplementary RUCA
      - COI
      - AHRF
6. Prefers census tract over ZIP where applicable, but only when the whole index
   family can be sourced at that level.
7. Searches recursively inside the project folder for the supplied reference files.

Notes on .numbers files
-----------------------
The uploaded reference bundle includes Apple Numbers files for at least some sources.
This script first looks for CSV/XLSX exports with the same basename. If none exist,
it will try to read the .numbers file only if `numbers_parser` is installed. If that
still is not available, the script records a note and falls back when possible.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import time
import zipfile
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
import requests

try:
    import gdown  # downloads public Google Drive folders at runtime
except ImportError:  # keeps local non-Drive/offline runs from crashing before download is needed
    gdown = None

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# 0. Defaults / constants
# ------------------------------------------------------------------
INPUT_FILE = Path("input_addresses.xlsx")
OUTPUT_PREFIX = Path("ses_indices_from_addresses")

# Large reference files should NOT live in GitHub.
# This public Google Drive folder is downloaded into reference_data/ when the script runs.
GOOGLE_DRIVE_FOLDER_URL = os.environ.get(
    "GOOGLE_DRIVE_FOLDER_URL",
    "https://drive.google.com/drive/folders/1SFsKP_0cyBpdkQv4YCS3PrNUKSvNPBgT?usp=drive_link",
)
REFERENCE_DATA_DIR = Path("reference_data")
REFERENCE_DATA_DOWNLOAD_MARKER = ".download_complete"
TEXT_IDENTIFIER_HEADERS = {
    "Unique ID", "Census Tract (FIPS)", "Block Group (FIPS)", "ZCTA (Zip Code)",
    "GeocodedAddress", "CensusTractFIPS", "BlockGroupFIPS", "ZCTA", "StateFIPS", "CountyFIPS",
    "GEO_ID", "STATE", "COUNTY", "TRACT", "CBSA",
    "GEOID10", "GEOID20", "STATEFP", "COUNTYFP", "TRACTCE", "BLKGRPCE",
}
GEOCODER_SLEEP_SECONDS = 0.10
REQUEST_TIMEOUT_SECONDS = 30
GEOCODER_BENCHMARK = "Public_AR_Census2020"
GEOCODER_VINTAGE = "Census2020_Census2020"
GEOCODE_CACHE_FILENAME = "census_geocode_cache.json"
GEOCODE_MODE_AUTO = "auto"
GEOCODE_MODE_LIVE = "live"
GEOCODE_MODE_OFFLINE = "offline"

MAIN_SHEET = "Main"
SUMMARY_SHEET = "Summary"
MAIN_DEFINITIONS_SHEET = "Main Definitions"
RUN_AUDIT_SHEET = "Run Audit"
NOTES_SHEET = "Notes"
NORMALIZED_BAD_SHEET = "Normalized Bad 0-100"

DISPLAY_LABEL_MAP = {
    "Unique ID": "Record ID",
    "Address (as inputted)": "Original Address",
    "Geocoded Address": "Matched / Standardized Address",
    "Census Tract (FIPS)": "Census Tract FIPS",
    "ZCTA (Zip Code)": "ZIP Code (ZCTA)",
    "Block Group (FIPS)": "Block Group FIPS",
    "ADI_NATRANK": "ADI National Rank",
    "ADI_STATERNK": "ADI State Rank",
    "National Risk Category": "ADI National Risk Category",
    "State Risk Category": "ADI State Risk Category",
    "SVI_National_RPL_THEME1": "SVI National Theme 1 Rank",
    "SVI_National_RPL_THEME2": "SVI National Theme 2 Rank",
    "SVI_National_RPL_THEME3": "SVI National Theme 3 Rank",
    "SVI_National_RPL_THEME4": "SVI National Theme 4 Rank",
    "SVI_National_RPL_THEMES": "SVI National Overall Rank",
    "SVI_State_RPL_THEME1": "SVI State Theme 1 Rank",
    "SVI_State_RPL_THEME2": "SVI State Theme 2 Rank",
    "SVI_State_RPL_THEME3": "SVI State Theme 3 Rank",
    "SVI_State_RPL_THEME4": "SVI State Theme 4 Rank",
    "SVI_State_RPL_THEMES": "SVI State Overall Rank",
    "VVI": "VVI Overall Score",
    "HealthCareAccess": "VVI Health Care Access",
    "CleanEnvironment": "VVI Clean Environment",
    "PrimaryRUCA": "RUCA Primary Code",
    "PrimaryRUCADescription": "RUCA Primary Description",
    "SecondaryRUCA": "RUCA Secondary Code",
    "SecondaryRUCADescription": "RUCA Secondary Description",
    "2019-2023 Distress Score": "DCI Distress Score (2019–2023)",
    "Quintile (5=Distressed)": "DCI Quintile",
    "Risk Categorization": "DCI Category",
    "RISK_SCORE": "NRI Risk Score",
    "RISK_RATNG": "NRI Risk Rating",
    "EAL_SCORE": "NRI Expected Annual Loss Score",
    "SOVI_SCORE": "NRI Social Vulnerability Score",
    "RESL_SCORE": "NRI Community Resilience Score",
    "HRSA_MUA_CENSUS_TRACT": "Indicates whether the census tract is designated as a medically underserved area, meaning residents may have fewer local health-care resources.",
    "POS_DIST_ED_TRACT": "Estimated distance from this tract to the nearest emergency department.",
    "POS_DIST_CLINIC_TRACT": "Estimated distance from this tract to the nearest outpatient clinic or similar care site.",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT": "Share of Medicare claims in the area that were paid at the full approved amount, used here as a health-system access / payment measure.",
    "POS_DIST_MEDSURG_ICU_TRACT": "Estimated distance from this tract to the nearest hospital with a medical-surgical intensive care unit.",
    "POS_DIST_TRAUMA_TRACT": "Estimated distance from this tract to the nearest trauma center.",
    "AHRF Source Year": "Year of the AHRF source file used for these health-care access measures.",
    "AHRF Geography Used": "Geographic level used to attach the AHRF values to this record.",
    "AHRF Missing Reason": "Why the AHRF values were missing for this record, if they could not be assigned.",
    "HRSA_MUA_CENSUS_TRACT": "AHRF MUA Flag (Census Tract)",
    "POS_DIST_ED_TRACT": "AHRF Distance to Emergency Department (Tract)",
    "POS_DIST_CLINIC_TRACT": "AHRF Distance to Clinic (Tract)",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT": "AHRF Percent Medicare Approved Full Amount",
    "POS_DIST_MEDSURG_ICU_TRACT": "AHRF Distance to Med-Surg ICU (Tract)",
    "POS_DIST_TRAUMA_TRACT": "AHRF Distance to Trauma Center (Tract)",
    "AHRF Source Year": "AHRF Source Year",
    "AHRF Geography Used": "AHRF Geography Used",
    "AHRF Missing Reason": "AHRF Missing Reason",
    "NatWalkInd": "National Walkability Index",
    "D2A_Ranked": "Walkability Household Mix Rank (D2A_Ranked)",
    "D2B_Ranked": "Walkability Employment Mix Rank (D2B_Ranked)",
    "D3B_Ranked": "Walkability Intersection Density Rank (D3B_Ranked)",
    "D4A_Ranked": "Walkability Transit / Destination Accessibility Rank (D4A_Ranked)",
    "Walkability Source Year": "Walkability Source Year",
    "Walkability Geography Used": "Walkability Geography Used",
    "Walkability Missing Reason": "Walkability Missing Reason",
    "Geocode Status": "Geocoding Status",
    "Geocode Method": "Geocoding Method",
    "Geocode Match Type": "Geocoding Match Type",
    "Geography Level Reached": "Most Detailed Geography Reached",
    "Tract-BlockGroup Consistency Flag": "Tract / Block Group Consistency",
    "Index Coverage Count": "Number of Index Families Populated",
}
DISPLAY_TO_RAW_LABEL_MAP = {v: k for k, v in DISPLAY_LABEL_MAP.items()}
TEXT_IDENTIFIER_RAW_HEADERS = set(TEXT_IDENTIFIER_HEADERS)

def prettify_column_label(label: str) -> str:
    if label in DISPLAY_LABEL_MAP:
        return DISPLAY_LABEL_MAP[label]
    raw = str(label).strip()
    if not raw:
        return raw
    raw = raw.replace('_', ' ')
    raw = re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', raw)
    replacements = {
        'Fips': 'FIPS', 'Zip': 'ZIP', 'Ruca': 'RUCA', 'Svi': 'SVI', 'Adi': 'ADI',
        'Vvi': 'VVI', 'Dci': 'DCI', 'Eji': 'EJI', 'Nri': 'NRI', 'Cre': 'CRE',
        'Fara': 'FARA', 'Ahrf': 'AHRF', 'Coi': 'COI', 'Eal': 'EAL', 'Resl': 'RESL', 'Walkability': 'Walkability',
        'Sovi': 'SOVI', 'Pred': 'PRED', 'Pe': 'PE', 'Rpl': 'RPL', 'Cbm': 'CBM',
        'Svm': 'SVM', 'Hvm': 'HVM', 'Ebm': 'EBM', 'Geo Id': 'GEO ID', 'Geoid': 'GEOID'
    }
    for old, new in replacements.items():
        raw = raw.replace(old, new)
    return raw

def definition_lookup_label(label: str) -> str:
    return DISPLAY_TO_RAW_LABEL_MAP.get(str(label).strip(), str(label).strip())

def get_field_definition(label: str) -> str:
    raw = definition_lookup_label(label)
    return MAIN_FIELD_DEFINITIONS.get(raw, MAIN_FIELD_DEFINITIONS.get(str(label).strip(), ''))

def is_text_identifier_header(label: str) -> bool:
    raw = definition_lookup_label(label)
    return raw in TEXT_IDENTIFIER_RAW_HEADERS or str(label).strip() in TEXT_IDENTIFIER_RAW_HEADERS

STATE_FIPS_TO_NAME = {
    "01": "Alabama", "02": "Alaska", "04": "Arizona", "05": "Arkansas", "06": "California",
    "08": "Colorado", "09": "Connecticut", "10": "Delaware", "11": "District of Columbia",
    "12": "Florida", "13": "Georgia", "15": "Hawaii", "16": "Idaho", "17": "Illinois",
    "18": "Indiana", "19": "Iowa", "20": "Kansas", "21": "Kentucky", "22": "Louisiana",
    "23": "Maine", "24": "Maryland", "25": "Massachusetts", "26": "Michigan", "27": "Minnesota",
    "28": "Mississippi", "29": "Missouri", "30": "Montana", "31": "Nebraska", "32": "Nevada",
    "33": "New Hampshire", "34": "New Jersey", "35": "New Mexico", "36": "New York",
    "37": "North Carolina", "38": "North Dakota", "39": "Ohio", "40": "Oklahoma", "41": "Oregon",
    "42": "Pennsylvania", "44": "Rhode Island", "45": "South Carolina", "46": "South Dakota",
    "47": "Tennessee", "48": "Texas", "49": "Utah", "50": "Vermont", "51": "Virginia",
    "53": "Washington", "54": "West Virginia", "55": "Wisconsin", "56": "Wyoming",
    "60": "American Samoa", "66": "Guam", "69": "Northern Mariana Islands", "72": "Puerto Rico",
    "74": "U.S. Minor Outlying Islands", "78": "U.S. Virgin Islands",
}
STATE_NAME_TO_ABBR = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR", "California": "CA",
    "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE", "District of Columbia": "DC",
    "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID", "Illinois": "IL",
    "Indiana": "IN", "Iowa": "IA", "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA",
    "Maine": "ME", "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN",
    "Mississippi": "MS", "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR",
    "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC", "South Dakota": "SD",
    "Tennessee": "TN", "Texas": "TX", "Utah": "UT", "Vermont": "VT", "Virginia": "VA",
    "Washington": "WA", "West Virginia": "WV", "Wisconsin": "WI", "Wyoming": "WY",
    "American Samoa": "AS", "Guam": "GU", "Northern Mariana Islands": "MP", "Puerto Rico": "PR",
    "U.S. Minor Outlying Islands": "UM", "U.S. Virgin Islands": "VI",
}

STATE_ABBR_TO_NAME = {abbr: name for name, abbr in STATE_NAME_TO_ABBR.items()}

TRACT_COLUMN_CANDIDATES = [
    "censustract", "census_tract", "tract", "tractid", "tract_id",
    "tractfips", "tract_fips", "geoid", "geoid20", "geoid_2020", "fips", "fips11",
]
ZIP_COLUMN_CANDIDATES = [
    "zipcode", "zip", "zip5", "zip_code", "postalcode", "postal_code", "zcta", "zcta5", "zcta5ce10", "zcta5ce20",
]
BLOCKGROUP_COLUMN_CANDIDATES = [
    "blockgroupfips", "blockgroup", "bg_fips", "bgfips", "geoid12", "fips12",
]
YEAR_COLUMN_CANDIDATES = ["year", "yr", "data_year", "coi_year"]

SVI_MAIN_METRICS = ["RPL_THEME1", "RPL_THEME2", "RPL_THEME3", "RPL_THEME4", "RPL_THEMES"]
VVI_MAIN_METRICS = [
    "VVI", "Economic", "Education", "HealthCareAccess", "Neighborhood",
    "Housing", "CleanEnvironment", "Social", "Transportation", "PublicSafety",
]
RUCA_MAIN_METRICS = [
    "PrimaryRUCA", "PrimaryRUCADescription", "SecondaryRUCA", "SecondaryRUCADescription",
]
DCI_MAIN_METRICS = ["2019-2023 Distress Score", "Quintile (5=Distressed)"]
ADI_MAIN_METRICS = ["UWADI_ADI_NATRANK", "UWADI_ADI_STATERNK"]
EJI_MAIN_METRICS = ["RPL_SER", "RPL_EJI", "RPL_EBM", "RPL_SVM", "RPL_HVM", "RPL_CBM", "RPL_EJI_CBM"]
CRE_MAIN_METRICS = ["PRED3_PE", "PRED12_PE", "PRED0_PE"]
FARA_MAIN_METRICS = ["LILATracts_1And10", "LILATracts_Vehicle", "LowIncomeTracts"]
NRI_MAIN_METRICS = ["RISK_SCORE", "RISK_RATNG", "EAL_SCORE", "SOVI_SCORE", "RESL_SCORE"]
AHRF_MAIN_METRICS = [
    "HRSA_MUA_CENSUS_TRACT",
    "POS_DIST_ED_TRACT",
    "POS_DIST_CLINIC_TRACT",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT",
    "POS_DIST_MEDSURG_ICU_TRACT",
    "POS_DIST_TRAUMA_TRACT",
]
WALK_MAIN_METRICS = ["NatWalkInd", "D2A_Ranked", "D2B_Ranked", "D3B_Ranked", "D4A_Ranked"]

COI_DOMAIN_OUTPUT_COLUMNS = [
    "year", "r_COI_nat", "r_COI_stt", "r_COI_met", "z_COI_nat", "z_COI_stt", "z_COI_met",
    "r_ED_nat", "r_ED_stt", "r_ED_met", "z_ED_nat", "z_ED_stt", "z_ED_met",
    "r_HE_nat", "r_HE_stt", "r_HE_met", "z_HE_nat", "z_HE_stt", "z_HE_met",
    "r_SE_nat", "r_SE_stt", "r_SE_met", "z_SE_nat", "z_SE_stt", "z_SE_met",
]
COI_SUBDOMAIN_OUTPUT_COLUMNS = [
    "r_ED_EC_nat", "r_ED_EC_stt", "r_ED_EC_met", "z_ED_EC_nat", "z_ED_EC_stt", "z_ED_EC_met",
    "r_ED_EL_nat", "r_ED_EL_stt", "r_ED_EL_met", "z_ED_EL_nat", "z_ED_EL_stt", "z_ED_EL_met",
    "r_ED_ER_nat", "r_ED_ER_stt", "r_ED_ER_met", "z_ED_ER_nat", "z_ED_ER_stt", "z_ED_ER_met",
    "r_ED_SP_nat", "r_ED_SP_stt", "r_ED_SP_met", "z_ED_SP_nat", "z_ED_SP_stt", "z_ED_SP_met",
    "r_HE_EP_nat", "r_HE_EP_stt", "r_HE_EP_met", "z_HE_EP_nat", "z_HE_EP_stt", "z_HE_EP_met",
    "r_HE_HR_nat", "r_HE_HR_stt", "r_HE_HR_met", "z_HE_HR_nat", "z_HE_HR_stt", "z_HE_HR_met",
    "r_HE_SE_nat", "r_HE_SE_stt", "r_HE_SE_met", "z_HE_SE_nat", "z_HE_SE_stt", "z_HE_SE_met",
    "r_HE_HE_nat", "r_HE_HE_stt", "r_HE_HE_met", "z_HE_HE_nat", "z_HE_HE_stt", "z_HE_HE_met",
    "r_SE_EI_nat", "r_SE_EI_stt", "r_SE_EI_met", "z_SE_EI_nat", "z_SE_EI_stt", "z_SE_EI_met",
    "r_SE_EO_nat", "r_SE_EO_stt", "r_SE_EO_met", "z_SE_EO_nat", "z_SE_EO_stt", "z_SE_EO_met",
    "r_SE_ER_nat", "r_SE_ER_stt", "r_SE_ER_met", "z_SE_ER_nat", "z_SE_ER_stt", "z_SE_ER_met",
    "r_SE_HQ_nat", "r_SE_HQ_stt", "r_SE_HQ_met", "z_SE_HQ_nat", "z_SE_HQ_stt", "z_SE_HQ_met",
    "r_SE_SR_nat", "r_SE_SR_stt", "r_SE_SR_met", "z_SE_SR_nat", "z_SE_SR_stt", "z_SE_SR_met",
    "r_SE_WL_nat", "r_SE_WL_stt", "r_SE_WL_met", "z_SE_WL_nat", "z_SE_WL_stt", "z_SE_WL_met",
]

MAIN_OUTPUT_COLUMN_ORDER = [
    "Unique ID",
    "Address (as inputted)",
    "Geocoded Address",
    "State / Territory",
    "Census Tract (FIPS)",
    "ZCTA (Zip Code)",
    "Block Group (FIPS)",
    "Latitude",
    "Longitude",
    "ADI_NATRANK",
    "ADI_STATERNK",
    "National Risk Category",
    "State Risk Category",
    "SVI_National_RPL_THEME1",
    "SVI_National_RPL_THEME2",
    "SVI_National_RPL_THEME3",
    "SVI_National_RPL_THEME4",
    "SVI_National_RPL_THEMES",
    "SVI_State_RPL_THEME1",
    "SVI_State_RPL_THEME2",
    "SVI_State_RPL_THEME3",
    "SVI_State_RPL_THEME4",
    "SVI_State_RPL_THEMES",
    "VVI",
    "Economic",
    "Education",
    "HealthCareAccess",
    "Neighborhood",
    "Housing",
    "CleanEnvironment",
    "Social",
    "Transportation",
    "PublicSafety",
    "VVI Risk Category",
    "PrimaryRUCA",
    "PrimaryRUCADescription",
    "SecondaryRUCA",
    "SecondaryRUCADescription",
    "2019-2023 Distress Score",
    "Quintile (5=Distressed)",
    "Risk Categorization",
    "RPL_SER",
    "RPL_EJI",
    "RPL_EBM",
    "RPL_SVM",
    "RPL_HVM",
    "RPL_CBM",
    "RPL_EJI_CBM",
    "PRED3_PE",
    "PRED12_PE",
    "PRED0_PE",
    "LILATracts_1And10",
    "LILATracts_Vehicle",
    "LowIncomeTracts",
    "RISK_SCORE",
    "RISK_RATNG",
    "EAL_SCORE",
    "SOVI_SCORE",
    "RESL_SCORE",
    "HRSA_MUA_CENSUS_TRACT",
    "POS_DIST_ED_TRACT",
    "POS_DIST_CLINIC_TRACT",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT",
    "POS_DIST_MEDSURG_ICU_TRACT",
    "POS_DIST_TRAUMA_TRACT",
    "NatWalkInd",
    "D2A_Ranked",
    "D2B_Ranked",
    "D3B_Ranked",
    "D4A_Ranked",
    "Geocode Status",
    "Geocode Method",
    "Geocode Match Type",
    "Geography Level Reached",
    "Geography Complete Flag",
    "Tract-BlockGroup Consistency Flag",
    "Index Coverage Count",
    "ADI Source",
    "ADI Source Year",
    "ADI Geography Used",
    "ADI Missing Reason",
    "SVI Source Year",
    "SVI Geography Used",
    "SVI Missing Reason",
    "VVI Source Year",
    "VVI Geography Used",
    "VVI Missing Reason",
    "RUCA Source Year",
    "RUCA Geography Used",
    "RUCA Missing Reason",
    "DCI Source Year",
    "DCI Geography Used",
    "DCI Missing Reason",
    "EJI Source Year",
    "EJI Geography Used",
    "EJI Missing Reason",
    "CRE Source Year",
    "CRE Geography Used",
    "CRE Missing Reason",
    "FARA Source Year",
    "FARA Geography Used",
    "FARA Missing Reason",
    "NRI Source Year",
    "NRI Geography Used",
    "NRI Missing Reason",
    "AHRF Source Year",
    "AHRF Geography Used",
    "AHRF Missing Reason",
    "Walkability Source Year",
    "Walkability Geography Used",
    "Walkability Missing Reason",
]

MAIN_FIELD_DEFINITIONS = {
    "Unique ID": "The row or patient identifier from your input file.",
    "Address (as inputted)": "The original address exactly as it was provided in the input file.",
    "Geocoded Address": "The standardized address returned by the geocoder after it interpreted the input address.",
    "State / Territory": "The U.S. state or territory linked to the geocoded address.",
    "Census Tract (FIPS)": "The census tract code for the address. This is a neighborhood-sized Census geography used for many tract-based indices.",
    "ZCTA (Zip Code)": "The ZIP Code Tabulation Area, which is the Census version of a ZIP code area.",
    "Block Group (FIPS)": "The census block group code for the address. This is smaller than a tract and is used for some block-group-level measures.",
    "Latitude": "The north-south map coordinate for the geocoded address.",
    "Longitude": "The east-west map coordinate for the geocoded address.",
    "ADI_NATRANK": "Area Deprivation Index national rank. Higher values generally mean more neighborhood disadvantage compared with the U.S. as a whole.",
    "ADI_STATERNK": "Area Deprivation Index state rank. Higher values generally mean more neighborhood disadvantage compared with places in the same state.",
    "National Risk Category": "Plain-language category for the national ADI score.",
    "State Risk Category": "Plain-language category for the state ADI score.",
    "SVI_National_RPL_THEME1": "Social Vulnerability Index national percentile rank for socioeconomic status.",
    "SVI_National_RPL_THEME2": "Social Vulnerability Index national percentile rank for household characteristics.",
    "SVI_National_RPL_THEME3": "Social Vulnerability Index national percentile rank for racial/ethnic minority status and language.",
    "SVI_National_RPL_THEME4": "Social Vulnerability Index national percentile rank for housing type and transportation.",
    "SVI_National_RPL_THEMES": "Overall national Social Vulnerability Index percentile rank. Higher values mean greater vulnerability.",
    "SVI_State_RPL_THEME1": "State-level Social Vulnerability Index percentile rank for socioeconomic status.",
    "SVI_State_RPL_THEME2": "State-level Social Vulnerability Index percentile rank for household characteristics.",
    "SVI_State_RPL_THEME3": "State-level Social Vulnerability Index percentile rank for racial/ethnic minority status and language.",
    "SVI_State_RPL_THEME4": "State-level Social Vulnerability Index percentile rank for housing type and transportation.",
    "SVI_State_RPL_THEMES": "Overall state-level Social Vulnerability Index percentile rank. Higher values mean greater vulnerability within that state.",
    "VVI": "Vaccine Vulnerability Index overall score. Higher values generally mean more vulnerability.",
    "Economic": "VVI economic subscore.",
    "Education": "VVI education subscore.",
    "HealthCareAccess": "VVI healthcare access subscore.",
    "Neighborhood": "VVI neighborhood subscore.",
    "Housing": "VVI housing subscore.",
    "CleanEnvironment": "VVI clean environment subscore.",
    "Social": "VVI social subscore.",
    "Transportation": "VVI transportation subscore.",
    "PublicSafety": "VVI public safety subscore.",
    "VVI Risk Category": "Plain-language category for the overall VVI score.",
    "PrimaryRUCA": "Primary Rural-Urban Commuting Area code for the location.",
    "PrimaryRUCADescription": "Plain-language description of the primary RUCA code.",
    "SecondaryRUCA": "Secondary Rural-Urban Commuting Area code for the location.",
    "SecondaryRUCADescription": "Plain-language description of the secondary RUCA code.",
    "2019-2023 Distress Score": "Distressed Communities Index score for the ZIP area. Higher values generally mean more economic distress.",
    "Quintile (5=Distressed)": "Distressed Communities Index quintile. A value of 5 means the area is in the most distressed group.",
    "Risk Categorization": "Plain-language category for the Distressed Communities Index quintile.",
    "RPL_SER": "EJI percentile rank for combined social and environmental burden. Higher values mean more burden.",
    "RPL_EJI": "EJI percentile rank for combined social, environmental, and health burden. Higher values mean more burden.",
    "RPL_EBM": "EJI percentile rank for environmental burden only.",
    "RPL_SVM": "EJI percentile rank for social vulnerability only.",
    "RPL_HVM": "EJI percentile rank for health vulnerability only.",
    "RPL_CBM": "EJI percentile rank for climate burden only.",
    "RPL_EJI_CBM": "EJI percentile rank for combined social, environmental, health, and climate burden.",
    "PRED3_PE": "CRE estimate of predicted 3-year prevalence or burden for this tract, expressed as a percent-style measure from the source file.",
    "PRED12_PE": "CRE estimate of predicted 12-year prevalence or burden for this tract, expressed as a percent-style measure from the source file.",
    "PRED0_PE": "CRE baseline predicted prevalence or burden for this tract, expressed as a percent-style measure from the source file.",
    "LILATracts_1And10": "FARA flag indicating a low-income, low-access tract using the 1-mile urban and 10-mile rural definition.",
    "LILATracts_Vehicle": "FARA flag indicating a low-income, low-access tract using the vehicle-access definition.",
    "LowIncomeTracts": "FARA flag indicating the tract is classified as low income.",
    "RISK_SCORE": "NRI overall risk score. Higher values generally mean greater overall natural hazard risk.",
    "RISK_RATNG": "Plain-language natural hazard risk rating from the NRI source.",
    "EAL_SCORE": "NRI expected annual loss score. Higher values mean more expected yearly loss from hazards.",
    "SOVI_SCORE": "NRI social vulnerability score. Higher values mean greater social vulnerability.",
    "RESL_SCORE": "NRI community resilience score. Higher values generally mean more resilience or capacity to recover.",
    "NatWalkInd": "National Walkability Index summary score for the Census block group. Higher values generally mean a more walkable built environment.",
    "D2A_Ranked": "Walkability submetric rank related to household mix and neighborhood composition from the national walkability source.",
    "D2B_Ranked": "Walkability submetric rank related to employment mix and destination diversity from the national walkability source.",
    "D3B_Ranked": "Walkability submetric rank related to street intersection density and connectivity.",
    "D4A_Ranked": "Walkability submetric rank related to transit or destination accessibility from the source framework.",
    "Walkability Source Year": "The year associated with the walkability source, if available from the file name.",
    "Walkability Geography Used": "The geography level used to assign walkability, which is expected to be census block group.",
    "Walkability Missing Reason": "If walkability is missing, this explains the most likely reason.",
    "EJI Source Year": "The year associated with the EJI source file.",
    "EJI Geography Used": "The geography level used to assign EJI.",
    "EJI Missing Reason": "If EJI is missing, this explains the most likely reason.",
    "CRE Source Year": "The year associated with the CRE source file.",
    "CRE Geography Used": "The geography level used to assign CRE.",
    "CRE Missing Reason": "If CRE is missing, this explains the most likely reason.",
    "FARA Source Year": "The year associated with the FARA source file.",
    "FARA Geography Used": "The geography level used to assign FARA.",
    "FARA Missing Reason": "If FARA is missing, this explains the most likely reason.",
    "NRI Source Year": "The year associated with the NRI source file.",
    "NRI Geography Used": "The geography level used to assign NRI.",
    "NRI Missing Reason": "If NRI is missing, this explains the most likely reason.",
    "ADI Source": "The source file or source type used for ADI.",
    "ADI Source Year": "The year associated with the ADI source file.",
    "ADI Geography Used": "The geography level used to assign ADI, such as block group.",
    "SVI Source Year": "The year associated with the SVI source file.",
    "SVI Geography Used": "The geography level used to assign SVI, such as tract or ZIP.",
    "VVI Source Year": "The year associated with the VVI source file.",
    "VVI Geography Used": "The geography level used to assign VVI.",
    "RUCA Source Year": "The year associated with the RUCA source file.",
    "RUCA Geography Used": "The geography level used to assign RUCA.",
    "DCI Source Year": "The year associated with the DCI source file.",
    "DCI Geography Used": "The geography level used to assign DCI.",
    "Geocode Status": "Whether geocoding succeeded, failed, or used fallback logic.",
    "Geocode Method": "How the geography was obtained, such as live Census geocoding, cache, or offline parsing.",
    "Geocode Match Type": "A simple label describing the type of geocoding match that occurred.",
    "Geography Level Reached": "The most detailed geography level successfully identified for that row.",
    "Geography Complete Flag": "Indicates whether the row has the key geography fields needed for downstream linkage.",
    "Tract-BlockGroup Consistency Flag": "Checks whether the block group belongs to the same tract it is paired with.",
    "Index Coverage Count": "How many major index families successfully populated for that row.",
    "ADI Missing Reason": "If ADI is missing, this explains the most likely reason.",
    "SVI Missing Reason": "If SVI is missing, this explains the most likely reason.",
    "VVI Missing Reason": "If VVI is missing, this explains the most likely reason.",
    "RUCA Missing Reason": "If RUCA is missing, this explains the most likely reason.",
    "DCI Missing Reason": "If DCI is missing, this explains the most likely reason.",
}

MAIN_FIELD_DEFINITIONS.update({
    "HRSA_MUA_CENSUS_TRACT": "Indicates whether the census tract falls within a HRSA-designated Medically Underserved Area or Population, meaning the area is recognized as having relatively limited primary care resources.",
    "POS_DIST_ED_TRACT": "Estimated distance in miles from the census tract to the nearest emergency department.",
    "POS_DIST_CLINIC_TRACT": "Estimated distance in miles from the census tract to the nearest clinic or outpatient care site, such as a community clinic or similar ambulatory care location.",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT": "Percentage of clinicians or services in the area that accepted Medicare's approved amount in full, used here as a marker of how accessible care may be for Medicare patients without extra balance billing.",
    "POS_DIST_MEDSURG_ICU_TRACT": "Estimated distance in miles from the census tract to the nearest hospital with a medical-surgical intensive care unit.",
    "POS_DIST_TRAUMA_TRACT": "Estimated distance in miles from the census tract to the nearest trauma center.",
    "AHRF Source Year": "Year of the AHRF source file used to populate these AHRF variables for the record.",
    "AHRF Geography Used": "Geographic level used to attach AHRF data to this record, such as census tract or block group.",
    "AHRF Missing Reason": "If AHRF is missing, this explains the most likely reason, such as missing geography or no source match."
})

def build_main_definitions_sheet() -> pd.DataFrame:
    rows = [
        {
            "Field Shown in Main": prettify_column_label(col),
            "Technical Field Name": col,
            "Plain-English Meaning": MAIN_FIELD_DEFINITIONS.get(col, "Definition not yet added."),
        }
        for col in MAIN_OUTPUT_COLUMN_ORDER
    ]
    return pd.DataFrame(rows)

# ------------------------------------------------------------------
# 1. Data classes
# ------------------------------------------------------------------
@dataclass
class ProjectPaths:
    root: Path
    svi_us_zip: Optional[Path]
    svi_us_ct: Optional[Path]
    svi_state_ct_dir: Optional[Path]
    svi_state_zip_dir: Optional[Path]
    vvi_ct: Optional[Path]
    vvi_zip: Optional[Path]
    ruca_ct: Optional[Path]
    ruca_zip: Optional[Path]
    dci_zip: Optional[Path]
    adi_blockgroup: Optional[Path]
    ahrf_blockgroup: Optional[Path]
    ahrf_tract: Optional[Path]
    coi_domains: Optional[Path]
    coi_subdomains: Optional[Path]
    cre_tract: Optional[Path]
    eji_tract: Optional[Path]
    fara_tract: Optional[Path]
    fara_lookup: Optional[Path]
    nri_tract: Optional[Path]
    walkability_gdb: Optional[Path]

# ------------------------------------------------------------------
# 2. Generic helpers
# ------------------------------------------------------------------
def normalize_name(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())

def normalize_state_token(value: object) -> str:
    token = normalize_name(value)
    return token.replace("zcta", "")

def to_tract11(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.split(".")[0]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    return text.zfill(11) if len(text) <= 11 else text[:11]

def to_blockgroup12(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.split(".")[0]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    return text.zfill(12) if len(text) <= 12 else text[:12]

def to_geoid15(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.split(".")[0]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    return text.zfill(15) if len(text) <= 15 else text[:15]

def geoid_like_to_tract11(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.split(".")[0]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    return text[-11:] if len(text) >= 11 else text.zfill(11)

def to_zip5(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"(\d{5})(?:-\d{4})?", text)
    if not match:
        return None
    return match.group(1)

def extract_zipcode_from_text(text: object) -> Optional[str]:
    if pd.isna(text):
        return None
    text = str(text).strip()
    if not text:
        return None

    end_patterns = [
        r"\b(\d{5})(?:-\d{4})?\b(?:\s*,?\s*(?:USA|United States))?\s*$",
        r"\b(?:AL|AK|AS|AZ|AR|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MP|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|PR|RI|SC|SD|TN|TX|UM|UT|VA|VI|VT|WA|WI|WV|WY)\s+(\d{5})(?:-\d{4})?\b(?:\s*,?\s*(?:USA|United States))?\s*$",
    ]
    for pat in end_patterns:
        match = re.search(pat, text, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    return None

def parse_state_from_address(text: object) -> tuple[Optional[str], Optional[str]]:
    if pd.isna(text):
        return None, None
    text = str(text).strip()
    if not text:
        return None, None

    abbr_match = re.search(
        r"\b(AL|AK|AS|AZ|AR|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MP|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|PR|RI|SC|SD|TN|TX|UM|UT|VA|VI|VT|WA|WI|WV|WY)\b\s+\d{5}(?:-\d{4})?\b",
        text,
        flags=re.IGNORECASE,
    )
    if abbr_match:
        abbr = abbr_match.group(1).upper()
        return STATE_ABBR_TO_NAME.get(abbr), abbr

    name_match = re.search(r",\s*([A-Za-z .]+?)\s+\d{5}(?:-\d{4})?\b", text)
    if name_match:
        candidate = re.sub(r"\s+", " ", name_match.group(1)).strip().title()
        if candidate in STATE_NAME_TO_ABBR:
            return candidate, STATE_NAME_TO_ABBR[candidate]

    return None, None

def looks_like_year(value: object) -> bool:
    if pd.isna(value):
        return False
    text = str(value).strip().split(".")[0]
    return text.isdigit() and 1900 <= int(text) <= 2100

def looks_like_address_text(value: object) -> bool:
    if pd.isna(value):
        return False
    text = str(value).strip()
    if not text:
        return False
    score = 0
    if extract_zipcode_from_text(text):
        score += 2
    if "," in text:
        score += 1
    if re.search(r"\d", text):
        score += 1
    if re.search(r"[A-Za-z]", text):
        score += 1
    return score >= 3

def series_match_fraction(series: pd.Series, parser) -> float:
    sample = series.dropna().astype(str).head(100)
    if sample.empty:
        return 0.0
    return float(sample.apply(lambda x: parser(x) is not None).mean())

def deduplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    counts: dict[str, int] = {}
    new_cols: list[str] = []
    for col in df.columns:
        base = str(col)
        if base not in counts:
            counts[base] = 0
            new_cols.append(base)
        else:
            counts[base] += 1
            new_cols.append(f"{base}_{counts[base]}")
    df = df.copy()
    df.columns = new_cols
    return df

def find_column_by_name(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    normalized_map = {normalize_name(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_name(candidate)
        if key in normalized_map:
            return normalized_map[key]
    for candidate in candidates:
        key = normalize_name(candidate)
        for norm_col, original_col in normalized_map.items():
            if key and (key in norm_col or norm_col in key):
                return original_col
    return None

def first_existing(paths: Iterable[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None

def first_existing_prefer_non_numbers(paths: Iterable[Path]) -> Optional[Path]:
    existing = [path for path in paths if path.exists()]
    if not existing:
        return None
    non_numbers = [path for path in existing if path.suffix.lower() != ".numbers"]
    return non_numbers[0] if non_numbers else existing[0]

def safe_float(value: object) -> Optional[float]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None

def normalize_address_for_cache(address: object) -> str:
    if pd.isna(address):
        return ""
    return re.sub(r"\s+", " ", str(address).strip()).upper()

def harmonize_geocode_record(record: dict[str, object]) -> dict[str, object]:
    out = dict(record) if isinstance(record, dict) else {}
    out.setdefault("geocoded_address", None)
    out["geoid"] = to_geoid15(out.get("geoid"))
    out["latitude"] = safe_float(out.get("latitude"))
    out["longitude"] = safe_float(out.get("longitude"))
    out["census_tract"] = to_tract11(out.get("census_tract"))
    out["blockgroup_fips"] = to_blockgroup12(out.get("blockgroup_fips"))
    out["county_fips"] = str(out.get("county_fips")).zfill(5) if out.get("county_fips") not in [None, "", float('nan')] else None
    out["state_fips"] = str(out.get("state_fips")).zfill(2) if out.get("state_fips") not in [None, "", float('nan')] else None
    out["state_name"] = out.get("state_name") or None
    out["state_abbr"] = out.get("state_abbr") or None
    out["matched_zipcode"] = to_zip5(out.get("matched_zipcode"))
    out["zipcode"] = to_zip5(out.get("zipcode"))
    out["geocode_status"] = out.get("geocode_status") or None
    out["geocode_source"] = out.get("geocode_source") or "cache"
    if not out.get("blockgroup_fips") and out.get("geoid"):
        out["blockgroup_fips"] = to_blockgroup12(out.get("geoid"))
    if not out.get("census_tract") and out.get("geoid"):
        out["census_tract"] = to_tract11(out.get("geoid"))
    if not out.get("county_fips") and out.get("geoid"):
        g = to_geoid15(out.get("geoid"))
        out["county_fips"] = g[:5] if g else None
    if not out.get("state_fips"):
        if out.get("geoid"):
            g = to_geoid15(out.get("geoid"))
            out["state_fips"] = g[:2] if g else None
        elif out.get("blockgroup_fips"):
            out["state_fips"] = str(out["blockgroup_fips"])[:2]
        elif out.get("census_tract"):
            out["state_fips"] = str(out["census_tract"])[:2]
    if not out.get("state_name") and out.get("state_fips"):
        out["state_name"] = STATE_FIPS_TO_NAME.get(out["state_fips"])
    if not out.get("state_abbr") and out.get("state_name"):
        out["state_abbr"] = STATE_NAME_TO_ABBR.get(out["state_name"])
    out["geocode_method"] = out.get("geocode_method") or out.get("geocode_source") or "cache"
    out["geocode_match_type"] = out.get("geocode_match_type") or infer_geocode_match_type(out.get("geocode_status"), out.get("geocode_source"))
    out["geography_level_reached"] = out.get("geography_level_reached") or infer_geography_level_reached(out)
    return out

def load_geocode_cache(cache_path: Path, notes: list[str]) -> dict[str, dict[str, object]]:
    if not cache_path.exists():
        return {}
    try:
        payload = json.loads(cache_path.read_text())
        if isinstance(payload, dict):
            return {k: harmonize_geocode_record(v) for k, v in payload.items()}
    except Exception as exc:
        notes.append(f"Could not read geocode cache {cache_path.name}: {exc}")
    return {}

def save_geocode_cache(cache_path: Path, cache: dict[str, dict[str, object]], notes: list[str]) -> None:
    try:
        cache_path.write_text(json.dumps(cache, indent=2, sort_keys=True))
    except Exception as exc:
        notes.append(f"Could not write geocode cache {cache_path.name}: {exc}")

def offline_geocode_record(address: str) -> dict[str, object]:
    raw_zip = extract_zipcode_from_text(address)
    state_name, state_abbr = parse_state_from_address(address)
    state_fips = None
    if state_name:
        for fips, name in STATE_FIPS_TO_NAME.items():
            if name == state_name:
                state_fips = fips
                break
    status = "offline_fallback" if address else "blank_address"
    source = "offline_fallback" if address else "blank"
    return {
        "geocoded_address": None,
        "geoid": None,
        "latitude": None,
        "longitude": None,
        "census_tract": None,
        "blockgroup_fips": None,
        "county_fips": None,
        "state_fips": state_fips,
        "state_name": state_name,
        "state_abbr": state_abbr,
        "matched_zipcode": raw_zip,
        "zipcode": raw_zip,
        "geocode_status": status,
        "geocode_source": source,
        "geocode_method": source,
        "geocode_match_type": infer_geocode_match_type(status, source),
        "geography_level_reached": "state_zip" if (state_fips or raw_zip) else "none",
    }

# ------------------------------------------------------------------
# 3. Input loader
# ------------------------------------------------------------------
def detect_input_columns(raw: pd.DataFrame) -> tuple[int, int]:
    best_addr_col = None
    best_addr_score = -1.0
    for idx in raw.columns:
        series = raw[idx].dropna().astype(str)
        if series.empty:
            continue
        score = float(series.head(200).apply(looks_like_address_text).mean())
        if score > best_addr_score:
            best_addr_score = score
            best_addr_col = idx

    if best_addr_col is None or best_addr_score < 0.20:
        return raw.columns[0], raw.columns[1]

    candidate_cols = [c for c in raw.columns if c != best_addr_col]
    left_candidates = [c for c in candidate_cols if c < best_addr_col]
    ordered_candidates = list(reversed(left_candidates)) + [c for c in candidate_cols if c > best_addr_col]

    best_id_col = None
    best_id_score = -1.0
    for idx in ordered_candidates:
        series = raw[idx].dropna().astype(str).str.strip()
        if series.empty:
            continue
        nonempty = series[series != ""]
        if nonempty.empty:
            continue
        uniqueness = float(nonempty.nunique() / len(nonempty))
        not_address = 1.0 - float(nonempty.head(200).apply(looks_like_address_text).mean())
        numericish = float(nonempty.head(200).apply(lambda x: bool(re.fullmatch(r"[A-Za-z0-9\-_]+", x))).mean())
        score = 0.5 * uniqueness + 0.3 * not_address + 0.2 * numericish
        if idx in left_candidates:
            score += 0.15
        if score > best_id_score:
            best_id_score = score
            best_id_col = idx

    if best_id_col is None:
        best_id_col = raw.columns[0] if raw.columns[0] != best_addr_col else raw.columns[1]

    return best_id_col, best_addr_col

def load_input_addresses(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    raw = pd.read_csv(path, dtype=str, header=None) if suffix == ".csv" else pd.read_excel(path, dtype=str, header=None)

    if raw.shape[1] < 2:
        raise ValueError("Input file must have at least 2 usable columns containing UniqueID and Address.")

    id_col, address_col = detect_input_columns(raw)
    df = raw.loc[:, [id_col, address_col]].copy()
    df.columns = ["UniqueID", "Address"]

    first_id = normalize_name(df.iloc[0, 0]) if not df.empty else ""
    first_address = normalize_name(df.iloc[0, 1]) if not df.empty else ""
    if first_id in {"uniqueid", "patientid", "id", "mrn", "studyid"} or first_address in {"address", "patientaddress", "homeaddress"}:
        df = df.iloc[1:].copy()

    df["UniqueID"] = df["UniqueID"].astype(str).replace("nan", "").str.strip()
    df["Address"] = df["Address"].astype(str).replace("nan", "").str.strip()
    df = df[(df["UniqueID"] != "") | (df["Address"] != "")].copy()

    header_id_tokens = {"uniqueid", "patientid", "id", "mrn", "studyid", "record", "row"}
    header_address_tokens = {"address", "patientaddress", "homeaddress", "streetaddress"}

    def is_obvious_header_row(row: pd.Series) -> bool:
        uid = normalize_name(row.get("UniqueID", ""))
        addr = normalize_name(row.get("Address", ""))
        return (uid in header_id_tokens and addr in header_address_tokens) or (uid == "" and addr in {"address", "addresses"})

    df = df[~df.apply(is_obvious_header_row, axis=1)].copy()
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        raise ValueError("No usable rows found in input file.")

    return df

# ------------------------------------------------------------------
# 4. Project-relative discovery and extraction
# ------------------------------------------------------------------
def extract_zip_archives(project_root: Path, notes: list[str]) -> None:
    """Extract ZIP archives anywhere under the project folder/reference_data.

    This matters for Streamlit/GitHub deployments because large reference bundles
    may be downloaded from Google Drive into reference_data/ instead of being
    committed to the repo.
    """
    for zip_path in project_root.rglob("*.zip"):
        if "__MACOSX" in zip_path.parts:
            continue
        target_dir = zip_path.parent / zip_path.stem
        if target_dir.exists():
            continue
        try:
            with zipfile.ZipFile(zip_path) as zf:
                zf.extractall(target_dir)
            notes.append(f"Extracted reference archive: {zip_path.name}")
        except Exception as exc:
            notes.append(f"Failed to extract {zip_path.name}: {exc}")


def find_recursive(project_root: Path, basename: str) -> Optional[Path]:
    matches = [p for p in project_root.rglob(basename) if "__MACOSX" not in p.parts]
    return matches[0] if matches else None

def find_recursive_dir(project_root: Path, dirname: str) -> Optional[Path]:
    matches = [p for p in project_root.rglob(dirname) if p.is_dir() and "__MACOSX" not in p.parts]
    return matches[0] if matches else None

def find_recursive_exact_names(project_root: Path, basenames: list[str]) -> list[Path]:
    wanted = {name.lower() for name in basenames}
    matches = [p for p in project_root.rglob("*") if p.is_file() and "__MACOSX" not in p.parts and p.name.lower() in wanted]
    return sorted(matches, key=lambda p: (len(p.parts), len(str(p))))

def find_recursive_contains(project_root: Path, required_tokens: list[str], suffixes: Optional[list[str]] = None) -> list[Path]:
    tokens = [t.lower() for t in required_tokens]
    suffix_set = {s.lower() for s in suffixes} if suffixes else None
    matches = []
    for p in project_root.rglob("*"):
        if not p.is_file() or "__MACOSX" in p.parts:
            continue
        name = p.name.lower()
        if suffix_set and p.suffix.lower() not in suffix_set:
            continue
        if all(tok in name for tok in tokens):
            matches.append(p)
    return sorted(matches, key=lambda p: (len(p.parts), len(str(p))))

def find_svi_national_ct(project_root: Path) -> Optional[Path]:
    exact = find_recursive_exact_names(project_root, [
        "SVI_2022_US_CT.csv",
        "SVI_2022_US_CT.xlsx",
        "SVI_2022_US_CT.xls",
        "SVI_2022_US_CT.numbers",
    ])
    if exact:
        return first_existing_prefer_non_numbers(exact)
    fuzzy = find_recursive_contains(project_root, ["svi", "2022", "us", "ct"], [".csv", ".xlsx", ".xls", ".numbers"])
    return first_existing_prefer_non_numbers(fuzzy)

def find_ahrf_blockgroup_ref(project_root: Path) -> Optional[Path]:
    exact = find_recursive_exact_names(project_root, [
        "clh_2023_blockgroup_2_0.csv",
        "clh_2023_blockgroup_2_0.xlsx",
        "clh_2023_blockgroup_2_0.xls",
        "clh_2023_blockgroup_2_0.numbers",
    ])
    if exact:
        return first_existing_prefer_non_numbers(exact)
    fuzzy = find_recursive_contains(project_root, ["clh", "2023", "blockgroup", "2_0"], [".csv", ".xlsx", ".xls", ".numbers"])
    return first_existing_prefer_non_numbers(fuzzy)

def find_ahrf_tract_ref(project_root: Path) -> Optional[Path]:
    exact = find_recursive_exact_names(project_root, [
        "clh_2023_tract_2_0.csv",
        "clh_2023_tract_2_0.xlsx",
        "clh_2023_tract_2_0.xls",
        "clh_2023_tract_2_0.numbers",
    ])
    if exact:
        return first_existing_prefer_non_numbers(exact)
    fuzzy = find_recursive_contains(project_root, ["clh", "2023", "tract", "2_0"], [".csv", ".xlsx", ".xls", ".numbers"])
    return first_existing_prefer_non_numbers(fuzzy)

def ensure_reference_files_from_drive(project_root: Path, notes: list[str]) -> None:
    """Download large reference files from the configured Google Drive folder.

    The downloaded files go into reference_data/, which should be listed in
    .gitignore so GitHub only stores code and small config files. The rest of
    this script already searches recursively inside project_root, so once the
    files are downloaded, locate_project_paths() can find them automatically.
    """
    target_dir = project_root / REFERENCE_DATA_DIR
    marker_file = target_dir / REFERENCE_DATA_DOWNLOAD_MARKER

    # Do not re-download on every run/session if the files are already present.
    if marker_file.exists():
        notes.append("Google Drive reference files already downloaded in reference_data/.")
        return

    if not GOOGLE_DRIVE_FOLDER_URL:
        notes.append("No Google Drive folder URL configured for reference files.")
        return

    if gdown is None:
        notes.append(
            "Google Drive reference files were not downloaded because gdown is not installed. "
            "Add gdown to requirements.txt."
        )
        return

    target_dir.mkdir(parents=True, exist_ok=True)
    notes.append("Downloading reference files from Google Drive into reference_data/...")

    try:
        downloaded = gdown.download_folder(
            url=GOOGLE_DRIVE_FOLDER_URL,
            output=str(target_dir),
            quiet=False,
            use_cookies=False,
        )
        marker_file.write_text("downloaded\n")
        count = len(downloaded) if downloaded else 0
        notes.append(f"Downloaded {count} Google Drive reference files into {target_dir}.")
    except Exception as exc:
        notes.append(f"Could not download Google Drive reference files: {exc}")


def locate_project_paths(project_root: Path, notes: list[str]) -> ProjectPaths:
    extract_zip_archives(project_root, notes)

    return ProjectPaths(
        root=project_root,
        svi_us_zip=find_recursive(project_root, "SVI_2022_US_ZipCode.csv"),
        svi_us_ct=find_svi_national_ct(project_root),
        svi_state_ct_dir=find_recursive_dir(project_root, "State - Census Tracts"),
        svi_state_zip_dir=find_recursive_dir(project_root, "State - Zip Codes"),
        vvi_ct=find_recursive(project_root, "VVI_CT_2025.xlsx"),
        vvi_zip=find_recursive(project_root, "VVI_zip_code_2025.xlsx"),
        ruca_ct=find_recursive(project_root, "RUCA-codes-2020-tract.csv"),
        ruca_zip=find_recursive(project_root, "RUCA-codes-2020-zipcode.csv"),
        dci_zip=find_recursive(project_root, "DCI-2019-2023-Scores-Only-11.xlsx"),
        adi_blockgroup=first_existing_prefer_non_numbers([p for p in [
            project_root / "US_2023_ADI_Census_Block_Group_v4_0_1.csv",
            project_root / "US_2023_ADI_Census_Block_Group_v4_0_1.xlsx",
            find_recursive(project_root, "US_2023_ADI_Census_Block_Group_v4_0_1.csv"),
            find_recursive(project_root, "US_2023_ADI_Census_Block_Group_v4_0_1.xlsx"),
            project_root / "US_2023_ADI_Census_Block_Group_v4_0_1.numbers",
            find_recursive(project_root, "US_2023_ADI_Census_Block_Group_v4_0_1.numbers"),
        ] if p is not None]),
        ahrf_blockgroup=find_ahrf_blockgroup_ref(project_root),
        ahrf_tract=find_ahrf_tract_ref(project_root),
        coi_domains=find_recursive(project_root, "data.csv") if find_recursive_dir(project_root, "2020 census tracts (COI 3.0-2023) Domains") else None,
        coi_subdomains=find_recursive(project_root, "data.csv") if find_recursive_dir(project_root, "2020 census tracts (COI 3.0-2023) Subdomains") else None,
        cre_tract=first_existing([p for p in [
            project_root / "CRE_24_Tract.csv",
            find_recursive(project_root, "CRE_24_Tract.csv"),
        ] if p is not None]),
        eji_tract=first_existing([p for p in [
            project_root / "EJI_2024_United_States.csv",
            find_recursive(project_root, "EJI_2024_United_States.csv"),
        ] if p is not None]),
        fara_tract=first_existing([p for p in [
            project_root / "Food Access Research Atlas.csv",
            find_recursive(project_root, "Food Access Research Atlas.csv"),
        ] if p is not None]),
        fara_lookup=first_existing([p for p in [
            project_root / "VariableLookup.csv",
            find_recursive(project_root, "VariableLookup.csv"),
        ] if p is not None]),
        nri_tract=first_existing([p for p in [
            project_root / "NRI_Table_CensusTracts.csv",
            find_recursive(project_root, "NRI_Table_CensusTracts.csv"),
        ] if p is not None]),
        walkability_gdb=first_existing([p for p in [
            find_recursive_dir(project_root, "Natl_WI.gdb"),
            find_recursive_dir(project_root, "National Walkability Index.gdb"),
            find_recursive_dir(project_root, "WalkabilityIndex.gdb"),
        ] if p is not None]),
    )

def patch_special_paths(paths: ProjectPaths) -> ProjectPaths:
    dom_dir = find_recursive_dir(paths.root, "2020 census tracts (COI 3.0-2023) Domains")
    sub_dir = find_recursive_dir(paths.root, "2020 census tracts (COI 3.0-2023) Subdomains")
    if dom_dir:
        paths.coi_domains = dom_dir / "data.csv"
    if sub_dir:
        paths.coi_subdomains = sub_dir / "data.csv"
    return paths

def try_convert_numbers_with_soffice(path: Path) -> Optional[Path]:
    outdir = path.parent
    candidates = [
        outdir / f"{path.stem}.xlsx",
        outdir / f"{path.stem}.csv",
    ]
    existing = first_existing(candidates)
    if existing:
        return existing

    soffice = shutil.which("soffice")
    if not soffice:
        return None

    commands = [
        [soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(outdir), str(path)],
        [soffice, "--headless", "--convert-to", "csv", "--outdir", str(outdir), str(path)],
    ]
    for cmd in commands:
        try:
            subprocess.run(cmd, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=20)
        except Exception:
            pass
        converted = first_existing(candidates)
        if converted:
            return converted
    return None

def resolve_numbers_or_export(path: Optional[Path]) -> Optional[Path]:
    if path is None:
        return None
    if path.suffix.lower() != ".numbers":
        return path if path.exists() else None
    siblings = [
        path.with_suffix(".csv"),
        path.with_suffix(".xlsx"),
        path.with_suffix(".xls"),
        path.parent / f"{path.stem}.csv",
        path.parent / f"{path.stem}.xlsx",
        path.parent / f"{path.stem}.xls",
    ]
    alt = first_existing(siblings)
    if alt:
        return alt
    converted = try_convert_numbers_with_soffice(path)
    return converted if converted else path

# ------------------------------------------------------------------
# 5. Read helpers
# ------------------------------------------------------------------
def read_numbers_table(path: Path) -> pd.DataFrame:
    try:
        from numbers_parser import Document  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            f"{path.name} is an Apple Numbers file. Export it once to CSV/XLSX or install numbers-parser."
        ) from exc

    doc = Document(str(path))
    sheet = doc.sheets[0]
    table = sheet.tables[0]
    rows = list(table.rows(values_only=True))
    if not rows:
        return pd.DataFrame()
    header = [str(x) if x is not None else "" for x in rows[0]]
    body = rows[1:]
    return deduplicate_columns(pd.DataFrame(body, columns=header))

def read_table(path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return deduplicate_columns(pd.read_csv(path, dtype=str, encoding_errors="replace"))
    if path.suffix.lower() == ".numbers":
        return read_numbers_table(path)
    if sheet_name is None:
        workbook = pd.read_excel(path, dtype=str, sheet_name=None)
        first_sheet = next(iter(workbook.values()))
        return deduplicate_columns(first_sheet)
    return deduplicate_columns(pd.read_excel(path, dtype=str, sheet_name=sheet_name))

def filter_dataframe_by_key(df: pd.DataFrame, key_col: str, key_parser, keys: set[str]) -> pd.DataFrame:
    if not keys:
        return df.iloc[0:0].copy()
    out = df.copy()
    out["merge_key"] = out[key_col].apply(key_parser)
    out = out[out["merge_key"].isin(keys)].copy()
    return out

def read_filtered_csv(path: Path, key_col: str, key_parser, keys: set[str], chunksize: int = 200_000) -> pd.DataFrame:
    if not keys:
        return pd.DataFrame(columns=[key_col, "merge_key"])
    chunks: list[pd.DataFrame] = []
    for chunk in pd.read_csv(path, dtype=str, chunksize=chunksize, encoding_errors="replace"):
        chunk = deduplicate_columns(chunk)
        chunk["merge_key"] = chunk[key_col].apply(key_parser)
        hit = chunk[chunk["merge_key"].isin(keys)].copy()
        if not hit.empty:
            chunks.append(hit)
    if not chunks:
        return pd.DataFrame(columns=[key_col, "merge_key"])
    return pd.concat(chunks, ignore_index=True)

def read_filtered_xlsx_data_sheet(path: Path, key_col: str, key_parser, keys: set[str], data_sheet: str = "Data") -> pd.DataFrame:
    if not keys:
        return pd.DataFrame(columns=[key_col, "merge_key"])
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[data_sheet]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    header = [str(x) if x is not None else "" for x in header]
    try:
        key_idx = header.index(key_col)
    except ValueError:
        raise KeyError(f"{key_col} not found in {path.name} sheet {data_sheet}")

    records = []
    for row in rows:
        raw_key = row[key_idx]
        merge_key = key_parser(raw_key)
        if merge_key in keys:
            record = {header[i]: row[i] for i in range(len(header))}
            record["merge_key"] = merge_key
            records.append(record)
    return pd.DataFrame(records, columns=header + ["merge_key"])

def find_key_column(df: pd.DataFrame, level: str) -> str:
    if level == "tract":
        by_name = find_column_by_name(df, TRACT_COLUMN_CANDIDATES)
        if by_name:
            return by_name
        best_col, best_score = None, -1.0
        for col in df.columns:
            score = series_match_fraction(df[col], to_tract11)
            if score > best_score:
                best_col, best_score = col, score
        return best_col or df.columns[0]
    if level == "zip":
        by_name = find_column_by_name(df, ZIP_COLUMN_CANDIDATES)
        if by_name:
            return by_name
        best_col, best_score = None, -1.0
        for col in df.columns:
            score = series_match_fraction(df[col], to_zip5)
            if score > best_score:
                best_col, best_score = col, score
        return best_col or df.columns[0]
    if level == "blockgroup":
        by_name = find_column_by_name(df, BLOCKGROUP_COLUMN_CANDIDATES)
        if by_name:
            return by_name
        best_col, best_score = None, -1.0
        for col in df.columns:
            score = series_match_fraction(df[col], to_blockgroup12)
            if score > best_score:
                best_col, best_score = col, score
        return best_col or df.columns[0]
    return df.columns[0]

def prefix_reference(df: pd.DataFrame, prefix: str, protected: Optional[set[str]] = None) -> pd.DataFrame:
    protected = protected or {"merge_key"}
    out = df.copy()
    rename_map = {col: f"{prefix}{col}" for col in out.columns if col not in protected}
    out = out.rename(columns=rename_map)
    return out

def strip_prefix_columns(columns: Iterable[str], prefix: str) -> list[str]:
    return [str(col)[len(prefix):] for col in columns if str(col).startswith(prefix)]

def combine_by_mask(preferred: pd.Series, fallback: pd.Series, use_preferred: pd.Series, use_fallback: Optional[pd.Series] = None) -> pd.Series:
    out = pd.Series(index=preferred.index, dtype=object)
    out = preferred.where(use_preferred)
    if use_fallback is None:
        use_fallback = ~use_preferred
    out = out.where(~use_fallback, fallback)
    return out

def ensure_output_columns(df: pd.DataFrame, ordered_columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in ordered_columns:
        if col not in out.columns:
            out[col] = pd.NA
    return out[ordered_columns]

def extract_year_from_path(path: Optional[Path]) -> Optional[str]:
    if not path:
        return None
    matches = re.findall(r"(?:19|20)\d{2}", path.name)
    return matches[-1] if matches else None

def infer_geography_level_reached(record: dict[str, object]) -> str:
    if to_geoid15(record.get("geoid")):
        return "block"
    if to_blockgroup12(record.get("blockgroup_fips")):
        return "blockgroup"
    if to_tract11(record.get("census_tract")):
        return "tract"
    if (record.get("state_fips") or record.get("state_abbr") or record.get("state_name")) and to_zip5(record.get("zipcode")):
        return "state_zip"
    if (record.get("state_fips") or record.get("state_abbr") or record.get("state_name")):
        return "state_only"
    if to_zip5(record.get("zipcode")):
        return "zip_only"
    return "none"

def infer_geocode_match_type(status: object, source: object) -> str:
    status_norm = normalize_name(status)
    source_norm = normalize_name(source)
    if status_norm == "blank_address" or source_norm == "blank":
        return "blank_address"
    if source_norm == "offline_fallback":
        return "parsed_state_zip"
    if source_norm == "census_live" and status_norm == "ok":
        return "census_top_match"
    if status_norm == "no_match":
        return "no_match"
    if status_norm == "request_error":
        return "request_error"
    if status_norm == "parse_error":
        return "parse_error"
    if source_norm == "cache":
        return "cache"
    return status_norm or source_norm or "unknown"

def read_table_header_only(path: Optional[Path], sheet_name: Optional[str] = None) -> list[str]:
    if not path or not path.exists():
        return []
    try:
        suffix = path.suffix.lower()
        if suffix == ".numbers":
            return []
        if suffix == ".csv":
            return list(pd.read_csv(path, nrows=0).columns)
        if suffix in {".xlsx", ".xls", ".xlsm"}:
            return list(pd.read_excel(path, sheet_name=sheet_name or 0, nrows=0).columns)
    except Exception:
        return []
    return []

def find_walkability_csv(project_root: Path) -> Optional[Path]:
    exact = find_recursive_exact_names(project_root, [
        "walkability_full_attributes.csv",
        "NationalWalkabilityIndex.csv",
        "National Walkability Index.csv",
        "Natl_WI.csv",
        "Natl_WI_Table.csv",
        "walkability.csv",
    ])
    if exact:
        return exact[0]
    fuzzy = find_recursive_contains(project_root, ["walkability"], [".csv"])
    if fuzzy:
        return fuzzy[0]
    return None

def validate_project_paths(paths: ProjectPaths, notes: list[str]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    def add_file_row(label: str, path: Optional[Path], expected_candidates: Optional[list[str]] = None, sheet_name: Optional[str] = None) -> None:
        if not path or not path.exists():
            rows.append({"Source": label, "Status": "Missing", "Path": "", "Detail": "not found"})
            return
        cols = read_table_header_only(path, sheet_name=sheet_name)
        detail = path.name
        if expected_candidates:
            found = [c for c in cols if normalize_name(c) in {normalize_name(x) for x in expected_candidates}]
            if found:
                detail = f"{path.name} | key cols: {', '.join(found[:4])}"
            elif cols:
                detail = f"{path.name} | header read, expected key column not confirmed"
            else:
                detail = f"{path.name} | header unavailable"
        rows.append({"Source": label, "Status": "OK", "Path": str(path), "Detail": detail})

    def add_dir_row(label: str, directory: Optional[Path]) -> None:
        if not directory or not directory.exists():
            rows.append({"Source": label, "Status": "Missing", "Path": "", "Detail": "not found"})
            return
        file_count = len([p for p in directory.rglob("*") if p.is_file() and not p.name.startswith("~$")])
        rows.append({"Source": label, "Status": "OK", "Path": str(directory), "Detail": f"{file_count} files"})

    add_file_row("ADI blockgroup", paths.adi_blockgroup, BLOCKGROUP_COLUMN_CANDIDATES)
    add_file_row("DCI ZIP", paths.dci_zip, ZIP_COLUMN_CANDIDATES)
    add_file_row("RUCA tract", paths.ruca_ct, TRACT_COLUMN_CANDIDATES)
    add_file_row("RUCA ZIP", paths.ruca_zip, ZIP_COLUMN_CANDIDATES)
    add_file_row("VVI tract", paths.vvi_ct, TRACT_COLUMN_CANDIDATES)
    add_file_row("VVI ZIP", paths.vvi_zip, ZIP_COLUMN_CANDIDATES)
    add_file_row("SVI national tract", paths.svi_us_ct, TRACT_COLUMN_CANDIDATES)
    add_file_row("SVI national ZIP", paths.svi_us_zip, ZIP_COLUMN_CANDIDATES)
    add_dir_row("SVI state tract directory", paths.svi_state_ct_dir)
    add_dir_row("SVI state ZIP directory", paths.svi_state_zip_dir)
    add_file_row("COI domains", paths.coi_domains, TRACT_COLUMN_CANDIDATES)
    add_file_row("COI subdomains", paths.coi_subdomains, TRACT_COLUMN_CANDIDATES)
    add_file_row("AHRF blockgroup", paths.ahrf_blockgroup, BLOCKGROUP_COLUMN_CANDIDATES)
    add_file_row("AHRF tract", paths.ahrf_tract, TRACT_COLUMN_CANDIDATES)
    add_file_row("CRE tract", paths.cre_tract, ["GEO_ID", "STATE", "COUNTY", "TRACT"])
    add_file_row("EJI tract", paths.eji_tract, ["GEOID", "AFFGEOID", "STATEFP", "COUNTYFP", "TRACTCE"])
    add_file_row("FARA tract", paths.fara_tract, ["CensusTract", "State", "County"])
    add_file_row("FARA lookup", paths.fara_lookup, ["Field", "LongName", "Description"])
    add_file_row("NRI tract", paths.nri_tract, ["TRACTFIPS", "RISK_SCORE", "RISK_RATNG", "EAL_SCORE", "SOVI_SCORE", "RESL_SCORE"])
    walk_csv = find_walkability_csv(paths.root)
    if walk_csv:
        add_file_row("Walkability CSV", walk_csv, ["GEOID10", "GEOID20", "NatWalkInd", "D2A_Ranked"])
    add_dir_row("Walkability geodatabase", paths.walkability_gdb)

    validation_df = pd.DataFrame(rows)
    missing_count = int((validation_df["Status"] == "Missing").sum()) if not validation_df.empty else 0
    notes.append(f"Reference validation completed: {len(validation_df) - missing_count} sources available, {missing_count} missing.")
    return validation_df

def build_summary_sheet(main_df: pd.DataFrame, validation_df: pd.DataFrame, notes: list[str]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    def is_effectively_missing(series: pd.Series) -> pd.Series:
        if series.dtype == object:
            return series.isna() | series.astype(str).str.strip().eq("")
        return series.isna()

    for order, col in enumerate(MAIN_OUTPUT_COLUMN_ORDER, start=1):
        series = main_df[col] if col in main_df.columns else pd.Series([pd.NA] * len(main_df), dtype="object")
        missing_mask = is_effectively_missing(series)
        populated = int((~missing_mask).sum())
        missing = int(missing_mask.sum())
        numeric = pd.to_numeric(series, errors="coerce").dropna()
        top_values = pd.NA
        if numeric.empty:
            vc = series[~missing_mask].astype(str).value_counts(dropna=False)
            if not vc.empty:
                top_values = "; ".join([f"{idx} ({int(val)})" for idx, val in vc.head(3).items()])
        rows.append({
            "Order": order,
            "Field in Main": prettify_column_label(col),
            "Technical Field Name": col,
            "Plain-English Meaning": MAIN_FIELD_DEFINITIONS.get(col, "Definition not yet added."),
            "Populated Rows": populated,
            "Missing Rows": missing,
            "Mean": round(float(numeric.mean()), 6) if not numeric.empty else pd.NA,
            "Min": round(float(numeric.min()), 6) if not numeric.empty else pd.NA,
            "Max": round(float(numeric.max()), 6) if not numeric.empty else pd.NA,
            "Median": round(float(numeric.median()), 6) if not numeric.empty else pd.NA,
            "IQR": round(float(numeric.quantile(0.75) - numeric.quantile(0.25)), 6) if not numeric.empty else pd.NA,
            "Top Values / Notes": top_values,
        })
    return pd.DataFrame(rows)

def build_run_audit_sheet(main_df: pd.DataFrame, validation_df: pd.DataFrame, notes: list[str]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    def add(section: str, metric: str, value: object) -> None:
        rows.append({"Section": section, "Metric": metric, "Value": value})

    add("Run", "Rows processed", len(main_df))
    add("Run", "Unique IDs", int(main_df["Unique ID"].nunique()) if "Unique ID" in main_df.columns else len(main_df))
    add("Run", "Unique nonblank addresses", int(main_df["Address (as inputted)"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if "Address (as inputted)" in main_df.columns else pd.NA)

    for col in ["Geocode Status", "Geocode Method", "Geography Level Reached"]:
        if col in main_df.columns:
            counts = main_df[col].fillna("<blank>").value_counts(dropna=False)
            for key, value in counts.items():
                add(col, key, int(value))

    if "Walkability Missing Reason" in main_df.columns:
        counts = main_df["Walkability Missing Reason"].fillna("<populated>").value_counts(dropna=False)
        for key, value in counts.items():
            add("Walkability status", key, int(value))

    if "State / Territory" in main_df.columns:
        state_counts = main_df["State / Territory"].fillna("<blank>").value_counts(dropna=False).head(15)
        for key, value in state_counts.items():
            add("Top states/territories", key, int(value))

    if validation_df is not None and not validation_df.empty:
        for _, row in validation_df.iterrows():
            add("Reference validation", row["Source"], f"{row['Status']} | {row['Detail']}")

    if notes:
        for idx, note in enumerate(notes, start=1):
            add("Notes", f"Note {idx}", note)

    return pd.DataFrame(rows)

# ------------------------------------------------------------------
# 6. Geocoding
# ------------------------------------------------------------------

def get_coordinates(session: requests.Session, address: str) -> tuple[Optional[float], Optional[float], Optional[str], str]:
    url = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress"
    params = {
        "address": f"{address}, USA" if "usa" not in address.lower() else address,
        "benchmark": GEOCODER_BENCHMARK,
        "format": "json",
    }
    try:
        response = session.get(url, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
        response.raise_for_status()
        payload = response.json()
        matches = payload.get("result", {}).get("addressMatches", [])
        if not matches:
            return None, None, None, "no_match"
        match = matches[0]
        coords = match.get("coordinates", {})
        return coords.get("y"), coords.get("x"), match.get("matchedAddress"), "ok"
    except requests.RequestException:
        return None, None, None, "request_error"
    except Exception:
        return None, None, None, "parse_error"

def get_census_geographies(session: requests.Session, lat: float, lon: float) -> dict[str, Optional[str]]:
    url = "https://geocoding.geo.census.gov/geocoder/geographies/coordinates"
    params = {
        "x": lon,
        "y": lat,
        "benchmark": GEOCODER_BENCHMARK,
        "vintage": GEOCODER_VINTAGE,
        "layers": "all",
        "format": "json",
    }
    try:
        response = session.get(url, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
        response.raise_for_status()
        payload = response.json()
        geos = payload.get("result", {}).get("geographies", {}) or {}
    except Exception:
        return {"tract_geoid": None, "block_group_geoid": None, "block_geoid": None}

    tract_geoid = None
    block_group_geoid = None
    block_geoid = None
    for key, items in geos.items():
        if not items:
            continue
        first = items[0] if isinstance(items, list) else None
        if not isinstance(first, dict):
            continue
        k = key.lower()
        geoid = first.get("GEOID")
        if geoid is None:
            continue
        if ("block groups" in k or "block group" in k) and not block_group_geoid:
            block_group_geoid = geoid
        elif ("blocks" in k or "block" in k) and ("group" not in k) and not block_geoid:
            block_geoid = geoid
        elif ("tracts" in k or "tract" in k) and not tract_geoid:
            tract_geoid = geoid
    return {
        "tract_geoid": tract_geoid,
        "block_group_geoid": block_group_geoid,
        "block_geoid": block_geoid,
    }

def geocode_addresses(
    df: pd.DataFrame,
    project_root: Path,
    notes: list[str],
    geocode_mode: str = GEOCODE_MODE_AUTO,
    cache_filename: str = GEOCODE_CACHE_FILENAME,
    refresh_cache: bool = False,
) -> pd.DataFrame:
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})
    cache_path = project_root / cache_filename
    persistent_cache = {} if refresh_cache else load_geocode_cache(cache_path, notes)
    run_cache: dict[str, dict[str, object]] = {}

    unique_addresses = df["Address"].fillna("").astype(str).str.strip().unique().tolist()
    allow_live = geocode_mode in {GEOCODE_MODE_AUTO, GEOCODE_MODE_LIVE}
    force_live = geocode_mode == GEOCODE_MODE_LIVE
    live_disabled_reason: Optional[str] = None
    cache_hits = 0
    live_hits = 0
    offline_hits = 0

    print(f"Unique addresses to resolve: {len(unique_addresses):,}")

    for i, address in enumerate(unique_addresses, start=1):
        norm_address = normalize_address_for_cache(address)

        if address == "":
            record = offline_geocode_record(address)
            record["geocode_status"] = "blank_address"
            record["geocode_source"] = "blank"
            record = harmonize_geocode_record(record)
            run_cache[address] = record
            continue

        if not refresh_cache and norm_address in persistent_cache:
            record = dict(persistent_cache[norm_address])
            record.setdefault("geocode_source", "cache")
            record = harmonize_geocode_record(record)
            run_cache[address] = record
            cache_hits += 1
            continue

        raw_zip = extract_zipcode_from_text(address)
        fallback_state_name, fallback_state_abbr = parse_state_from_address(address)

        if not allow_live:
            record = offline_geocode_record(address)
            record = harmonize_geocode_record(record)
            record = harmonize_geocode_record(record)
            run_cache[address] = record
            persistent_cache[norm_address] = record
            offline_hits += 1
            continue

        lat, lon, matched_address, status = get_coordinates(session, address)
        if status == "request_error":
            if force_live:
                raise RuntimeError(
                    "Census geocoder request failed in live-only mode. Check internet access and try again."
                )
            allow_live = False
            live_disabled_reason = "Census geocoder unreachable; switched to offline fallback mode."
            record = offline_geocode_record(address)
            run_cache[address] = record
            persistent_cache[norm_address] = record
            offline_hits += 1
            continue

        tract = None
        block = None
        blockgroup_geoid = None
        matched_zip = extract_zipcode_from_text(matched_address)

        if lat is not None and lon is not None:
            geo = get_census_geographies(session, lat, lon)
            tract = to_tract11(geo.get("tract_geoid") or geo.get("block_group_geoid") or geo.get("block_geoid"))
            block = to_geoid15(geo.get("block_geoid"))
            blockgroup_geoid = to_blockgroup12(geo.get("block_group_geoid") or geo.get("block_geoid"))
            time.sleep(GEOCODER_SLEEP_SECONDS)

        blockgroup_fips = blockgroup_geoid
        state_fips = tract[:2] if tract else (blockgroup_fips[:2] if blockgroup_fips else None)
        county_fips = tract[:5] if tract else (blockgroup_fips[:5] if blockgroup_fips else None)
        state_name = STATE_FIPS_TO_NAME.get(state_fips) if state_fips else fallback_state_name
        state_abbr = STATE_NAME_TO_ABBR.get(state_name) if state_name else fallback_state_abbr

        if status == "no_match" and not tract and not blockgroup_fips:
            record = offline_geocode_record(address)
            record["geocode_status"] = "no_match"
            record["geocode_source"] = "offline_fallback"
        else:
            record = {
                "geocoded_address": matched_address,
                "geoid": block,
                "latitude": lat,
                "longitude": lon,
                "census_tract": tract,
                "blockgroup_fips": blockgroup_fips,
                "county_fips": county_fips,
                "state_fips": state_fips,
                "state_name": state_name,
                "state_abbr": state_abbr,
                "matched_zipcode": matched_zip,
                "zipcode": matched_zip or raw_zip,
                "geocode_status": status if matched_address else (status or "failed"),
                "geocode_source": "census_live",
            }

        # Preserve simple address-derived state/ZIP context whenever live geocoding
        # did not provide it.
        record["state_name"] = record.get("state_name") or fallback_state_name
        record["state_abbr"] = record.get("state_abbr") or fallback_state_abbr
        if not record.get("state_fips") and record.get("state_abbr"):
            for fips, name in STATE_FIPS_TO_NAME.items():
                if name == STATE_ABBR_TO_NAME.get(record["state_abbr"]):
                    record["state_fips"] = fips
                    break
        record["matched_zipcode"] = to_zip5(record.get("matched_zipcode"))
        record["zipcode"] = to_zip5(record.get("zipcode")) or raw_zip

        record = harmonize_geocode_record(record)
        run_cache[address] = record
        persistent_cache[norm_address] = record
        if record.get("geocode_source") == "census_live":
            live_hits += 1
        else:
            offline_hits += 1

        if i % 50 == 0:
            print(f"Processed {i:,} / {len(unique_addresses):,} unique addresses...")

    save_geocode_cache(cache_path, persistent_cache, notes)

    if cache_hits:
        notes.append(f"Geocode cache hits: {cache_hits}")
    if live_hits:
        notes.append(f"Census live geocodes this run: {live_hits}")
    if offline_hits:
        notes.append(f"Offline fallback geocodes this run: {offline_hits}")
    if live_disabled_reason:
        notes.append(live_disabled_reason)

    geocode_df = pd.DataFrame.from_dict(run_cache, orient="index").reset_index().rename(columns={"index": "Address"})
    required_geocode_cols = [
        "geocoded_address", "geoid", "latitude", "longitude", "census_tract", "blockgroup_fips",
        "county_fips", "state_fips", "state_name", "state_abbr", "matched_zipcode", "zipcode",
        "geocode_status", "geocode_source", "geocode_method", "geocode_match_type", "geography_level_reached",
    ]
    for col in required_geocode_cols:
        if col not in geocode_df.columns:
            geocode_df[col] = None

    out = df.merge(geocode_df, on="Address", how="left")
    for col in required_geocode_cols:
        if col not in out.columns:
            out[col] = None
    out["geoid"] = out["geoid"].astype(object)
    out["zipcode"] = out["zipcode"].apply(to_zip5)
    out["matched_zipcode"] = out["matched_zipcode"].apply(to_zip5)
    out["census_tract"] = out["census_tract"].apply(to_tract11)
    out["blockgroup_fips"] = out["blockgroup_fips"].apply(to_blockgroup12)
    return out

# ------------------------------------------------------------------
# 7. Reference loaders

# ------------------------------------------------------------------
def load_vvi(paths: ProjectPaths, tract_keys: set[str], zip_keys: set[str], notes: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    vvi_ct = pd.DataFrame(columns=["merge_key"])
    vvi_zip = pd.DataFrame(columns=["merge_key"])

    if paths.vvi_ct and paths.vvi_ct.exists():
        df = read_table(paths.vvi_ct)
        key_col = find_key_column(df, "tract")
        vvi_ct = filter_dataframe_by_key(df, key_col, to_tract11, tract_keys)
    else:
        notes.append("VVI census tract reference not found.")

    if paths.vvi_zip and paths.vvi_zip.exists():
        df = read_table(paths.vvi_zip)
        key_col = find_key_column(df, "zip")
        vvi_zip = filter_dataframe_by_key(df, key_col, to_zip5, zip_keys)
    else:
        notes.append("VVI ZIP reference not found.")

    return vvi_ct, vvi_zip

def load_ruca(paths: ProjectPaths, tract_keys: set[str], zip_keys: set[str], notes: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    ruca_ct = pd.DataFrame(columns=["merge_key"])
    ruca_zip = pd.DataFrame(columns=["merge_key"])

    if paths.ruca_ct and paths.ruca_ct.exists():
        df = read_filtered_csv(paths.ruca_ct, "TractFIPS20", to_tract11, tract_keys)
        ruca_ct = df
    else:
        notes.append("RUCA census tract reference not found.")

    if paths.ruca_zip and paths.ruca_zip.exists():
        df = read_filtered_csv(paths.ruca_zip, "ZIPCode", to_zip5, zip_keys)
        ruca_zip = df
        if "PrimaryRUCADescription" not in ruca_zip.columns and "PrimaryRUCA" in ruca_zip.columns:
            primary_desc = {
                "1": "Metropolitan core", "2": "Metropolitan high commuting", "3": "Metropolitan low commuting",
                "4": "Micropolitan core", "5": "Micropolitan high commuting", "6": "Micropolitan low commuting",
                "7": "Small town core", "8": "Small town high commuting", "9": "Small town low commuting", "10": "Rural",
            }
            secondary_desc = {
                "1": "Metropolitan core, no addtional code", "2": "Metropolitan high commuting", "3": "Metropolitan low commuting",
                "4": "Micropolitan core", "5": "Micropolitan high commuting", "6": "Micropolitan low commuting",
                "7": "Small town core", "8": "Small town high commuting", "9": "Small town low commuting", "10": "Rural",
            }
            ruca_zip["PrimaryRUCADescription"] = ruca_zip["PrimaryRUCA"].map(primary_desc)
            ruca_zip["SecondaryRUCADescription"] = ruca_zip["SecondaryRUCA"].map(secondary_desc)
    else:
        notes.append("RUCA ZIP reference not found.")

    return ruca_ct, ruca_zip

def load_dci(paths: ProjectPaths, zip_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not zip_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.dci_zip or not paths.dci_zip.exists():
        notes.append("DCI reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    df = read_table(paths.dci_zip, sheet_name="Zips")
    key_col = find_column_by_name(df, ["Zip Code"]) or df.columns[0]
    df = filter_dataframe_by_key(df, key_col, to_zip5, zip_keys)
    return df

def load_adi_blockgroup(paths: ProjectPaths, blockgroup_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not blockgroup_keys:
        return pd.DataFrame(columns=["merge_key"])
    path = resolve_numbers_or_export(paths.adi_blockgroup)
    if path is None or not path.exists():
        notes.append("ADI reference not found.")
        return pd.DataFrame(columns=["merge_key"])

    try:
        # Fast path for CSV: avoid loading the entire ADI file into memory.
        if path.suffix.lower() == ".csv":
            header = pd.read_csv(path, dtype=str, nrows=0, encoding_errors="replace")
            cols = list(header.columns)
            key_col = find_column_by_name(header, ["FIPS", "BLOCKGROUPFIPS", "GEOID12", "GEOID"])  # type: ignore[arg-type]
            if key_col is None:
                # Fall back to common ADI naming conventions if present.
                for candidate in ["FIPS", "GEOID12", "GEOID", "BLOCKGROUPFIPS"]:
                    if candidate in cols:
                        key_col = candidate
                        break
            if key_col is None:
                # Last resort: pick the first column (still chunked).
                key_col = cols[0]

            out = read_filtered_csv(path, key_col, to_blockgroup12, blockgroup_keys)

            nat_col = find_column_by_name(out, ["ADI_NATRANK", "UWADI_ADI_NATRANK", "ADI NATRANK", "National Rank"])
            state_col = find_column_by_name(out, ["ADI_STATERNK", "UWADI_ADI_STATERNK", "STATERNK", "State Rank"])
            rename_map = {}
            if nat_col and nat_col != "UWADI_ADI_NATRANK":
                rename_map[nat_col] = "UWADI_ADI_NATRANK"
            if state_col and state_col != "UWADI_ADI_STATERNK":
                rename_map[state_col] = "UWADI_ADI_STATERNK"
            if rename_map:
                out = out.rename(columns=rename_map)
            return out

        # Default path for XLSX/Numbers/etc.
        df = read_table(path)
        key_col = find_column_by_name(df, ["FIPS", "BLOCKGROUPFIPS", "GEOID12", "GEOID"])
        if key_col is None:
            best_col, best_score = None, -1.0
            for col in df.columns:
                score = series_match_fraction(df[col], to_blockgroup12)
                if score > best_score:
                    best_col, best_score = col, score
            key_col = best_col or df.columns[0]

        nat_col = find_column_by_name(df, ["ADI_NATRANK", "UWADI_ADI_NATRANK", "ADI NATRANK", "National Rank"])
        state_col = find_column_by_name(df, ["ADI_STATERNK", "UWADI_ADI_STATERNK", "STATERNK", "State Rank"])

        out = filter_dataframe_by_key(df, key_col, to_blockgroup12, blockgroup_keys)
        rename_map = {}
        if nat_col and nat_col != "UWADI_ADI_NATRANK":
            rename_map[nat_col] = "UWADI_ADI_NATRANK"
        if state_col and state_col != "UWADI_ADI_STATERNK":
            rename_map[state_col] = "UWADI_ADI_STATERNK"
        if rename_map:
            out = out.rename(columns=rename_map)
        return out
    except Exception as exc:
        notes.append(f"Failed to load ADI reference ({path.name}): {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_ahrf_blockgroup(paths: ProjectPaths, blockgroup_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not blockgroup_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.ahrf_blockgroup or not paths.ahrf_blockgroup.exists():
        notes.append("AHRF blockgroup reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        if paths.ahrf_blockgroup.suffix.lower() == ".csv":
            notes.append(f"Using AHRF blockgroup CSV reference: {paths.ahrf_blockgroup.name}")
            df = read_filtered_csv(paths.ahrf_blockgroup, "BLOCKGROUPFIPS", to_blockgroup12, blockgroup_keys)
        else:
            df = read_filtered_xlsx_data_sheet(paths.ahrf_blockgroup, "BLOCKGROUPFIPS", to_blockgroup12, blockgroup_keys)
        return df if not df.empty else pd.DataFrame(columns=["merge_key"])
    except Exception as exc:
        notes.append(f"Failed to load AHRF blockgroup reference: {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_ahrf_tract_selected(paths: ProjectPaths, tract_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not tract_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.ahrf_tract or not paths.ahrf_tract.exists():
        notes.append("AHRF tract reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        if paths.ahrf_tract.suffix.lower() == ".csv":
            notes.append(f"Using AHRF tract CSV reference: {paths.ahrf_tract.name}")
            df = read_filtered_csv(paths.ahrf_tract, "TRACTFIPS", to_tract11, tract_keys)
        else:
            df = read_filtered_xlsx_data_sheet(paths.ahrf_tract, "TRACTFIPS", to_tract11, tract_keys)
        if df.empty:
            return df
        desired = ["merge_key"] + [col for col in AHRF_MAIN_METRICS if col in df.columns]
        return df[[c for c in desired if c in df.columns]].copy()
    except Exception as exc:
        notes.append(f"Failed to load AHRF tract reference: {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_cre_tract(paths: ProjectPaths, tract_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not tract_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.cre_tract or not paths.cre_tract.exists():
        notes.append("CRE tract reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        df = read_filtered_csv(paths.cre_tract, "GEO_ID", geoid_like_to_tract11, tract_keys)
        if not df.empty:
            return df
        full = read_table(paths.cre_tract)
        if all(col in full.columns for col in ["STATE", "COUNTY", "TRACT"]):
            full["merge_key"] = (
                full["STATE"].astype(str).str.zfill(2)
                + full["COUNTY"].astype(str).str.zfill(3)
                + full["TRACT"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
            )
            return full[full["merge_key"].isin(tract_keys)].copy()
        return pd.DataFrame(columns=["merge_key"])
    except Exception as exc:
        notes.append(f"Failed to load CRE tract reference: {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_eji_tract(paths: ProjectPaths, tract_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not tract_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.eji_tract or not paths.eji_tract.exists():
        notes.append("EJI tract reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        for candidate in ["GEOID", "GEOID_2020", "AFFGEOID"]:
            try:
                df = read_filtered_csv(paths.eji_tract, candidate, geoid_like_to_tract11, tract_keys)
                if not df.empty:
                    return df
            except Exception:
                pass
        full = read_table(paths.eji_tract)
        if all(col in full.columns for col in ["STATEFP", "COUNTYFP", "TRACTCE"]):
            full["merge_key"] = (
                full["STATEFP"].astype(str).str.zfill(2)
                + full["COUNTYFP"].astype(str).str.zfill(3)
                + full["TRACTCE"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
            )
            return full[full["merge_key"].isin(tract_keys)].copy()
        key_col = find_key_column(full, ["GEOID", "GEOID_2020", "AFFGEOID", "TRACT"])
        return filter_dataframe_by_key(full, key_col, geoid_like_to_tract11, tract_keys)
    except Exception as exc:
        notes.append(f"Failed to load EJI tract reference: {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_fara_tract(paths: ProjectPaths, tract_keys: set[str], notes: list[str]) -> pd.DataFrame:
    empty_cols = ["merge_key", "CensusTract", "LILATracts_1And10", "LILATracts_Vehicle", "LowIncomeTracts"]
    if not paths.fara_tract or not paths.fara_tract.exists():
        notes.append("FARA tract reference not found.")
        return pd.DataFrame(columns=empty_cols)
    try:
        if not tract_keys:
            notes.append("FARA file found, but no tract keys were available for matching in this run.")
            return pd.DataFrame(columns=empty_cols)
        for candidate in ["CensusTract", "tract", "Tract", "Census_Tract"]:
            try:
                df = read_filtered_csv(paths.fara_tract, candidate, geoid_like_to_tract11, tract_keys)
                if not df.empty:
                    return df
            except Exception:
                pass
        full = read_table(paths.fara_tract)
        key_col = find_key_column(full, ["CensusTract", "tract", "Tract", "Census_Tract", "GEOID"])
        out = filter_dataframe_by_key(full, key_col, geoid_like_to_tract11, tract_keys)
        if out.empty:
            notes.append("FARA file loaded, but no matching tract rows were found for the current cohort.")
        return out
    except Exception as exc:
        notes.append(f"Failed to load FARA tract reference: {exc}")
        return pd.DataFrame(columns=empty_cols)

def load_nri_tract(paths: ProjectPaths, tract_keys: set[str], notes: list[str]) -> pd.DataFrame:
    if not tract_keys:
        return pd.DataFrame(columns=["merge_key"])
    if not paths.nri_tract or not paths.nri_tract.exists():
        notes.append("NRI tract reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        usecols = ["TRACTFIPS", "RISK_SCORE", "RISK_RATNG", "EAL_SCORE", "SOVI_SCORE", "RESL_SCORE", "STCOFIPS", "TRACT", "STATEFIPS", "COUNTYFIPS"]
        available_cols = read_table_header_only(paths.nri_tract)
        selected = [c for c in usecols if c in available_cols]
        if selected:
            full = pd.read_csv(paths.nri_tract, dtype=str, usecols=selected)
        else:
            full = read_table(paths.nri_tract)
        if "TRACTFIPS" in full.columns:
            full["merge_key"] = full["TRACTFIPS"].apply(to_tract11)
            out = full[full["merge_key"].isin(tract_keys)].copy()
            if not out.empty:
                return out
        if all(col in full.columns for col in ["STCOFIPS", "TRACT"]):
            tract_part = full["TRACT"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
            full["merge_key"] = full["STCOFIPS"].astype(str).str.zfill(5) + tract_part
            out = full[full["merge_key"].isin(tract_keys)].copy()
            if not out.empty:
                return out
        if all(col in full.columns for col in ["STATEFIPS", "COUNTYFIPS", "TRACT"]):
            tract_part = full["TRACT"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
            full["merge_key"] = full["STATEFIPS"].astype(str).str.zfill(2) + full["COUNTYFIPS"].astype(str).str.zfill(3) + tract_part
            out = full[full["merge_key"].isin(tract_keys)].copy()
            if not out.empty:
                return out
        key_col = find_key_column(full, ["TRACTFIPS", "TRACT", "GEOID", "AFFGEOID"])
        return filter_dataframe_by_key(full, key_col, geoid_like_to_tract11, tract_keys)
    except Exception as exc:
        notes.append(f"Failed to load NRI tract reference: {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_walkability_bg(paths: ProjectPaths, blockgroup_keys: set[str], notes: list[str]) -> tuple[pd.DataFrame, str]:
    if not blockgroup_keys:
        return pd.DataFrame(columns=["merge_key"]), "no_keys"

    def _build_merge_key(props: dict) -> Optional[str]:
        merge_key = to_blockgroup12(props.get("GEOID20") or props.get("GEOID10") or props.get("GEOID") or props.get("BLKGRP") or props.get("merge_key"))
        if not merge_key and all(props.get(k) is not None for k in ["STATEFP", "COUNTYFP", "TRACTCE", "BLKGRPCE"]):
            merge_key = to_blockgroup12(f"{str(props.get('STATEFP')).zfill(2)}{str(props.get('COUNTYFP')).zfill(3)}{str(props.get('TRACTCE')).zfill(6)}{str(props.get('BLKGRPCE')).zfill(1)}")
        if not merge_key and all(props.get(k) is not None for k in ["STATE", "COUNTY", "TRACT", "BLKGRP"]):
            merge_key = to_blockgroup12(f"{str(props.get('STATE')).zfill(2)}{str(props.get('COUNTY')).zfill(3)}{str(props.get('TRACT')).zfill(6)}{str(props.get('BLKGRP')).zfill(1)}")
        return merge_key

    def _load_csv_walkability(csv_path: Path) -> tuple[pd.DataFrame, str]:
        try:
            header = pd.read_csv(csv_path, dtype=str, nrows=0, encoding_errors="replace")
            cols = deduplicate_columns(header).columns.tolist()
            key_col = None
            for cand in ["GEOID20", "GEOID10", "GEOID", "merge_key"]:
                if cand in cols:
                    key_col = cand
                    break
            if key_col:
                walk = read_filtered_csv(csv_path, key_col, to_blockgroup12, blockgroup_keys)
            else:
                walk = read_table(csv_path)
                if walk.empty:
                    notes.append(f"Walkability CSV exists but loaded empty: {csv_path.name}")
                    return pd.DataFrame(columns=["merge_key"]), "csv_empty"
                if all(c in walk.columns for c in ["STATEFP", "COUNTYFP", "TRACTCE", "BLKGRPCE"]):
                    walk["merge_key"] = (
                        walk["STATEFP"].astype(str).str.zfill(2)
                        + walk["COUNTYFP"].astype(str).str.zfill(3)
                        + walk["TRACTCE"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
                        + walk["BLKGRPCE"].astype(str).str.zfill(1)
                    ).apply(to_blockgroup12)
                elif all(c in walk.columns for c in ["STATE", "COUNTY", "TRACT", "BLKGRP"]):
                    walk["merge_key"] = (
                        walk["STATE"].astype(str).str.zfill(2)
                        + walk["COUNTY"].astype(str).str.zfill(3)
                        + walk["TRACT"].astype(str).str.replace(".", "", regex=False).str.zfill(6)
                        + walk["BLKGRP"].astype(str).str.zfill(1)
                    ).apply(to_blockgroup12)
                else:
                    notes.append(f"Walkability CSV found but no block-group key columns detected: {csv_path.name}")
                    return pd.DataFrame(columns=["merge_key"]), "csv_no_key"
                walk = walk[walk["merge_key"].isin(blockgroup_keys)].copy()
            if not walk.empty:
                notes.append(f"Loaded Walkability rows for {walk['merge_key'].nunique():,} matched block groups from CSV: {csv_path.name}")
                return deduplicate_columns(walk), "csv_ok"
            notes.append(f"Walkability CSV loaded but none of the cohort block groups matched: {csv_path.name}")
            return pd.DataFrame(columns=["merge_key"]), "no_match"
        except Exception as exc:
            notes.append(f"Walkability CSV could not be processed: {exc}")
            return pd.DataFrame(columns=["merge_key"]), "csv_error"

    csv_path = find_walkability_csv(paths.root)
    if csv_path and csv_path.exists():
        return _load_csv_walkability(csv_path)

    if paths.walkability_gdb and paths.walkability_gdb.exists():
        try:
            import fiona
            layers = list(fiona.listlayers(paths.walkability_gdb))
            layer = "NationalWalkabilityIndex" if "NationalWalkabilityIndex" in layers else (layers[0] if layers else None)
            if layer:
                records = []
                with fiona.open(paths.walkability_gdb, layer=layer) as src:
                    for feat in src:
                        props = dict(feat.get("properties") or {})
                        merge_key = _build_merge_key(props)
                        if merge_key in blockgroup_keys:
                            props["merge_key"] = merge_key
                            records.append(props)
                if records:
                    walk = pd.DataFrame(records)
                    notes.append(f"Loaded Walkability rows for {walk['merge_key'].nunique():,} matched block groups.")
                    return deduplicate_columns(walk), "gdb_ok"
            else:
                notes.append("Walkability geodatabase contains no readable layers.")
        except Exception as exc:
            notes.append(f"Walkability fiona read path unavailable, trying pyogrio fallback: {exc}")

        try:
            import pyogrio
            layers = pyogrio.list_layers(paths.walkability_gdb)
            layer_names = [row[0] for row in layers] if len(layers) else []
            layer = "NationalWalkabilityIndex" if "NationalWalkabilityIndex" in layer_names else (layer_names[0] if layer_names else None)
            if layer:
                cols = ["GEOID10", "GEOID20", "STATEFP", "COUNTYFP", "TRACTCE", "BLKGRPCE", "NatWalkInd", "D2A_Ranked", "D2B_Ranked", "D3B_Ranked", "D4A_Ranked"]
                walk = pyogrio.read_dataframe(paths.walkability_gdb, layer=layer, columns=cols, read_geometry=False)
                walk["merge_key"] = walk.apply(lambda r: _build_merge_key(r.to_dict()), axis=1)
                walk = walk[walk["merge_key"].isin(blockgroup_keys)].copy()
                if not walk.empty:
                    notes.append(f"Loaded Walkability rows for {walk['merge_key'].nunique():,} matched block groups.")
                    return deduplicate_columns(walk), "gdb_ok"
                notes.append("Walkability geodatabase loaded but none of the cohort block groups matched.")
            else:
                notes.append("Walkability geodatabase contains no readable layers.")
        except Exception as exc:
            notes.append(f"Walkability geodatabase could not be processed: {exc}")

    if paths.walkability_gdb and paths.walkability_gdb.exists():
        notes.append("Walkability geodatabase was found, but the preferred CSV source was not available and the geodatabase could not be read.")
        return pd.DataFrame(columns=["merge_key"]), "reader_missing"

    notes.append("Walkability source not found. Provide walkability_full_attributes.csv in the same project folder as the walkability geodatabase.")
    return pd.DataFrame(columns=["merge_key"]), "source_missing"

def load_svi_state_files(state_dir: Optional[Path], states_needed: set[str], keys_needed: set[str], level: str, notes: list[str]) -> pd.DataFrame:
    if not states_needed or not keys_needed:
        return pd.DataFrame(columns=["merge_key"])
    if state_dir is None or not state_dir.exists():
        notes.append(f"SVI state {level} directory not found.")
        return pd.DataFrame(columns=["merge_key"])

    file_map: dict[str, Path] = {}
    for path in state_dir.glob("*.csv"):
        if path.name.startswith("."):
            continue
        file_map[normalize_state_token(path.stem)] = path

    frames: list[pd.DataFrame] = []
    for state_name in sorted(states_needed):
        token = normalize_state_token(state_name)
        path = file_map.get(token)
        if path is None:
            notes.append(f"SVI state {level} file not found for {state_name}.")
            continue
        if level == "tract":
            df = read_filtered_csv(path, "FIPS", to_tract11, keys_needed)
        else:
            df = read_filtered_csv(path, "FIPS", to_zip5, keys_needed)
        frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["merge_key"])
    return pd.concat(frames, ignore_index=True)

def load_svi_national(path: Optional[Path], keys_needed: set[str], level: str, notes: list[str]) -> pd.DataFrame:
    if not keys_needed:
        return pd.DataFrame(columns=["merge_key"])
    path = resolve_numbers_or_export(path)
    if path is None or not path.exists():
        notes.append(f"SVI national {level} reference not found.")
        return pd.DataFrame(columns=["merge_key"])

    try:
        if path.suffix.lower() == ".csv":
            if level == "tract":
                return read_filtered_csv(path, "FIPS", to_tract11, keys_needed)
            return read_filtered_csv(path, "FIPS", to_zip5, keys_needed)

        df = read_table(path)
        key_col = find_key_column(df, level)
        parser = to_tract11 if level == "tract" else to_zip5
        return filter_dataframe_by_key(df, key_col, parser, keys_needed)
    except Exception as exc:
        notes.append(f"Failed to load SVI national {level} reference ({path.name}): {exc}")
        return pd.DataFrame(columns=["merge_key"])

def load_svi(paths: ProjectPaths, base_df: pd.DataFrame, notes: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    tract_keys = {x for x in base_df["census_tract"].dropna().astype(str) if x}
    zip_keys = {x for x in base_df["zipcode"].dropna().astype(str) if x}
    states_needed = {x for x in base_df["state_name"].dropna().astype(str) if x}

    nat_ct = load_svi_national(paths.svi_us_ct, tract_keys, "tract", notes)
    nat_zip = load_svi_national(paths.svi_us_zip, zip_keys, "zip", notes)
    state_ct = load_svi_state_files(paths.svi_state_ct_dir, states_needed, tract_keys, "tract", notes)
    state_zip = load_svi_state_files(paths.svi_state_zip_dir, states_needed, zip_keys, "zip", notes)
    if nat_ct.empty and not state_ct.empty:
        notes.append("National tract SVI reference unavailable; using tract-level state SVI rows as the tract-capable national fallback for this run.")
        nat_ct = state_ct.copy()
    return nat_ct, nat_zip, state_ct, state_zip

def load_coi_table(path: Optional[Path], tract_keys: set[str], notes: list[str], label: str) -> pd.DataFrame:
    if not tract_keys:
        return pd.DataFrame(columns=["merge_key"])
    if path is None or not path.exists():
        notes.append(f"{label} reference not found.")
        return pd.DataFrame(columns=["merge_key"])
    try:
        return read_filtered_csv(path, "geoid20", to_tract11, tract_keys)
    except Exception as exc:
        notes.append(f"Failed to load {label}: {exc}")
        return pd.DataFrame(columns=["merge_key"])

# ------------------------------------------------------------------
# 8. Categorization rules
# ------------------------------------------------------------------
def categorize_adi_rank(value: object) -> Optional[str]:
    number = safe_float(value)
    if number is None:
        return None
    fraction = number / 100.0 if number > 1 else number
    if fraction < 0:
        return None
    if fraction <= 0.25:
        return "Low"
    if fraction <= 0.50:
        return "Low-Medium"
    if fraction <= 0.75:
        return "Medium-High"
    return "High"

def categorize_vvi(value: object) -> Optional[str]:
    number = safe_float(value)
    if number is None:
        return None
    if number < -1:
        return "Low"
    if number <= 1:
        return "Average"
    if number <= 2:
        return "Medium"
    return "High"

def categorize_dci_quintile(value: object) -> Optional[str]:
    number = safe_float(value)
    if number is None:
        return None
    quintile = int(number)
    # User-confirmed convention: Quintile 5=Distressed, 4=At Risk, 3=Mid-Tier, 2=Comfortable, 1=Prosperous.
    mapping = {
        1: "Prosperous",
        2: "Comfortable",
        3: "Mid-Tier",
        4: "At Risk",
        5: "Distressed",
    }
    return mapping.get(quintile)

# ------------------------------------------------------------------
# 9. Merge and selection logic
# ------------------------------------------------------------------
def merge_with_prefix(base_df: pd.DataFrame, ref_df: pd.DataFrame, left_on: str, prefix: str) -> pd.DataFrame:
    if ref_df.empty:
        return base_df.copy()
    ref = prefix_reference(ref_df, prefix)
    out = base_df.drop(columns=[c for c in base_df.columns if c.startswith("merge_key")], errors="ignore").merge(
        ref, left_on=left_on, right_on="merge_key", how="left"
    )
    return out.drop(columns=["merge_key"], errors="ignore")

def build_outputs(base_df: pd.DataFrame, paths: ProjectPaths, notes: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    tract_keys = {x for x in base_df["census_tract"].dropna().astype(str) if x}
    zip_keys = {x for x in base_df["zipcode"].dropna().astype(str) if x}
    blockgroup_keys = {x for x in base_df["blockgroup_fips"].dropna().astype(str) if x}

    nat_ct, nat_zip, state_ct, state_zip = load_svi(paths, base_df, notes)

    vvi_ct, vvi_zip = load_vvi(paths, tract_keys, zip_keys, notes)

    ruca_ct, ruca_zip = load_ruca(paths, tract_keys, zip_keys, notes)

    dci_zip = load_dci(paths, zip_keys, notes)

    adi_bg = load_adi_blockgroup(paths, blockgroup_keys, notes)

    ahrf_bg = load_ahrf_blockgroup(paths, blockgroup_keys, notes)

    ahrf_ct = load_ahrf_tract_selected(paths, tract_keys, notes)

    coi_domains = load_coi_table(paths.coi_domains, tract_keys, notes, "COI domains")

    coi_subdomains = load_coi_table(paths.coi_subdomains, tract_keys, notes, "COI subdomains")

    cre_tract = load_cre_tract(paths, tract_keys, notes)

    eji_tract = load_eji_tract(paths, tract_keys, notes)

    fara_tract = load_fara_tract(paths, tract_keys, notes)

    nri_tract = load_nri_tract(paths, tract_keys, notes)

    walk_bg, walk_status = load_walkability_bg(paths, blockgroup_keys, notes)

    merged = base_df.copy()
    merged = merge_with_prefix(merged, nat_ct, "census_tract", "svi_nat_ct__")
    merged = merge_with_prefix(merged, nat_zip, "zipcode", "svi_nat_zip__")
    merged = merge_with_prefix(merged, state_ct, "census_tract", "svi_state_ct__")
    merged = merge_with_prefix(merged, state_zip, "zipcode", "svi_state_zip__")
    merged = merge_with_prefix(merged, vvi_ct, "census_tract", "vvi_ct__")
    merged = merge_with_prefix(merged, vvi_zip, "zipcode", "vvi_zip__")
    merged = merge_with_prefix(merged, ruca_ct, "census_tract", "ruca_ct__")
    merged = merge_with_prefix(merged, ruca_zip, "zipcode", "ruca_zip__")
    merged = merge_with_prefix(merged, dci_zip, "zipcode", "dci_zip__")
    merged = merge_with_prefix(merged, adi_bg, "blockgroup_fips", "adi_bg__")
    merged = merge_with_prefix(merged, ahrf_bg, "blockgroup_fips", "ahrf_bg__")
    merged = merge_with_prefix(merged, ahrf_ct, "census_tract", "ahrf_ct__")
    merged = merge_with_prefix(merged, cre_tract, "census_tract", "cre_ct__")
    merged = merge_with_prefix(merged, eji_tract, "census_tract", "eji_ct__")
    merged = merge_with_prefix(merged, fara_tract, "census_tract", "fara_ct__")
    merged = merge_with_prefix(merged, nri_tract, "census_tract", "nri_ct__")
    merged = merge_with_prefix(merged, walk_bg, "blockgroup_fips", "walk_bg__")

    can_svi_nat_ct = merged.get("svi_nat_ct__RPL_THEMES", pd.Series(index=merged.index)).notna()
    can_svi_nat_zip = merged.get("svi_nat_zip__RPL_THEMES", pd.Series(index=merged.index)).notna()
    use_svi_nat_ct = can_svi_nat_ct
    use_svi_nat_zip = (~use_svi_nat_ct) & can_svi_nat_zip
    can_svi_state_ct = merged.get("svi_state_ct__RPL_THEMES", pd.Series(index=merged.index)).notna()
    can_svi_state_zip = merged.get("svi_state_zip__RPL_THEMES", pd.Series(index=merged.index)).notna()
    use_svi_state_ct = can_svi_state_ct
    use_svi_state_zip = (~use_svi_state_ct) & can_svi_state_zip

    can_vvi_ct = merged.get("vvi_ct__VVI", pd.Series(index=merged.index)).notna()
    can_vvi_zip = merged.get("vvi_zip__VVI", pd.Series(index=merged.index)).notna()
    use_vvi_ct = can_vvi_ct
    use_vvi_zip = (~use_vvi_ct) & can_vvi_zip

    can_ruca_ct = merged.get("ruca_ct__PrimaryRUCA", pd.Series(index=merged.index)).notna()
    can_ruca_zip = merged.get("ruca_zip__PrimaryRUCA", pd.Series(index=merged.index)).notna()
    use_ruca_ct = can_ruca_ct
    use_ruca_zip = (~use_ruca_ct) & can_ruca_zip

    base_cols = [
        "UniqueID", "Address", "geocoded_address", "state_name", "state_abbr", "state_fips",
        "county_fips", "census_tract", "blockgroup_fips", "zipcode", "latitude", "longitude", "geocode_status",
        "geocode_method", "geocode_match_type", "geography_level_reached",
    ]
    for col in base_cols:
        if col not in merged.columns:
            merged[col] = None
    main = merged[base_cols].copy()
    main = main.rename(columns={
        "geocoded_address": "GeocodedAddress",
        "state_name": "StateTerritory",
        "state_abbr": "StateAbbreviation",
        "state_fips": "StateFIPS",
        "county_fips": "CountyFIPS",
        "census_tract": "CensusTractFIPS",
        "blockgroup_fips": "BlockGroupFIPS",
        "zipcode": "ZCTA",
        "latitude": "Latitude",
        "longitude": "Longitude",
        "geocode_status": "Geocode Status",
        "geocode_method": "Geocode Method",
        "geocode_match_type": "Geocode Match Type",
        "geography_level_reached": "Geography Level Reached",
    })

    adi_nat = merged.get("adi_bg__UWADI_ADI_NATRANK", pd.Series(index=merged.index))
    adi_state = merged.get("adi_bg__UWADI_ADI_STATERNK", pd.Series(index=merged.index))
    ahrf_adi_nat = merged.get("ahrf_bg__UWADI_ADI_NATRANK", pd.Series(index=merged.index))
    ahrf_adi_state = merged.get("ahrf_bg__UWADI_ADI_STATERNK", pd.Series(index=merged.index))
    use_direct_adi = adi_nat.notna()
    use_ahrf_adi = (~use_direct_adi) & ahrf_adi_nat.notna()

    main["ADI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[use_direct_adi | use_ahrf_adi, "ADI_GeographyUsed"] = "BlockGroup"
    main["ADI_Source"] = pd.Series(index=main.index, dtype=object)
    main.loc[use_direct_adi, "ADI_Source"] = "ADI_File"
    main.loc[use_ahrf_adi, "ADI_Source"] = "AHRF_Fallback"
    main["ADI_NATRANK"] = combine_by_mask(adi_nat, ahrf_adi_nat, use_direct_adi, use_ahrf_adi)
    main["ADI_STATERNK"] = combine_by_mask(adi_state, ahrf_adi_state, use_direct_adi, use_ahrf_adi)
    main["ADI_NationalRiskCategory"] = main["ADI_NATRANK"].apply(categorize_adi_rank)
    main["ADI_StateRiskCategory"] = main["ADI_STATERNK"].apply(categorize_adi_rank)

    main["SVI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    nat_geo = pd.Series(pd.NA, index=main.index, dtype=object)
    nat_geo.loc[use_svi_nat_ct] = "CensusTract"
    nat_geo.loc[use_svi_nat_zip] = "ZIPCode"
    state_geo = pd.Series(pd.NA, index=main.index, dtype=object)
    state_geo.loc[use_svi_state_ct] = "CensusTract"
    state_geo.loc[use_svi_state_zip] = "ZIPCode"
    same_geo = nat_geo.notna() & (nat_geo == state_geo)
    mixed_geo = nat_geo.notna() & state_geo.notna() & (nat_geo != state_geo)
    main.loc[same_geo, "SVI_GeographyUsed"] = nat_geo.loc[same_geo]
    main.loc[mixed_geo, "SVI_GeographyUsed"] = "Mixed: National " + nat_geo.loc[mixed_geo].astype(str) + ", State " + state_geo.loc[mixed_geo].astype(str)
    main.loc[main["SVI_GeographyUsed"].isna() & nat_geo.notna(), "SVI_GeographyUsed"] = "National " + nat_geo[nat_geo.notna()].astype(str)
    main.loc[main["SVI_GeographyUsed"].isna() & state_geo.notna(), "SVI_GeographyUsed"] = "State " + state_geo[state_geo.notna()].astype(str)
    for metric in SVI_MAIN_METRICS:
        main[f"SVI_National_{metric}"] = combine_by_mask(
            merged.get(f"svi_nat_ct__{metric}", pd.Series(index=merged.index)),
            merged.get(f"svi_nat_zip__{metric}", pd.Series(index=merged.index)),
            use_svi_nat_ct,
            use_svi_nat_zip,
        )
        main[f"SVI_State_{metric}"] = combine_by_mask(
            merged.get(f"svi_state_ct__{metric}", pd.Series(index=merged.index)),
            merged.get(f"svi_state_zip__{metric}", pd.Series(index=merged.index)),
            use_svi_state_ct,
            use_svi_state_zip,
        )

    main["VVI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[use_vvi_ct, "VVI_GeographyUsed"] = "CensusTract"
    main.loc[use_vvi_zip, "VVI_GeographyUsed"] = "ZIPCode"
    for metric in VVI_MAIN_METRICS:
        main[f"VVI_{metric}"] = combine_by_mask(
            merged.get(f"vvi_ct__{metric}", pd.Series(index=merged.index)),
            merged.get(f"vvi_zip__{metric}", pd.Series(index=merged.index)),
            use_vvi_ct,
            use_vvi_zip,
        )
    main["VVI_RiskCategory"] = main["VVI_VVI"].apply(categorize_vvi)

    main["RUCA_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[use_ruca_ct, "RUCA_GeographyUsed"] = "CensusTract"
    main.loc[use_ruca_zip, "RUCA_GeographyUsed"] = "ZIPCode"
    for metric in RUCA_MAIN_METRICS:
        main[f"RUCA_{metric}"] = combine_by_mask(
            merged.get(f"ruca_ct__{metric}", pd.Series(index=merged.index)),
            merged.get(f"ruca_zip__{metric}", pd.Series(index=merged.index)),
            use_ruca_ct,
            use_ruca_zip,
        )

    main["DCI_2019_2023_DistressScore"] = merged.get("dci_zip__2019-2023 Distress Score")
    main["DCI_Quintile"] = merged.get("dci_zip__Quintile (5=Distressed)")
    main["DCI_Category"] = main["DCI_Quintile"].apply(categorize_dci_quintile)

    for metric in EJI_MAIN_METRICS:
        main[metric] = merged.get(f"eji_ct__{metric}")
    for metric in CRE_MAIN_METRICS:
        main[metric] = merged.get(f"cre_ct__{metric}")
    for metric in FARA_MAIN_METRICS:
        main[metric] = merged.get(f"fara_ct__{metric}")
    for metric in NRI_MAIN_METRICS:
        main[metric] = merged.get(f"nri_ct__{metric}", pd.Series(index=merged.index, dtype=object))
    for metric in AHRF_MAIN_METRICS:
        main[metric] = merged.get(f"ahrf_ct__{metric}", pd.Series(index=merged.index, dtype=object))
    for metric in WALK_MAIN_METRICS:
        main[metric] = merged.get(f"walk_bg__{metric}", pd.Series(index=merged.index, dtype=object))

    # Source years / provenance.
    main["ADI_SourceYear"] = extract_year_from_path(paths.adi_blockgroup) or extract_year_from_path(paths.ahrf_blockgroup)
    main["SVI_SourceYear"] = extract_year_from_path(paths.svi_us_ct) or extract_year_from_path(paths.svi_us_zip)
    main["VVI_SourceYear"] = extract_year_from_path(paths.vvi_ct) or extract_year_from_path(paths.vvi_zip)
    main["RUCA_SourceYear"] = extract_year_from_path(paths.ruca_ct) or extract_year_from_path(paths.ruca_zip)
    main["DCI_SourceYear"] = extract_year_from_path(paths.dci_zip)
    main["DCI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["DCI_2019_2023_DistressScore"].notna(), "DCI_GeographyUsed"] = "ZIPCode"
    main["EJI_SourceYear"] = extract_year_from_path(paths.eji_tract)
    main["EJI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["RPL_EJI"].notna(), "EJI_GeographyUsed"] = "CensusTract"
    main["CRE_SourceYear"] = extract_year_from_path(paths.cre_tract)
    main["CRE_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["PRED3_PE"].notna() | main["PRED12_PE"].notna() | main["PRED0_PE"].notna(), "CRE_GeographyUsed"] = "CensusTract"
    main["FARA_SourceYear"] = extract_year_from_path(paths.fara_tract)
    main["FARA_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["LILATracts_1And10"].notna() | main["LILATracts_Vehicle"].notna() | main["LowIncomeTracts"].notna(), "FARA_GeographyUsed"] = "CensusTract"
    main["NRI_SourceYear"] = extract_year_from_path(paths.nri_tract)
    main["NRI_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["RISK_SCORE"].notna() | main["RISK_RATNG"].notna() | main["EAL_SCORE"].notna() | main["SOVI_SCORE"].notna() | main["RESL_SCORE"].notna(), "NRI_GeographyUsed"] = "CensusTract"
    main["AHRF_SourceYear"] = extract_year_from_path(paths.ahrf_tract)
    main["AHRF_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[
        main["HRSA_MUA_CENSUS_TRACT"].notna() | main["POS_DIST_ED_TRACT"].notna() | main["POS_DIST_CLINIC_TRACT"].notna() |
        main["PC_PCT_MEDICARE_APPRVD_FULL_AMT"].notna() | main["POS_DIST_MEDSURG_ICU_TRACT"].notna() | main["POS_DIST_TRAUMA_TRACT"].notna(),
        "AHRF_GeographyUsed"
    ] = "CensusTract"
    main["Walkability_SourceYear"] = extract_year_from_path(find_walkability_csv(paths.root)) or extract_year_from_path(paths.walkability_gdb)
    main["Walkability_GeographyUsed"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["NatWalkInd"].notna() | main["D2A_Ranked"].notna() | main["D2B_Ranked"].notna() | main["D3B_Ranked"].notna() | main["D4A_Ranked"].notna(), "Walkability_GeographyUsed"] = "BlockGroup"

    # QC / missingness.
    main["GeographyCompleteFlag"] = (
        main["StateTerritory"].notna() & main["CensusTractFIPS"].notna() & main["BlockGroupFIPS"].notna() & main["ZCTA"].notna()
    ).map({True: "Yes", False: "No"})
    tract_series = main["CensusTractFIPS"].fillna("").astype(str)
    bg_series = main["BlockGroupFIPS"].fillna("").astype(str)
    consistency = pd.Series(pd.NA, index=main.index, dtype=object)
    both_geo = tract_series.ne("") & bg_series.ne("")
    consistency.loc[both_geo] = (bg_series.loc[both_geo].str[:11] == tract_series.loc[both_geo]).map({True: "Pass", False: "Fail"})
    main["TractBlockGroupConsistencyFlag"] = consistency
    main["IndexCoverageCount"] = pd.concat([
        main["ADI_NATRANK"].notna().astype(int),
        main["SVI_National_RPL_THEMES"].notna().astype(int),
        main["VVI_VVI"].notna().astype(int),
        main["RUCA_PrimaryRUCA"].notna().astype(int),
        main["DCI_2019_2023_DistressScore"].notna().astype(int),
        main["RPL_EJI"].notna().astype(int),
        main["PRED3_PE"].notna().astype(int),
        main["LILATracts_1And10"].notna().astype(int),
        main["RISK_SCORE"].notna().astype(int),
        main["HRSA_MUA_CENSUS_TRACT"].notna().astype(int),
        main["NatWalkInd"].notna().astype(int),
    ], axis=1).sum(axis=1)

    main["ADI_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["ADI_NATRANK"].isna() & main["BlockGroupFIPS"].isna(), "ADI_MissingReason"] = "Missing block group"
    main.loc[main["ADI_NATRANK"].isna() & main["BlockGroupFIPS"].notna() & main["ADI_MissingReason"].isna(), "ADI_MissingReason"] = "No ADI match found for block group"

    main["SVI_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["SVI_National_RPL_THEMES"].isna() & main["CensusTractFIPS"].isna() & main["ZCTA"].isna(), "SVI_MissingReason"] = "Missing tract and ZIP"
    main.loc[main["SVI_National_RPL_THEMES"].isna() & main["SVI_MissingReason"].isna(), "SVI_MissingReason"] = "No complete national+state SVI row found"

    main["VVI_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["VVI_VVI"].isna() & main["CensusTractFIPS"].isna() & main["ZCTA"].isna(), "VVI_MissingReason"] = "Missing tract and ZIP"
    main.loc[main["VVI_VVI"].isna() & main["VVI_MissingReason"].isna(), "VVI_MissingReason"] = "No VVI match found"

    main["RUCA_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["RUCA_PrimaryRUCA"].isna() & main["CensusTractFIPS"].isna() & main["ZCTA"].isna(), "RUCA_MissingReason"] = "Missing tract and ZIP"
    main.loc[main["RUCA_PrimaryRUCA"].isna() & main["RUCA_MissingReason"].isna(), "RUCA_MissingReason"] = "No RUCA match found"

    main["DCI_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[main["DCI_2019_2023_DistressScore"].isna() & main["ZCTA"].isna(), "DCI_MissingReason"] = "Missing ZIP"
    main.loc[main["DCI_2019_2023_DistressScore"].isna() & main["DCI_MissingReason"].isna(), "DCI_MissingReason"] = "No DCI match found"

    ahrf_any = main[["HRSA_MUA_CENSUS_TRACT", "POS_DIST_ED_TRACT", "POS_DIST_CLINIC_TRACT", "PC_PCT_MEDICARE_APPRVD_FULL_AMT", "POS_DIST_MEDSURG_ICU_TRACT", "POS_DIST_TRAUMA_TRACT"]].notna().any(axis=1)
    main["AHRF_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[~ahrf_any & main["CensusTractFIPS"].isna(), "AHRF_MissingReason"] = "Missing tract"
    main.loc[~ahrf_any & main["AHRF_MissingReason"].isna(), "AHRF_MissingReason"] = "No AHRF tract match found"

    walk_any = main[["NatWalkInd", "D2A_Ranked", "D2B_Ranked", "D3B_Ranked", "D4A_Ranked"]].notna().any(axis=1)
    main["Walkability_MissingReason"] = pd.Series(index=main.index, dtype=object)
    main.loc[~walk_any & main["BlockGroupFIPS"].isna(), "Walkability_MissingReason"] = "Missing block group"
    if walk_status == "reader_missing":
        main.loc[~walk_any & main["Walkability_MissingReason"].isna(), "Walkability_MissingReason"] = "Walkability CSV source could not be used"
    elif walk_status == "source_missing":
        main.loc[~walk_any & main["Walkability_MissingReason"].isna(), "Walkability_MissingReason"] = "Walkability source file not found"
    elif walk_status in {"csv_no_key", "csv_error", "csv_empty"}:
        main.loc[~walk_any & main["Walkability_MissingReason"].isna(), "Walkability_MissingReason"] = "Walkability CSV source could not be used"
    else:
        main.loc[~walk_any & main["Walkability_MissingReason"].isna(), "Walkability_MissingReason"] = "No walkability match found for block group"

    # Supplementary SVI: full chosen national+state rows, but only one geography per patient.
    nat_ct_cols = sorted(strip_prefix_columns(merged.columns, "svi_nat_ct__"))
    nat_zip_cols = sorted(strip_prefix_columns(merged.columns, "svi_nat_zip__"))
    state_ct_cols = sorted(strip_prefix_columns(merged.columns, "svi_state_ct__"))
    state_zip_cols = sorted(strip_prefix_columns(merged.columns, "svi_state_zip__"))
    nat_all = sorted(set(nat_ct_cols) | set(nat_zip_cols))
    state_all = sorted(set(state_ct_cols) | set(state_zip_cols))
    svi_extra = {"SVI_GeographyUsed": main["SVI_GeographyUsed"]}
    for col in nat_all:
        svi_extra[f"National_{col}"] = combine_by_mask(
            merged.get(f"svi_nat_ct__{col}", pd.Series(index=merged.index)),
            merged.get(f"svi_nat_zip__{col}", pd.Series(index=merged.index)),
            use_svi_nat_ct,
            use_svi_nat_zip,
        )
    for col in state_all:
        svi_extra[f"State_{col}"] = combine_by_mask(
            merged.get(f"svi_state_ct__{col}", pd.Series(index=merged.index)),
            merged.get(f"svi_state_zip__{col}", pd.Series(index=merged.index)),
            use_svi_state_ct,
            use_svi_state_zip,
        )
    svi_supp = pd.concat([main.copy(), pd.DataFrame(svi_extra, index=main.index)], axis=1)
    svi_supp = svi_supp.copy()

    # Supplementary RUCA: selected row only.
    ruca_ct_cols = sorted(strip_prefix_columns(merged.columns, "ruca_ct__"))
    ruca_zip_cols = sorted(strip_prefix_columns(merged.columns, "ruca_zip__"))
    ruca_all = sorted(set(ruca_ct_cols) | set(ruca_zip_cols))
    ruca_extra = {"RUCA_GeographyUsed": main["RUCA_GeographyUsed"]}
    for col in ruca_all:
        if col == "merge_key":
            continue
        ruca_extra[col] = combine_by_mask(
            merged.get(f"ruca_ct__{col}", pd.Series(index=merged.index)),
            merged.get(f"ruca_zip__{col}", pd.Series(index=merged.index)),
            use_ruca_ct,
            use_ruca_zip,
        )
    ruca_supp = pd.concat([main.copy(), pd.DataFrame(ruca_extra, index=main.index)], axis=1)
    ruca_supp = ruca_supp.copy()

    # AHRF supplementary.
    ahrf_cols = sorted(strip_prefix_columns(merged.columns, "ahrf_bg__"))
    ahrf_extra = {}
    for col in ahrf_cols:
        if col == "merge_key":
            continue
        ahrf_extra[col] = merged.get(f"ahrf_bg__{col}")
    ahrf_supp = pd.concat([main.copy(), pd.DataFrame(ahrf_extra, index=main.index)], axis=1)
    ahrf_supp = ahrf_supp.copy()

    # COI outputs.
    for col in ["geoid"]:
        if col not in base_df.columns:
            base_df[col] = None
    coi_base = base_df[[
        "UniqueID", "Address", "geocoded_address", "state_name", "census_tract", "zipcode", "latitude", "longitude"
    ]].copy().rename(columns={
        "geocoded_address": "GeocodedAddress",
        "state_name": "StateTerritory",
        "census_tract": "CensusTractFIPS",
        "zipcode": "ZCTA",
        "latitude": "Latitude",
        "longitude": "Longitude",
    })

    coi_domains_out = coi_base.merge(coi_domains, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not coi_domains.empty else coi_base.copy()
    coi_sub_out = coi_base.merge(coi_subdomains, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not coi_subdomains.empty else coi_base.copy()

    desired_domain_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "ZCTA", "Latitude", "Longitude"]
    desired_sub_cols = desired_domain_cols.copy()
    for col in COI_DOMAIN_OUTPUT_COLUMNS:
        if col in coi_domains_out.columns:
            desired_domain_cols.append(col)
    for col in COI_SUBDOMAIN_OUTPUT_COLUMNS:
        if col in coi_sub_out.columns:
            desired_sub_cols.append(col)
    if "year" in coi_sub_out.columns and "year" not in desired_sub_cols:
        desired_sub_cols.insert(len(desired_domain_cols), "year")

    coi_domains_out = coi_domains_out[[c for c in desired_domain_cols if c in coi_domains_out.columns]].copy()
    coi_sub_out = coi_sub_out[[c for c in desired_sub_cols if c in coi_sub_out.columns]].copy()
    if "year" in coi_domains_out.columns:
        coi_domains_out = coi_domains_out.sort_values(["UniqueID", "year"], kind="stable")
    if "year" in coi_sub_out.columns:
        coi_sub_out = coi_sub_out.sort_values(["UniqueID", "year"], kind="stable")

    # CRE output.
    cre_base = base_df[[
        "UniqueID", "Address", "geocoded_address", "state_name", "census_tract", "zipcode", "latitude", "longitude"
    ]].copy().rename(columns={
        "geocoded_address": "GeocodedAddress",
        "state_name": "StateTerritory",
        "census_tract": "CensusTractFIPS",
        "zipcode": "ZCTA",
        "latitude": "Latitude",
        "longitude": "Longitude",
    })
    cre_out = cre_base.merge(cre_tract, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not cre_tract.empty else cre_base.copy()
    cre_desired_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "ZCTA", "Latitude", "Longitude"]
    for col in cre_out.columns:
        if col not in cre_desired_cols and col != "merge_key":
            cre_desired_cols.append(col)
    cre_out = cre_out[[c for c in cre_desired_cols if c in cre_out.columns]].copy()

    # EJI output.
    eji_base = cre_base.copy()
    eji_out = eji_base.merge(eji_tract, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not eji_tract.empty else eji_base.copy()
    eji_desired_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "ZCTA", "Latitude", "Longitude"]
    for col in eji_out.columns:
        if col not in eji_desired_cols and col != "merge_key":
            eji_desired_cols.append(col)
    eji_out = eji_out[[c for c in eji_desired_cols if c in eji_out.columns]].copy()

    # FARA output.
    fara_base = cre_base.copy()
    fara_out = fara_base.merge(fara_tract, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not fara_tract.empty else fara_base.copy()
    fara_desired_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "ZCTA", "Latitude", "Longitude"]
    for col in fara_out.columns:
        if col not in fara_desired_cols and col != "merge_key":
            fara_desired_cols.append(col)
    fara_out = fara_out[[c for c in fara_desired_cols if c in fara_out.columns]].copy()

    # NRI output.
    nri_base = cre_base.copy()
    nri_out = nri_base.merge(nri_tract, left_on="CensusTractFIPS", right_on="merge_key", how="left") if not nri_tract.empty else nri_base.copy()
    nri_desired_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "ZCTA", "Latitude", "Longitude"]
    for col in nri_out.columns:
        if col not in nri_desired_cols and col != "merge_key":
            nri_desired_cols.append(col)
    nri_out = nri_out[[c for c in nri_desired_cols if c in nri_out.columns]].copy()

    # Walkability output.
    walk_base = base_df[[
        "UniqueID", "Address", "geocoded_address", "state_name", "census_tract", "blockgroup_fips", "zipcode", "latitude", "longitude"
    ]].copy().rename(columns={
        "geocoded_address": "GeocodedAddress",
        "state_name": "StateTerritory",
        "census_tract": "CensusTractFIPS",
        "blockgroup_fips": "BlockGroupFIPS",
        "zipcode": "ZCTA",
        "latitude": "Latitude",
        "longitude": "Longitude",
    })
    walk_out = walk_base.merge(walk_bg, left_on="BlockGroupFIPS", right_on="merge_key", how="left") if not walk_bg.empty else walk_base.copy()
    walk_desired_cols = ["UniqueID", "Address", "GeocodedAddress", "StateTerritory", "CensusTractFIPS", "BlockGroupFIPS", "ZCTA", "Latitude", "Longitude"]
    for col in walk_out.columns:
        if col not in walk_desired_cols and col != "merge_key":
            walk_desired_cols.append(col)
    walk_out = walk_out[[c for c in walk_desired_cols if c in walk_out.columns]].copy()

    main = main.rename(columns={
        "UniqueID": "Unique ID",
        "Address": "Address (as inputted)",
        "GeocodedAddress": "Geocoded Address",
        "StateTerritory": "State / Territory",
        "CensusTractFIPS": "Census Tract (FIPS)",
        "ZCTA": "ZCTA (Zip Code)",
        "BlockGroupFIPS": "Block Group (FIPS)",
        "ADI_NationalRiskCategory": "National Risk Category",
        "ADI_StateRiskCategory": "State Risk Category",
        "VVI_VVI": "VVI",
        "VVI_Economic": "Economic",
        "VVI_Education": "Education",
        "VVI_HealthCareAccess": "HealthCareAccess",
        "VVI_Neighborhood": "Neighborhood",
        "VVI_Housing": "Housing",
        "VVI_CleanEnvironment": "CleanEnvironment",
        "VVI_Social": "Social",
        "VVI_Transportation": "Transportation",
        "VVI_PublicSafety": "PublicSafety",
        "VVI_RiskCategory": "VVI Risk Category",
        "RUCA_PrimaryRUCA": "PrimaryRUCA",
        "RUCA_PrimaryRUCADescription": "PrimaryRUCADescription",
        "RUCA_SecondaryRUCA": "SecondaryRUCA",
        "RUCA_SecondaryRUCADescription": "SecondaryRUCADescription",
        "DCI_2019_2023_DistressScore": "2019-2023 Distress Score",
        "DCI_Quintile": "Quintile (5=Distressed)",
        "DCI_Category": "Risk Categorization",
        "ADI_Source": "ADI Source",
        "ADI_SourceYear": "ADI Source Year",
        "ADI_GeographyUsed": "ADI Geography Used",
        "SVI_SourceYear": "SVI Source Year",
        "SVI_GeographyUsed": "SVI Geography Used",
        "VVI_SourceYear": "VVI Source Year",
        "VVI_GeographyUsed": "VVI Geography Used",
        "RUCA_SourceYear": "RUCA Source Year",
        "RUCA_GeographyUsed": "RUCA Geography Used",
        "DCI_SourceYear": "DCI Source Year",
        "DCI_GeographyUsed": "DCI Geography Used",
        "GeographyCompleteFlag": "Geography Complete Flag",
        "TractBlockGroupConsistencyFlag": "Tract-BlockGroup Consistency Flag",
        "IndexCoverageCount": "Index Coverage Count",
        "ADI_MissingReason": "ADI Missing Reason",
        "SVI_MissingReason": "SVI Missing Reason",
        "VVI_MissingReason": "VVI Missing Reason",
        "RUCA_MissingReason": "RUCA Missing Reason",
        "DCI_MissingReason": "DCI Missing Reason",
        "EJI_SourceYear": "EJI Source Year",
        "EJI_GeographyUsed": "EJI Geography Used",
        "EJI_MissingReason": "EJI Missing Reason",
        "CRE_SourceYear": "CRE Source Year",
        "CRE_GeographyUsed": "CRE Geography Used",
        "CRE_MissingReason": "CRE Missing Reason",
        "FARA_SourceYear": "FARA Source Year",
        "FARA_GeographyUsed": "FARA Geography Used",
        "FARA_MissingReason": "FARA Missing Reason",
        "NRI_SourceYear": "NRI Source Year",
        "NRI_GeographyUsed": "NRI Geography Used",
        "NRI_MissingReason": "NRI Missing Reason",
        "AHRF_SourceYear": "AHRF Source Year",
        "AHRF_GeographyUsed": "AHRF Geography Used",
        "AHRF_MissingReason": "AHRF Missing Reason",
        "Walkability_SourceYear": "Walkability Source Year",
        "Walkability_GeographyUsed": "Walkability Geography Used",
        "Walkability_MissingReason": "Walkability Missing Reason",
    })
    main = ensure_output_columns(main, MAIN_OUTPUT_COLUMN_ORDER)

    tract_count = int(main["Census Tract (FIPS)"].notna().sum()) if "Census Tract (FIPS)" in main.columns else 0
    bg_count = int(main["Block Group (FIPS)"].notna().sum()) if "Block Group (FIPS)" in main.columns else 0
    notes.append(f"Geography coverage: tract populated for {tract_count}/{len(main)} rows; block group populated for {bg_count}/{len(main)} rows.")
    print(f"Geography coverage -> tract: {tract_count}/{len(main)}, block group: {bg_count}/{len(main)}")

    return main, svi_supp, ruca_supp, coi_domains_out, coi_sub_out, ahrf_supp, cre_out, eji_out, fara_out, nri_out, walk_out

# ------------------------------------------------------------------
# 10. Workbook writing / formatting
# ------------------------------------------------------------------
def auto_fit_worksheet(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_idx = col_cells[0].column
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)

NORMALIZED_BAD_COPY_HEADERS = [
    "Record ID",
    "Original Address",
    "Matched / Standardized Address",
    "State / Territory",
    "Census Tract FIPS",
    "ZIP Code (ZCTA)",
    "Block Group FIPS",
    "Latitude",
    "Longitude",
]

def safe_float(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except Exception:
        return None

def clip(value, lo, hi):
    if value is None:
        return None
    return min(max(value, lo), hi)

def round_or_blank(value, digits: int = 6):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return round(float(value), digits)

def bad_from_good(value):
    numeric = safe_float(value)
    if numeric is None:
        return None
    return 100.0 - numeric

def get_normalized_bad_note(spec: dict, fallback_note: str = "") -> str:
    if spec.get("copy_through"):
        return fallback_note or spec.get("note", "") or "Copied unchanged from Main."
    if spec.get("raw_copy"):
        return spec.get("note", "")
    if "Proxy" in spec.get("out_header", ""):
        return f"Bad-oriented proxy derived from {spec['source_header']}; higher values indicate a less favorable or more remote category."
    return f"Bad-oriented 0-100 transformation of {spec['source_header']}; higher values indicate worse conditions, burden, vulnerability, distress, or risk."

def normalize_header_key(value: str) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).casefold()

def map_nri_rating(value):
    text = normalize_header_key(value) if value is not None else ""
    mapping = {
        "very low": 100,
        "relatively low": 75,
        "relatively moderate": 50,
        "relatively high": 25,
        "very high": 0,
    }
    return mapping.get(text)

def map_ruca_proxy(value):
    numeric = safe_float(value)
    if numeric is None:
        return None
    code = int(round(numeric))
    if code in {1, 2, 3}:
        return 100
    if code in {4, 5, 6}:
        return 67
    if code in {7, 8, 9}:
        return 33
    if code == 10:
        return 0
    return None

def map_binary_reverse(value):
    numeric = safe_float(value)
    if numeric is None:
        return None
    code = int(round(numeric))
    if code == 0:
        return 100
    if code == 1:
        return 0
    return None

def map_dci_quintile(value):
    numeric = safe_float(value)
    if numeric is None:
        return None
    code = int(round(numeric))
    return {1: 100, 2: 75, 3: 50, 4: 25, 5: 0}.get(code)

def map_coi_level(value):
    text = normalize_header_key(value) if value is not None else ""
    mapping = {
        "very low": 0,
        "low": 25,
        "moderate": 50,
        "high": 75,
        "very high": 100,
    }
    return mapping.get(text)

def get_sheet_header_map(ws) -> dict[str, int]:
    header_map: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        if header is None:
            continue
        header_text = str(header).strip()
        if header_text:
            header_map[header_text] = col_idx
    return header_map

def get_main_cell_value(main_ws, header_map: dict[str, int], row_idx: int, header: str):
    col_idx = header_map.get(header)
    if not col_idx:
        return None
    return main_ws.cell(row=row_idx, column=col_idx).value

def find_optional_main_header(headers: list[str], preferred: list[str], contains_all: Optional[list[str]] = None) -> Optional[str]:
    by_key = {normalize_header_key(header): header for header in headers}
    for candidate in preferred:
        hit = by_key.get(normalize_header_key(candidate))
        if hit:
            return hit
    if contains_all:
        for header in headers:
            header_key = normalize_header_key(header)
            if all(token in header_key for token in contains_all):
                if "source year" in header_key or "geography used" in header_key or "missing reason" in header_key:
                    continue
                return header
    return None

def add_normalized_bad_sheet(wb) -> None:
    if MAIN_SHEET not in wb.sheetnames:
        return

    if NORMALIZED_BAD_SHEET in wb.sheetnames:
        del wb[NORMALIZED_BAD_SHEET]

    main_ws = wb[MAIN_SHEET]
    insert_idx = wb.sheetnames.index(MAIN_SHEET) + 1
    ws = wb.create_sheet(title=NORMALIZED_BAD_SHEET, index=insert_idx)

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    bold_font = Font(bold=True)
    note_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
    italic_font = Font(italic=True)

    main_header_map = get_sheet_header_map(main_ws)
    main_headers = list(main_header_map.keys())

    fixed_specs = [
        {
            "out_header": "ADI National Rank (Bad 0-100)",
            "source_header": "ADI National Rank",
            "note": "ADI national rank scaled so 100 = least disadvantaged and 0 = most disadvantaged.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (100.0 - safe_float(x)) / 99.0,
        },
        {
            "out_header": "ADI State Rank (Bad 0-100)",
            "source_header": "ADI State Rank",
            "note": "ADI state decile scaled so 100 = least disadvantaged and 0 = most disadvantaged.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (10.0 - safe_float(x)) / 9.0,
        },
        {
            "out_header": "SVI National Theme 1 Rank (Bad 0-100)",
            "source_header": "SVI National Theme 1 Rank",
            "note": "SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI National Theme 2 Rank (Bad 0-100)",
            "source_header": "SVI National Theme 2 Rank",
            "note": "SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI National Theme 3 Rank (Bad 0-100)",
            "source_header": "SVI National Theme 3 Rank",
            "note": "SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI National Theme 4 Rank (Bad 0-100)",
            "source_header": "SVI National Theme 4 Rank",
            "note": "SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI National Overall Rank (Bad 0-100)",
            "source_header": "SVI National Overall Rank",
            "note": "SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI State Theme 1 Rank (Bad 0-100)",
            "source_header": "SVI State Theme 1 Rank",
            "note": "State SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI State Theme 2 Rank (Bad 0-100)",
            "source_header": "SVI State Theme 2 Rank",
            "note": "State SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI State Theme 3 Rank (Bad 0-100)",
            "source_header": "SVI State Theme 3 Rank",
            "note": "State SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI State Theme 4 Rank (Bad 0-100)",
            "source_header": "SVI State Theme 4 Rank",
            "note": "State SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "SVI State Overall Rank (Bad 0-100)",
            "source_header": "SVI State Overall Rank",
            "note": "State SVI percentile rank inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "VVI Overall Score (Bad 0-100)",
            "source_header": "VVI Overall Score",
            "note": "VVI standardized score clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Economic (Bad 0-100)",
            "source_header": "Economic",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Education (Bad 0-100)",
            "source_header": "Education",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "VVI Health Care Access (Bad 0-100)",
            "source_header": "VVI Health Care Access",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Neighborhood (Bad 0-100)",
            "source_header": "Neighborhood",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Housing (Bad 0-100)",
            "source_header": "Housing",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "VVI Clean Environment (Bad 0-100)",
            "source_header": "VVI Clean Environment",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Social (Bad 0-100)",
            "source_header": "Social",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Transportation (Bad 0-100)",
            "source_header": "Transportation",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "Public Safety (Bad 0-100)",
            "source_header": "Public Safety",
            "note": "VVI subscore clipped to [-2, 2]; higher vulnerability becomes lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (2.0 - clip(safe_float(x), -2.0, 2.0)) / 4.0,
        },
        {
            "out_header": "RUCA Primary Code (Bad 0-100 Proxy)",
            "source_header": "RUCA Primary Code",
            "note": "RUCA grouped as an urban-to-rural proxy: 1-3 = 100, 4-6 = 67, 7-9 = 33, 10 = 0.",
            "transform": map_ruca_proxy,
        },
        {
            "out_header": "RUCA Secondary Code (Bad 0-100 Proxy)",
            "source_header": "RUCA Secondary Code",
            "note": "RUCA grouped as an urban-to-rural proxy: 1-3 = 100, 4-6 = 67, 7-9 = 33, 10 = 0.",
            "transform": map_ruca_proxy,
        },
        {
            "out_header": "DCI Distress Score (2019–2023) (Bad 0-100)",
            "source_header": "DCI Distress Score (2019–2023)",
            "note": "DCI distress score inverted so higher distress becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 - safe_float(x),
        },
        {
            "out_header": "DCI Quintile (Bad 0-100)",
            "source_header": "DCI Quintile",
            "note": "DCI quintile mapped as 1 = 100, 2 = 75, 3 = 50, 4 = 25, 5 = 0.",
            "transform": map_dci_quintile,
        },
        {
            "out_header": "RPL SER (Bad 0-100)",
            "source_header": "RPL SER",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL EJI (Bad 0-100)",
            "source_header": "RPL EJI",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL EBM (Bad 0-100)",
            "source_header": "RPL EBM",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL SVM (Bad 0-100)",
            "source_header": "RPL SVM",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL HVM (Bad 0-100)",
            "source_header": "RPL HVM",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL CBM (Bad 0-100)",
            "source_header": "RPL CBM",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "RPL EJI CBM (Bad 0-100)",
            "source_header": "RPL EJI CBM",
            "note": "EJI percentile rank inverted so higher burden becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (1.0 - safe_float(x)),
        },
        {
            "out_header": "PRED3 PE (Bad 0-100)",
            "source_header": "PRED3 PE",
            "note": "Percent in the high CRE vulnerability group; higher raw values are worse.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 - safe_float(x),
        },
        {
            "out_header": "PRED0 PE (Bad 0-100)",
            "source_header": "PRED0 PE",
            "note": "Percent in the low CRE vulnerability group; higher raw values are better.",
            "transform": lambda x: safe_float(x),
        },
        {
            "out_header": "PRED12 PE (Raw %)",
            "source_header": "PRED12 PE",
            "note": "Raw percent in the middle CRE vulnerability groups; copied without normalization.",
            "transform": safe_float,
            "raw_copy": True,
        },
        {
            "out_header": "LILATracts 1 And10 (Bad 0-100)",
            "source_header": "LILATracts 1 And10",
            "note": "Binary reverse mapping: 0 = 100 and 1 = 0.",
            "transform": map_binary_reverse,
        },
        {
            "out_header": "LILATracts Vehicle (Bad 0-100)",
            "source_header": "LILATracts Vehicle",
            "note": "Binary reverse mapping: 0 = 100 and 1 = 0.",
            "transform": map_binary_reverse,
        },
        {
            "out_header": "Low Income Tracts (Bad 0-100)",
            "source_header": "Low Income Tracts",
            "note": "Binary reverse mapping: 0 = 100 and 1 = 0.",
            "transform": map_binary_reverse,
        },
        {
            "out_header": "NRI Risk Score (Bad 0-100)",
            "source_header": "NRI Risk Score",
            "note": "NRI risk score inverted so higher risk becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 - safe_float(x),
        },
        {
            "out_header": "NRI Risk Rating (Bad 0-100)",
            "source_header": "NRI Risk Rating",
            "note": "NRI rating mapped as Very Low = 100, Relatively Low = 75, Relatively Moderate = 50, Relatively High = 25, Very High = 0.",
            "transform": map_nri_rating,
        },
        {
            "out_header": "NRI Expected Annual Loss Score (Bad 0-100)",
            "source_header": "NRI Expected Annual Loss Score",
            "note": "NRI expected annual loss score inverted so higher risk becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 - safe_float(x),
        },
        {
            "out_header": "NRI Social Vulnerability Score (Bad 0-100)",
            "source_header": "NRI Social Vulnerability Score",
            "note": "NRI social vulnerability score inverted so higher vulnerability becomes a lower Good score.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 - safe_float(x),
        },
        {
            "out_header": "NRI Community Resilience Score (Bad 0-100)",
            "source_header": "NRI Community Resilience Score",
            "note": "NRI community resilience score retained so higher resilience remains better.",
            "transform": lambda x: safe_float(x),
        },
        {
            "out_header": "National Walkability Index (Bad 0-100)",
            "source_header": "National Walkability Index",
            "note": "Walkability rank scaled from 1-20 so 1 = 0 and 20 = 100.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 19.0,
        },
        {
            "out_header": "Walkability Household Mix Rank (D2A_Ranked) (Bad 0-100)",
            "source_header": "Walkability Household Mix Rank (D2A_Ranked)",
            "note": "Walkability rank scaled from 1-20 so 1 = 0 and 20 = 100.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 19.0,
        },
        {
            "out_header": "Walkability Employment Mix Rank (D2B_Ranked) (Bad 0-100)",
            "source_header": "Walkability Employment Mix Rank (D2B_Ranked)",
            "note": "Walkability rank scaled from 1-20 so 1 = 0 and 20 = 100.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 19.0,
        },
        {
            "out_header": "Walkability Intersection Density Rank (D3B_Ranked) (Bad 0-100)",
            "source_header": "Walkability Intersection Density Rank (D3B_Ranked)",
            "note": "Walkability rank scaled from 1-20 so 1 = 0 and 20 = 100.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 19.0,
        },
        {
            "out_header": "Walkability Transit / Destination Accessibility Rank (D4A_Ranked) (Bad 0-100)",
            "source_header": "Walkability Transit / Destination Accessibility Rank (D4A_Ranked)",
            "note": "Walkability rank scaled from 1-20 so 1 = 0 and 20 = 100.",
            "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 19.0,
        },
    ]

    output_specs: list[dict] = []
    for header in NORMALIZED_BAD_COPY_HEADERS:
        output_specs.append(
            {
                "out_header": header,
                "source_header": header,
                "note": main_ws.cell(row=2, column=main_header_map[header]).value if header in main_header_map else "Copied unchanged from Main.",
                "copy_through": True,
            }
        )
    output_specs.extend(fixed_specs)

    coi_score_header = find_optional_main_header(
        main_headers,
        preferred=["COI Overall Score", "COI Score", "Overall COI Score"],
        contains_all=["coi", "score"],
    )
    if coi_score_header:
        output_specs.append(
            {
                "out_header": "COI Overall Score (Bad 0-100)",
                "source_header": coi_score_header,
                "note": "COI score scaled from 1-100 so lower opportunity becomes a lower Good score.",
                "transform": lambda x: None if safe_float(x) is None else 100.0 * (safe_float(x) - 1.0) / 99.0,
            }
        )
    coi_level_header = find_optional_main_header(
        main_headers,
        preferred=["COI Level", "COI Category", "COI Overall Category", "COI Overall Level"],
        contains_all=["coi", "level"],
    ) or find_optional_main_header(
        main_headers,
        preferred=["COI Level", "COI Category", "COI Overall Category", "COI Overall Level"],
        contains_all=["coi", "category"],
    )
    if coi_level_header:
        output_specs.append(
            {
                "out_header": "COI Level / Category (Bad 0-100)",
                "source_header": coi_level_header,
                "note": "COI level mapped as Very low = 0, Low = 25, Moderate = 50, High = 75, Very high = 100.",
                "transform": map_coi_level,
            }
        )

    for col_idx, spec in enumerate(output_specs, start=1):
        header_cell = ws.cell(row=1, column=col_idx, value=spec["out_header"])
        header_cell.fill = header_fill
        header_cell.font = bold_font
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fallback_note = main_ws.cell(row=2, column=main_header_map[spec["source_header"]]).value if spec.get("copy_through") and spec["source_header"] in main_header_map else ""
        note_cell = ws.cell(row=2, column=col_idx, value=get_normalized_bad_note(spec, fallback_note=fallback_note))
        note_cell.fill = note_fill
        note_cell.font = italic_font
        note_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for row_idx in range(3, main_ws.max_row + 1):
        for col_idx, spec in enumerate(output_specs, start=1):
            source_value = get_main_cell_value(main_ws, main_header_map, row_idx, spec["source_header"])
            if spec.get("copy_through"):
                out_value = "" if source_value is None else source_value
            else:
                transformed = spec["transform"](source_value)
                if spec.get("raw_copy"):
                    out_value = "" if transformed is None else transformed
                else:
                    out_value = round_or_blank(bad_from_good(transformed), digits=6)
            ws.cell(row=row_idx, column=col_idx, value=out_value)

    ws.freeze_panes = "A3"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    text_col_indexes: set[int] = set()
    for col_idx in range(1, ws.max_column + 1):
        header_text = str(ws.cell(row=1, column=col_idx).value).strip()
        if is_text_identifier_header(header_text):
            text_col_indexes.add(col_idx)
    for col_idx in text_col_indexes:
        for row in ws.iter_rows(min_row=3, min_col=col_idx, max_col=col_idx):
            row[0].number_format = "@"

    auto_fit_worksheet(ws)

def format_workbook(path: Path, add_main_comments: bool = False) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    bold_font = Font(bold=True)
    note_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
    italic_font = Font(italic=True)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        text_col_indexes: set[int] = set()
        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = bold_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if add_main_comments and ws.title == MAIN_SHEET:
                    definition = get_field_definition(str(cell.value).strip())
                    if definition:
                        cell.comment = Comment(definition, "ChatGPT")
                if is_text_identifier_header(str(cell.value).strip()):
                    text_col_indexes.add(cell.column)
        if ws.title == MAIN_SHEET and ws.max_row >= 1:
            ws.insert_rows(2)
            for col_idx, cell in enumerate(ws[1], start=1):
                description = get_field_definition(str(cell.value).strip())
                desc_cell = ws.cell(row=2, column=col_idx, value=description)
                desc_cell.fill = note_fill
                desc_cell.font = italic_font
                desc_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.freeze_panes = "A3"
            ws.auto_filter.ref = ws.dimensions
        start_text_row = 3 if ws.title == MAIN_SHEET else 2
        for col_idx in text_col_indexes:
            for row in ws.iter_rows(min_row=start_text_row, min_col=col_idx, max_col=col_idx):
                row[0].number_format = "@"
        auto_fit_worksheet(ws)

    add_normalized_bad_sheet(wb)
    wb.save(path)

def write_workbook(path: Path, sheets: dict[str, pd.DataFrame], notes: list[str], add_main_glossary: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        wrote_any = False
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            out_df = df.copy()
            raw_cols = list(out_df.columns)
            for col in raw_cols:
                if col in TEXT_IDENTIFIER_RAW_HEADERS:
                    out_df[col] = out_df[col].apply(lambda x: "" if pd.isna(x) else str(x))
            if safe_name == SUMMARY_SHEET and "Metric" in out_df.columns:
                out_df["Metric"] = out_df["Metric"].apply(lambda x: prettify_column_label(str(x)) if isinstance(x, str) else x)
            out_df = out_df.rename(columns=lambda c: prettify_column_label(str(c)))
            out_df.to_excel(writer, sheet_name=safe_name, index=False)
            wrote_any = True
        if add_main_glossary:
            build_main_definitions_sheet().to_excel(writer, sheet_name=MAIN_DEFINITIONS_SHEET[:31], index=False)
        pd.DataFrame({"Notes": notes if notes else ["Run completed without warnings."]}).to_excel(
            writer, sheet_name=NOTES_SHEET[:31], index=False
        )
        if not wrote_any:
            pd.DataFrame().to_excel(writer, sheet_name=MAIN_SHEET[:31], index=False)
    format_workbook(path, add_main_comments=add_main_glossary)

# ------------------------------------------------------------------
# 11. CLI / main
# ------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Geocode addresses and append SES / vulnerability indices.")
    parser.add_argument("--input", type=Path, default=INPUT_FILE, help="Input Excel/CSV file containing usable UniqueID and Address columns")
    parser.add_argument("--output-prefix", type=Path, default=OUTPUT_PREFIX, help="Base output prefix; five workbooks will be written")
    parser.add_argument("--project-root", type=Path, default=None, help="Folder containing this program and the reference files / ZIP archives")
    parser.add_argument("--geocode-mode", choices=[GEOCODE_MODE_AUTO, GEOCODE_MODE_LIVE, GEOCODE_MODE_OFFLINE], default=GEOCODE_MODE_AUTO, help="Use Census live geocoding when available, require it, or force offline fallback")
    parser.add_argument("--geocode-cache", default=GEOCODE_CACHE_FILENAME, help="JSON cache filename stored inside the project root")
    parser.add_argument("--refresh-geocode-cache", action="store_true", help="Ignore any saved geocode cache and rebuild it during this run")
    return parser

def resolve_input_path(project_root: Path, requested_input: Path) -> Path:
    input_path = requested_input if requested_input.is_absolute() else project_root / requested_input
    if input_path.exists():
        return input_path
    if requested_input == INPUT_FILE:
        candidates = sorted(
            [p for p in project_root.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".csv"} and not p.name.startswith("~$")],
            key=lambda p: p.name.lower(),
        )
        if len(candidates) == 1:
            return candidates[0]
    raise FileNotFoundError(f"Input file not found: {input_path}")

def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    notes: list[str] = []

    project_root = args.project_root.resolve() if args.project_root else Path(__file__).resolve().parent
    input_path = resolve_input_path(project_root, args.input)
    output_prefix = args.output_prefix if args.output_prefix.is_absolute() else project_root / args.output_prefix

    ensure_reference_files_from_drive(project_root, notes)

    print("Locating project reference files...")
    paths = patch_special_paths(locate_project_paths(project_root, notes))

    validation_df = validate_project_paths(paths, notes)

    if paths.svi_us_ct:
        notes.append(f"Using SVI national tract source: {paths.svi_us_ct.name}")
    if paths.walkability_gdb:
        walk_csv = find_walkability_csv(paths.root)
        if walk_csv:
            notes.append(f"Using walkability source: {walk_csv.name}")
    if not validation_df.empty:
        print("Reference validation summary:")
        for _, row in validation_df.iterrows():
            print(f"- {row['Source']}: {row['Status']} ({row['Detail']})")

    print("Loading input file...")
    base_df = load_input_addresses(input_path)
    print(f"Loaded {len(base_df):,} rows from {input_path}")

    print("Geocoding unique addresses and deriving tract / blockgroup / ZIP...")
    base_df = geocode_addresses(
        base_df,
        project_root=project_root,
        notes=notes,
        geocode_mode=args.geocode_mode,
        cache_filename=args.geocode_cache,
        refresh_cache=args.refresh_geocode_cache,
    )

    print("Loading reference tables and building outputs...")
    main_df, svi_supp, ruca_supp, coi_domains, coi_sub, ahrf_supp, cre_out, eji_out, fara_out, nri_out, walk_out = build_outputs(base_df, paths, notes)

    summary_df = build_summary_sheet(main_df, validation_df, notes)

    run_audit_df = build_run_audit_sheet(main_df, validation_df, notes)

    main_path = output_prefix.with_name(f"{output_prefix.name}_main.xlsx")
    svi_path = output_prefix.with_name(f"{output_prefix.name}_supplementary_svi.xlsx")
    ruca_path = output_prefix.with_name(f"{output_prefix.name}_supplementary_ruca.xlsx")
    coi_path = output_prefix.with_name(f"{output_prefix.name}_coi.xlsx")
    ahrf_path = output_prefix.with_name(f"{output_prefix.name}_ahrf.xlsx")
    cre_path = output_prefix.with_name(f"{output_prefix.name}_cre.xlsx")
    eji_path = output_prefix.with_name(f"{output_prefix.name}_eji.xlsx")
    fara_path = output_prefix.with_name(f"{output_prefix.name}_fara.xlsx")
    nri_path = output_prefix.with_name(f"{output_prefix.name}_nri.xlsx")
    walk_path = output_prefix.with_name(f"{output_prefix.name}_walkability.xlsx")

    print("Writing output workbooks...")
    write_workbook(main_path, {SUMMARY_SHEET: summary_df, RUN_AUDIT_SHEET: run_audit_df, MAIN_SHEET: main_df}, notes, add_main_glossary=True)
    write_workbook(svi_path, {"SVI Data": svi_supp}, notes)
    write_workbook(ruca_path, {"RUCA Data": ruca_supp}, notes)
    write_workbook(coi_path, {"COI Domain Scores": coi_domains, "COI Subdomain Scores": coi_sub}, notes)
    write_workbook(ahrf_path, {"AHRF Data": ahrf_supp}, notes)
    write_workbook(cre_path, {"CRE Data": cre_out}, notes)
    write_workbook(eji_path, {"EJI Data": eji_out}, notes)
    write_workbook(fara_path, {"FARA Data": fara_out}, notes)
    write_workbook(nri_path, {"NRI Data": nri_out}, notes)
    write_workbook(walk_path, {"Walkability Data": walk_out}, notes)

    print("✓ Done")
    print(f"Main workbook: {main_path}")
    print(f"SVI supplementary workbook: {svi_path}")
    print(f"RUCA supplementary workbook: {ruca_path}")
    print(f"COI workbook: {coi_path}")
    print(f"AHRF workbook: {ahrf_path}")
    print(f"CRE workbook: {cre_path}")
    print(f"EJI workbook: {eji_path}")
    print(f"FARA workbook: {fara_path}")
    print(f"NRI workbook: {nri_path}")
    print(f"Walkability workbook: {walk_path}")

if __name__ == "__main__":
    main()
