import importlib.util
import io
import logging
import os
import re
import sys
import tempfile
import textwrap
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    _HAS_PLOTLY = True
except Exception:
    _HAS_PLOTLY = False


SCRIPT_PATH = Path(__file__).resolve().parent / "ses_indices_from_addresses_drive_enabled.py"
LOG_PATH = SCRIPT_PATH.parent / "dashboard.log"

logger = logging.getLogger("ses_dashboard")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setLevel(logging.INFO)
    sh = logging.StreamHandler()
    sh.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh.setFormatter(fmt)
    sh.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(sh)

INDEX_OPTIONS = [
    "ADI", "SVI", "VVI", "RUCA", "DCI", "AHRF", "CRE", "EJI", "FARA", "NRI", "Walkability", "COI",
]

OUTPUT_NAMES = [
    "Main",
    "Supplementary_SVI",
    "Supplementary_RUCA",
    "COI_Domains",
    "COI_Subdomains",
    "AHRF_Supplementary",
    "CRE_Output",
    "EJI_Output",
    "FARA_Output",
    "NRI_Output",
    "Walkability_Output",
]

GEOCODE_DISPLAY_CANDIDATES = [
    "Record ID", "Original Address", "Matched / Standardized Address", "State / Territory",
    "Census Tract FIPS", "ZIP Code (ZCTA)", "Block Group FIPS", "Latitude", "Longitude",
    "Census Tract (FIPS)", "ZCTA (Zip Code)", "Block Group (FIPS)",
]

FAMILY_RULES = {
    "ADI": {
        "prefixes": ["ADI ", "ADI_"],
        "contains": [],
        "extra": ["National Risk Category", "State Risk Category", "ADI Source", "ADI Source Year", "ADI Geography Used", "ADI Missing Reason"],
    },
    "SVI": {
        "prefixes": ["SVI National ", "SVI State ", "SVI_"],
        "extra": ["SVI Source Year", "SVI Geography Used", "SVI Missing Reason"],
    },
    "VVI": {
        "prefixes": ["VVI ", "VVI_"],
        "extra": [
            "Economic", "Education", "HealthCareAccess", "Neighborhood", "Housing", "CleanEnvironment", "Social",
            "Transportation", "PublicSafety", "VVI Risk Category", "VVI Source Year", "VVI Geography Used", "VVI Missing Reason",
            "VVI Overall Score", "Clean Environment", "Public Safety", "VVI Health Care Access",
        ],
    },
    "RUCA": {
        "prefixes": ["RUCA ", "RUCA_"],
        "extra": ["RUCA Source Year", "RUCA Geography Used", "RUCA Missing Reason"],
    },
    "DCI": {
        "prefixes": ["DCI ", "DCI_"],
        "contains": ["Distress Score", "Quintile", "DCI Category", "Risk Categorization"],
        "extra": ["DCI Source Year", "DCI Geography Used", "DCI Missing Reason"],
    },
    "AHRF": {
        "prefixes": ["AHRF ", "AHRF_"],
        "contains": ["Emergency Department", "Clinic", "Med-Surg ICU", "Trauma Center", "Medicare Approved", "MUA"],
        "extra": ["AHRF Source Year", "AHRF Geography Used", "AHRF Missing Reason"],
    },
    "CRE": {
        "prefixes": ["CRE ", "CRE_", "PRED"],
        "extra": ["CRE Source Year", "CRE Geography Used", "CRE Missing Reason"],
    },
    "EJI": {
        "prefixes": ["EJI ", "EJI_", "RPL "],
        "contains": ["RPL_", "RPL "],
        "extra": ["EJI Source Year", "EJI Geography Used", "EJI Missing Reason"],
    },
    "FARA": {
        "prefixes": ["FARA ", "FARA_", "LILATracts", "LowIncomeTracts", "Low Income Tracts"],
        "extra": ["FARA Source Year", "FARA Geography Used", "FARA Missing Reason"],
    },
    "NRI": {
        "prefixes": ["NRI ", "NRI_"],
        "contains": ["Risk Score", "Expected Annual Loss", "Social Vulnerability", "Community Resilience", "Risk Rating"],
        "extra": ["NRI Source Year", "NRI Geography Used", "NRI Missing Reason"],
    },
    "Walkability": {
        "prefixes": ["Walkability ", "Walkability_", "National Walkability Index"],
        "contains": ["D2A_Ranked", "D2B_Ranked", "D3B_Ranked", "D4A_Ranked", "Intersection Density", "Destination Accessibility"],
        "extra": ["Walkability Source Year", "Walkability Geography Used", "Walkability Missing Reason"],
    },
    "COI": {
        "prefixes": ["COI ", "COI_", "r_COI", "z_COI", "r_ED", "z_ED", "r_HE", "z_HE", "r_SE", "z_SE"],
        "extra": ["COI Source Year", "COI Geography Used", "COI Missing Reason"],
    },
}

APP_RED = "#f64f59"
GOOD_COLOR = "#12c2e9"
MID_COLOR = "#c471ed"
BAD_COLOR = "#f64f59"
CHART_HEIGHT = 360


NATIVE_VISUAL_CONFIG = {
    "ADI": {"metrics": {"ADI_NATRANK": {"min": 1, "max": 100, "higher_is_good": False, "fmt": "int"}, "ADI_STATERNK": {"min": 1, "max": 10, "higher_is_good": False, "fmt": "int"}}, "qualitative_fields": ["National Risk Category", "State Risk Category"]},
    "SVI": {"metrics": {
        "SVI_National_RPL_THEME1": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_National_RPL_THEME2": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_National_RPL_THEME3": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_National_RPL_THEME4": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_National_RPL_THEMES": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_State_RPL_THEME1": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_State_RPL_THEME2": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_State_RPL_THEME3": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_State_RPL_THEME4": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"},
        "SVI_State_RPL_THEMES": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"}}},
    "VVI": {"metrics": {k: {"min": -2, "max": 2, "higher_is_good": False, "fmt": "decimal2"} for k in ["VVI","Economic","Education","HealthCareAccess","Neighborhood","Housing","CleanEnvironment","Social","Transportation","PublicSafety"]}, "qualitative_fields": ["VVI Risk Category"]},
    "COI": {"metrics": {"r_COI": {"min": 1, "max": 100, "higher_is_good": True, "fmt": "int"}, "r_ED": {"min": 1, "max": 100, "higher_is_good": True, "fmt": "int"}, "r_HE": {"min": 1, "max": 100, "higher_is_good": True, "fmt": "int"}, "r_SE": {"min": 1, "max": 100, "higher_is_good": True, "fmt": "int"}}},
    "RUCA": {"metrics": {"PrimaryRUCA": {"min": 1, "max": 10, "higher_is_good": False, "fmt": "int"}}, "qualitative_fields": ["PrimaryRUCADescription"]},
    "DCI": {"metrics": {"2019-2023 Distress Score": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "decimal1"}}, "qualitative_fields": ["Risk Categorization"]},
    "EJI": {"metrics": {k: {"min": 0, "max": 1, "higher_is_good": False, "fmt": "decimal2"} for k in ["RPL_EJI","RPL_SER","RPL_EBM","RPL_SVM","RPL_HVM","RPL_CBM","RPL_EJI_CBM"]}, "table_only_metrics": True},
    "CRE": {"metrics": {"PRED3_PE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "percent_or_number"}, "PRED12_PE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "percent_or_number"}, "PRED0_PE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "percent_or_number"}}},
    "FARA": {"metrics": {"LILATracts_1And10": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "binary"}, "LILATracts_Vehicle": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "binary"}, "LowIncomeTracts": {"min": 0, "max": 1, "higher_is_good": False, "fmt": "binary"}}},
    "NRI": {"metrics": {"RISK_SCORE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "decimal1"}, "EAL_SCORE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "decimal1"}, "SOVI_SCORE": {"min": 0, "max": 100, "higher_is_good": False, "fmt": "decimal1"}, "RESL_SCORE": {"min": 0, "max": 100, "higher_is_good": True, "fmt": "decimal1"}}, "qualitative_fields": ["RISK_RATNG"]},
    "AHRF": {"metrics": {"POS_DIST_ED_TRACT": {"min": 0, "max": 25, "higher_is_good": False, "fmt": "decimal2"}, "POS_DIST_CLINIC_TRACT": {"min": 0, "max": 25, "higher_is_good": False, "fmt": "decimal2"}, "PC_PCT_MEDICARE_APPRVD_FULL_AMT": {"min": 0, "max": 100, "higher_is_good": True, "fmt": "decimal1"}, "POS_DIST_MEDSURG_ICU_TRACT": {"min": 0, "max": 25, "higher_is_good": False, "fmt": "decimal2"}, "POS_DIST_TRAUMA_TRACT": {"min": 0, "max": 25, "higher_is_good": False, "fmt": "decimal2"}}, "qualitative_fields": []},
    "Walkability": {"metrics": {"NatWalkInd": {"min": 1, "max": 20, "higher_is_good": True, "fmt": "int"}, "D2A_Ranked": {"min": 1, "max": 20, "higher_is_good": True, "fmt": "int"}, "D2B_Ranked": {"min": 1, "max": 20, "higher_is_good": True, "fmt": "int"}, "D3B_Ranked": {"min": 1, "max": 20, "higher_is_good": True, "fmt": "int"}, "D4A_Ranked": {"min": 1, "max": 20, "higher_is_good": True, "fmt": "int"}}},
}


DEFINITIONS_CANDIDATES = [
    Path(__file__).resolve().parent / "ses_indices_from_addresses_main(1).xlsx",
    Path(__file__).resolve().parent / "ses_indices_from_addresses_main.xlsx",
]

PASTED_DEFINITIONS_CANDIDATES = [
    Path(__file__).resolve().parent / "Pasted text.txt",
]

# Hard fallback so labels work even if the definitions file is not found at runtime.
HARDCODED_MAIN_DEFINITIONS = {
    "Record ID": {"shown": "Record ID", "technical": "Unique ID", "meaning": "The row or patient identifier from your input file."},
    "Original Address": {"shown": "Original Address", "technical": "Address (as inputted)", "meaning": "The original address exactly as it was provided in the input file."},
    "Matched / Standardized Address": {"shown": "Matched / Standardized Address", "technical": "Geocoded Address", "meaning": "The standardized address returned by the geocoder after it interpreted the input address."},
    "State / Territory": {"shown": "State / Territory", "technical": "State / Territory", "meaning": "The U.S. state or territory linked to the geocoded address."},
    "Census Tract FIPS": {"shown": "Census Tract FIPS", "technical": "Census Tract (FIPS)", "meaning": "The census tract code for the address."},
    "ZIP Code (ZCTA)": {"shown": "ZIP Code (ZCTA)", "technical": "ZCTA (Zip Code)", "meaning": "The ZIP Code Tabulation Area."},
    "Block Group FIPS": {"shown": "Block Group FIPS", "technical": "Block Group (FIPS)", "meaning": "The census block group code for the address."},
    "Latitude": {"shown": "Latitude", "technical": "Latitude", "meaning": "The north-south map coordinate for the geocoded address."},
    "Longitude": {"shown": "Longitude", "technical": "Longitude", "meaning": "The east-west map coordinate for the geocoded address."},
    "ADI National Rank": {"shown": "ADI National Rank", "technical": "ADI_NATRANK", "meaning": "Area Deprivation Index national rank. Higher values generally mean more neighborhood disadvantage compared with places nationwide."},
    "ADI State Rank": {"shown": "ADI State Rank", "technical": "ADI_STATERNK", "meaning": "Area Deprivation Index state rank. Higher values generally mean more neighborhood disadvantage compared with places in the same state."},
    "ADI National Risk Category": {"shown": "ADI National Risk Category", "technical": "National Risk Category", "meaning": "Plain-language category for the national ADI score."},
    "ADI State Risk Category": {"shown": "ADI State Risk Category", "technical": "State Risk Category", "meaning": "Plain-language category for the state ADI score."},
    "SVI National Theme 1 Rank": {"shown": "SVI National Theme 1 Rank", "technical": "SVI_National_RPL_THEME1", "meaning": "Social Vulnerability Index national percentile rank for socioeconomic status."},
    "SVI National Theme 2 Rank": {"shown": "SVI National Theme 2 Rank", "technical": "SVI_National_RPL_THEME2", "meaning": "Social Vulnerability Index national percentile rank for household characteristics."},
    "SVI National Theme 3 Rank": {"shown": "SVI National Theme 3 Rank", "technical": "SVI_National_RPL_THEME3", "meaning": "Social Vulnerability Index national percentile rank for minority status and language."},
    "SVI National Theme 4 Rank": {"shown": "SVI National Theme 4 Rank", "technical": "SVI_National_RPL_THEME4", "meaning": "Social Vulnerability Index national percentile rank for housing type and transportation."},
    "SVI National Overall Rank": {"shown": "SVI National Overall Rank", "technical": "SVI_National_RPL_THEMES", "meaning": "Social Vulnerability Index national overall percentile rank."},
    "SVI State Theme 1 Rank": {"shown": "SVI State Theme 1 Rank", "technical": "SVI_State_RPL_THEME1", "meaning": "Social Vulnerability Index state percentile rank for socioeconomic status."},
    "SVI State Theme 2 Rank": {"shown": "SVI State Theme 2 Rank", "technical": "SVI_State_RPL_THEME2", "meaning": "Social Vulnerability Index state percentile rank for household characteristics."},
    "SVI State Theme 3 Rank": {"shown": "SVI State Theme 3 Rank", "technical": "SVI_State_RPL_THEME3", "meaning": "Social Vulnerability Index state percentile rank for minority status and language."},
    "SVI State Theme 4 Rank": {"shown": "SVI State Theme 4 Rank", "technical": "SVI_State_RPL_THEME4", "meaning": "Social Vulnerability Index state percentile rank for housing type and transportation."},
    "SVI State Overall Rank": {"shown": "SVI State Overall Rank", "technical": "SVI_State_RPL_THEMES", "meaning": "Social Vulnerability Index state overall percentile rank."},
    "VVI Overall Score": {"shown": "VVI Overall Score", "technical": "VVI", "meaning": "Vaccine Vulnerability"},
    "Economic": {"shown": "Economic", "technical": "Economic", "meaning": "Economic"},
    "Education": {"shown": "Education", "technical": "Education", "meaning": "Education"},
    "VVI Health Care Access": {"shown": "VVI Health Care Access", "technical": "HealthCareAccess", "meaning": "Healthcare Access"},
    "Neighborhood": {"shown": "Neighborhood", "technical": "Neighborhood", "meaning": "Neighborhood"},
    "Housing": {"shown": "Housing", "technical": "Housing", "meaning": "Housing"},
    "VVI Clean Environment": {"shown": "VVI Clean Environment", "technical": "CleanEnvironment", "meaning": "Clean Environment"},
    "Social": {"shown": "Social", "technical": "Social", "meaning": "Social"},
    "Transportation": {"shown": "Transportation", "technical": "Transportation", "meaning": "Transportation"},
    "Public Safety": {"shown": "Public Safety", "technical": "PublicSafety", "meaning": "Public Safety"},
    "VVI Risk Category": {"shown": "VVI Risk Category", "technical": "VVI Risk Category", "meaning": "VVI Risk Category"},
    "RUCA Primary Code": {"shown": "RUCA Primary Code", "technical": "PrimaryRUCA", "meaning": "Primary Rural-Urban Commuting Area code for the location."},
    "RUCA Primary Description": {"shown": "RUCA Primary Description", "technical": "PrimaryRUCADescription", "meaning": "Plain-language description of the primary RUCA code."},
    "RUCA Secondary Code": {"shown": "RUCA Secondary Code", "technical": "SecondaryRUCA", "meaning": "Secondary Rural-Urban Commuting Area code for the location."},
    "RUCA Secondary Description": {"shown": "RUCA Secondary Description", "technical": "SecondaryRUCADescription", "meaning": "Plain-language description of the secondary RUCA code."},
    "DCI Distress Score (2019–2023)": {"shown": "DCI Distress Score (2019–2023)", "technical": "2019-2023 Distress Score", "meaning": "Distressed Communities Index score for the ZIP area. Higher values generally mean more economic distress."},
    "DCI Quintile": {"shown": "DCI Quintile", "technical": "Quintile (5=Distressed)", "meaning": "Distressed Communities Index quintile. A value of 5 means the area is in the most distressed group."},
    "DCI Category": {"shown": "DCI Category", "technical": "Risk Categorization", "meaning": "Plain-language category for the Distressed Communities Index quintile."},
    "RPL SER": {"shown": "RPL SER", "technical": "RPL_SER", "meaning": "Social & Environmental"},
    "RPL EJI": {"shown": "RPL EJI", "technical": "RPL_EJI", "meaning": "Combined (social, environmental, health)"},
    "RPL EBM": {"shown": "RPL EBM", "technical": "RPL_EBM", "meaning": "Environmental"},
    "RPL SVM": {"shown": "RPL SVM", "technical": "RPL_SVM", "meaning": "Social Vulnerability"},
    "RPL HVM": {"shown": "RPL HVM", "technical": "RPL_HVM", "meaning": "Health Vulnerability"},
    "RPL CBM": {"shown": "RPL CBM", "technical": "RPL_CBM", "meaning": "Climate Burden"},
    "RPL EJI CBM": {"shown": "RPL EJI CBM", "technical": "RPL_EJI_CBM", "meaning": "Combined (social, environmental, health, climate)"},
    "PRED3 PE": {"shown": "PRED3 PE", "technical": "PRED3_PE", "meaning": "CRE estimate of predicted 3-year prevalence or burden for this tract."},
    "PRED12 PE": {"shown": "PRED12 PE", "technical": "PRED12_PE", "meaning": "CRE estimate of predicted 12-year prevalence or burden for this tract."},
    "PRED0 PE": {"shown": "PRED0 PE", "technical": "PRED0_PE", "meaning": "CRE baseline predicted prevalence or burden for this tract."},
    "LILATracts 1 And10": {"shown": "LILATracts 1 And10", "technical": "LILATracts_1And10", "meaning": "FARA flag indicating a low-income, low-access tract using the 1-mile urban and 10-mile rural definition."},
    "LILATracts Vehicle": {"shown": "LILATracts Vehicle", "technical": "LILATracts_Vehicle", "meaning": "FARA flag indicating a low-income, low-access tract using the vehicle-access definition."},
    "Low Income Tracts": {"shown": "Low Income Tracts", "technical": "LowIncomeTracts", "meaning": "FARA flag indicating the tract is classified as low income."},
    "NRI Risk Score": {"shown": "NRI Risk Score", "technical": "RISK_SCORE", "meaning": "NRI overall risk score. Higher values generally mean greater overall natural hazard risk."},
    "NRI Risk Rating": {"shown": "NRI Risk Rating", "technical": "RISK_RATNG", "meaning": "Plain-language natural hazard risk rating from the NRI source."},
    "NRI Expected Annual Loss Score": {"shown": "NRI Expected Annual Loss Score", "technical": "EAL_SCORE", "meaning": "NRI expected annual loss score. Higher values mean more expected yearly loss from hazards."},
    "NRI Social Vulnerability Score": {"shown": "NRI Social Vulnerability Score", "technical": "SOVI_SCORE", "meaning": "NRI social vulnerability score. Higher values mean greater social vulnerability."},
    "NRI Community Resilience Score": {"shown": "NRI Community Resilience Score", "technical": "RESL_SCORE", "meaning": "NRI community resilience score. Higher values generally mean more resilience or capacity to recover."},
    "AHRF MUA Flag (Census Tract)": {"shown": "AHRF MUA Flag (Census Tract)", "technical": "HRSA_MUA_CENSUS_TRACT", "meaning": "Indicates whether the census tract falls within a HRSA-designated Medically Underserved Area or Population."},
    "AHRF Distance to Emergency Department (Tract)": {"shown": "AHRF Distance to Emergency Department (Tract)", "technical": "POS_DIST_ED_TRACT", "meaning": "Estimated distance in miles from the census tract to the nearest emergency department."},
    "AHRF Distance to Clinic (Tract)": {"shown": "AHRF Distance to Clinic (Tract)", "technical": "POS_DIST_CLINIC_TRACT", "meaning": "Estimated distance in miles from the census tract to the nearest clinic or outpatient care site."},
    "AHRF Percent Medicare Approved Full Amount": {"shown": "AHRF Percent Medicare Approved Full Amount", "technical": "PC_PCT_MEDICARE_APPRVD_FULL_AMT", "meaning": "Percentage of clinicians or services in the area that accepted Medicare's approved amount in full."},
    "AHRF Distance to Med-Surg ICU (Tract)": {"shown": "AHRF Distance to Med-Surg ICU (Tract)", "technical": "POS_DIST_MEDSURG_ICU_TRACT", "meaning": "Estimated distance in miles from the census tract to the nearest hospital with a medical-surgical intensive care unit."},
    "AHRF Distance to Trauma Center (Tract)": {"shown": "AHRF Distance to Trauma Center (Tract)", "technical": "POS_DIST_TRAUMA_TRACT", "meaning": "Estimated distance in miles from the census tract to the nearest trauma center."},
    "National Walkability Index": {"shown": "National Walkability Index", "technical": "NatWalkInd", "meaning": "National Walkability Index summary score for the Census block group. Higher values generally mean a more walkable built environment."},
    "Walkability Household Mix Rank (D2A_Ranked)": {"shown": "Walkability Household Mix Rank (D2A_Ranked)", "technical": "D2A_Ranked", "meaning": "Walkability submetric rank related to household mix and neighborhood composition."},
    "Walkability Employment Mix Rank (D2B_Ranked)": {"shown": "Walkability Employment Mix Rank (D2B_Ranked)", "technical": "D2B_Ranked", "meaning": "Walkability submetric rank related to employment mix and destination diversity."},
    "Walkability Intersection Density Rank (D3B_Ranked)": {"shown": "Walkability Intersection Density Rank (D3B_Ranked)", "technical": "D3B_Ranked", "meaning": "Walkability submetric rank related to street intersection density and connectivity."},
    "Walkability Transit / Destination Accessibility Rank (D4A_Ranked)": {"shown": "Walkability Transit / Destination Accessibility Rank (D4A_Ranked)", "technical": "D4A_Ranked", "meaning": "Walkability submetric rank related to transit or destination accessibility."},
}

SHORT_LABEL_OVERRIDES = {
    "ADI National Rank": "National Rank",
    "ADI_NATRANK": "National Rank",
    "ADI State Rank": "State Rank",
    "ADI_STATERNK": "State Rank",
    "SVI National Theme 1 Rank": "National - SES",
    "SVI_National_RPL_THEME1": "National - SES",
    "SVI National Theme 2 Rank": "National - Household",
    "SVI_National_RPL_THEME2": "National - Household",
    "SVI National Theme 3 Rank": "National - Minority Status & Language",
    "SVI_National_RPL_THEME3": "National - Minority Status & Language",
    "SVI National Theme 4 Rank": "National - Housing & Transportation",
    "SVI_National_RPL_THEME4": "National - Housing & Transportation",
    "SVI National Overall Rank": "National - Overall",
    "SVI_National_RPL_THEMES": "National - Overall",
    "SVI State Theme 1 Rank": "State - SES",
    "SVI_State_RPL_THEME1": "State - SES",
    "SVI State Theme 2 Rank": "State - Household",
    "SVI_State_RPL_THEME2": "State - Household",
    "SVI State Theme 3 Rank": "State - Minority Status & Language",
    "SVI_State_RPL_THEME3": "State - Minority Status & Language",
    "SVI State Theme 4 Rank": "State - Housing & Transportation",
    "SVI_State_RPL_THEME4": "State - Housing & Transportation",
    "SVI State Overall Rank": "State - Overall",
    "SVI_State_RPL_THEMES": "State - Overall",
    "VVI Overall Score": "Vaccine Vulnerability",
    "VVI": "Vaccine Vulnerability",
    "Economic": "Economic",
    "Education": "Education",
    "VVI Health Care Access": "Healthcare Access",
    "HealthCareAccess": "Healthcare Access",
    "Neighborhood": "Neighborhood",
    "Housing": "Housing",
    "VVI Clean Environment": "Clean Environment",
    "CleanEnvironment": "Clean Environment",
    "Social": "Social",
    "Transportation": "Transportation",
    "Public Safety": "Public Safety",
    "PublicSafety": "Public Safety",
    "RUCA Primary Code": "Primary RUCA Code",
    "PrimaryRUCA": "Primary RUCA Code",
    "RUCA Primary Description": "Primary RUCA Description",
    "PrimaryRUCADescription": "Primary RUCA Description",
    "RUCA Secondary Code": "Secondary RUCA Code",
    "SecondaryRUCA": "Secondary RUCA Code",
    "RUCA Secondary Description": "Secondary RUCA Description",
    "SecondaryRUCADescription": "Secondary RUCA Description",
    "DCI Distress Score (2019–2023)": "Distress Score",
    "2019-2023 Distress Score": "Distress Score",
    "DCI Quintile": "Quintile",
    "Quintile (5=Distressed)": "Quintile",
    "DCI Category": "Category",
    "Risk Categorization": "Category",
    "RPL SER": "Social & Environmental",
    "RPL_SER": "Social & Environmental",
    "RPL EJI": "Combined",
    "RPL_EJI": "Combined",
    "RPL EBM": "Environmental",
    "RPL_EBM": "Environmental",
    "RPL SVM": "Social Vulnerability",
    "RPL_SVM": "Social Vulnerability",
    "RPL HVM": "Health Vulnerability",
    "RPL_HVM": "Health Vulnerability",
    "RPL CBM": "Climate Burden",
    "RPL_CBM": "Climate Burden",
    "RPL EJI CBM": "Combined + Climate",
    "RPL_EJI_CBM": "Combined + Climate",
    "PRED3 PE": "3-Year CRE",
    "PRED3_PE": "3-Year CRE",
    "PRED12 PE": "12-Year CRE",
    "PRED12_PE": "12-Year CRE",
    "PRED0 PE": "Baseline CRE",
    "PRED0_PE": "Baseline CRE",
    "LILATracts 1 And10": "1/10-Mile Access",
    "LILATracts_1And10": "1/10-Mile Access",
    "LILATracts Vehicle": "Vehicle Access",
    "LILATracts_Vehicle": "Vehicle Access",
    "Low Income Tracts": "Low Income",
    "LowIncomeTracts": "Low Income",
    "NRI Risk Score": "Overall Risk",
    "RISK_SCORE": "Overall Risk",
    "NRI Risk Rating": "Risk Rating",
    "RISK_RATNG": "Risk Rating",
    "NRI Expected Annual Loss Score": "Annual Loss",
    "EAL_SCORE": "Annual Loss",
    "NRI Social Vulnerability Score": "Social Vulnerability",
    "SOVI_SCORE": "Social Vulnerability",
    "NRI Community Resilience Score": "Community Resilience",
    "RESL_SCORE": "Community Resilience",
    "AHRF MUA Flag (Census Tract)": "MUA Flag",
    "HRSA_MUA_CENSUS_TRACT": "MUA Flag",
    "AHRF Distance to Emergency Department (Tract)": "ED Distance",
    "POS_DIST_ED_TRACT": "ED Distance",
    "AHRF Distance to Clinic (Tract)": "Clinic Distance",
    "POS_DIST_CLINIC_TRACT": "Clinic Distance",
    "AHRF Percent Medicare Approved Full Amount": "Medicare Approved %",
    "PC_PCT_MEDICARE_APPRVD_FULL_AMT": "Medicare Approved %",
    "AHRF Distance to Med-Surg ICU (Tract)": "ICU Distance",
    "POS_DIST_MEDSURG_ICU_TRACT": "ICU Distance",
    "AHRF Distance to Trauma Center (Tract)": "Trauma Distance",
    "POS_DIST_TRAUMA_TRACT": "Trauma Distance",
    "National Walkability Index": "Walkability",
    "NatWalkInd": "Walkability",
    "Walkability Household Mix Rank (D2A_Ranked)": "Household Mix",
    "D2A_Ranked": "Household Mix",
    "Walkability Employment Mix Rank (D2B_Ranked)": "Employment Mix",
    "D2B_Ranked": "Employment Mix",
    "Walkability Intersection Density Rank (D3B_Ranked)": "Intersection Density",
    "D3B_Ranked": "Intersection Density",
    "Walkability Transit / Destination Accessibility Rank (D4A_Ranked)": "Transit / Access",
    "D4A_Ranked": "Transit / Access",
}

def normalize_definition_key(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())

def _definition_keys(shown: str, technical: str) -> set[str]:
    keys = set()
    for item in [shown, technical]:
        item = str(item or "").strip()
        if not item:
            continue
        variants = {
            item,
            item.replace("_", " "),
            item.replace("-", " "),
            item.replace("/", " "),
            re.sub(r"(?<!^)(?=[A-Z])", " ", item),
        }
        for v in variants:
            nk = normalize_definition_key(v)
            if nk:
                keys.add(nk)
    return keys

def load_main_definitions() -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}

    def add_row(shown: str, technical: str, meaning: str):
        meta = {"shown": str(shown).strip(), "technical": str(technical).strip(), "meaning": str(meaning).strip()}
        for nk in _definition_keys(meta["shown"], meta["technical"]):
            out[nk] = meta

    for meta in HARDCODED_MAIN_DEFINITIONS.values():
        add_row(meta.get("shown", ""), meta.get("technical", ""), meta.get("meaning", ""))

    for path in DEFINITIONS_CANDIDATES:
        if path.exists():
            try:
                defs = pd.read_excel(path, sheet_name="Main Definitions").fillna("")
                for _, row in defs.iterrows():
                    add_row(row.get("Field Shown in Main", ""), row.get("Technical Field Name", ""), row.get("Plain-English Meaning", ""))
                break
            except Exception:
                pass

    for path in PASTED_DEFINITIONS_CANDIDATES:
        if path.exists():
            try:
                defs = pd.read_csv(path, sep="\t").fillna("")
                for _, row in defs.iterrows():
                    add_row(row.get("Field Shown in Main", ""), row.get("Technical Field Name", ""), row.get("Plain-English Meaning", ""))
                break
            except Exception:
                pass

    return out

MAIN_DEFINITIONS = load_main_definitions()

def lookup_definition(field_name: str) -> dict[str, str]:
    raw = str(field_name).strip()
    candidates = {
        raw,
        raw.replace("__", " "),
        raw.replace("_", " "),
        raw.replace("-", " "),
        raw.replace("/", " "),
        re.sub(r"(?<!^)(?=[A-Z])", " ", raw),
        prettify_fallback_label(raw),
    }
    for candidate in candidates:
        meta = MAIN_DEFINITIONS.get(normalize_definition_key(candidate))
        if meta:
            return meta
    return {}
def lookup_definition(field_name: str) -> dict[str, str]:
    raw = str(field_name).strip()
    candidates = [
        raw,
        raw.replace("__", " "),
        raw.replace("_", " "),
        raw.replace("-", " "),
        raw.replace("  ", " "),
        prettify_fallback_label(raw),
    ]
    for candidate in candidates:
        meta = MAIN_DEFINITIONS.get(normalize_definition_key(candidate))
        if meta:
            return meta
    return {}

def short_label_from_definition(field_name: str, technical_name: str, meaning: str) -> str:
    field = str(field_name).strip()
    technical = str(technical_name).strip()

    for candidate in (field, technical):
        if candidate in SHORT_LABEL_OVERRIDES:
            return SHORT_LABEL_OVERRIDES[candidate]
        norm = normalize_definition_key(candidate)
        for src, out in SHORT_LABEL_OVERRIDES.items():
            if normalize_definition_key(src) == norm:
                return out

    pretty = prettify_fallback_label(field or technical)
    return pretty

GRAPH_EXPLANATIONS = {
    "SVI": {
        "SVI SES": "SES = socioeconomic status.",
        "SVI SES Score": "SES = socioeconomic status.",
        "SVI SES Percentile": "SES = socioeconomic status; higher percentile indicates greater vulnerability.",
        "SVI Household": "Household = household composition and disability.",
        "SVI Household Score": "Household = household composition and disability.",
        "SVI Household Percentile": "Household = household composition and disability.",
        "SVI Minority/Language": "Minority/Language = racial and ethnic minority status and English language proficiency.",
        "SVI Minority/Language Score": "Minority/Language = racial and ethnic minority status and English language proficiency.",
        "SVI Minority/Language Percentile": "Minority/Language = racial and ethnic minority status and English language proficiency.",
        "SVI Housing/Transport": "Housing/Transport = housing type and transportation.",
        "SVI Housing/Transport Score": "Housing/Transport = housing type and transportation.",
        "SVI Housing/Transport Percentile": "Housing/Transport = housing type and transportation.",
    }
}



def prettify_fallback_label(name: str) -> str:
    text = str(name).replace("__", " ").replace("_", " ").strip()
    text = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", text)
    text = " ".join(text.split())
    return text

def display_label(name: str) -> str:
    raw = str(name).strip()
    if raw == "UniqueID":
        return "Individual ID"
    if raw == "Patient ID":
        return "Individual ID"
    meta = lookup_definition(raw)
    if meta:
        return short_label_from_definition(
            meta.get("shown", raw),
            meta.get("technical", raw),
            meta.get("meaning", ""),
        )
    return prettify_fallback_label(raw)

def value_to_score(value):
    num = coerce_numeric(value)
    if num is not None:
        if num <= 1:
            return float(num)
        if num <= 100:
            return float(num) / 100.0
    text = textify(value).lower()
    if not text:
        return None
    if any(tok in text for tok in ["low", "good", "advantaged", "accessible", "resilient"]):
        return 0.0
    if any(tok in text for tok in ["medium", "moderate", "mid", "average"]):
        return 0.5
    if any(tok in text for tok in ["high", "bad", "poor", "distressed", "deprivation", "vulnerab", "limited"]):
        return 1.0
    return None


def interpolate_hex(c1: str, c2: str, t: float) -> str:
    c1 = c1.lstrip("#")
    c2 = c2.lstrip("#")
    a = tuple(int(c1[i:i+2], 16) for i in (0, 2, 4))
    b = tuple(int(c2[i:i+2], 16) for i in (0, 2, 4))
    rgb = tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return "#%02x%02x%02x" % rgb


def gradient_color(score) -> str:
    if score is None:
        return "transparent"
    score = max(0.0, min(1.0, float(score)))
    if score <= 0.5:
        return interpolate_hex(GOOD_COLOR, MID_COLOR, score / 0.5)
    return interpolate_hex(MID_COLOR, BAD_COLOR, (score - 0.5) / 0.5)


def rename_display_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rename_map = {}
    for col in out.columns:
        label = str(col)
        if label == "UniqueID":
            rename_map[col] = "Individual ID"
        elif label == "Patient ID":
            rename_map[col] = "Individual ID"
        else:
            rename_map[col] = display_label(label)
    out = out.rename(columns=rename_map)
    if "Individual ID" in out.columns and list(out.columns).count("Individual ID") > 1:
        out = out.loc[:, ~out.columns.duplicated()]
    return out


def get_metric_base_name(name: str) -> str:
    return str(name).split("__")[0].strip()


def table_cell_color(value, family: str | None, column_name: str) -> str | None:
    base = get_metric_base_name(column_name)
    if not family:
        return None
    if family == "RUCA":
        return None
    cfg = resolve_metric_config(family, base)
    if not cfg:
        return None
    if cfg.get("fmt") == "binary" or cfg.get("no_table_color"):
        return None
    num = coerce_numeric(value)
    if num is None:
        return None
    frac = score_fraction(float(num), float(cfg["min"]), float(cfg["max"]))
    badness = 1.0 - frac if bool(cfg.get("higher_is_good", False)) else frac
    return gradient_color(badness)


def style_gradient_table(df: pd.DataFrame, family: str | None = None):
    display_df = rename_display_columns(df)
    styles = pd.DataFrame("", index=display_df.index, columns=display_df.columns)
    original_cols = list(df.columns)
    for idx, orig_col in enumerate(original_cols):
        disp_col = display_df.columns[idx]
        for row_idx in display_df.index:
            color = table_cell_color(df.loc[row_idx, orig_col], family, orig_col)
            if color:
                styles.loc[row_idx, disp_col] = f"background-color: {color}; color: white;"
    return display_df.style.apply(lambda _: styles, axis=None)

def graph_explanation_lines(family: str, chart_df: pd.DataFrame) -> list[str]:
    out = []
    for metric in chart_df["metric"].tolist():
        metric_name = str(metric).strip()
        meta = lookup_definition(metric_name)
        label = display_label(metric_name)
        explanation = (meta or {}).get("meaning", "")
        if explanation:
            out.append(f"**{label}**: {explanation}")
    uniq = []
    seen = set()
    for item in out:
        if item not in seen:
            uniq.append(item)
            seen.add(item)
    return uniq

def ensure_pipeline_exists() -> None:
    if not SCRIPT_PATH.exists():
        st.error(f"Could not find pipeline file: {SCRIPT_PATH}")
        st.stop()


def load_pipeline_module(path: Path):
    name = "ses_pipeline_dashboard"
    spec = importlib.util.spec_from_file_location(name, str(path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load module from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)  # type: ignore
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module


@st.cache_resource(show_spinner=False)
def get_pipeline(pipeline_mtime_ns: int):
    logger.info("Loading pipeline module from %s", SCRIPT_PATH)
    t0 = time.perf_counter()
    mod = load_pipeline_module(SCRIPT_PATH)
    logger.info("Pipeline loaded in %.2fs", time.perf_counter() - t0)
    return mod


@st.cache_resource(show_spinner=False)
def get_project_paths(pipeline_mtime_ns: int):
    mod = get_pipeline(pipeline_mtime_ns)
    project_root = SCRIPT_PATH.parent
    notes: list[str] = []

    # The drive-enabled pipeline downloads the large Google Drive reference
    # folder into reference_data/ before locating files. This lets the
    # dashboard run without storing the large reference folders in GitHub.
    if hasattr(mod, "ensure_reference_files_from_drive"):
        mod.ensure_reference_files_from_drive(project_root, notes)

    paths = mod.locate_project_paths(project_root, notes)
    return mod.patch_special_paths(paths)


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    seen: dict[str, int] = {}
    new_cols: list[str] = []
    for col in out.columns:
        base = str(col)
        seen[base] = seen.get(base, 0) + 1
        new_cols.append(base if seen[base] == 1 else f"{base}__{seen[base]}")
    out.columns = new_cols
    return out


def serialize_df(df: pd.DataFrame) -> str:
    return make_unique_columns(df).to_json(orient="records", date_format="iso")


def deserialize_df(value: str) -> pd.DataFrame:
    if not value:
        return pd.DataFrame()
    return pd.read_json(io.StringIO(value), orient="records")


def _clean_geo(value, width=None):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text == "" or text.lower() in {"nan", "none"}:
        return None
    text = text.split(".")[0].strip()
    text = "".join(ch for ch in text if ch.isdigit()) if width is not None else text
    if width is not None:
        if text == "":
            return None
        text = text.zfill(width)[-width:]
    return text or None


def normalize_geo_keys(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "geoid" in out.columns:
        out["geoid"] = out["geoid"].apply(lambda x: _clean_geo(x, 15))
    if "census_tract" in out.columns:
        out["census_tract"] = out["census_tract"].apply(lambda x: _clean_geo(x, 11))
    if "zipcode" in out.columns:
        out["zipcode"] = out["zipcode"].apply(lambda x: _clean_geo(x, 5))
    if "matched_zipcode" in out.columns:
        out["matched_zipcode"] = out["matched_zipcode"].apply(lambda x: _clean_geo(x, 5))
    if "blockgroup_fips" in out.columns:
        out["blockgroup_fips"] = out["blockgroup_fips"].apply(lambda x: _clean_geo(x, 12))
    if "state_fips" in out.columns:
        out["state_fips"] = out["state_fips"].apply(lambda x: _clean_geo(x, 2))
    if "census_tract" in out.columns and "state_fips" in out.columns:
        out.loc[out["state_fips"].isna() & out["census_tract"].notna(), "state_fips"] = out["census_tract"].astype(str).str[:2]
    if "blockgroup_fips" in out.columns and "census_tract" in out.columns:
        out.loc[out["census_tract"].isna() & out["blockgroup_fips"].notna(), "census_tract"] = out["blockgroup_fips"].astype(str).str[:11]
    return out


@st.cache_data(show_spinner=False)
def cached_geocode(input_df_json: str, pipeline_mtime_ns: int):
    mod = get_pipeline(pipeline_mtime_ns)
    input_df = deserialize_df(input_df_json)
    notes: list[str] = []
    project_root = SCRIPT_PATH.parent

    t0 = time.perf_counter()
    geocoded = mod.geocode_addresses(
        input_df,
        project_root=project_root,
        notes=notes,
        geocode_mode=getattr(mod, "GEOCODE_MODE_AUTO", "auto"),
    )
    geocoded = normalize_geo_keys(make_unique_columns(geocoded))
    logger.info("geocode_addresses completed in %.2fs", time.perf_counter() - t0)
    return serialize_df(geocoded), notes


@st.cache_data(show_spinner=False)
def cached_build_outputs(geocoded_json: str, pipeline_mtime_ns: int):
    mod = get_pipeline(pipeline_mtime_ns)
    paths = get_project_paths(pipeline_mtime_ns)
    geocoded = normalize_geo_keys(deserialize_df(geocoded_json))
    notes: list[str] = []

    t0 = time.perf_counter()
    outputs = mod.build_outputs(geocoded, paths, notes)
    logger.info("build_outputs completed in %.2fs", time.perf_counter() - t0)

    serialized_outputs = []
    for item in outputs:
        serialized_outputs.append(serialize_df(item) if isinstance(item, pd.DataFrame) else "")
    return serialized_outputs, notes



def sanitize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean values before writing with openpyxl.
    Excel/openpyxl can fail with SerialisationError when cells contain invalid XML
    control characters or strings longer than Excel's 32,767-character cell limit.
    """
    out = make_unique_columns(df).copy()
    illegal_xml_re = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

    def clean_value(value):
        if isinstance(value, str):
            value = illegal_xml_re.sub("", value)
            if len(value) > 32767:
                value = value[:32760] + "..."
        return value

    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].map(clean_value)
    return out

def template_excel_bytes() -> bytes:
    df = pd.DataFrame({
        "UniqueID": ["P001", "P002", "P003"],
        "Address": [
            "123 Main St, Springfield, VA 22150",
            "456 Oak Ave, Brambleton, VA 20148",
            "789 Pine Rd, Seattle, WA 98101",
        ],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Addresses", index=False)
    buf.seek(0)
    return buf.getvalue()


def excel_bytes_from_df(df: pd.DataFrame, sheet_name: str = "Results") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sanitize_for_excel(df).to_excel(writer, sheet_name=sheet_name, index=False)
    buf.seek(0)
    return buf.getvalue()


def zip_bytes_from_outputs(outputs: list[pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, df in zip(OUTPUT_NAMES, outputs):
            inner = excel_bytes_from_df(df, sheet_name=name[:31] or "Sheet1")
            zf.writestr(f"{name}.xlsx", inner)
    buf.seek(0)
    return buf.getvalue()


def textify(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def coerce_numeric(v):
    if v is None or isinstance(v, bool):
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    text = str(v).strip().replace(",", "")
    if text == "" or text.lower() in {"nan", "none"}:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except Exception:
        return None


def pretty_metric_name(name: str, width: int = 28) -> str:
    clean = display_label(str(name)).replace("_", " ").strip()
    clean = " ".join(clean.split())
    return "<br>".join(textwrap.wrap(clean, width=width)) or str(name)


def available_columns(df: pd.DataFrame, family: str) -> list[str]:
    config = NATIVE_VISUAL_CONFIG.get(family, {})
    metric_keys = set(config.get("metrics", {}).keys())
    qualitative_keys = set(config.get("qualitative_fields", []))

    if config.get("table_only_metrics"):
        out = []
        for col in df.columns:
            base = str(col).split("__")[0]
            if base in metric_keys:
                out.append(str(col))
                continue
            meta = lookup_definition(base)
            shown = meta.get("shown", "") if meta else ""
            technical = meta.get("technical", "") if meta else ""
            if shown in metric_keys or technical in metric_keys:
                out.append(str(col))
        seen = set()
        uniq = []
        for c in out:
            if c not in seen:
                uniq.append(c)
                seen.add(c)
        return uniq

    rules = FAMILY_RULES[family]
    out: list[str] = []
    for col in df.columns:
        sc = str(col)
        if any(sc.startswith(pref) for pref in rules.get("prefixes", [])):
            out.append(sc)
            continue
        if any(token in sc for token in rules.get("contains", [])):
            out.append(sc)
            continue
        if sc in rules.get("extra", []):
            out.append(sc)
            continue
    for col in df.columns:
        base = str(col).split("__")[0]
        if base in metric_keys or base in qualitative_keys:
            out.append(str(col))
            continue
        meta = lookup_definition(base)
        shown = meta.get("shown", "") if meta else ""
        technical = meta.get("technical", "") if meta else ""
        if shown in metric_keys or technical in metric_keys or shown in qualitative_keys or technical in qualitative_keys:
            out.append(str(col))
    seen = set()
    uniq = []
    for c in out:
        if c not in seen:
            uniq.append(c)
            seen.add(c)
    return uniq

def split_family_table(df: pd.DataFrame, family: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    cols = available_columns(df, family)
    if not cols:
        return pd.DataFrame({"Note": [f"No output columns found for {family}."]}), pd.DataFrame()
    fam = df[cols].copy()
    numeric_cols = [c for c in fam.columns if fam[c].map(coerce_numeric).notna().any()]
    nonempty_cols = [c for c in fam.columns if fam[c].astype(str).replace({"nan": "", "None": ""}).str.strip().ne("").any()]
    fam = fam[nonempty_cols] if nonempty_cols else fam
    num = fam[numeric_cols].copy() if numeric_cols else pd.DataFrame(index=fam.index)
    return fam, num


PLOT_EXCLUDE_TOKENS = [
    "year", "source", "missing reason", "geography used", "geocode", "record id", "address",
    "latitude", "longitude", "fips", "zip code", "zip", "tract", "block group", "state / territory",
]


def should_exclude_from_plot(col: str) -> bool:
    sc = str(col).strip().lower()
    return any(tok in sc for tok in PLOT_EXCLUDE_TOKENS)


def resolve_metric_config(family: str, metric_name: str) -> dict:
    base = str(metric_name).split("__")[0]
    cfg = NATIVE_VISUAL_CONFIG.get(family, {}).get("metrics", {})
    if base in cfg:
        return cfg[base]
    meta = lookup_definition(base)
    for candidate in [meta.get("shown", ""), meta.get("technical", "")]:
        if candidate in cfg:
            return cfg[candidate]
    return {}


def format_native_value(value: float, fmt: str) -> str:
    if fmt == "int":
        return f"{int(round(value))}"
    if fmt == "decimal1":
        return f"{value:.1f}"
    if fmt == "decimal2":
        return f"{value:.2f}"
    if fmt == "binary":
        return f"{int(round(value))}"
    if fmt == "percent_or_number":
        return f"{value:.0f}%"
    return f"{value:g}"


def format_scale_label(value: float, fmt: str) -> str:
    if fmt in {"int", "binary"}:
        return f"{int(round(value))}"
    if fmt == "percent_or_number":
        return f"{value:.0f}%"
    if fmt == "decimal1":
        return f"{value:.1f}"
    if fmt == "decimal2":
        return f"{value:.1f}"
    return f"{value:g}"


def score_fraction(value: float, min_v: float, max_v: float) -> float:
    if max_v == min_v:
        return 0.5
    frac = (value - min_v) / (max_v - min_v)
    return max(0.0, min(1.0, frac))


def metric_table_to_chart_df(df: pd.DataFrame, family: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["metric", "metric_label", "value"])
    row = df.iloc[0].to_dict()
    rows = []
    for k, v in row.items():
        if should_exclude_from_plot(str(k)):
            continue
        num = coerce_numeric(v)
        if num is None:
            continue
        cfg = resolve_metric_config(family, str(k))
        if not cfg:
            continue
        min_v = float(cfg["min"])
        max_v = float(cfg["max"])
        higher_is_good = bool(cfg.get("higher_is_good", False))
        frac = score_fraction(float(num), min_v, max_v)
        color_frac = frac if not higher_is_good else 1.0 - frac
        rows.append({
            "metric": str(k),
            "metric_label": display_label(str(k)),
            "metric_display": display_label(str(k)),
            "raw_value": float(num),
            "value": float(num),
            "value_text": format_native_value(float(num), cfg.get("fmt", "decimal2")),
            "min": min_v,
            "mid": (min_v + max_v) / 2.0,
            "max": max_v,
            "higher_is_good": higher_is_good,
            "position_frac": frac,
            "fmt": cfg.get("fmt", "decimal2"),
        })
    if not rows:
        return pd.DataFrame(columns=["metric", "metric_label", "value"])
    return pd.DataFrame(rows)

def classify_level_from_text(text: str) -> str:
    s = text.strip().lower()
    if not s:
        return "Unknown"
    low_tokens = ["low", "urban", "accessible", "advantaged", "resilient", "good", "close", "small", "none"]
    high_tokens = ["high", "rural", "distressed", "underserved", "deprivation", "vulnerab", "far", "limited", "poor"]
    med_tokens = ["medium", "moderate", "suburban", "mixed", "average", "mid"]
    if any(tok in s for tok in high_tokens):
        return "High"
    if any(tok in s for tok in med_tokens):
        return "Medium"
    if any(tok in s for tok in low_tokens):
        return "Low"
    return "Medium"


def summarize_qualitative_family(fam_df: pd.DataFrame, family: str) -> tuple[str, list[tuple[str, str]]]:
    if fam_df.empty:
        return "Unknown", []
    row = fam_df.iloc[0].to_dict()
    config = NATIVE_VISUAL_CONFIG.get(family, {})
    qual_fields = config.get("qualitative_fields", [])
    pairs: list[tuple[str, str]] = []
    for qf in qual_fields:
        for k, v in row.items():
            base = str(k).split("__")[0]
            meta = lookup_definition(base)
            candidates = {base, str(k), meta.get("shown", "") if meta else "", meta.get("technical", "") if meta else ""}
            if qf in candidates:
                sv = textify(v)
                if sv:
                    pairs.append((str(k), sv))
    headline = pairs[0][1] if pairs else "Unknown"
    return headline, pairs[:4]

def render_index_cards(main_df: pd.DataFrame, selected: list[str]) -> None:
    st.subheader("Index overview")
    cols = st.columns(4)
    for i, family in enumerate(selected):
        fam, numeric = split_family_table(main_df, family)
        if fam.empty:
            populated = 0
            total = 0
            geo = ""
            miss = ""
        else:
            total = len(fam.columns)
            populated = sum(textify(fam.iloc[0][c]) != "" for c in fam.columns)
            geo_col = next((c for c in fam.columns if "Geography Used" in c), None)
            miss_col = next((c for c in fam.columns if "Missing Reason" in c), None)
            geo = textify(fam.iloc[0][geo_col]) if geo_col else ""
            miss = textify(fam.iloc[0][miss_col]) if miss_col else ""
        with cols[i % 4]:
            with st.container(border=True):
                st.markdown(f"**{family}**")
                st.metric("Populated fields", f"{populated}/{total}" if total else "0/0")
                if geo:
                    st.caption(f"Geography: {geo}")
                elif miss:
                    st.caption(miss)
                elif total == 0:
                    st.caption("No family columns were found in Main.")
                else:
                    st.caption("Computed, but no geography label returned.")


def render_gradient_legend() -> None:
    st.markdown(
        f"""
        <div style="margin: 0 0 1rem 0; width: 360px;">
            <div style="font-weight: 600; margin-bottom: 0.45rem;">Scale</div>
            <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; font-size:0.95rem; margin-bottom:0.3rem;">
                <div style="text-align:left;">Good</div>
                <div style="text-align:center;">Middle</div>
                <div style="text-align:right;">Bad</div>
            </div>
            <div style="width:100%; height:14px; border-radius:999px; background: linear-gradient(90deg, {GOOD_COLOR} 0%, {MID_COLOR} 50%, {BAD_COLOR} 100%); border: 1px solid #e5e7eb;"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
def render_metric_charts(main_df: pd.DataFrame, selected: list[str]) -> None:
    st.subheader("Visualizations")
    chart_items = []
    qualitative_items = []
    for family in selected:
        fam_df, _ = split_family_table(main_df, family)
        chart_df = pd.DataFrame()
        if family not in {"RUCA"}:
            chart_df = metric_table_to_chart_df(fam_df, family)
        if not chart_df.empty:
            chart_items.append((family, chart_df))
        level, details = summarize_qualitative_family(fam_df, family)
        if details and family not in {"AHRF", "EJI"}:
            qualitative_items.append((family, level, details))

    if chart_items:
        left_col, right_col = st.columns(2)
        column_heights = [0, 0]
        columns = [left_col, right_col]
        for family, chart_df in sorted(chart_items, key=lambda x: len(x[1]), reverse=True):
            target = 0 if column_heights[0] <= column_heights[1] else 1
            column_heights[target] += max(2, len(chart_df))
            with columns[target]:
                with st.container(border=True):
                    st.markdown(f"**{family}**")
                    rows_html = []
                    for _, row in chart_df.iterrows():
                        tick_left = format_scale_label(row["min"], row["fmt"])
                        tick_mid = format_scale_label(row["mid"], row["fmt"])
                        tick_right = format_scale_label(row["max"], row["fmt"])
                        gradient_css = f"linear-gradient(90deg, {GOOD_COLOR} 0%, {MID_COLOR} 50%, {BAD_COLOR} 100%)"
                        if bool(row["higher_is_good"]):
                            gradient_css = f"linear-gradient(90deg, {BAD_COLOR} 0%, {MID_COLOR} 50%, {GOOD_COLOR} 100%)"
                        pos_pct = max(0.0, min(100.0, float(row["position_frac"]) * 100.0))
                        rows_html.append(f"""
                        <div style='margin: 0 0 1rem 0;'>
                            <div style='display:flex; justify-content:space-between; gap:12px; align-items:flex-end; margin-bottom:0.2rem;'>
                                <div style='font-size:0.95rem; font-weight:500; line-height:1.2;'>{row['metric_display']}</div>
                                <div style='font-size:0.95rem; font-weight:700; white-space:nowrap;'>{row['value_text']}</div>
                            </div>
                            <div style='position:relative; height:14px; border-radius:999px; background:{gradient_css}; border:1px solid #e5e7eb;'>
                                <div style='position:absolute; left:calc({pos_pct}% - 2px); top:-3px; width:4px; height:18px; background:#111827; border-radius:2px;'></div>
                            </div>
                            <div style='display:flex; justify-content:space-between; font-size:0.8rem; color:#6b7280; margin-top:0.25rem;'>
                                <span>{tick_left}</span>
                                <span>{tick_mid}</span>
                                <span>{tick_right}</span>
                            </div>
                        </div>
                        """)
                    st.markdown("".join(rows_html), unsafe_allow_html=True)
                    explanations = graph_explanation_lines(family, chart_df)
                    if explanations:
                        with st.expander("More info"):
                            for line in explanations:
                                st.markdown(line)

    if qualitative_items:
        st.markdown("**Qualitative summaries**")
        cols = st.columns(3)
        for i, (family, level, details) in enumerate(qualitative_items):
            with cols[i % 3]:
                with st.container(border=True):
                    st.markdown(f"**{family}**")
                    for label, value in details:
                        st.caption(f"{display_label(label)}: {value}")

    if not chart_items and not qualitative_items:
        st.caption("No chartable numeric scores or qualitative summaries were available for the selected index families.")


def render_normalized_chart(main_df: pd.DataFrame, selected: list[str]) -> None:
    st.subheader("Normalization")
    all_rows = []
    for family in selected:
        if family == "RUCA":
            continue
        fam_df, _ = split_family_table(main_df, family)
        chart_df = metric_table_to_chart_df(fam_df, family)
        if chart_df.empty:
            continue
        temp = chart_df.copy()
        temp["family"] = family
        temp["normalized_100"] = (temp["position_frac"] * 100.0).round().astype(int)
        all_rows.append(temp)

    if not all_rows:
        st.caption("No chartable numeric scores were available for normalization.")
        return

    combined = pd.concat(all_rows, ignore_index=True)
    order_map = {name: i for i, name in enumerate(INDEX_OPTIONS)}
    combined["family_sort"] = combined["family"].map(order_map).fillna(999)
    combined = combined.sort_values(["family_sort", "metric_display"]).reset_index(drop=True)

    rows_html = []
    for _, row in combined.iterrows():
        pos_pct = max(0.0, min(100.0, float(row["normalized_100"])))
        gradient_css = f"linear-gradient(90deg, {GOOD_COLOR} 0%, {MID_COLOR} 50%, {BAD_COLOR} 100%)"
        if bool(row["higher_is_good"]):
            gradient_css = f"linear-gradient(90deg, {BAD_COLOR} 0%, {MID_COLOR} 50%, {GOOD_COLOR} 100%)"
        rows_html.append(f"""
        <div style='margin: 0 0 1rem 0;'>
            <div style='display:flex; justify-content:space-between; gap:12px; align-items:flex-end; margin-bottom:0.2rem;'>
                <div style='font-size:0.95rem; font-weight:500; line-height:1.2;'>{row['family']} — {row['metric_display']}</div>
                <div style='font-size:0.95rem; font-weight:700; white-space:nowrap;'>{int(row['normalized_100'])}</div>
            </div>
            <div style='position:relative; height:14px; border-radius:999px; background:{gradient_css}; border:1px solid #e5e7eb;'>
                <div style='position:absolute; left:calc({pos_pct}% - 2px); top:-3px; width:4px; height:18px; background:#111827; border-radius:2px;'></div>
            </div>
            <div style='display:flex; justify-content:space-between; font-size:0.8rem; color:#6b7280; margin-top:0.25rem;'>
                <span>0</span>
                <span>50</span>
                <span>100</span>
            </div>
        </div>
        """)

    with st.container(border=True):
        st.markdown("".join(rows_html), unsafe_allow_html=True)

def build_context_table(main_df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in GEOCODE_DISPLAY_CANDIDATES if c in main_df.columns]
    seen = set()
    ordered = []
    for c in cols:
        if c not in seen:
            ordered.append(c)
            seen.add(c)
    out = main_df[ordered].copy() if ordered else main_df.head(1).copy()
    if "Patient ID" in out.columns and "UniqueID" in out.columns:
        out = out.drop(columns=["Patient ID"])
    return rename_display_columns(out)


def run_single(unique_id: str, address: str):
    mtime_ns = SCRIPT_PATH.stat().st_mtime_ns
    input_df = pd.DataFrame([{"UniqueID": str(unique_id or "").strip(), "Address": str(address or "").strip()}])
    geocoded_json, geocode_notes = cached_geocode(serialize_df(input_df), mtime_ns)
    output_jsons, build_notes = cached_build_outputs(geocoded_json, mtime_ns)
    geocoded_df = deserialize_df(geocoded_json)
    outputs = [deserialize_df(x) for x in output_jsons]
    main_df = outputs[0] if outputs else pd.DataFrame()
    return geocoded_df, main_df, outputs, geocode_notes + build_notes


def run_batch(upload_path: Path):
    mtime_ns = SCRIPT_PATH.stat().st_mtime_ns
    mod = get_pipeline(mtime_ns)
    input_df = mod.load_input_addresses(upload_path)
    geocoded_json, geocode_notes = cached_geocode(serialize_df(input_df), mtime_ns)
    output_jsons, build_notes = cached_build_outputs(geocoded_json, mtime_ns)
    geocoded_df = deserialize_df(geocoded_json)
    outputs = [deserialize_df(x) for x in output_jsons]
    main_df = outputs[0] if outputs else pd.DataFrame()
    return input_df, geocoded_df, main_df, outputs, geocode_notes + build_notes


def filtered_main_download(main_df: pd.DataFrame, selected: list[str]) -> pd.DataFrame:
    keep = list(build_context_table(main_df).columns)
    reverse_map = {display_label(c): c for c in main_df.columns}
    raw_keep = [reverse_map.get(c, c) for c in keep]
    for fam in selected:
        cols = available_columns(main_df, fam)
        for c in cols:
            if c not in raw_keep:
                raw_keep.append(c)
    out = main_df[raw_keep].copy() if raw_keep else main_df.copy()
    if "Patient ID" in out.columns and "UniqueID" in out.columns:
        out = out.drop(columns=["Patient ID"])
    return rename_display_columns(out)


def render_family_sections(main_df: pd.DataFrame, selected: list[str]) -> None:
    for fam in selected:
        fam_df, _ = split_family_table(main_df, fam)
        with st.expander(f"{fam} details", expanded=False):
            st.dataframe(style_gradient_table(fam_df, fam), width="stretch", hide_index=True)


def main():
    st.set_page_config(page_title="SES Indices Dashboard", layout="wide")
    st.markdown(
        """
        <style>
        html, body, [class*="css"], [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stSidebar"], .stApp, .stMarkdown, .stDataFrame, .stTabs, .stMetric, .stTextInput input, .stSelectbox, .stMultiSelect, .stButton button, .stExpander, .stExpander summary, .stExpander details {
            font-family: Helvetica, Arial, sans-serif !important;
        }
        .material-symbols-rounded, .material-symbols-outlined, .material-icons, [class*="material-symbol"] {
            font-family: "Material Symbols Rounded", "Material Symbols Outlined", "Material Icons" !important;
            font-variation-settings: "FILL" 0, "wght" 400, "GRAD" 0, "opsz" 24;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    ensure_pipeline_exists()

    st.title("SES Indices Dashboard")
    st.caption("Runs the Drive-enabled SES index pipeline, loads large reference files from Google Drive, shows the main output clearly, and packages all supplementary outputs for download.")

    mtime_ns = SCRIPT_PATH.stat().st_mtime_ns
    _ = get_pipeline(mtime_ns)
    # Do NOT call get_project_paths() here. It can trigger a large Google Drive
    # download before the user even clicks Run. The reference files are located
    # lazily inside cached_build_outputs() only after an actual run starts.

    defaults = {
        "single_main_df_json": "",
        "single_geocoded_df_json": "",
        "single_notes": [],
        "single_outputs_jsons": [],
        "batch_main_df_json": "",
        "batch_notes": [],
        "batch_outputs_jsons": [],
        "single_main_xlsx_bytes": None,
        "single_all_zip_bytes": None,
        "batch_main_xlsx_bytes": None,
        "batch_all_zip_bytes": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    tab_single, tab_batch = st.tabs(["Individual", "Batch Upload"])

    with tab_single:
        st.markdown("Enter one individual, run the SES pipeline once, and then explore the selected index families in a cleaner layout.")
        with st.form("single_form"):
            c1, c2 = st.columns([1, 2])
            with c1:
                unique_id = st.text_input("Individual ID", key="single_pid")
            with c2:
                address = st.text_input("Address (street, city, state ZIP)", key="single_address")
            selected_indices = st.multiselect(
                "Index families to display",
                INDEX_OPTIONS,
                default=["ADI", "SVI", "VVI", "DCI", "NRI", "Walkability"],
                key="single_indices",
            )
            submit_single = st.form_submit_button("Run", width="stretch")

        if submit_single:
            if not address.strip():
                st.error("Please enter an address.")
            else:
                try:
                    with st.spinner("Running pipeline..."):
                        geocoded_df, main_df, outputs, notes = run_single(unique_id, address)
                    st.session_state.single_geocoded_df_json = serialize_df(geocoded_df)
                    st.session_state.single_main_df_json = serialize_df(main_df)
                    st.session_state.single_notes = notes
                    st.session_state.single_outputs_jsons = [serialize_df(df) for df in outputs]
                    st.session_state.single_main_xlsx_bytes = None
                    st.session_state.single_all_zip_bytes = None
                except Exception as exc:
                    logger.exception("Single run failed: %s", exc)
                    st.exception(exc)

        if st.session_state.single_main_df_json:
            main_df = deserialize_df(st.session_state.single_main_df_json)
            outputs = [deserialize_df(x) for x in st.session_state.single_outputs_jsons]
            notes = st.session_state.single_notes

            sub_a, sub_b, sub_norm, sub_c = st.tabs(["Overview", "Visualizations", "Normalization", "Tables & Downloads"])

            with sub_a:
                st.subheader("Geocode and Context")
                st.dataframe(style_gradient_table(build_context_table(main_df)), width="stretch", hide_index=True)
                render_index_cards(main_df, selected_indices)
    
            with sub_b:
                render_metric_charts(main_df, selected_indices)

            with sub_norm:
                render_normalized_chart(main_df, selected_indices)

            with sub_c:
                render_family_sections(main_df, selected_indices)
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Prepare displayed Main results", key="prepare_single_main"):
                        try:
                            st.session_state.single_main_xlsx_bytes = excel_bytes_from_df(
                                filtered_main_download(main_df, selected_indices),
                                sheet_name="Main_Selected",
                            )
                        except Exception as exc:
                            logger.exception("Single Main Excel export failed: %s", exc)
                            st.error(f"Could not prepare the Excel file: {exc}")
                    if st.session_state.single_main_xlsx_bytes:
                        st.download_button(
                            "Download displayed Main results",
                            data=st.session_state.single_main_xlsx_bytes,
                            file_name="ses_indices_single_main_selected.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_single_main",
                        )
                with col2:
                    st.caption("All-output ZIPs can be large, so this is only generated after you click Prepare.")
                    if st.button("Prepare all pipeline outputs", key="prepare_single_all"):
                        try:
                            st.session_state.single_all_zip_bytes = zip_bytes_from_outputs(outputs)
                        except Exception as exc:
                            logger.exception("Single all-output ZIP export failed: %s", exc)
                            st.error(f"Could not prepare the ZIP file: {exc}")
                    if st.session_state.single_all_zip_bytes:
                        st.download_button(
                            "Download all pipeline outputs",
                            data=st.session_state.single_all_zip_bytes,
                            file_name="ses_indices_single_all_outputs.zip",
                            mime="application/zip",
                            key="download_single_all",
                        )

    with tab_batch:
        st.markdown("Upload a CSV or Excel file with a record ID column and an address column. The dashboard runs the same SES pipeline and returns the main results plus all supplementary outputs.")
        st.download_button(
            "Download input template",
            data=template_excel_bytes(),
            file_name="ses_indices_input_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Expected input format", expanded=True):
            st.dataframe(
                pd.DataFrame({
                    "Individual ID": ["P001", "P002"],
                    "Address": [
                        "123 Main St, Springfield, VA 22150",
                        "456 Oak Ave, Brambleton, VA 20148",
                    ],
                }),
                width="stretch",
                hide_index=True,
            )

        with st.form("batch_form"):
            uploaded = st.file_uploader("Upload file", type=["csv", "xlsx", "xls"])
            batch_indices = st.multiselect(
                "Index families to emphasize in the preview",
                INDEX_OPTIONS,
                default=["ADI", "SVI", "VVI", "DCI", "NRI", "Walkability"],
                key="batch_indices",
            )
            submit_batch = st.form_submit_button("Run batch", width="stretch")

        if submit_batch:
            if uploaded is None:
                st.error("Please upload a file.")
            else:
                suffix = Path(uploaded.name).suffix.lower() or ".xlsx"
                fd, tmp_name = tempfile.mkstemp(suffix=suffix)
                tmp_path = Path(tmp_name)
                try:
                    os.write(fd, uploaded.getvalue())
                finally:
                    os.close(fd)
                try:
                    with st.spinner("Running batch pipeline..."):
                        _, _, main_df, outputs, notes = run_batch(tmp_path)
                    st.session_state.batch_main_df_json = serialize_df(main_df)
                    st.session_state.batch_notes = notes
                    st.session_state.batch_outputs_jsons = [serialize_df(df) for df in outputs]
                    st.session_state.batch_main_xlsx_bytes = None
                    st.session_state.batch_all_zip_bytes = None
                except Exception as exc:
                    logger.exception("Batch run failed: %s", exc)
                    st.exception(exc)
                finally:
                    try:
                        os.unlink(tmp_path)
                    except OSError:
                        pass

        if st.session_state.batch_main_df_json:
            main_df = deserialize_df(st.session_state.batch_main_df_json)
            outputs = [deserialize_df(x) for x in st.session_state.batch_outputs_jsons]
            notes = st.session_state.batch_notes
            preview_df = filtered_main_download(main_df, batch_indices)

            st.success(f"Completed: {len(main_df)} row(s)")
            st.subheader("Results preview")
            st.dataframe(style_gradient_table(preview_df), width="stretch", height=500, hide_index=True)

            c1, c2 = st.columns(2)
            with c1:
                if st.button("Prepare previewed Main results", key="prepare_batch_main"):
                    try:
                        st.session_state.batch_main_xlsx_bytes = excel_bytes_from_df(preview_df, sheet_name="Main_Selected")
                    except Exception as exc:
                        logger.exception("Batch Main Excel export failed: %s", exc)
                        st.error(f"Could not prepare the Excel file: {exc}")
                if st.session_state.batch_main_xlsx_bytes:
                    st.download_button(
                        "Download previewed Main results",
                        data=st.session_state.batch_main_xlsx_bytes,
                        file_name="ses_indices_batch_main_selected.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_batch_main",
                    )
            with c2:
                st.caption("All-output ZIPs can be large, so this is only generated after you click Prepare.")
                if st.button("Prepare all pipeline outputs", key="prepare_batch_all"):
                    try:
                        st.session_state.batch_all_zip_bytes = zip_bytes_from_outputs(outputs)
                    except Exception as exc:
                        logger.exception("Batch all-output ZIP export failed: %s", exc)
                        st.error(f"Could not prepare the ZIP file: {exc}")
                if st.session_state.batch_all_zip_bytes:
                    st.download_button(
                        "Download all pipeline outputs",
                        data=st.session_state.batch_all_zip_bytes,
                        file_name="ses_indices_batch_all_outputs.zip",
                        mime="application/zip",
                        key="download_batch_all",
                    )


if __name__ == "__main__":
    main()