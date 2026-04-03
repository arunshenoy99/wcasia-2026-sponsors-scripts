"""
Module for reading and processing sponsor data from Excel/CSV files
"""
import pandas as pd
import os
from typing import List, Dict, Optional
from config import (
    SPONSOR_TYPE_TO_TEMPLATE,
    EXCEL_COLUMNS,
    FILTER_BY_OUTREACH,
    OUTREACH_FILTER_VALUE,
    TEMPLATE_NAME_COLUMN,
)


class SponsorReader:
    """Reads and processes sponsor data from spreadsheets"""
    
    def __init__(self, file_path: str):
        """
        Initialize the SponsorReader
        
        Args:
            file_path: Path to the Excel or CSV file
        """
        self.file_path = file_path
        self.df = None
        self.sponsor_type_column = None
        
    def read_file(self) -> pd.DataFrame:
        """
        Read the Excel or CSV file
        
        Returns:
            DataFrame with sponsor data
        """
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        
        # Try to read as Excel first, then CSV
        try:
            if self.file_path.endswith('.xlsx') or self.file_path.endswith('.xls'):
                self.df = pd.read_excel(self.file_path)
            else:
                self.df = pd.read_csv(self.file_path)
        except Exception as e:
            raise ValueError(f"Error reading file: {e}")
        
        return self.df
    
    def identify_sponsor_type_column(self) -> Optional[str]:
        """
        Identify the column containing sponsor type labels
        
        Returns:
            Column name if found, None otherwise
        """
        if self.df is None:
            self.read_file()
        
        # Look for columns that might contain sponsor types
        sponsor_type_values = set(SPONSOR_TYPE_TO_TEMPLATE.keys())
        
        for col in self.df.columns:
            # Check if this column contains sponsor type values
            unique_values = set(self.df[col].dropna().astype(str).unique())
            if sponsor_type_values.intersection(unique_values):
                self.sponsor_type_column = col
                return col
        
        # If not found, check for common column names (prioritize "Sales Strategy")
        possible_names = ['Sales Strategy', 'Sponsor Type', 'Type', 'Category', 'Sponsor Category', 'Status']
        for name in possible_names:
            if name in self.df.columns:
                # Check if it contains any sponsor type values
                unique_values = set(self.df[name].dropna().astype(str).unique())
                if sponsor_type_values.intersection(unique_values):
                    self.sponsor_type_column = name
                    return name
        
        # If "Sales Strategy" column exists, use it even if values don't match exactly
        # (in case of slight variations in naming)
        if 'Sales Strategy' in self.df.columns:
            self.sponsor_type_column = 'Sales Strategy'
            return 'Sales Strategy'
        
        return None
    
    def map_sponsor_type_to_template(self, sponsor_type: str) -> Optional[str]:
        """
        Map sponsor type to FreeScout template name
        
        Args:
            sponsor_type: The sponsor type from Excel
            
        Returns:
            FreeScout template name or None if not found
        """
        # Try exact match first
        if sponsor_type in SPONSOR_TYPE_TO_TEMPLATE:
            return SPONSOR_TYPE_TO_TEMPLATE[sponsor_type]
        
        # Try case-insensitive match
        for key, template in SPONSOR_TYPE_TO_TEMPLATE.items():
            if key.lower() == sponsor_type.lower():
                return template
        
        return None
    
    def get_sponsors(self) -> List[Dict]:
        """
        Get list of sponsors with mapped template information
        
        Returns:
            List of sponsor dictionaries with all required fields
        """
        if self.df is None:
            self.read_file()

        # Round CSV mode: file has "Template Name" column (output of extract_round_leads.py)
        if TEMPLATE_NAME_COLUMN in self.df.columns:
            return self._get_sponsors_from_round_csv()

        # Standard mode: identify sponsor type column and map to template
        if self.sponsor_type_column is None:
            if EXCEL_COLUMNS.get("SPONSOR_TYPE"):
                if EXCEL_COLUMNS["SPONSOR_TYPE"] in self.df.columns:
                    self.sponsor_type_column = EXCEL_COLUMNS["SPONSOR_TYPE"]
                else:
                    self.identify_sponsor_type_column()
            else:
                self.identify_sponsor_type_column()

        if self.sponsor_type_column is None:
            raise ValueError("Could not identify sponsor type column. Please ensure the Excel file contains a 'Sales Strategy' column with one of the expected sponsor types.")

        sponsors = []
        for _, row in self.df.iterrows():
            if FILTER_BY_OUTREACH and EXCEL_COLUMNS.get("INITIAL_OUTREACH_BY"):
                outreach_col = EXCEL_COLUMNS["INITIAL_OUTREACH_BY"]
                if outreach_col in self.df.columns:
                    outreach_value = str(row.get(outreach_col, "")).strip() if not pd.isna(row.get(outreach_col)) else ""
                    if outreach_value != OUTREACH_FILTER_VALUE:
                        continue
            if pd.isna(row.get(EXCEL_COLUMNS["EMAIL"])) or not row.get(EXCEL_COLUMNS["EMAIL"]):
                continue
            sponsor_type = str(row[self.sponsor_type_column]) if not pd.isna(row.get(self.sponsor_type_column)) else None
            template_name = self.map_sponsor_type_to_template(sponsor_type) if sponsor_type else None
            if not template_name:
                print(f"Warning: No template found for sponsor type '{sponsor_type}' for {row.get(EXCEL_COLUMNS['EMAIL'])}")
                continue
            sponsor = {
                "email": str(row[EXCEL_COLUMNS["EMAIL"]]).strip(),
                "company_name": str(row[EXCEL_COLUMNS["COMPANY_NAME"]]).strip() if not pd.isna(row.get(EXCEL_COLUMNS["COMPANY_NAME"])) else "",
                "contact_person": str(row[EXCEL_COLUMNS["CONTACT_PERSON"]]).strip() if not pd.isna(row.get(EXCEL_COLUMNS["CONTACT_PERSON"])) else "",
                "sponsor_type": sponsor_type,
                "template_name": template_name,
                "row_data": row.to_dict(),
            }
            sponsors.append(sponsor)
        return sponsors

    def _get_sponsors_from_round_csv(self) -> List[Dict]:
        """Get sponsors when file has Template Name column (round CSV from extract_round_leads.py)."""
        sponsors = []
        email_col = "Email"
        company_col = "Company Name"
        contact_col = "Contact Person"
        template_col = TEMPLATE_NAME_COLUMN
        for col in [email_col, company_col, contact_col, template_col]:
            if col not in self.df.columns:
                raise ValueError(f"Round CSV missing column '{col}'. Required: Email, Company Name, Contact Person, Template Name.")
        outreach_col = EXCEL_COLUMNS.get("INITIAL_OUTREACH_BY") if FILTER_BY_OUTREACH else None
        for _, row in self.df.iterrows():
            if outreach_col and outreach_col in self.df.columns:
                outreach_value = str(row.get(outreach_col, "")).strip() if not pd.isna(row.get(outreach_col)) else ""
                if outreach_value != OUTREACH_FILTER_VALUE:
                    continue
            email_val = row.get(email_col)
            if pd.isna(email_val) or not str(email_val).strip():
                continue
            template_name = str(row[template_col]).strip() if not pd.isna(row.get(template_col)) else ""
            if not template_name:
                continue
            sponsor = {
                "email": str(row[email_col]).strip(),
                "company_name": str(row[company_col]).strip() if not pd.isna(row.get(company_col)) else "",
                "contact_person": str(row[contact_col]).strip() if not pd.isna(row.get(contact_col)) else "",
                "sponsor_type": row.get("Status", ""),
                "template_name": template_name,
                "row_data": row.to_dict(),
            }
            sponsors.append(sponsor)
        return sponsors
    
    def get_sponsor_count(self) -> int:
        """Get total number of sponsors that can be processed"""
        sponsors = self.get_sponsors()
        return len(sponsors)

