r"""
================================================================================
 DCT ALL-TEMPLATE VALIDATOR  v3.0
================================================================================
 Validates EVERY DCT upload template - loans, deposits, membership, land,
 collateral, transactions, agents - not just loan files.

 WHAT CHANGED FROM v2.0
 ----------------------
 1. 14 of your 16 templates used to be skipped. Now every DCT template type is
    recognised and validated: Membership, LandDetails, KCC, LT/MT, OtherLoan,
    SB & Others, FD, FD Non-Cumulative, RD, Pigmy, GoldCollaterals,
    LoanTransactions, Agent, ShareTransactions.
 2. Three severity levels instead of one. A whole file is no longer destroyed
    by one column-wide data convention:
        HARD REJECT  the portal will refuse this row  -> kept out of Clean Data
        REVIEW       a calculation or consistency mismatch to check -> listed,
                     but the row still goes to Clean Data
        INFO         file-level observation (e.g. a whole column is negative)
 3. Column-wide sign convention is detected once at file level, not reported
    4,829 times as per-row errors.
 4. Clean Data is written in the EXACT DCT template column order, with the
    template's own header names, dates as DD-MM-YYYY, amounts as numbers and
    identifiers as text. No SourceRow, no RuleID, no error text in it - it is
    ready to upload as-is.
 5. FirstInstallmentDate and InstallmentAmount are now derived and checked.
 6. Full deposit validation: maturity date, maturity amount, term, rate,
    installment consistency, RD and Pigmy specific checks.

--------------------------------------------------------------------------------
 REQUIREMENTS
--------------------------------------------------------------------------------
    pip install pandas openpyxl xlrd

 Windows, Python 3.13 / 3.14.

--------------------------------------------------------------------------------
 RUN
--------------------------------------------------------------------------------
    GUI  :  python dct_all_template_validator.py
    CLI  :  python dct_all_template_validator.py "C:\PACS_Data\PANCHAGANGA VSS"

 Output goes to  <selected folder>\Validated_DD-MM-YYYY_HH-MM-SS\
 Source files are opened read-only and are never modified.
================================================================================
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is required.  Run:  pip install pandas openpyxl xlrd")

try:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  Run:  pip install pandas openpyxl xlrd")

APP_NAME = "DCT All-Template Validator"
APP_VERSION = "3.1"
DATA_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".txt"}

HARD, REVIEW, INFO = "HARD REJECT", "REVIEW", "INFO"

# These thresholds can be changed in dct_validator_settings.json. Keep the
# defaults aligned with the documented validator rules.
DEFAULT_RULE_SETTINGS = {
    "negative_sign_convention_share": 0.80,
    "date_tolerance_days": 2,
    "installment_amount_tolerance_pct": 0.05,
    "deposit_amount_tolerance_pct": 0.02,
    "rd_total_paid_tolerance_pct": 0.02,
}


def load_rule_settings() -> dict:
    settings = dict(DEFAULT_RULE_SETTINGS)
    base = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    path = base / "dct_validator_settings.json"
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        for key, default in DEFAULT_RULE_SETTINGS.items():
            value = float(raw.get(key, default))
            if key.endswith("_pct") or key.endswith("_share"):
                if not 0 <= value <= 1:
                    continue
            elif value < 0 or value > 366:
                continue
            settings[key] = value
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        pass
    return settings


RULE_SETTINGS = load_rule_settings()


# ==============================================================================
# 1. TEMPLATE REGISTRY
#    Column orders copied from the DCT suite's MASTER_TEMPLATES so Clean Data
#    comes out in the exact shape the portal expects.
# ==============================================================================

TEMPLATES: dict[str, list[str]] = {
    "KCC": ["SlNo", "ProductDescription", "AdmissionNo", "LoanNo", "Scheme",
            "CropDescription", "SanctionedDate", "SanctionedAmount", "LoanPeriod",
            "DueDate", "ROIPercentage", "PenalROIPercentage", "IODPercentage",
            "OutstandingPrincipal", "OutstandingInterest",
            "OutstandingPenalInterest", "PacsId", "BranchId"],
    "MTLT": ["SlNo", "Name", "AdmissionNo", "ProductDescription", "LoanNo",
             "Scheme", "PurposeDescription", "SanctionedDate", "SanctionedAmount",
             "RepaymentFrequency", "LoanPeriod", "GestationPeriod",
             "FirstInstallmentDate", "InstallmentAmount", "DueDate",
             "ROIPercentage", "PenalROIPercentage", "IODPercentage", "LANo",
             "LCNo", "LCDate", "DCCBLoanAccountNo", "OutstandingPrincipal",
             "OutstandingInterest", "OutstandingPenalInterest", "PacsId",
             "BranchId"],
    "OTHERS": ["SlNo", "ProductDescription", "AdmissionNo", "LoanNo",
               "PurposeDescription", "SanctionedDate", "SanctionedAmount",
               "LoanPeriod", "DueDate", "Scheme", "RepaymentFrequency",
               "FirstInstallmentDate", "InstallmentAmount", "ROI", "PenalROI",
               "IOD", "LedgerFolioNo", "OutstandingPrincipal",
               "OutstandingInterest", "OutstandingPenalInterest", "PacsId",
               "BranchId"],
    "SB": ["NAME", "AdmissionNo", "ProductDescription", "AccountNo",
           "DepositDate", "IntroducerNo", "ChequeOption", "OperationTypeDesc",
           "IsOrganisation", "RegisterSlNo", "Balance", "InterestBalance",
           "LedgerFolioNo", "JointAdmissionNo", "PacsIDPkey", "BranchId"],
    "FD": ["ProductDescription", "AdmissionNo", "AccountNo", "OperationTypeDesc",
           "DepositDate", "DepositAmount", "TerminDays", "TerminMonths",
           "MaturityDate", "MaturityAmount", "RateOfInterest", "Status",
           "PacsIDPkey", "LedgerFolioNo", "DepositTypeDesc", "InterestAmount",
           "InstallmentAmount", "InterestPaymentModeDesc", "InstallmentsPaid",
           "TotalInstallmentAmountPaid", "LastPaidInstallDate", "IsInterestPosted",
           "chkIsInterestPostingToCB", "LastInterestPostingDate",
           "TotalInterestAmount", "JointAdmissionNo", "CERTINO", "BranchId",
           "PacsId"],
    "RD": ["NAME", "ProductDescription", "AdmissionNo", "AccountNo",
           "OperationTypeDesc", "DepositDate", "DepositAmount", "TerminDays",
           "TerminMonths", "MaturityDate", "MaturityAmount", "RateOfInterest",
           "Status", "PacsIDPkey", "LedgerFolioNo", "DepositTypeDesc",
           "InterestAmount", "InstallmentAmount", "InterestPaymentModeDesc",
           "InstallmentsPaid", "TotalInstallmentAmountPaid", "LastPaidInstallDate",
           "IsInterestPosted", "chkIsInterestPostingToCB",
           "LastInterestPostingDate", "TotalInterestAmount", "JointAdmissionNo",
           "BranchId", "PacsId"],
    "PIGMY": ["NAME", "AdmissionNo", "ProductDescription", "AccountNo",
              "DepositDate", "OperationTypeDesc", "InstallmentAmount",
              "PeriodinMonths", "PeriodinYears", "MaturityDate", "TotalAmount",
              "RateOfInterest", "AgentNo", "FrequencyDescription",
              "LedgerFolioNo", "IsPayable", "totalInterestAmount", "PacsIDPKey",
              "BranchId"],
    "TRANSACTIONS": ["AdmissionNo", "ProductDescription", "LoanNo",
                     "TransactionDate", "LedgerFolioNo", "DisbursementAmount",
                     "OtherChargesDebitAmount", "CollectedPrincipal",
                     "CollectedInterest", "CollectedPenalInterest",
                     "CollectedIOD", "CollectedOthers", "BalancePrincipal",
                     "BalanceInterest", "BalancePenalInterest", "BalanceIOD",
                     "Charges", "PacsId", "BranchId", "PurposeDescription"],
    # Templates below have no fixed order in the suite: the source order is kept.
    "MEMBERSHIP": [],
    "LAND": [],
    "COLLATERAL": [],
    "AGENT": [],
    "SHARETRANS": [],
}

CATEGORY_LABELS = {
    "KCC": "Short Term KCC Loans",
    "MTLT": "Medium / Long Term Loans (LT & MT)",
    "OTHERS": "Other Loans (Gold / Deposit Cert / Produce / Housing)",
    "SB": "SB, PD, Staff & Customer Security Deposits",
    "FD": "FD / Cash Certificate Deposits",
    "FDNONCUM": "FD Non-Cumulative Deposits",
    "RD": "Recurring Deposits",
    "PIGMY": "Pigmy Deposits",
    "MEMBERSHIP": "Membership / Personal Details",
    "LAND": "Land Details",
    "COLLATERAL": "Gold / Collateral Details",
    "TRANSACTIONS": "Loan Transaction Details",
    "AGENT": "Pigmy Agent Master",
    "SHARETRANS": "Share Transactions",
    "REFERENCE": "Reference file (not uploaded - skipped)",
    "UNKNOWN": "Unrecognised file",
}

LOAN_CATEGORIES = {"KCC", "MTLT", "OTHERS"}
DEPOSIT_CATEGORIES = {"FD", "FDNONCUM", "RD", "PIGMY", "SB"}
SKIP_CATEGORIES = {"REFERENCE"}

# FDNONCUM shares the FD template shape.
TEMPLATES["FDNONCUM"] = TEMPLATES["FD"]


# ---- filename detection, extended from the DCT suite's FILENAME_PATTERNS -----
FILENAME_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"balance[\s_-]*sheet|balsheet", re.I), "REFERENCE"),
    (re.compile(r"\bvillages?(?:[-_]?list)?\b", re.I), "REFERENCE"),
    (re.compile(r"pacs.*id.*detail", re.I), "REFERENCE"),
    (re.compile(r"invalid[_ ]?date", re.I), "REFERENCE"),
    (re.compile(r"membership|personal[\s_-]*detail|member[\s_-]*detail", re.I), "MEMBERSHIP"),
    (re.compile(r"land[\s_-]*detail", re.I), "LAND"),
    (re.compile(r"kcc", re.I), "KCC"),
    (re.compile(r"lt\s*and\s*mt|mt\s*and\s*lt|ltandmt|loanmt|loan\s*mt|mtlt", re.I), "MTLT"),
    (re.compile(r"other[\s_-]*loan|loanoth", re.I), "OTHERS"),
    (re.compile(r"collateral|gold", re.I), "COLLATERAL"),
    (re.compile(r"sb[\s_-]*and[\s_-]*other|sbandother|sb[\s_-]*other", re.I), "SB"),
    (re.compile(r"fd[\s_-]*noncum|non[\s_-]*cumulative", re.I), "FDNONCUM"),
    (re.compile(r"fd[\s_-]*template|akshaya|cash[\s_-]*certificate", re.I), "FD"),
    (re.compile(r"(^|[^a-z])rd(template)?([^a-z]|$)", re.I), "RD"),
    (re.compile(r"pigmy", re.I), "PIGMY"),
    (re.compile(r"transaction[\s_-]*detail|loantrans", re.I), "TRANSACTIONS"),
    (re.compile(r"^\d*[._ ]*agent", re.I), "AGENT"),
    (re.compile(r"sharetrans|share\s*trans", re.I), "SHARETRANS"),
]

# ---- header signatures: (category, must-have canonical columns, score) -------
HEADER_SIGNATURES: list[tuple[str, set[str]]] = [
    ("TRANSACTIONS", {"LoanNo", "TransactionDate", "CollectedPrincipal"}),
    ("TRANSACTIONS", {"LoanNo", "TransactionDate", "DisbursementAmount"}),
    ("KCC", {"CropDescription", "SanctionedAmount", "LoanNo"}),
    ("MTLT", {"GestationPeriod", "SanctionedAmount", "LoanNo"}),
    ("OTHERS", {"SanctionedAmount", "LoanNo", "PurposeDescription"}),
    ("PIGMY", {"AgentNo", "InstallmentAmount", "AccountNo"}),
    ("RD", {"MaturityAmount", "InstallmentAmount", "AccountNo", "NAME"}),
    ("FD", {"MaturityAmount", "DepositAmount", "AccountNo", "CERTINO"}),
    ("FD", {"MaturityAmount", "DepositAmount", "AccountNo"}),
    ("SB", {"AccountNo", "Balance", "RegisterSlNo"}),
    ("SB", {"AccountNo", "Balance", "OperationTypeDesc"}),
    ("COLLATERAL", {"LoanNo", "GrossWeight"}),
    ("COLLATERAL", {"LoanNo", "CollateralDescription"}),
    ("LAND", {"AdmissionNo", "SurveyNo"}),
    ("AGENT", {"AgentNo", "AgentName"}),
    ("MEMBERSHIP", {"AdmissionNo", "MemberName", "AdmissionDate"}),
    ("MEMBERSHIP", {"AdmissionNo", "MemberSurName", "DOB"}),
]


# ---- value domains taken from the DCT suite ---------------------------------
VALID_CASTE_VALUES = {
    "Not Available", "Schedule Caste", "Schedule Tribe", "Backward Caste A",
    "Backward Caste B", "Backward Caste C", "Backward Caste D",
    "Other Backward Caste", "General", "Minorities",
}
VALID_GENDER_VALUES = {"Male", "Female", "Others"}
VALID_LAND_TYPE = {"Wet", "Dry", "Not Available"}
VALID_IRRIGATION = {"Tank", "Well", "Canal", "Not Available", "Others"}
VALID_OWNERSHIP = {"Owner", "Tenant Farmers", "ShareCroppers", "Oral Lessees"}

# Portal spelling is exact: "Halfyearly", one word, lowercase y.
FREQUENCY_MONTHS = {"monthly": 1, "quarterly": 3, "halfyearly": 6, "yearly": 12}
PORTAL_FREQUENCY_SPELLING = {1: "Monthly", 3: "Quarterly",
                             6: "Halfyearly", 12: "Yearly"}

ROI_MIN, ROI_MAX = 0.0, 50.0
PENAL_MIN, PENAL_MAX = 0.0, 50.0
DEPOSIT_RATE_MAX = 15.0          # above this is a REVIEW, not a rejection


# ---- per-category mandatory fields ------------------------------------------
REQUIRED_FIELDS: dict[str, list[str]] = {
    "KCC": ["AdmissionNo", "LoanNo", "ProductDescription", "SanctionedDate",
            "SanctionedAmount", "LoanPeriod"],
    "MTLT": ["AdmissionNo", "LoanNo", "ProductDescription", "SanctionedDate",
             "SanctionedAmount", "LoanPeriod", "RepaymentFrequency",
             "GestationPeriod"],
    "OTHERS": ["AdmissionNo", "LoanNo", "ProductDescription", "SanctionedDate",
               "SanctionedAmount", "LoanPeriod", "RepaymentFrequency"],
    "FD": ["AdmissionNo", "AccountNo", "ProductDescription", "DepositDate",
           "DepositAmount", "MaturityAmount"],
    "FDNONCUM": ["AdmissionNo", "AccountNo", "ProductDescription", "DepositDate",
                 "DepositAmount"],
    "RD": ["AdmissionNo", "AccountNo", "ProductDescription", "DepositDate",
           "InstallmentAmount"],
    "PIGMY": ["AdmissionNo", "AccountNo", "ProductDescription", "DepositDate",
              "InstallmentAmount"],
    "SB": ["AdmissionNo", "AccountNo", "ProductDescription"],
    "MEMBERSHIP": ["AdmissionNo"],
    "LAND": ["AdmissionNo"],
    "COLLATERAL": ["LoanNo"],
    "TRANSACTIONS": ["LoanNo", "TransactionDate"],
    "AGENT": ["AgentNo"],
    "SHARETRANS": ["AdmissionNo"],
}

# ---- field type handling for the Clean Data sheet ---------------------------
DATE_FIELDS = {
    "SanctionedDate", "DueDate", "FirstInstallmentDate", "LCDate", "DepositDate",
    "MaturityDate", "LastPaidInstallDate", "LastInterestPostingDate",
    "TransactionDate", "AdmissionDate", "DOB", "GPAIssueDate",
    "GoldStockRegDate", "AppraisalDate",
}
AMOUNT_FIELDS = {
    "SanctionedAmount", "InstallmentAmount", "OutstandingPrincipal",
    "OutstandingInterest", "OutstandingPenalInterest", "DepositAmount",
    "MaturityAmount", "InterestAmount", "TotalInterestAmount",
    "totalInterestAmount", "TotalInstallmentAmountPaid", "TotalAmount",
    "Balance", "InterestBalance", "DisbursementAmount",
    "OtherChargesDebitAmount", "CollectedPrincipal", "CollectedInterest",
    "CollectedPenalInterest", "CollectedIOD", "CollectedOthers",
    "BalancePrincipal", "BalanceInterest", "BalancePenalInterest", "BalanceIOD",
    "Charges", "ShareBalance", "Thrift",
}
RATE_FIELDS = {"ROI", "PenalROI", "IOD", "ROIPercentage", "PenalROIPercentage",
               "IODPercentage", "RateOfInterest"}
INT_FIELDS = {"SlNo", "LoanPeriod", "GestationPeriod", "TerminDays",
              "TerminMonths", "InstallmentsPaid", "PeriodinMonths",
              "PeriodinYears", "PacsId", "BranchId", "PacsIDPkey", "PacsIDPKey",
              "Age"}
TEXT_ID_FIELDS = {"AdmissionNo", "LoanNo", "AccountNo", "LedgerFolioNo",
                  "JointAdmissionNo", "CERTINO", "LANo", "LCNo",
                  "DCCBLoanAccountNo", "AgentNo", "RegisterSlNo",
                  "IntroducerNo", "SurveyNo"}

# Columns whose sign convention is worth detecting at file level.
SIGN_WATCH_FIELDS = ["OutstandingPrincipal", "OutstandingInterest",
                     "OutstandingPenalInterest", "Balance", "DepositAmount",
                     "BalancePrincipal"]


# ---- header aliases ---------------------------------------------------------
COLUMN_ALIASES: dict[str, list[str]] = {
    "SlNo": ["slno", "srno", "serialno"],
    "Name": ["name", "membername", "customername"],
    "NAME": [],
    "AdmissionNo": ["admissionno", "admno", "memberno", "membernumber"],
    "LoanNo": ["loanno", "loannumber", "loanacno", "loanaccountno"],
    "AccountNo": ["accountno", "acno", "acctno", "depositaccountno"],
    "LedgerFolioNo": ["ledgerfoliono", "folio", "folionumber"],
    "ProductDescription": ["productdescription", "product", "productdesc"],
    "PurposeDescription": ["purposedescription", "purpose", "loanpurposename"],
    "CropDescription": ["cropdescription", "crop"],
    "Scheme": ["scheme", "schemename"],
    "SanctionedDate": ["sanctioneddate", "sanctiondate", "loandate"],
    "DueDate": ["duedate", "maturitydateloan"],
    "FirstInstallmentDate": ["firstinstallmentdate", "firstinstalmentdate", "firstemidate"],
    "SanctionedAmount": ["sanctionedamount", "sanctionamount", "loanamount"],
    "InstallmentAmount": ["installmentamount", "instalmentamount", "emiamount"],
    "OutstandingPrincipal": ["outstandingprincipal", "principaloutstanding", "oustandingprincipal"],
    "OutstandingInterest": ["outstandinginterest", "interestoutstanding"],
    "OutstandingPenalInterest": ["outstandingpenalinterest", "penalinterest"],
    "LoanPeriod": ["loanperiod", "period", "tenure", "tenuremonths"],
    "RepaymentFrequency": ["repaymentfrequency", "frequency", "repayfreq"],
    "GestationPeriod": ["gestationperiod", "gestation", "gestationmonths"],
    "ROIPercentage": ["roipercentage"],
    "PenalROIPercentage": ["penalroipercentage"],
    "IODPercentage": ["iodpercentage"],
    "ROI": ["roi", "rateofinterestloan"],
    "PenalROI": ["penalroi", "penalrate"],
    "IOD": ["iod"],
    "LANo": ["lano"], "LCNo": ["lcno"], "LCDate": ["lcdate"],
    "DCCBLoanAccountNo": ["dccbloanaccountno", "dccbloanacno"],
    "PacsId": ["pacsid", "pacscode"],
    "PacsIDPkey": ["pacsidpkey"],
    "BranchId": ["branchid", "branchcode"],
    # deposits
    "DepositDate": ["depositdate", "openingdate", "startdate"],
    "DepositAmount": ["depositamount", "principalamount"],
    "MaturityDate": ["maturitydate"],
    "MaturityAmount": ["maturityamount"],
    "TerminDays": ["termindays", "termdays", "perioddays"],
    "TerminMonths": ["terminmonths", "termmonths", "periodmonths"],
    "RateOfInterest": ["rateofinterest", "interestrate", "intrate"],
    "DepositTypeDesc": ["deposittypedesc", "deposittype"],
    "OperationTypeDesc": ["operationtypedesc", "operationtype"],
    "InterestAmount": ["interestamount"],
    "TotalInterestAmount": ["totalinterestamount"],
    "totalInterestAmount": [],
    "InstallmentsPaid": ["installmentspaid", "installme9ntspaid", "noofinstallmentspaid"],
    "TotalInstallmentAmountPaid": ["totalinstallmentamountpaid", "totinstallmentspaid"],
    "LastPaidInstallDate": ["lastpaidinstalldate", "lastpaidinstldate"],
    "InterestPaymentModeDesc": ["interestpaymentmodedesc"],
    "IsInterestPosted": ["isinterestposted"],
    "chkIsInterestPostingToCB": ["chkisinterestpostingtocb"],
    "LastInterestPostingDate": ["lastinterestpostingdate"],
    "JointAdmissionNo": ["jointadmissionno"],
    "CERTINO": ["certino", "certificateno"],
    "Status": ["status"],
    "Balance": ["balance", "balanceamount", "accountbalance", "currentbalance",
                "closingbalance", "availablebalance", "sbbalance", "ledgerbalance"],
    "InterestBalance": ["interestbalance"],
    "IntroducerNo": ["introducerno"],
    "ChequeOption": ["chequeoption"],
    "IsOrganisation": ["isorganisation", "isorganization"],
    "RegisterSlNo": ["registerslno"],
    "PeriodinMonths": ["periodinmonths"],
    "PeriodinYears": ["periodinyears"],
    "TotalAmount": ["totalamount"],
    "AgentNo": ["agentno", "agentcode"],
    "AgentName": ["agentname"],
    "FrequencyDescription": ["frequencydescription"],
    "IsPayable": ["ispayable"],
    # transactions
    "TransactionDate": ["transactiondate", "trandate", "txndate"],
    "DisbursementAmount": ["disbursementamount"],
    "OtherChargesDebitAmount": ["otherchargesdebitamount"],
    "CollectedPrincipal": ["collectedprincipal"],
    "CollectedInterest": ["collectedinterest"],
    "CollectedPenalInterest": ["collectedpenalinterest"],
    "CollectedIOD": ["collectediod"],
    "CollectedOthers": ["collectedothers"],
    "BalancePrincipal": ["balanceprincipal"],
    "BalanceInterest": ["balanceinterest"],
    "BalancePenalInterest": ["balancepenalinterest"],
    "BalanceIOD": ["balanceiod"],
    "Charges": ["charges"],
    # membership / land / collateral
    "MemberName": ["membername"],
    "MemberSurName": ["membersurname"],
    "AdmissionDate": ["admissiondate", "joiningdate"],
    "DOB": ["dob", "dateofbirth"],
    "Age": ["age"],
    "GenderDescription": ["genderdescription", "gender"],
    "CasteDescription": ["castedescription", "caste"],
    "MaritalStatusDesc": ["maritalstatusdesc", "maritalstatus"],
    "CommunityDescription": ["communitydescription", "community"],
    "ContactNo": ["contactno", "mobileno", "phoneno"],
    "ShareBalance": ["sharebalance"],
    "Thrift": ["thrift"],
    "AadhaarNo": ["aadhaarno", "aadharno", "uidno"],
    "PANNo": ["panno", "pan"],
    "SurveyNo": ["surveyno", "survayno"],
    "LandType": ["landtype"],
    "IrrigationType": ["irrigationtype"],
    "LandOwnershipType": ["landownershiptype", "ownershiptype"],
    "SoilType": ["soiltype"],
    "GrossWeight": ["grossweight", "gwt"],
    "NetWeight": ["netweight", "nwt"],
    "CollateralDescription": ["collateraldescription", "collateraldesc",
                              "goldornamentdesc", "ornamentdescription"],
}


# ==============================================================================
# 2. RULE CATALOGUE  (drives the Explanation sheet)
# ==============================================================================

RULE_CATALOG: dict[str, dict] = {
    # ---- generic ----
    "G001": {"sev": HARD, "title": "Mandatory field is blank",
             "meaning": "A column the portal treats as compulsory has no value in this row.",
             "why": "The upload parser rejects the record outright.",
             "fix": "Fill the value in the source file. If it genuinely does not exist, remove the row from this batch."},
    "G002": {"sev": HARD, "title": "Identifier stored as a number",
             "meaning": "An identifier such as AdmissionNo or LoanNo ends in '.0', meaning Excel treated it as a number.",
             "why": "Numeric conversion drops leading zeros and rounds long account numbers, so the uploaded identifier no longer matches the core system.",
             "fix": "Select the column in Excel, Format Cells > Text, then re-paste the values."},
    "G003": {"sev": HARD, "title": "Date cannot be read",
             "meaning": "The date does not match any recognised format.",
             "why": "An unparseable date is rejected, and a wrongly guessed one is worse than a rejection.",
             "fix": "Use DD-MM-YYYY. Watch for text-formatted dates and day/month swaps."},
    "G004": {"sev": HARD, "title": "Duplicate key",
             "meaning": "The same identifier appears in more than one row of this file.",
             "why": "The portal keys on this identifier; the second row either overwrites the first or is rejected.",
             "fix": "Decide which row is correct and remove the other. Never merge two records by adding their amounts."},
    "G005": {"sev": REVIEW, "title": "Amount is not a number",
             "meaning": "An amount column holds text that cannot be read as a figure.",
             "why": "Non-numeric amounts are rejected at upload.",
             "fix": "Remove currency symbols, letters and notes from the cell; keep only the figure."},
    "G006": {"sev": INFO, "title": "Column-wide negative sign convention",
             "meaning": "Almost every row in this column is negative, so this is an export convention, not a row-level error.",
             "why": "Reported once at file level instead of once per row, which would otherwise bury every real error.",
             "fix": "Confirm with the source export whether credit balances are exported negative. If so, correct the sign for the whole column before upload; the portal expects positive balances."},
    "G007": {"sev": REVIEW, "title": "Isolated negative amount",
             "meaning": "This amount is negative while the rest of the column is positive.",
             "why": "A one-off negative in an otherwise positive column is normally a data-entry or calculation error, not a convention.",
             "fix": "Check this account's ledger. Note that negative OutstandingInterest and OutstandingPenalInterest are rejected outright by the portal."},
    "G008": {"sev": HARD, "title": "Negative outstanding interest",
             "meaning": "OutstandingInterest or OutstandingPenalInterest is below zero.",
             "why": "Confirmed portal rejection: 'Invalid Outstanding Interest Amount: -1018'.",
             "fix": "Correct the source calculation for this loan."},
    # ---- loans ----
    "L001": {"sev": HARD, "title": "ROI out of the accepted range",
             "meaning": f"Rate of interest must be greater than {ROI_MIN:g} and less than {ROI_MAX:g}.",
             "why": "The portal validates the interest band and rejects anything outside it.",
             "fix": "Correct the rate. A blank or 0 usually means it was never captured in the society software."},
    "L002": {"sev": HARD, "title": "Penal ROI out of the accepted range",
             "meaning": f"Penal rate must be greater than {PENAL_MIN:g} and less than {PENAL_MAX:g}.",
             "why": "Confirmed portal message: 'Invalid Penal ROI Percentage: 0, It should be less than 50 and greater than 0'. Zero is not accepted even where no penalty applies in practice.",
             "fix": "Use the society's actual penal rate, or the DCCB standard rate, never 0."},
    "L003": {"sev": HARD, "title": "Loan period is not a valid whole number of months",
             "meaning": "LoanPeriod must be a positive whole number.",
             "why": "The portal treats the period as an integer month count.",
             "fix": "Enter the tenure in whole months."},
    "L004": {"sev": HARD, "title": "Repayment frequency not recognised",
             "meaning": "The frequency text does not match a portal value.",
             "why": "The portal spelling is exact: Monthly, Quarterly, Halfyearly, Yearly. 'Half Yearly', 'HalfYearly' and 'half-yearly' are all rejected.",
             "fix": "Replace with the exact spelling, one word, only the first letter capital."},
    "L005": {"sev": HARD, "title": "Loan period does not match the repayment frequency",
             "meaning": "LoanPeriod is not an exact multiple of the frequency's period length.",
             "why": "Confirmed portal message: 'Invalid Loan Period: 13, Loan period should match the Repayment frequency'.",
             "fix": "Make the tenure a multiple of the cycle (12/24/36 for Yearly, 6/12/18 for Halfyearly), or change the frequency."},
    "L006": {"sev": HARD, "title": "Gestation period missing or invalid",
             "meaning": "GestationPeriod is blank, negative, or not shorter than the loan period.",
             "why": "Confirmed portal behaviour: GestationPeriod is required on MT/LT (agri term loan) uploads and the row is rejected when blank.",
             "fix": "Enter the gestation in months. It must be less than LoanPeriod."},
    "L007": {"sev": HARD, "title": "Sanctioned amount is zero or less",
             "meaning": "SanctionedAmount must be greater than 0.",
             "why": "A loan with no sanctioned amount is not a loan record.",
             "fix": "Fill the sanctioned figure from the loan ledger."},
    "L008": {"sev": REVIEW, "title": "DueDate does not match SanctionedDate + LoanPeriod",
             "meaning": "The maturity date is more than 2 days away from the derived value.",
             "why": "SanctionedDate is the fixed anchor; every other loan date is derived from it.",
             "fix": "Correct DueDate, not SanctionedDate. DueDate = SanctionedDate + LoanPeriod months."},
    "L009": {"sev": HARD, "title": "DueDate is on or before SanctionedDate",
             "meaning": "The loan matures before or on the day it was given.",
             "why": "The tenure becomes meaningless and the portal rejects it.",
             "fix": "Recalculate DueDate from SanctionedDate plus LoanPeriod months."},
    "L010": {"sev": REVIEW, "title": "FirstInstallmentDate does not match the formula",
             "meaning": "FirstInstallmentDate should be SanctionedDate plus one repayment cycle (Monthly +1, Quarterly +3, Halfyearly +6, Yearly +12 months). On MT/LT loans the gestation period is added on top.",
             "why": "A first instalment that does not follow the sanction date and frequency produces a wrong repayment schedule in the core system, even when the upload itself succeeds.",
             "fix": "Recompute from SanctionedDate. The expected value is shown in the Expected column."},
    "L011": {"sev": HARD, "title": "FirstInstallmentDate outside the loan period",
             "meaning": "The first instalment falls before the sanction date, or after the loan's own DueDate.",
             "why": "An instalment cannot fall outside the life of the loan; the portal rejects the schedule.",
             "fix": "Check RepaymentFrequency, GestationPeriod and LoanPeriod for this loan."},
    "L012": {"sev": REVIEW, "title": "InstallmentAmount does not match the formula",
             "meaning": "InstallmentAmount should be SanctionedAmount divided by the number of instalments, where the count is LoanPeriod split into RepaymentFrequency cycles. A 5% tolerance is allowed.",
             "why": "A wrong instalment amount gives the member a wrong repayment schedule. Changing RepaymentFrequency changes this expected amount.",
             "fix": "Recompute, or confirm the society deliberately uses a different instalment (e.g. interest-inclusive EMI). The expected value is shown in the Expected column."},
    "L013": {"sev": REVIEW, "title": "InstallmentAmount is missing or zero",
             "meaning": "The loan has a repayment frequency but no instalment amount.",
             "why": "The core system cannot build a repayment schedule without it.",
             "fix": "Compute as SanctionedAmount / number of instalments, or take it from the loan ledger."},
    "L014": {"sev": REVIEW, "title": "Outstanding principal exceeds sanctioned amount",
             "meaning": "The balance is larger than the amount ever sanctioned.",
             "why": "Usually a column mix-up, or interest folded into the principal column.",
             "fix": "Confirm against the ledger and split interest into OutstandingInterest."},
    "L015": {"sev": HARD, "title": "Sanctioned date is in the future",
             "meaning": "The loan is dated after today.",
             "why": "A loan cannot be sanctioned in the future; normally a year typo.",
             "fix": "Check the year on this row."},
    "L016": {"sev": HARD, "title": "Loan sanctioned before the member joined",
             "meaning": "SanctionedDate falls before that member's AdmissionDate in the membership file.",
             "why": "Confirmed portal rejection: 'Admission Date is greater than Sanction Date or invalid Admission Date'.",
             "fix": "Check which of the two dates is wrong for this member."},
    # ---- deposits ----
    "D001": {"sev": HARD, "title": "MaturityAmount is zero or blank",
             "meaning": "The deposit has no maturity value.",
             "why": "Confirmed portal behaviour: the deposit is rejected outright.",
             "fix": "Compute as DepositAmount + InterestAmount, or take it from the deposit register."},
    "D002": {"sev": HARD, "title": "DepositAmount is not less than MaturityAmount",
             "meaning": "For a lump-sum deposit the maturity value must exceed the principal.",
             "why": "A maturity that does not exceed the principal makes no sense for an interest-bearing deposit and is rejected.",
             "fix": "Check whether InterestAmount was left out of MaturityAmount."},
    "D003": {"sev": REVIEW, "title": "MaturityDate does not match DepositDate + term",
             "meaning": "MaturityDate should be DepositDate plus TerminDays, or plus TerminMonths.",
             "why": "On renewed deposits, DepositDate often holds the renewal date rather than the true start date.",
             "fix": "Compare against the reverse-calculated DepositDate shown in the Expected column and correct whichever is wrong."},
    "D004": {"sev": REVIEW, "title": "MaturityAmount does not match the formula",
             "meaning": "MaturityAmount = DepositAmount + InterestAmount, or (InstallmentAmount x TerminMonths) + InterestAmount for installment-type deposits. A 2% tolerance is allowed.",
             "why": "A mismatch means either the interest or the principal figure is wrong in the source.",
             "fix": "Recompute from the deposit register; the expected value is shown."},
    "D005": {"sev": HARD, "title": "Deposit term is missing",
             "meaning": "Neither TerminDays nor TerminMonths carries a positive value.",
             "why": "The core system cannot open a term deposit with no term.",
             "fix": "Fill TerminMonths (or TerminDays) from the deposit register."},
    "D006": {"sev": REVIEW, "title": "Interest rate outside the usual range",
             "meaning": f"RateOfInterest is below 0 or above {DEPOSIT_RATE_MAX:g}%.",
             "why": "Not an automatic rejection, but a rate outside this band is nearly always a data error.",
             "fix": "Verify against the society's rate chart for this product."},
    "D007": {"sev": HARD, "title": "RD or Pigmy installment amount missing",
             "meaning": "An installment-type deposit has no positive InstallmentAmount.",
             "why": "An RD or Pigmy account with no instalment cannot be operated.",
             "fix": "Fill from the deposit register."},
    "D008": {"sev": REVIEW, "title": "TotalInstallmentAmountPaid does not match",
             "meaning": "Expected InstallmentAmount x InstallmentsPaid.",
             "why": "A mismatch means either the count or the paid total is wrong, and the opening balance will be wrong after upload.",
             "fix": "Recompute from the passbook or deposit register."},
    "D009": {"sev": HARD, "title": "MaturityDate is on or before DepositDate",
             "meaning": "The deposit matures before or on the day it was opened.",
             "why": "Impossible sequence; the portal rejects it.",
             "fix": "Check both dates for this account."},
    "D010": {"sev": REVIEW, "title": "SB Yes/No field not normalised",
             "meaning": "ChequeOption or IsOrganisation holds something other than Yes or No.",
             "why": "The portal accepts only Yes or No in these fields.",
             "fix": "Replace with exactly Yes or No."},
    "D011": {"sev": HARD, "title": "Pigmy agent number missing",
             "meaning": "AgentNo is blank on a Pigmy account.",
             "why": "Pigmy collection is agent-based; the account cannot be assigned without it.",
             "fix": "Fill AgentNo, and make sure that agent exists in the Agent master."},
    # ---- membership / land / collateral / transactions ----
    "P001": {"sev": HARD, "title": "Duplicate AdmissionNo in membership",
             "meaning": "The same member number appears more than once.",
             "why": "AdmissionNo is the member key; a duplicate creates or overwrites the wrong member.",
             "fix": "Keep the higher-priority member type and remove the other, per the DCT suite's duplicate handling."},
    "P002": {"sev": REVIEW, "title": "Caste description not in the allowed list",
             "meaning": "CasteDescription must be one of the portal's fixed values.",
             "why": "Anything outside the list is rejected or silently blanked.",
             "fix": "Map to one of: " + ", ".join(sorted(VALID_CASTE_VALUES)) + "."},
    "P003": {"sev": REVIEW, "title": "Gender not normalised",
             "meaning": "GenderDescription must be Male, Female or Others.",
             "why": "The portal accepts only these three values.",
             "fix": "Map M/F/T and free text to Male, Female or Others."},
    "P004": {"sev": HARD, "title": "AdmissionDate is in the future",
             "meaning": "The member joined after today.",
             "why": "Confirmed portal rejection alongside the sanction-date check.",
             "fix": "Check the year on this row."},
    "P005": {"sev": REVIEW, "title": "Date of birth inconsistent with age",
             "meaning": "DOB and Age disagree by more than one year, or DOB is after AdmissionDate.",
             "why": "The core system derives eligibility from these; a mismatch causes downstream problems.",
             "fix": "Correct whichever is wrong. The DCT suite can derive DOB as AdmissionDate minus 20 years when it is missing."},
    "P006": {"sev": REVIEW, "title": "Duplicate Aadhaar or PAN",
             "meaning": "The same Aadhaar or PAN appears against more than one member.",
             "why": "Normally two records for the same person, which will create a duplicate member.",
             "fix": "Confirm whether the two members are the same person before uploading."},
    "P007": {"sev": REVIEW, "title": "Contact number is not 10 digits",
             "meaning": "ContactNo does not look like a valid Indian mobile number.",
             "why": "Not an upload blocker, but SMS and KYC follow-up will fail.",
             "fix": "Correct or blank the number; a wrong number is worse than none."},
    "X001": {"sev": HARD, "title": "AdmissionNo has no matching member",
             "meaning": "This AdmissionNo does not appear in the membership file in this folder.",
             "why": "Confirmed portal rejection: 'Could not find Application Details'. The portal attaches every record to an existing member.",
             "fix": "Either the member is missing from the membership template, or the AdmissionNo is wrong here. Check both."},
    "X002": {"sev": HARD, "title": "LoanNo has no matching loan master",
             "meaning": "This LoanNo does not appear in any loan template in this folder.",
             "why": "Transactions and collateral attach to an existing loan account; an orphan row is rejected or silently dropped.",
             "fix": "Include the master file for that product, or correct the LoanNo. Matching is on LoanNo only, never on amount."},
    "X003": {"sev": HARD, "title": "Transaction dated before the loan was sanctioned",
             "meaning": "TransactionDate is earlier than the SanctionedDate on the matching master.",
             "why": "Money cannot move on a loan that did not exist yet. Usually an opening-balance row or a date typo.",
             "fix": "Opening balances should not be uploaded as transactions. Otherwise correct the date."},
    "X004": {"sev": REVIEW, "title": "Transaction has no amount",
             "meaning": "Every amount column on this transaction row is zero or blank.",
             "why": "A zero-value transaction carries no information; the DCT suite drops these before upload.",
             "fix": "Remove the row, or fill the correct amount."},
    "X005": {"sev": REVIEW, "title": "Transaction AdmissionNo differs from the loan master",
             "meaning": "The member on this transaction is not the member who owns the loan.",
             "why": "Confirmed cause of a whole product's 'Could not find Application Details' rejections: the source transaction's own AdmissionNo did not match the loan owner in the master.",
             "fix": "The master is authoritative. Correct the transaction, or confirm the loan's true owner."},
    "X006": {"sev": REVIEW, "title": "LoanNo exists in more than one loan master",
             "meaning": "The same LoanNo is used by different products for different members.",
             "why": "LoanNo is not guaranteed unique across products. Picking the wrong one misassigns a real transaction to the wrong product.",
             "fix": "Disambiguate using the transaction's own AdmissionNo, or renumber one of the loans."},
}


# ==============================================================================
# 3. FILE READING
# ==============================================================================

def sniff_real_format(path: Path) -> str:
    """True format from magic bytes. Portal downloads named .xls are often TSV."""
    with path.open("rb") as fh:
        head = fh.read(8)
    if head[:2] == b"PK":
        return "xlsx"
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "xls"
    return "text"


def _sniff_delimiter(path: Path) -> str:
    sample = path.read_text(encoding="utf-8", errors="ignore")[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
    except csv.Error:
        first = sample.splitlines()[0] if sample.splitlines() else ""
        return max(",\t;|", key=first.count)


def read_any(path: Path) -> tuple[pd.DataFrame, str]:
    """Load with every column as text, so identifiers are never coerced."""
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if path.stat().st_size == 0:
        raise ValueError("File is empty (0 bytes).")

    fmt = sniff_real_format(path)
    if fmt == "xlsx":
        return pd.read_excel(path, sheet_name=0, dtype=str,
                             engine="openpyxl"), "Excel (.xlsx)"
    if fmt == "xls":
        try:
            return pd.read_excel(path, sheet_name=0, dtype=str,
                                 engine="xlrd"), "Legacy Excel (.xls)"
        except ImportError:
            raise RuntimeError("Genuine legacy .xls file. Run: pip install xlrd") from None

    delim = _sniff_delimiter(path)
    df = pd.read_csv(path, sep=delim, dtype=str, encoding="utf-8",
                     encoding_errors="ignore", keep_default_na=False,
                     na_values=[""])
    shown = {"\t": "TAB", ",": "COMMA", ";": "SEMICOLON", "|": "PIPE"}
    note = f"Delimited text ({shown.get(delim, repr(delim))}-separated)"
    if path.suffix.lower() in (".xls", ".xlsx", ".xlsm"):
        note += "  <-- named as Excel but is NOT a real Excel file"
    return df, note


# ==============================================================================
# 4. HEADERS AND CATEGORY DETECTION
# ==============================================================================

def _norm(name) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


_ALIAS_LOOKUP: dict[str, str] = {}
for _canon, _aliases in COLUMN_ALIASES.items():
    _ALIAS_LOOKUP.setdefault(_norm(_canon), _canon)
    for _a in _aliases:
        _ALIAS_LOOKUP.setdefault(_norm(_a), _canon)


def map_headers(df: pd.DataFrame) -> tuple[pd.DataFrame, dict, list[str]]:
    """Canonicalise headers. Unknown columns are kept untouched."""
    rename, unmapped = {}, []
    for col in df.columns:
        canonical = _ALIAS_LOOKUP.get(_norm(col))
        if canonical and canonical not in rename.values():
            rename[col] = canonical
        elif canonical:
            unmapped.append(f"{col} (second column matching {canonical})")
        else:
            unmapped.append(str(col))
    return df.rename(columns=rename), rename, unmapped


def detect_category(path: Path, columns: set[str]) -> tuple[str, str]:
    """Return (category, how_it_was_decided).

    The file name is tried first, because DCT templates are named after the
    template they are ("4.LTandMTLoanTemplate", "10.RDTemplate") and that is a
    stronger signal than headers alone - several deposit templates share almost
    identical columns and are only distinguishable by name. Header signatures
    are the fallback for a file that was renamed or exported generically, and
    a disagreement between the two is reported rather than silently resolved.
    """
    cols_ci = {str(c).lower() for c in columns}

    by_header = None
    for cat, signature in HEADER_SIGNATURES:
        if {s.lower() for s in signature} <= cols_ci:
            by_header = (cat, f"header signature ({', '.join(sorted(signature))})")
            break

    for pattern, cat in FILENAME_PATTERNS:
        if pattern.search(path.stem):
            how = f"file name (/{pattern.pattern}/)"
            if by_header and by_header[0] != cat:
                how += f"; note: headers looked like {by_header[0]}"
            return cat, how

    if by_header:
        return by_header[0], by_header[1] + " (file name matched nothing)"
    return "UNKNOWN", "no file name pattern or header signature matched"


# ==============================================================================
# 5. PARSERS
# ==============================================================================

DATE_PATTERNS = ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d/%m/%y",
                 "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%m/%d/%Y"]
_BLANK = {"", "nan", "nat", "none", "null", "-", "na", "n/a"}


def is_blank(value) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().lower() in _BLANK


def parse_date(value):
    """Return (date|None, error|None)."""
    if is_blank(value):
        return None, "empty"
    text = str(value).strip()
    if re.fullmatch(r"\d{5}(\.\d+)?", text):            # Excel serial
        serial = float(text)
        if 20000 <= serial <= 60000:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=serial)).date(), None
    text = text.split(" ")[0].split("T")[0]
    for fmt in DATE_PATTERNS:
        try:
            return datetime.strptime(text, fmt).date(), None
        except ValueError:
            continue
    return None, f"unrecognised date format: {value!r}"


def parse_number(value):
    """Return (float|None, error|None). Sign is preserved, never abs()'d."""
    if is_blank(value):
        return None, "empty"
    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"(?i)^rs\.?", "", text)
    text = re.sub(r"[()₹$,%\s]", "", text)
    try:
        num = float(text)
    except ValueError:
        return None, f"not a number: {value!r}"
    return (-num if negative else num), None


def freq_months(value):
    if is_blank(value):
        return None, "empty"
    key = re.sub(r"[^a-z]", "", str(value).lower())
    if key in FREQUENCY_MONTHS:
        return FREQUENCY_MONTHS[key], None
    return None, f"unknown RepaymentFrequency {value!r}"


def add_months(d: date | None, months) -> date | None:
    """Calendar month arithmetic, clamping the day to the target month's length."""
    if d is None or months in (None, ""):
        return None
    try:
        months = int(round(float(months)))
    except (TypeError, ValueError):
        return None
    total = d.month - 1 + months
    year = d.year + total // 12
    month = total % 12 + 1
    if month == 12:
        last = 31
    else:
        last = (date(year + (month == 12), month % 12 + 1, 1) - timedelta(days=1)).day
    return date(year, month, min(d.day, last))


def digits_only(value) -> str:
    return re.sub(r"\D", "", str(value or ""))


# ==============================================================================
# 6. ISSUE COLLECTION
# ==============================================================================

class Issue:
    __slots__ = ("row", "rule", "field", "value", "expected", "detail")

    def __init__(self, row, rule, field, value=None, expected=None, detail=""):
        self.row = row
        self.rule = rule
        self.field = field
        self.value = value
        self.expected = expected
        self.detail = detail

    @property
    def severity(self) -> str:
        return RULE_CATALOG.get(self.rule, {}).get("sev", REVIEW)


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, date):
        return v.strftime("%d-%m-%Y")
    if isinstance(v, float):
        return f"{v:,.2f}"
    return str(v)


# ==============================================================================
# 7. VALIDATORS
# ==============================================================================

def check_common(df, cols, cat, issues, sign_flags):
    """Mandatory fields, identifier typing, date parsing, and sign convention."""
    required = [c for c in REQUIRED_FIELDS.get(cat, []) if c in cols]

    for col in required:
        blanks = df.index[df[col].map(is_blank)]
        for i in blanks:
            issues.append(Issue(i, "G001", col, detail=f"{col} is blank"))

    for col in [c for c in TEXT_ID_FIELDS if c in cols]:
        bad = df.index[df[col].astype(str).str.strip().str.endswith(".0")]
        for i in bad:
            issues.append(Issue(i, "G002", col, df.at[i, col],
                                detail=f"{col} looks numeric-converted"))

    for col in [c for c in DATE_FIELDS if c in cols]:
        for i, raw in df[col].items():
            dt, err = parse_date(raw)
            if err and err != "empty":
                issues.append(Issue(i, "G003", col, raw, detail=err))

    # Column-wide sign convention, judged once per column, not per row.
    for col in [c for c in SIGN_WATCH_FIELDS if c in cols]:
        nums = [parse_number(v)[0] for v in df[col]]
        nums = [n for n in nums if n is not None and n != 0]
        if not nums:
            continue
        neg = sum(1 for n in nums if n < 0)
        if neg == 0:
            continue
        share = neg / len(nums)
        if share >= RULE_SETTINGS["negative_sign_convention_share"]:
            sign_flags.append((col, neg, len(nums), share))
            issues.append(Issue(None, "G006", col, f"{neg:,} of {len(nums):,} rows",
                                detail=f"{share*100:.1f}% of non-zero values in "
                                       f"{col} are negative"))
        else:
            for i, v in df[col].items():
                n, _ = parse_number(v)
                if n is not None and n < 0:
                    rule = "G008" if col in ("OutstandingInterest",
                                             "OutstandingPenalInterest") else "G007"
                    issues.append(Issue(i, rule, col, n,
                                        detail=f"{col} is negative"))


def check_duplicates(df, cols, cat, issues):
    """Identifier-based duplicates only. Amounts are never used as a key."""
    key_map = {
        "MEMBERSHIP": ["AdmissionNo"],
        "AGENT": ["AgentNo"],
    }
    if cat in LOAN_CATEGORIES:
        keys = ["LoanNo"]
    elif cat in ("FD", "FDNONCUM", "RD", "PIGMY", "SB"):
        keys = ["AccountNo"]
    else:
        keys = key_map.get(cat, [])

    dup_index = set()
    for key in [k for k in keys if k in cols]:
        series = df[key].astype(str).str.strip().str.upper()
        mask = series.duplicated(keep=False) & (series != "")
        for i in df.index[mask]:
            dup_index.add(i)
            issues.append(Issue(i, "G004", key, df.at[i, key],
                                detail=f"{key} appears more than once in this file"))
    return dup_index


def check_loans(df, cols, cat, issues, member_dates):
    for i, row in df.iterrows():
        sd, _ = parse_date(row.get("SanctionedDate")) if "SanctionedDate" in cols else (None, None)
        dd, _ = parse_date(row.get("DueDate")) if "DueDate" in cols else (None, None)
        fid, _ = parse_date(row.get("FirstInstallmentDate")) if "FirstInstallmentDate" in cols else (None, None)

        # --- rates (column name differs by template) ------------------------
        roi_col = "ROIPercentage" if "ROIPercentage" in cols else "ROI"
        penal_col = "PenalROIPercentage" if "PenalROIPercentage" in cols else "PenalROI"

        if roi_col in cols:
            roi, err = parse_number(row.get(roi_col))
            if err and err != "empty":
                issues.append(Issue(i, "G005", roi_col, row.get(roi_col), detail=err))
            elif roi is not None and not (ROI_MIN < roi < ROI_MAX):
                issues.append(Issue(i, "L001", roi_col, roi,
                                    f"{ROI_MIN:g} < value < {ROI_MAX:g}"))

        if penal_col in cols:
            penal, err = parse_number(row.get(penal_col))
            if err and err != "empty":
                issues.append(Issue(i, "G005", penal_col, row.get(penal_col), detail=err))
            elif penal is not None and not (PENAL_MIN < penal < PENAL_MAX):
                issues.append(Issue(i, "L002", penal_col, penal,
                                    f"{PENAL_MIN:g} < value < {PENAL_MAX:g}"))

        # --- period and frequency -------------------------------------------
        period = None
        if "LoanPeriod" in cols:
            val, err = parse_number(row.get("LoanPeriod"))
            if err and err != "empty":
                issues.append(Issue(i, "L003", "LoanPeriod", row.get("LoanPeriod"), detail=err))
            elif val is not None:
                if val != int(val) or val <= 0:
                    issues.append(Issue(i, "L003", "LoanPeriod", val, "whole months > 0"))
                else:
                    period = int(val)

        months = None
        if "RepaymentFrequency" in cols and not is_blank(row.get("RepaymentFrequency")):
            months, err = freq_months(row.get("RepaymentFrequency"))
            if err:
                issues.append(Issue(
                    i, "L004", "RepaymentFrequency", row.get("RepaymentFrequency"),
                    "Monthly / Quarterly / Halfyearly / Yearly"))
            else:
                exact = PORTAL_FREQUENCY_SPELLING[months]
                if str(row.get("RepaymentFrequency")).strip() != exact:
                    issues.append(Issue(
                        i, "L004", "RepaymentFrequency",
                        row.get("RepaymentFrequency"), exact,
                        "recognised, but the portal spelling must match exactly"))

        if period and months and period % months != 0:
            issues.append(Issue(
                i, "L005", "LoanPeriod", period,
                f"a multiple of {months}",
                f"LoanPeriod {period} is not an exact multiple of the "
                f"{PORTAL_FREQUENCY_SPELLING[months]} cycle ({months} months)"))

        # --- gestation (MT/LT only) -----------------------------------------
        gest = None
        if cat == "MTLT":
            raw_gest = row.get("GestationPeriod") if "GestationPeriod" in cols else None
            if is_blank(raw_gest):
                issues.append(Issue(i, "L006", "GestationPeriod", raw_gest, "required"))
            else:
                gest, err = parse_number(raw_gest)
                if err:
                    issues.append(Issue(i, "L006", "GestationPeriod", raw_gest, detail=err))
                elif gest < 0 or (period is not None and gest >= period):
                    issues.append(Issue(i, "L006", "GestationPeriod", gest,
                                        f"0 to {period - 1 if period else '?'}"))

        # --- amounts ---------------------------------------------------------
        sa, _ = parse_number(row.get("SanctionedAmount")) if "SanctionedAmount" in cols else (None, None)
        if sa is not None and sa <= 0:
            issues.append(Issue(i, "L007", "SanctionedAmount", sa, "> 0"))

        op, _ = parse_number(row.get("OutstandingPrincipal")) if "OutstandingPrincipal" in cols else (None, None)
        if sa and op is not None and abs(op) > abs(sa):
            issues.append(Issue(i, "L014", "OutstandingPrincipal", op,
                                f"not more than {sa:,.2f}"))

        # --- dates derived from SanctionedDate (the fixed anchor) ------------
        expected_dd = add_months(sd, period) if (sd and period) else None
        if expected_dd and dd and abs((expected_dd - dd).days) > RULE_SETTINGS["date_tolerance_days"]:
            issues.append(Issue(i, "L008", "DueDate", dd, expected_dd,
                                "DueDate = SanctionedDate + LoanPeriod months"))
        if sd and dd and dd <= sd:
            issues.append(Issue(i, "L009", "DueDate", dd, f"after {sd:%d-%m-%Y}"))
        if sd and sd > date.today():
            issues.append(Issue(i, "L015", "SanctionedDate", sd, "not after today"))

        # --- FirstInstallmentDate -------------------------------------------
        # Base rule from the DCT suite: SanctionedDate + one repayment cycle.
        # On MT/LT loans the gestation period is served first, so it is added.
        if sd and months:
            offset = months + int(gest) if (cat == "MTLT" and gest) else months
            expected_fid = add_months(sd, offset)
        if fid and abs((expected_fid - fid).days) > RULE_SETTINGS["date_tolerance_days"]:
                issues.append(Issue(
                    i, "L010", "FirstInstallmentDate", fid, expected_fid,
                    f"SanctionedDate + {offset} month(s) "
                    f"({PORTAL_FREQUENCY_SPELLING[months]}"
                    + (f" + {int(gest)} gestation" if (cat == "MTLT" and gest) else "")
                    + ")"))
            if not fid and "FirstInstallmentDate" in cols:
                issues.append(Issue(i, "L010", "FirstInstallmentDate", None,
                                    expected_fid, "blank - derive from SanctionedDate"))
        if fid and sd and fid < sd:
            issues.append(Issue(i, "L011", "FirstInstallmentDate", fid,
                                f"on or after {sd:%d-%m-%Y}"))
        if fid and expected_dd and fid > expected_dd:
            issues.append(Issue(i, "L011", "FirstInstallmentDate", fid,
                                f"not after {expected_dd:%d-%m-%Y}",
                                "first instalment falls after the loan's own DueDate"))

        # --- InstallmentAmount ----------------------------------------------
        if "InstallmentAmount" in cols:
            ia, err = parse_number(row.get("InstallmentAmount"))
            if err and err != "empty":
                issues.append(Issue(i, "G005", "InstallmentAmount",
                                    row.get("InstallmentAmount"), detail=err))
            if sa and period and months:
                n_inst = max(1, round(period / months))
                expected_ia = round(sa / n_inst, 2)
                if ia is None or ia == 0:
                    issues.append(Issue(i, "L013", "InstallmentAmount", ia,
                                        expected_ia,
                                        f"{n_inst} instalment(s) over {period} month(s)"))
        elif abs(ia - expected_ia) > max(1.0, RULE_SETTINGS["installment_amount_tolerance_pct"] * expected_ia):
                    issues.append(Issue(
                        i, "L012", "InstallmentAmount", ia, expected_ia,
                        f"SanctionedAmount / {n_inst} instalment(s) "
                        f"({PORTAL_FREQUENCY_SPELLING[months]} over {period} months)"))

        # --- member cross-check ----------------------------------------------
        if member_dates and "AdmissionNo" in cols:
            key = digits_only(row.get("AdmissionNo"))
            if key:
                if key not in member_dates:
                    issues.append(Issue(i, "X001", "AdmissionNo",
                                        row.get("AdmissionNo"),
                                        detail="not found in the membership file"))
                else:
                    adm = member_dates[key]
                    if adm and sd and adm > sd:
                        issues.append(Issue(i, "L016", "SanctionedDate", sd,
                                            f"on or after {adm:%d-%m-%Y}",
                                            "member's AdmissionDate is later than "
                                            "this loan's SanctionedDate"))


def check_deposits(df, cols, cat, issues, member_dates):
    for i, row in df.iterrows():
        dep_date, _ = parse_date(row.get("DepositDate")) if "DepositDate" in cols else (None, None)
        mat_date, _ = parse_date(row.get("MaturityDate")) if "MaturityDate" in cols else (None, None)

        dep_amt, _ = parse_number(row.get("DepositAmount")) if "DepositAmount" in cols else (None, None)
        mat_amt, _ = parse_number(row.get("MaturityAmount")) if "MaturityAmount" in cols else (None, None)
        int_amt, _ = parse_number(row.get("InterestAmount")) if "InterestAmount" in cols else (None, None)
        inst_amt, _ = parse_number(row.get("InstallmentAmount")) if "InstallmentAmount" in cols else (None, None)
        td, _ = parse_number(row.get("TerminDays")) if "TerminDays" in cols else (None, None)
        tm, _ = parse_number(row.get("TerminMonths")) if "TerminMonths" in cols else (None, None)
        dep_type = str(row.get("DepositTypeDesc", "") or "")
        installment_type = "install" in dep_type.lower() or cat in ("RD", "PIGMY")

        # --- term ------------------------------------------------------------
        if cat in ("FD", "FDNONCUM", "RD"):
            if not (td and td > 0) and not (tm and tm > 0):
                issues.append(Issue(i, "D005", "TerminMonths", row.get("TerminMonths"),
                                    "> 0", "neither TerminDays nor TerminMonths is set"))

        # --- maturity date ----------------------------------------------------
        expected_md = None
        if dep_date and td and td > 0:
            expected_md = dep_date + timedelta(days=float(td))
        elif dep_date and tm and tm > 0:
            expected_md = add_months(dep_date, tm)

        if expected_md and mat_date and abs((expected_md - mat_date).days) > RULE_SETTINGS["date_tolerance_days"]:
            # Reverse calc helps when DepositDate holds a renewal date.
            reverse = None
            if mat_date and td and td > 0:
                reverse = mat_date - timedelta(days=float(td))
            elif mat_date and tm and tm > 0:
                reverse = add_months(mat_date, -float(tm))
            issues.append(Issue(i, "D003", "MaturityDate", mat_date, expected_md,
                                "DepositDate + term"
                                + (f"; or DepositDate should be {reverse:%d-%m-%Y} "
                                   f"if this deposit was renewed" if reverse else "")))
        if dep_date and mat_date and mat_date <= dep_date:
            issues.append(Issue(i, "D009", "MaturityDate", mat_date,
                                f"after {dep_date:%d-%m-%Y}"))

        # --- maturity amount --------------------------------------------------
        if cat in ("FD", "FDNONCUM", "RD"):
            if mat_amt is None or mat_amt <= 0:
                issues.append(Issue(i, "D001", "MaturityAmount", mat_amt, "> 0"))
            else:
                principal = (inst_amt * float(tm or 0)) if (installment_type and inst_amt) else dep_amt
                if principal and int_amt is not None:
                    expected_ma = principal + int_amt
        if abs(mat_amt - expected_ma) > max(1.0, RULE_SETTINGS["deposit_amount_tolerance_pct"] * abs(expected_ma)):
                        issues.append(Issue(
                            i, "D004", "MaturityAmount", mat_amt, round(expected_ma, 2),
                            "(InstallmentAmount x TerminMonths) + InterestAmount"
                            if installment_type else
                            "DepositAmount + InterestAmount"))
                if not installment_type and dep_amt and dep_amt >= mat_amt:
                    issues.append(Issue(i, "D002", "DepositAmount", dep_amt,
                                        f"less than {mat_amt:,.2f}"))

        # --- rate -------------------------------------------------------------
        if "RateOfInterest" in cols:
            rate, err = parse_number(row.get("RateOfInterest"))
            if err and err != "empty":
                issues.append(Issue(i, "G005", "RateOfInterest",
                                    row.get("RateOfInterest"), detail=err))
            elif rate is not None and (rate < 0 or rate > DEPOSIT_RATE_MAX):
                issues.append(Issue(i, "D006", "RateOfInterest", rate,
                                    f"0 to {DEPOSIT_RATE_MAX:g}"))

        # --- installment-type specifics ---------------------------------------
        if cat in ("RD", "PIGMY"):
            if inst_amt is None or inst_amt <= 0:
                issues.append(Issue(i, "D007", "InstallmentAmount", inst_amt, "> 0"))
            paid_n, _ = parse_number(row.get("InstallmentsPaid")) if "InstallmentsPaid" in cols else (None, None)
            total_paid, _ = parse_number(row.get("TotalInstallmentAmountPaid")) if "TotalInstallmentAmountPaid" in cols else (None, None)
            if inst_amt and paid_n and total_paid:
                expected_tp = inst_amt * paid_n
        if abs(total_paid - expected_tp) > max(1.0, RULE_SETTINGS["rd_total_paid_tolerance_pct"] * expected_tp):
                    issues.append(Issue(i, "D008", "TotalInstallmentAmountPaid",
                                        total_paid, round(expected_tp, 2),
                                        f"InstallmentAmount x InstallmentsPaid "
                                        f"({inst_amt:,.2f} x {paid_n:g})"))

        if cat == "PIGMY" and "AgentNo" in cols and is_blank(row.get("AgentNo")):
            issues.append(Issue(i, "D011", "AgentNo", None, "required"))

        # --- SB Yes/No --------------------------------------------------------
        if cat == "SB":
            for f in ("ChequeOption", "IsOrganisation"):
                if f in cols and str(row.get(f, "")).strip() not in ("Yes", "No"):
                    issues.append(Issue(i, "D010", f, row.get(f), "Yes or No"))

        # --- member cross-check ------------------------------------------------
        if member_dates and "AdmissionNo" in cols:
            key = digits_only(row.get("AdmissionNo"))
            if key and key not in member_dates:
                issues.append(Issue(i, "X001", "AdmissionNo", row.get("AdmissionNo"),
                                    detail="not found in the membership file"))


def check_membership(df, cols, issues):
    for i, row in df.iterrows():
        if "AdmissionDate" in cols:
            adm, err = parse_date(row.get("AdmissionDate"))
            if adm and adm > date.today():
                issues.append(Issue(i, "P004", "AdmissionDate", adm, "not after today"))
        else:
            adm = None

        if "DOB" in cols:
            dob, _ = parse_date(row.get("DOB"))
            if dob:
                if adm and dob > adm:
                    issues.append(Issue(i, "P005", "DOB", dob,
                                        f"before {adm:%d-%m-%Y}",
                                        "date of birth is after the admission date"))
                if "Age" in cols:
                    age, _ = parse_number(row.get("Age"))
                    if age:
                        ref = adm or date.today()
                        derived = ref.year - dob.year - ((ref.month, ref.day) < (dob.month, dob.day))
                        if abs(derived - age) > 1:
                            issues.append(Issue(i, "P005", "Age", age, derived,
                                                "age does not agree with DOB"))

        if "GenderDescription" in cols:
            g = str(row.get("GenderDescription", "") or "").strip()
            if g and g not in VALID_GENDER_VALUES:
                issues.append(Issue(i, "P003", "GenderDescription", g,
                                    " / ".join(sorted(VALID_GENDER_VALUES))))

        if "CasteDescription" in cols:
            c = str(row.get("CasteDescription", "") or "").strip()
            if c and c not in VALID_CASTE_VALUES:
                issues.append(Issue(i, "P002", "CasteDescription", c,
                                    "one of the portal's fixed caste values"))

        if "ContactNo" in cols:
            contact = digits_only(row.get("ContactNo"))
            if contact and len(contact) != 10:
                issues.append(Issue(i, "P007", "ContactNo", row.get("ContactNo"),
                                    "10 digits"))

    for col, label in (("AadhaarNo", "Aadhaar"), ("PANNo", "PAN")):
        if col not in cols:
            continue
        series = df[col].astype(str).str.strip().str.upper()
        mask = series.duplicated(keep=False) & ~series.map(is_blank)
        for i in df.index[mask]:
            issues.append(Issue(i, "P006", col, df.at[i, col],
                                detail=f"{label} shared with another member"))


def check_land(df, cols, issues, member_dates):
    for i, row in df.iterrows():
        if member_dates and "AdmissionNo" in cols:
            key = digits_only(row.get("AdmissionNo"))
            if key and key not in member_dates:
                issues.append(Issue(i, "X001", "AdmissionNo", row.get("AdmissionNo"),
                                    detail="not found in the membership file"))
        for col, allowed in (("LandType", VALID_LAND_TYPE),
                             ("IrrigationType", VALID_IRRIGATION),
                             ("LandOwnershipType", VALID_OWNERSHIP)):
            if col in cols:
                v = str(row.get(col, "") or "").strip()
                if v and v not in allowed:
                    issues.append(Issue(i, "P002", col, v, " / ".join(sorted(allowed)),
                                        f"{col} is not one of the portal's values"))


def check_collateral(df, cols, issues, loan_index):
    for i, row in df.iterrows():
        if loan_index and "LoanNo" in cols:
            key = digits_only(row.get("LoanNo"))
            if key and key not in loan_index:
                issues.append(Issue(i, "X002", "LoanNo", row.get("LoanNo"),
                                    detail="no loan master in this folder carries this LoanNo"))


TXN_AMOUNT_COLS = ["DisbursementAmount", "OtherChargesDebitAmount",
                   "CollectedPrincipal", "CollectedInterest",
                   "CollectedPenalInterest", "CollectedIOD", "CollectedOthers",
                   "Charges"]


def check_transactions(df, cols, issues, loan_index):
    present_amounts = [c for c in TXN_AMOUNT_COLS if c in cols]
    for i, row in df.iterrows():
        txn_date, _ = parse_date(row.get("TransactionDate")) if "TransactionDate" in cols else (None, None)

        if present_amounts:
            has_amount = False
            for c in present_amounts:
                n, _ = parse_number(row.get(c))
                if n is not None and n != 0:
                    has_amount = True
                    break
            if not has_amount:
                issues.append(Issue(i, "X004", "amounts", None, "at least one non-zero",
                                    "every amount column on this row is zero or blank"))

        if not loan_index or "LoanNo" not in cols:
            continue
        key = digits_only(row.get("LoanNo"))
        if not key:
            continue
        candidates = loan_index.get(key)
        if not candidates:
            issues.append(Issue(i, "X002", "LoanNo", row.get("LoanNo"),
                                detail="no loan master in this folder carries this LoanNo"))
            continue

        txn_admno = digits_only(row.get("AdmissionNo")) if "AdmissionNo" in cols else ""
        if len(candidates) > 1:
            exact = [c for c in candidates if txn_admno and c["admission_no"] == txn_admno]
            if len(exact) == 1:
                candidates = exact
            else:
                issues.append(Issue(
                    i, "X006", "LoanNo", row.get("LoanNo"),
                    detail="LoanNo used by: "
                           + "; ".join(f"{c['product']} (member {c['admission_no']})"
                                       for c in candidates[:4])))
                continue

        owner = candidates[0]
        if txn_admno and owner["admission_no"] and txn_admno != owner["admission_no"]:
            issues.append(Issue(i, "X005", "AdmissionNo", row.get("AdmissionNo"),
                                owner["admission_no"],
                                f"the master says this loan belongs to member "
                                f"{owner['admission_no']} ({owner['product']})"))
        if txn_date and owner["sanctioned"] and txn_date < owner["sanctioned"]:
            issues.append(Issue(i, "X003", "TransactionDate", txn_date,
                                f"on or after {owner['sanctioned']:%d-%m-%Y}"))


# ==============================================================================
# 8. PER-FILE RESULT
# ==============================================================================

class FileResult:
    def __init__(self):
        self.path: Path | None = None
        self.name = ""
        self.category = "UNKNOWN"
        self.detected_by = ""
        self.format_note = ""
        self.total = 0
        self.clean = pd.DataFrame()
        self.rejected = pd.DataFrame()
        self.review = pd.DataFrame()
        self.duplicates = pd.DataFrame()
        self.rule_counts: dict[str, int] = {}
        self.sign_flags: list = []
        self.unmapped_columns: list[str] = []
        self.mapped: dict = {}
        self.missing_template_cols: list[str] = []
        self.output_path: Path | None = None
        self.failure: str | None = None
        self.skipped_reason: str | None = None

    @property
    def n_clean(self):
        return len(self.clean)

    @property
    def n_rejected(self):
        return len(self.rejected)

    @property
    def n_review(self):
        return len(self.review)

    @property
    def pass_pct(self):
        return (self.n_clean / self.total * 100) if self.total else 0.0

    @property
    def family(self):
        if self.category in LOAN_CATEGORIES:
            return "Loans"
        if self.category in DEPOSIT_CATEGORIES:
            return "Deposits"
        if self.category == "TRANSACTIONS":
            return "Transactions"
        return "Membership & Others"


# ==============================================================================
# 9. CLEAN-DATA FORMATTING
# ==============================================================================

def format_clean(df: pd.DataFrame, cols: set[str], cat: str) -> tuple[pd.DataFrame, list[str]]:
    """Return the clean rows in template order, typed for upload.

    Values are only re-presented, never re-derived: a date becomes DD-MM-YYYY,
    an amount becomes a number rounded to 2 decimals, an identifier stays text.
    Nothing is recalculated or substituted.
    """
    template = TEMPLATES.get(cat) or []
    # Match template columns case-insensitively: the DCT templates mix "NAME"
    # and "Name" for the same field, and a case-only difference must not make
    # the tool think the column is missing and blank it out.
    ci = {str(c).lower(): c for c in df.columns}
    for want in template:
        actual = ci.get(want.lower())
        if actual is not None and actual != want:
            df = df.rename(columns={actual: want})
    missing = [c for c in template if c not in df.columns]
    order = [c for c in template if c in df.columns] if template else list(df.columns)
    extras = [c for c in df.columns if c not in order]
    out = df[order + extras].copy() if template else df.copy()

    for col in out.columns:
        if col in DATE_FIELDS:
            out[col] = out[col].map(
                lambda v: (parse_date(v)[0].strftime("%d-%m-%Y")
                           if parse_date(v)[0] else ""))
        elif col in AMOUNT_FIELDS or col in RATE_FIELDS:
            out[col] = out[col].map(
                lambda v: (round(parse_number(v)[0], 2)
                           if parse_number(v)[0] is not None else ""))
        elif col in INT_FIELDS:
            out[col] = out[col].map(
                lambda v: (int(round(parse_number(v)[0]))
                           if parse_number(v)[0] is not None else ""))
        elif col in TEXT_ID_FIELDS:
            out[col] = out[col].map(
                lambda v: "" if is_blank(v) else re.sub(r"\.0$", "", str(v).strip()))
        else:
            out[col] = out[col].map(lambda v: "" if is_blank(v) else str(v).strip())

    # Add any template column the source did not have, so the shape is right.
    for col in missing:
        out[col] = ""
    if template:
        out = out[[c for c in template] + extras]
    return out, missing


# ==============================================================================
# 10. VALIDATE ONE FILE
# ==============================================================================

def validate_one(path: Path, member_dates: dict, loan_index: dict,
                 progress=None) -> FileResult:
    res = FileResult()
    res.path, res.name = path, path.name

    raw, res.format_note = read_any(path)
    df, res.mapped, res.unmapped_columns = map_headers(raw)
    df = df.reset_index(drop=True)
    cols = set(df.columns)
    res.category, res.detected_by = detect_category(path, cols)
    res.total = len(df)

    if res.category in SKIP_CATEGORIES:
        res.skipped_reason = ("Reference file (balance sheet, villages list, "
                              "PACS ID details) - not uploaded, nothing to validate.")
        return res
    if res.category == "UNKNOWN":
        raise ValueError(
            "Could not identify the DCT template type.\n"
            "Headers found: " + ", ".join(map(str, raw.columns))[:400])
    if df.empty:
        raise ValueError("No data rows below the header.")

    issues: list[Issue] = []
    check_common(df, cols, res.category, issues, res.sign_flags)
    dup_index = check_duplicates(df, cols, res.category, issues)

    if res.category in LOAN_CATEGORIES:
        check_loans(df, cols, res.category, issues, member_dates)
    elif res.category in DEPOSIT_CATEGORIES:
        check_deposits(df, cols, res.category, issues, member_dates)
    elif res.category == "MEMBERSHIP":
        check_membership(df, cols, issues)
    elif res.category == "LAND":
        check_land(df, cols, issues, member_dates)
    elif res.category == "COLLATERAL":
        check_collateral(df, cols, issues, loan_index)
    elif res.category == "TRANSACTIONS":
        check_transactions(df, cols, issues, loan_index)

    if progress:
        progress(res.total, res.total)

    for iss in issues:
        res.rule_counts[iss.rule] = res.rule_counts.get(iss.rule, 0) + 1

    # --- split by severity ---------------------------------------------------
    hard_by_row: dict[int, list[Issue]] = {}
    review_rows = []
    for iss in issues:
        if iss.severity == INFO:
            continue
        if iss.severity == HARD and iss.row is not None:
            hard_by_row.setdefault(iss.row, []).append(iss)
        else:
            review_rows.append(iss)

    hard_index = set(hard_by_row)

    rejected = df.loc[sorted(hard_index)].copy() if hard_index else pd.DataFrame()
    if not rejected.empty:
        rejected.insert(0, "SourceRow", [i + 2 for i in rejected.index])
        rejected.insert(1, "RuleIDs",
                        [", ".join(sorted({x.rule for x in hard_by_row[i]}))
                         for i in rejected.index])
        rejected.insert(2, "RejectionReasons",
                        [" | ".join(
                            f"{x.field}: {RULE_CATALOG[x.rule]['title']}"
                            + (f" (found {_fmt(x.value)}"
                               + (f", expected {_fmt(x.expected)}" if x.expected is not None else "")
                               + ")" if x.value is not None or x.expected is not None else "")
                            + (f" - {x.detail}" if x.detail else "")
                            for x in hard_by_row[i])
                         for i in rejected.index])
    res.rejected = rejected

    review_records = []
    for iss in review_rows:
        info = RULE_CATALOG.get(iss.rule, {})
        record = {
            "SourceRow": (iss.row + 2) if iss.row is not None else "(whole column)",
            "RuleID": iss.rule,
            "Severity": iss.severity,
            "Field": iss.field,
            "Found": _fmt(iss.value),
            "Expected": _fmt(iss.expected),
            "Problem": info.get("title", iss.rule),
            "Detail": iss.detail,
        }
        for key in ("AdmissionNo", "LoanNo", "AccountNo", "ProductDescription"):
            if key in cols:
                record[key] = (df.at[iss.row, key] if iss.row is not None else "")
        review_records.append(record)
    res.review = pd.DataFrame(review_records)

    clean_index = [i for i in df.index if i not in hard_index]
    clean_raw = df.loc[clean_index].copy()
    res.clean, res.missing_template_cols = format_clean(clean_raw, cols, res.category)

    dups = df.loc[sorted(dup_index)].copy() if dup_index else pd.DataFrame()
    if not dups.empty:
        dups.insert(0, "SourceRow", [i + 2 for i in dups.index])
    res.duplicates = dups

    return res


# ==============================================================================
# 11. CROSS-FILE INDEXES
# ==============================================================================

def build_indexes(paths: list[Path], log=print) -> tuple[dict, dict, list[str]]:
    """Membership AdmissionNo -> AdmissionDate, and LoanNo -> [loan owners].

    Built before per-file validation so cross-file checks can run. A LoanNo can
    legitimately repeat across products, so every candidate is kept - never
    overwritten by whichever file was read last.
    """
    member_dates: dict[str, date | None] = {}
    loan_index: dict[str, list[dict]] = {}
    notes: list[str] = []

    for path in paths:
        try:
            raw, _ = read_any(path)
            df, _, _ = map_headers(raw)
            cols = set(df.columns)
            cat, _how = detect_category(path, cols)

            if cat == "MEMBERSHIP" and "AdmissionNo" in cols:
                for _, row in df.iterrows():
                    key = digits_only(row.get("AdmissionNo"))
                    if not key:
                        continue
                    adm = parse_date(row.get("AdmissionDate"))[0] if "AdmissionDate" in cols else None
                    member_dates.setdefault(key, adm)
                notes.append(f"{path.name}: {len(member_dates):,} member(s) indexed")

            elif cat in LOAN_CATEGORIES and "LoanNo" in cols:
                count = 0
                for _, row in df.iterrows():
                    key = digits_only(row.get("LoanNo"))
                    if not key:
                        continue
                    loan_index.setdefault(key, []).append({
                        "product": row.get("ProductDescription", ""),
                        "admission_no": digits_only(row.get("AdmissionNo")),
                        "sanctioned": parse_date(row.get("SanctionedDate"))[0]
                        if "SanctionedDate" in cols else None,
                        "category": cat,
                    })
                    count += 1
                notes.append(f"{path.name}: {count:,} loan account(s) indexed ({cat})")
        except Exception as exc:                        # noqa: BLE001
            notes.append(f"{path.name}: could not index ({exc})")

    collisions = sum(1 for v in loan_index.values() if len(v) > 1)
    if collisions:
        notes.append(f"{collisions:,} LoanNo value(s) appear in more than one "
                     f"loan master - disambiguated by AdmissionNo where possible")
    return member_dates, loan_index, notes


# ==============================================================================
# 12. REPORT WRITING
# ==============================================================================

HDR = PatternFill("solid", fgColor="1F4E78")
ERRF = PatternFill("solid", fgColor="FFC7CE")
REVF = PatternFill("solid", fgColor="FFEB9C")
OKF = PatternFill("solid", fgColor="C6EFCE")


def _style(ws, header_fill=None, wrap_cols=()):
    if ws.max_row < 1:
        return
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill or HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    sample = min(ws.max_row, 200)
    for idx, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=sample), 1):
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 10), 55)
    for letter in wrap_cols:
        ws.column_dimensions[letter].width = 70
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.column_letter == letter:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")


def explanation_frame(rule_counts: dict[str, int]) -> pd.DataFrame:
    rows = []
    order = {HARD: 0, REVIEW: 1, INFO: 2}
    for rule_id, count in sorted(
            rule_counts.items(),
            key=lambda kv: (order.get(RULE_CATALOG.get(kv[0], {}).get("sev", REVIEW), 3),
                            -kv[1])):
        info = RULE_CATALOG.get(rule_id, {})
        rows.append({
            "Rule ID": rule_id,
            "Severity": info.get("sev", REVIEW),
            "Rows affected": count,
            "Problem": info.get("title", rule_id),
            "What it means": info.get("meaning", ""),
            "Why it matters": info.get("why", ""),
            "How to fix it": info.get("fix", ""),
        })
    if not rows:
        rows = [{"Rule ID": "-", "Severity": "-", "Rows affected": 0,
                 "Problem": "No rule was triggered",
                 "What it means": "Every row passed every check.",
                 "Why it matters": "",
                 "How to fix it": "Nothing to do. Clean Data is ready to upload."}]
    return pd.DataFrame(rows)


def summary_frame(res: FileResult) -> pd.DataFrame:
    rows = [
        ("DCT VALIDATION SUMMARY", ""),
        ("Tool", f"{APP_NAME} v{APP_VERSION}"),
        ("Generated", datetime.now().strftime("%d-%m-%Y %H:%M:%S")),
        ("Source file", res.name),
        ("Template detected", f"{res.category} - {CATEGORY_LABELS.get(res.category, '')}"),
        ("Detected by", res.detected_by),
        ("Format detected", res.format_note),
        ("", ""),
        ("RECORD COUNTS", ""),
        ("Total input records", res.total),
        ("Clean records (ready to upload)", res.n_clean),
        ("Rejected records (portal will refuse)", res.n_rejected),
        ("Review items (check, but not blocking)", res.n_review),
        ("Duplicate records", len(res.duplicates)),
        ("Pass rate", f"{res.pass_pct:.1f}%"),
        ("", ""),
    ]
    if res.sign_flags:
        rows.append(("FILE-LEVEL OBSERVATIONS", ""))
        for col, neg, total, share in res.sign_flags:
            rows.append((f"{col} sign convention",
                         f"{neg:,} of {total:,} non-zero values ({share*100:.1f}%) "
                         f"are negative - treated as an export convention, not "
                         f"{neg:,} separate errors"))
        rows.append(("", ""))

    rows.append(("RULES TRIGGERED", "Rows affected"))
    order = {HARD: 0, REVIEW: 1, INFO: 2}
    for rule_id, count in sorted(
            res.rule_counts.items(),
            key=lambda kv: (order.get(RULE_CATALOG.get(kv[0], {}).get("sev", REVIEW), 3), -kv[1])):
        info = RULE_CATALOG.get(rule_id, {})
        rows.append((f"[{info.get('sev','')}] {rule_id}  {info.get('title', rule_id)}", count))

    rows += [("", ""), ("TEMPLATE SHAPE", "")]
    template = TEMPLATES.get(res.category) or []
    rows.append(("Template columns", len(template) if template else "source order kept"))
    if res.missing_template_cols:
        rows.append(("Missing from source (added blank)",
                     ", ".join(res.missing_template_cols)))
    if res.unmapped_columns:
        rows.append(("Extra columns kept after the template",
                     ", ".join(res.unmapped_columns[:25])))

    rows += [
        ("", ""),
        ("ACTIVE THRESHOLDS", ""),
        ("ROI / PenalROI", f"greater than 0, less than {ROI_MAX:g}"),
        ("LoanPeriod", "must divide evenly by the repayment frequency months"),
        ("Portal frequency spelling", "Monthly / Quarterly / Halfyearly / Yearly"),
        ("FirstInstallmentDate", "SanctionedDate + one cycle (+ gestation on MT/LT)"),
        ("Negative sign convention", f"{RULE_SETTINGS['negative_sign_convention_share']:.0%} of non-zero values"),
        ("Date consistency", f"{RULE_SETTINGS['date_tolerance_days']:g} day tolerance"),
        ("InstallmentAmount", f"SanctionedAmount / instalment count, {RULE_SETTINGS['installment_amount_tolerance_pct']:.0%} tolerance"),
        ("MaturityAmount", f"DepositAmount + InterestAmount, {RULE_SETTINGS['deposit_amount_tolerance_pct']:.0%} tolerance"),
        ("Deposit rate", f"0 to {DEPOSIT_RATE_MAX:g}% flagged for review outside this"),
        ("", ""),
        ("NOTE", "The source file was opened read-only and not modified. No "
                 "account number, amount, date or rate was changed. Clean Data "
                 "is re-presented in template order only - dates as DD-MM-YYYY, "
                 "amounts as numbers, identifiers as text. Nothing is recalculated."),
    ]
    return pd.DataFrame(rows, columns=["Item", "Value"])


def write_file_report(res: FileResult, out_dir: Path) -> Path:
    out = out_dir / f"{Path(res.name).stem}__VALIDATED.xlsx"

    def or_note(df, note):
        return df if not df.empty else pd.DataFrame({"Info": [note]})

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        or_note(res.clean, "No clean records.").to_excel(
            writer, sheet_name="Clean Data", index=False)
        or_note(res.rejected, "No rejected rows.").to_excel(
            writer, sheet_name="Rejected Rows", index=False)
        or_note(res.review, "No review items.").to_excel(
            writer, sheet_name="Review Items", index=False)
        or_note(res.duplicates, "No duplicates found.").to_excel(
            writer, sheet_name="Duplicates", index=False)
        explanation_frame(res.rule_counts).to_excel(
            writer, sheet_name="Error Explanation", index=False)
        summary_frame(res).to_excel(writer, sheet_name="Summary", index=False)

        wb = writer.book
        _style(wb["Clean Data"], OKF if res.n_clean else None)
        _style(wb["Rejected Rows"], ERRF if res.n_rejected else None, wrap_cols=("C",))
        _style(wb["Review Items"], REVF if res.n_review else None, wrap_cols=("H",))
        _style(wb["Duplicates"])
        _style(wb["Error Explanation"], wrap_cols=("E", "F", "G"))
        _style(wb["Summary"])

        ws = wb["Error Explanation"]
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 42

        ws = wb["Summary"]
        ws.column_dimensions["A"].width = 48
        ws.column_dimensions["B"].width = 70
        for row in ws.iter_rows(min_row=2, max_col=2):
            if isinstance(row[0].value, str) and row[0].value and row[0].value.isupper():
                row[0].font = Font(bold=True)
            if row[1].value and isinstance(row[1].value, str) and len(row[1].value) > 60:
                row[1].alignment = Alignment(wrap_text=True, vertical="top")

    res.output_path = out
    return out


def _family_frame(results: list[FileResult], family: str) -> pd.DataFrame:
    rows = []
    for r in results:
        if r.family != family or r.skipped_reason:
            continue
        top = ""
        hard = {k: v for k, v in r.rule_counts.items()
                if RULE_CATALOG.get(k, {}).get("sev") == HARD}
        if hard:
            worst = max(hard, key=hard.get)
            top = f"{worst} {RULE_CATALOG[worst]['title']} ({hard[worst]:,})"
        rows.append({
            "File": r.name,
            "Template": r.category,
            "Total records": r.total,
            "Clean": r.n_clean,
            "Rejected": r.n_rejected,
            "Review items": r.n_review,
            "Pass rate": f"{r.pass_pct:.1f}%",
            "Top blocking problem": top or "None",
            "Report file": r.output_path.name if r.output_path else "",
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        [{"File": f"No {family.lower()} files found in this folder."}])


def _family_totals(results: list[FileResult], family: str) -> pd.DataFrame:
    subset = [r for r in results if r.family == family and not r.skipped_reason]
    t = sum(r.total for r in subset)
    c = sum(r.n_clean for r in subset)
    rej = sum(r.n_rejected for r in subset)
    rev = sum(r.n_review for r in subset)
    return pd.DataFrame(
        [("Files validated", len(subset)),
         ("Total input records", t),
         ("Clean records (ready to upload)", c),
         ("Rejected records", rej),
         ("Review items", rev),
         ("Pass rate", f"{(c / t * 100) if t else 0:.1f}%")],
        columns=["Item", "Value"])


def _family_rules(results: list[FileResult], family: str) -> dict:
    merged: dict[str, int] = {}
    for r in results:
        if r.family != family or r.skipped_reason:
            continue
        for rid, count in r.rule_counts.items():
            merged[rid] = merged.get(rid, 0) + count
    return merged


FAMILIES = ["Loans", "Deposits", "Transactions", "Membership & Others"]


def write_consolidated(results, failures, out_dir: Path, folder: Path,
                       index_notes: list[str]) -> Path:
    out = out_dir / "_VALIDATION_SUMMARY.xlsx"
    live = [r for r in results if not r.skipped_reason]

    overview = [
        ("DCT VALIDATION RUN", ""),
        ("Tool", f"{APP_NAME} v{APP_VERSION}"),
        ("Generated", datetime.now().strftime("%d-%m-%Y %H:%M:%S")),
        ("Folder validated", str(folder)),
        ("Output folder", out_dir.name),
        ("", ""),
        ("FILE COUNTS", ""),
        ("Files found", len(results) + len(failures)),
        ("Files validated", len(live)),
        ("Reference files skipped", sum(1 for r in results if r.skipped_reason)),
        ("Files that could not be read", len(failures)),
        ("", ""),
        ("RECORD TOTALS", ""),
        ("Total records", sum(r.total for r in live)),
        ("Clean records (ready to upload)", sum(r.n_clean for r in live)),
        ("Rejected records", sum(r.n_rejected for r in live)),
        ("Review items", sum(r.n_review for r in live)),
        ("", ""),
        ("TEMPLATES FOUND", ""),
    ]
    for r in sorted(live, key=lambda x: x.category):
        overview.append((r.name, f"{r.category} - {CATEGORY_LABELS.get(r.category, '')}"))
    overview += [
        ("", ""),
        ("NOTE", "Source files were opened read-only and not modified."),
    ]

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(overview, columns=["Item", "Value"]).to_excel(
            writer, sheet_name="Run Overview", index=False)

        for family in FAMILIES:
            short = {"Membership & Others": "Others"}.get(family, family)
            _family_frame(results, family).to_excel(
                writer, sheet_name=short, index=False)
            explanation_frame(_family_rules(results, family)).to_excel(
                writer, sheet_name=f"{short} Explanation"[:31], index=False)
            _family_totals(results, family).to_excel(
                writer, sheet_name=f"{short} Summary"[:31], index=False)

        skipped = [{"File": r.name, "Why": r.skipped_reason}
                   for r in results if r.skipped_reason]
        skipped += [{"File": r.name, "Why": r.failure} for r in failures]
        pd.DataFrame(skipped or [{"File": "None", "Why": "-"}]).to_excel(
            writer, sheet_name="Skipped & Failed", index=False)
        pd.DataFrame({"Cross-file index log": index_notes or ["-"]}).to_excel(
            writer, sheet_name="Index Log", index=False)

        wb = writer.book
        for name in wb.sheetnames:
            if name.endswith("Explanation"):
                _style(wb[name], wrap_cols=("E", "F", "G"))
                wb[name].column_dimensions["A"].width = 10
                wb[name].column_dimensions["B"].width = 14
                wb[name].column_dimensions["C"].width = 14
                wb[name].column_dimensions["D"].width = 42
            else:
                _style(wb[name], wrap_cols=("B",) if name.endswith(
                    ("Overview", "Summary", "Failed", "Log")) else ())

        for name in wb.sheetnames:
            if name.endswith(("Overview", "Summary")):
                ws = wb[name]
                ws.column_dimensions["A"].width = 44
                ws.column_dimensions["B"].width = 66
                for row in ws.iter_rows(min_row=2, max_col=1):
                    if isinstance(row[0].value, str) and row[0].value and row[0].value.isupper():
                        row[0].font = Font(bold=True)
    return out


# ==============================================================================
# 13. FOLDER RUN
# ==============================================================================

def discover(folder: Path, recursive: bool) -> list[Path]:
    it = folder.rglob("*") if recursive else folder.glob("*")
    found = []
    for p in it:
        if not p.is_file() or p.suffix.lower() not in DATA_SUFFIXES:
            continue
        if p.name.startswith("~$"):
            continue
        rel = str(p.relative_to(folder))
        if "Validated_" in rel:
            continue
        if p.name.endswith("__VALIDATED.xlsx") or p.name == "_VALIDATION_SUMMARY.xlsx":
            continue
        found.append(p)
    return sorted(found)


def run_folder(folder: Path, recursive: bool = False, log=print, progress=None):
    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"Not a folder: {folder}")
    files = discover(folder, recursive)
    if not files:
        raise FileNotFoundError(f"No data files found in {folder}")

    stamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    out_dir = folder / f"Validated_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=False)
    log(f"Output folder: {out_dir.name}")

    log("Pass 1 of 2 - building membership and loan indexes...")
    member_dates, loan_index, index_notes = build_indexes(files, log)
    log(f"  {len(member_dates):,} member(s), {len(loan_index):,} loan account(s) indexed.")

    log("Pass 2 of 2 - validating each template...")
    results: list[FileResult] = []
    failures: list[FileResult] = []

    for n, path in enumerate(files, 1):
        log(f"[{n}/{len(files)}] {path.name}")
        try:
            res = validate_one(path, member_dates, loan_index, progress)
            if res.skipped_reason:
                log(f"    {res.category}: {res.skipped_reason}")
                results.append(res)
                continue
            write_file_report(res, out_dir)
            results.append(res)
            log(f"    {res.category} ({CATEGORY_LABELS.get(res.category,'')}): "
                f"{res.total:,} rows | clean {res.n_clean:,} | "
                f"rejected {res.n_rejected:,} | review {res.n_review:,} "
                f"({res.pass_pct:.1f}% pass)")
            for col, neg, total, share in res.sign_flags:
                log(f"    note: {col} is negative on {neg:,}/{total:,} rows "
                    f"({share*100:.0f}%) - flagged once as a sign convention")
        except Exception as exc:                        # noqa: BLE001
            bad = FileResult()
            bad.path, bad.name, bad.failure = path, path.name, str(exc)
            failures.append(bad)
            log(f"    COULD NOT READ: {exc}")

    summary_path = write_consolidated(results, failures, out_dir, folder, index_notes)
    log(f"Consolidated summary: {summary_path.name}")
    return out_dir, results, failures


def open_folder(target: Path):
    try:
        os.startfile(target)                            # Windows
    except AttributeError:
        subprocess.run(["xdg-open", str(target)], check=False)


# ==============================================================================
# 14. CLI
# ==============================================================================

def run_cli(argv) -> int:
    ap = argparse.ArgumentParser(description=f"{APP_NAME} v{APP_VERSION}")
    ap.add_argument("folder", help="Folder containing the DCT templates")
    ap.add_argument("--recursive", action="store_true", help="Include sub-folders")
    args = ap.parse_args(argv)
    try:
        out_dir, results, failures = run_folder(Path(args.folder), args.recursive)
    except Exception as exc:                            # noqa: BLE001
        print(f"FAILED: {exc}", file=sys.stderr)
        return 1
    live = [r for r in results if not r.skipped_reason]
    print("=" * 74)
    print(f"  Files validated : {len(live)}   (skipped {len(results) - len(live)}, "
          f"unreadable {len(failures)})")
    print(f"  Total records   : {sum(r.total for r in live):,}")
    print(f"  Clean           : {sum(r.n_clean for r in live):,}")
    print(f"  Rejected        : {sum(r.n_rejected for r in live):,}")
    print(f"  Review items    : {sum(r.n_review for r in live):,}")
    print(f"  Output          : {out_dir}")
    print("=" * 74)
    return 0


# ==============================================================================
# 15. GUI
# ==============================================================================

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"{APP_NAME} v{APP_VERSION}")
    root.geometry("880x640")
    root.minsize(820, 600)

    folder_var = tk.StringVar()
    recursive_var = tk.BooleanVar(value=False)
    open_done_var = tk.BooleanVar(value=True)
    last_out: list[Path] = []

    frm = ttk.Frame(root, padding=14)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text=APP_NAME, font=("Segoe UI", 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(frm, text="Validates every DCT template in a folder - loans, deposits, "
                        "membership, land, collateral, transactions and agents. Reports go "
                        "to a new Validated_<date>_<time> folder inside it. Source files "
                        "are never modified.",
              foreground="#555", wraplength=820, justify="left").grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(2, 12))

    ttk.Label(frm, text="Data folder").grid(row=2, column=0, sticky="w", pady=4)
    entry = ttk.Entry(frm, textvariable=folder_var)
    entry.grid(row=2, column=1, sticky="ew", padx=8, pady=4)

    found_var = tk.StringVar(value="No folder selected.")

    def refresh():
        f = folder_var.get().strip()
        if not f:
            found_var.set("No folder selected.")
            return
        try:
            files = discover(Path(f), recursive_var.get())
            found_var.set(f"{len(files)} data file(s) found"
                          + (f": {', '.join(p.name for p in files[:4])}"
                             + (" ..." if len(files) > 4 else "") if files else ""))
        except Exception as exc:                        # noqa: BLE001
            found_var.set(f"Could not read folder: {exc}")

    def pick():
        d = filedialog.askdirectory(title="Select the folder containing DCT templates")
        if d:
            folder_var.set(d)
            refresh()

    ttk.Button(frm, text="Select folder...", command=pick).grid(row=2, column=2, pady=4)
    ttk.Label(frm, textvariable=found_var, foreground="#777",
              wraplength=820, justify="left").grid(
        row=3, column=1, columnspan=2, sticky="w", padx=8)

    opts = ttk.LabelFrame(frm, text="Options", padding=10)
    opts.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(14, 6))
    ttk.Checkbutton(opts, text="Include sub-folders", variable=recursive_var,
                    command=refresh).grid(row=0, column=0, sticky="w", padx=6)
    ttk.Checkbutton(opts, text="Open the output folder when finished",
                    variable=open_done_var).grid(row=1, column=0, sticky="w", padx=6)
    ttk.Label(opts, text="Template type is detected automatically from the column "
                         "headers, with the file name as a fallback.",
              foreground="#777").grid(row=2, column=0, sticky="w", padx=6, pady=(6, 0))

    bar = ttk.Progressbar(frm, mode="determinate")
    bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 4))

    status = tk.StringVar(value="Ready.")
    ttk.Label(frm, textvariable=status, foreground="#1F4E78").grid(
        row=6, column=0, columnspan=3, sticky="w")

    log_box = tk.Text(frm, height=16, wrap="word", font=("Consolas", 9))
    log_box.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(6, 8))
    log_box.configure(state="disabled")
    frm.rowconfigure(7, weight=1)

    def say(msg):
        log_box.configure(state="normal")
        log_box.insert("end", str(msg) + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")

    actions = ttk.Frame(frm)
    actions.grid(row=8, column=0, columnspan=3, sticky="ew")
    start_btn = ttk.Button(actions, text="Validate folder")
    start_btn.pack(side="left")

    def open_out():
        if not last_out:
            messagebox.showinfo("Nothing yet", "Run a validation first.")
            return
        open_folder(last_out[-1])

    ttk.Button(actions, text="Open output folder", command=open_out).pack(side="left", padx=8)
    ttk.Button(actions, text="Close", command=root.destroy).pack(side="right")

    def worker():
        folder = Path(folder_var.get().strip())
        try:
            def gui_log(m):
                root.after(0, say, m)

            def gui_progress(done, total):
                root.after(0, bar.configure, {"maximum": total, "value": done})

            root.after(0, status.set, "Validating...")
            out_dir, results, failures = run_folder(
                folder, recursive_var.get(), gui_log, gui_progress)
            last_out.append(out_dir)
            live = [r for r in results if not r.skipped_reason]
            root.after(0, say, "-" * 66)
            root.after(0, say,
                       f"FINISHED\n"
                       f"  files validated : {len(live)} "
                       f"(skipped {len(results)-len(live)}, unreadable {len(failures)})\n"
                       f"  total records   : {sum(r.total for r in live):,}\n"
                       f"  clean records   : {sum(r.n_clean for r in live):,}\n"
                       f"  rejected        : {sum(r.n_rejected for r in live):,}\n"
                       f"  review items    : {sum(r.n_review for r in live):,}\n"
                       f"  output folder   : {out_dir.name}")
            root.after(0, status.set, f"Finished. Reports in {out_dir.name}")
            if open_done_var.get():
                root.after(0, open_folder, out_dir)
            root.after(0, lambda: messagebox.showinfo(
                "Validation complete",
                f"{len(live)} template(s) validated.\n"
                f"{sum(r.n_clean for r in live):,} of "
                f"{sum(r.total for r in live):,} records are clean.\n\n"
                f"Reports saved in:\n{out_dir}"))
        except Exception as exc:                        # noqa: BLE001
            root.after(0, say, f"FAILED: {exc}")
            root.after(0, status.set, "Failed.")
            root.after(0, lambda e=exc: messagebox.showerror("Validation failed", str(e)))
        finally:
            root.after(0, start_btn.configure, {"state": "normal"})

    def start():
        if not folder_var.get().strip():
            messagebox.showwarning("No folder", "Select a data folder first.")
            return
        start_btn.configure(state="disabled")
        bar.configure(value=0)
        say("=" * 66)
        threading.Thread(target=worker, daemon=True).start()

    start_btn.configure(command=start)
    entry.bind("<FocusOut>", lambda _e: refresh())
    root.mainloop()


# ==============================================================================
def main() -> int:
    if len(sys.argv) > 1:
        return run_cli(sys.argv[1:])
    try:
        run_gui()
    except Exception:                                   # noqa: BLE001
        traceback.print_exc()
        print("\nGUI could not start. Use the command line:\n"
              "  python dct_all_template_validator.py \"C:\\path\\to\\folder\"")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
