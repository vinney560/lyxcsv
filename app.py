from flask import Flask, render_template, request, send_file, jsonify, session
import pandas as pd
import io
import re
from datetime import datetime
import openpyxl
from werkzeug.security import check_password_hash
import warnings
import json
import hashlib
import os
from functools import wraps
import numpy as np
from scipy import stats
from decouple import config

# Suppress warnings
warnings.filterwarnings("ignore", category=UserWarning)

app = Flask(__name__)

# ===== CONFIGURATION FROM .env =====
app.secret_key = config('SECRET_KEY', default='test_your_secret_key')
app.config['MAX_CONTENT_LENGTH'] = int(config('MAX_CONTENT_LENGTH', default=100 * 1024 * 1024))
app.config['CURRENT_DF'] = None
app.config['CURRENT_ORIGINAL_DF'] = None
app.config['CLEANING_HISTORY'] = []

PASSWORD_HASH = config('PASSWORD_HASH', default='scrypt:32768:8:1$e7lR8kAjb5FO5UPO$7baf61e4819f40495e885fdab107595473232a79ae1d08a6cf7abe12518f1e83045e7cdc09b0736a783d0f12411e2a8cd53f8132498ff2653b07808f51d944e8')

# ===================== LOGIN DECORATOR =====================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            return jsonify({'error': 'Authentication required'}), 401
        return f(*args, **kwargs)
    return decorated_function

# ===================== EMAIL FORMATTER CLASS =====================
class EmailFormatter:
    """Smart email formatting and completion"""
    
    DOMAIN_PATTERNS = {
        'gmail': ['gmail', 'gmal', 'gmail.', 'google', 'googl', 'gamil'],
        'yahoo': ['yahoo', 'yaho', 'yahho', 'yhoo', 'yho'],
        'hotmail': ['hotmail', 'hotmal', 'hotmail.', 'outlook'],
        'outlook': ['outlook', 'outlok', 'outlook.', 'outl'],
        'aol': ['aol', 'aol.', 'american online'],
        'icloud': ['icloud', 'iclod', 'me.com'],
        'protonmail': ['protonmail', 'protomail', 'proton'],
        'live': ['live', 'live.', 'live.com'],
        'mail': ['mail', 'mail.', 'mail.com'],
        'gmx': ['gmx', 'gmx.', 'gmx.com']
    }
    
    DOMAIN_TYPOS = {
        'gmai.com': 'gmail.com',
        'gmal.com': 'gmail.com',
        'gamil.com': 'gmail.com',
        'yahooo.com': 'yahoo.com',
        'yaho.com': 'yahoo.com',
        'yahho.com': 'yahoo.com',
        'hotmial.com': 'hotmail.com',
        'hotmil.com': 'hotmail.com',
        'outlok.com': 'outlook.com',
        'outlk.com': 'outlook.com',
        'protonmil.com': 'protonmail.com',
        'protomail.com': 'protonmail.com'
    }
    
    @staticmethod
    def detect_domain_pattern(emails):
        domain_counts = {}
        for email in emails:
            if pd.isna(email) or email == '':
                continue
            email_str = str(email).lower().strip()
            
            if '@' in email_str:
                parts = email_str.split('@')
                if len(parts) == 2:
                    domain_part = parts[1]
                    domain_base = domain_part.split('.')[0] if '.' in domain_part else domain_part
                    domain_counts[domain_base] = domain_counts.get(domain_base, 0) + 1
            else:
                domain_counts['no_domain'] = domain_counts.get('no_domain', 0) + 1
        
        if not domain_counts:
            return None, 0
        
        most_common = max(domain_counts.items(), key=lambda x: x[1])
        return most_common[0], most_common[1] / len(emails) if len(emails) > 0 else 0
    
    @staticmethod
    def format_email(email, default_domain='gmail.com'):
        if pd.isna(email) or email == '':
            return ''
        
        email_str = str(email).strip()
        if not email_str:
            return ''
        
        email_str = email_str.replace(' ', '')
        
        for typo, correct in EmailFormatter.DOMAIN_TYPOS.items():
            if typo in email_str:
                email_str = email_str.replace(typo, correct)
        
        if '@' in email_str:
            parts = email_str.split('@')
            if len(parts) == 2:
                username, domain = parts[0], parts[1]
                
                if '.' not in domain:
                    for known_domain, patterns in EmailFormatter.DOMAIN_PATTERNS.items():
                        if domain in patterns:
                            return f"{username}@{known_domain}.com"
                    return f"{username}@{domain}.com"
                else:
                    domain_parts = domain.split('.')
                    if len(domain_parts) == 2:
                        domain_base, domain_ext = domain_parts
                        for known_domain, patterns in EmailFormatter.DOMAIN_PATTERNS.items():
                            if domain_base in patterns:
                                return f"{username}@{known_domain}.{domain_ext}"
                        if domain_ext.lower() not in ['com', 'org', 'net', 'edu', 'gov', 'io', 'co', 'uk']:
                            return f"{username}@{domain_base}.com"
                    return email_str
            return email_str
        
        username = email_str
        
        if '.' in username:
            parts = username.split('.')
            if len(parts) == 2:
                for known_domain, patterns in EmailFormatter.DOMAIN_PATTERNS.items():
                    if parts[1].lower() in patterns:
                        return f"{parts[0]}@{known_domain}.com"
        
        return f"{username}@{default_domain}"
    
    @staticmethod
    def format_column(df, column_name, default_domain='gmail.com'):
        if column_name not in df.columns:
            return df, {'error': f'Column "{column_name}" not found'}
        
        emails = df[column_name].tolist()
        common_domain, ratio = EmailFormatter.detect_domain_pattern(emails)
        
        if common_domain and common_domain != 'no_domain':
            for known_domain, patterns in EmailFormatter.DOMAIN_PATTERNS.items():
                if common_domain in patterns:
                    default_domain = f"{known_domain}.com"
                    break
            else:
                default_domain = f"{common_domain}.com" if '.' not in common_domain else common_domain
        
        formatted = []
        changes = 0
        for email in emails:
            formatted_email = EmailFormatter.format_email(email, default_domain)
            formatted.append(formatted_email)
            if str(email).strip() != str(formatted_email).strip():
                changes += 1
        
        df[column_name] = formatted
        
        return df, {
            'changes': changes,
            'total': len(emails),
            'default_domain': default_domain,
            'detected_domain': common_domain
        }

# ===================== DATA CLEANER CLASS =====================
class DataCleaner:
    """Comprehensive data cleaning class"""
    
    @staticmethod
    def get_cleaning_report(df, columns=None):
        if columns is None:
            columns = df.columns.tolist()
        else:
            columns = [col for col in columns if col in df.columns]
        
        report = {
            'total_rows': int(len(df)),
            'total_columns': int(len(columns)),
            'missing_values': {},
            'duplicate_rows': int(df.duplicated().sum()),
            'duplicate_examples': [],
            'empty_columns': [],
            'columns_with_issues': [],
            'data_types': {},
            'outliers': {},
            'unique_counts': {},
            'text_issues': {},
            'numeric_issues': {},
            'mixed_types': {},
            'suspicious_values': {},
            'cell_issues': {},
            'year_columns': [],
            'numeric_columns': [],
            'date_columns': [],
            'boolean_columns': [],
            'email_columns': []
        }
        
        for col in columns:
            if col not in df.columns:
                continue
            
            if df[col].dtype == 'datetime64[ns]':
                report['date_columns'].append(str(col))
                continue
            
            sample = df[col].dropna().astype(str).head(20)
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            email_count = sample.str.match(email_pattern).sum()
            partial_email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+$'
            partial_email_count = sample.str.match(partial_email_pattern).sum()
            
            if email_count > len(sample) * 0.3 or partial_email_count > len(sample) * 0.3:
                report['email_columns'].append(str(col))
                continue
            
            if df[col].dtype == 'object' or df[col].dtype == 'bool':
                sample = df[col].dropna().astype(str).head(20)
                bool_patterns = ['true', 'false', 'yes', 'no', 'y', 'n', 't', 'f', '1', '0']
                bool_count = sample.str.lower().str.strip().isin(bool_patterns).sum()
                if bool_count > len(sample) * 0.8:
                    pattern_counts = {}
                    for val in sample.str.lower().str.strip():
                        if val in bool_patterns:
                            pattern_counts[val] = pattern_counts.get(val, 0) + 1
                    
                    if pattern_counts:
                        dominant = max(pattern_counts.items(), key=lambda x: x[1])
                        if dominant[0] in ['yes', 'no']:
                            report['boolean_columns'].append(f"{col} (Yes/No)")
                        elif dominant[0] in ['true', 'false']:
                            report['boolean_columns'].append(f"{col} (True/False)")
                        elif dominant[0] in ['y', 'n']:
                            report['boolean_columns'].append(f"{col} (Y/N)")
                        elif dominant[0] in ['t', 'f']:
                            report['boolean_columns'].append(f"{col} (T/F)")
                        elif dominant[0] in ['1', '0']:
                            report['boolean_columns'].append(f"{col} (1/0)")
                        else:
                            report['boolean_columns'].append(f"{col} (Mixed)")
            
            if df[col].dtype in ['int64', 'float64']:
                non_null = df[col].dropna()
                if len(non_null) > 0:
                    year_values = non_null[(non_null >= 1900) & (non_null <= 2100)]
                    if len(year_values) > len(non_null) * 0.5:
                        report['year_columns'].append(str(col))
                    report['numeric_columns'].append(str(col))
        
        if report['duplicate_rows'] > 0:
            dup_mask = df.duplicated(keep=False)
            dup_df = df[dup_mask]
            for idx, row in dup_df.head(5).iterrows():
                row_data = {}
                for col in df.columns[:3]:
                    if col in df.columns:
                        val = row[col]
                        row_data[str(col)] = str(val) if pd.notna(val) else ''
                report['duplicate_examples'].append(row_data)
        
        for col in columns:
            if col not in df.columns:
                continue
            missing_count = df[col].isnull().sum()
            if df[col].dtype == 'object':
                empty_count = (df[col].astype(str).str.strip() == '').sum()
                missing_count += empty_count
            if missing_count > 0:
                report['missing_values'][col] = int(missing_count)
                missing_rows = df[df[col].isnull() | (df[col].astype(str).str.strip() == '')].index.tolist()
                report['cell_issues'][col] = report['cell_issues'].get(col, {})
                report['cell_issues'][col]['missing'] = [int(r) for r in missing_rows[:10]]
        
        for col in columns:
            if col not in df.columns:
                continue
            if df[col].isnull().all() or (df[col].dtype == 'object' and (df[col].str.strip() == '').all()):
                report['empty_columns'].append(str(col))
        
        for col in columns:
            if col not in df.columns:
                continue
            if df[col].dtype == 'object':
                numeric_pattern = r'^\s*[-+]?\s*[\d,]+(?:\.\d+)?\s*$'
                non_empty = df[col].astype(str).str.strip() != ''
                string_numbers = df[col].astype(str).str.match(numeric_pattern).fillna(False) & non_empty
                if string_numbers.sum() > 0 and string_numbers.sum() < non_empty.sum():
                    report['mixed_types'][col] = {
                        'string_numbers': int(string_numbers.sum()),
                        'total': int(len(df[col]))
                    }
                    numeric_rows = df[string_numbers].index.tolist()
                    report['cell_issues'][col] = report['cell_issues'].get(col, {})
                    report['cell_issues'][col]['numeric_strings'] = [int(r) for r in numeric_rows[:10]]
                
                weird_pattern = r'[^\x00-\x7F]'
                weird_chars = df[col].astype(str).str.contains(weird_pattern, na=False).sum()
                if weird_chars > 0:
                    report['text_issues'][col] = {
                        'weird_characters': int(weird_chars),
                        'examples': [str(x) for x in df[col].astype(str)[df[col].astype(str).str.contains(weird_pattern, na=False)].head(3).tolist()]
                    }
                    weird_rows = df[df[col].astype(str).str.contains(weird_pattern, na=False)].index.tolist()
                    report['cell_issues'][col] = report['cell_issues'].get(col, {})
                    report['cell_issues'][col]['weird_characters'] = [int(r) for r in weird_rows[:10]]
                
                spaces = df[col].astype(str).str.match(r'^\s+|\s+$', na=False).sum()
                if spaces > 0:
                    if col not in report['text_issues']:
                        report['text_issues'][col] = {}
                    report['text_issues'][col]['leading_trailing_spaces'] = int(spaces)
                    space_rows = df[df[col].astype(str).str.match(r'^\s+|\s+$', na=False)].index.tolist()
                    report['cell_issues'][col] = report['cell_issues'].get(col, {})
                    report['cell_issues'][col]['spaces'] = [int(r) for r in space_rows[:10]]
                
                empty_strings = (df[col].astype(str).str.strip() == '').sum()
                if empty_strings > 0:
                    if col not in report['text_issues']:
                        report['text_issues'][col] = {}
                    report['text_issues'][col]['empty_strings'] = int(empty_strings)
        
        for col in df.select_dtypes(include=[np.number]).columns:
            if col not in columns:
                continue
            if len(df[col].dropna()) > 0:
                Q1 = df[col].quantile(0.25)
                Q3 = df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
                if len(outliers) > 0:
                    report['outliers'][col] = int(len(outliers))
                    report['cell_issues'][col] = report['cell_issues'].get(col, {})
                    report['cell_issues'][col]['outliers'] = [int(r) for r in outliers.index.tolist()[:10]]
        
        for col in columns:
            if col in df.columns:
                report['unique_counts'][col] = int(df[col].nunique())
        
        return report
    
    @staticmethod
    def clean_numeric_strings(df, columns=None):
        if columns is None:
            columns = df.select_dtypes(include=['object']).columns
        
        changes = {}
        for col in columns:
            if col not in df.columns:
                continue
            
            try:
                if pd.api.types.is_numeric_dtype(df[col]):
                    continue
                
                non_empty = df[col].astype(str).str.strip() != ''
                non_null_values = df[col].dropna().astype(str)
                non_null_values = non_null_values[non_null_values.str.strip() != '']
                
                if len(non_null_values) == 0:
                    continue
                    
                sample = non_null_values.head(min(20, len(non_null_values)))
                numeric_pattern = r'^\s*[-+]?\s*[\d,]+(?:\.\d+)?\s*$'
                matches = sample.str.match(numeric_pattern).fillna(False)
                numeric_count = matches.sum()
                
                if numeric_count > len(sample) * 0.6:
                    original_col = df[col].copy()
                    
                    def extract_valid_number(s):
                        s = str(s).strip()
                        if not s or s == '':
                            return ""
                        sign = ""
                        if s and s[0] in '+-':
                            sign = s[0]
                            s = s[1:]
                        digits = []
                        decimal_added = False
                        for c in s:
                            if c.isdigit():
                                digits.append(c)
                            elif c == '.' and not decimal_added:
                                digits.append('.')
                                decimal_added = True
                        if not digits or digits == ['.']:
                            return ""
                        result = sign + ''.join(digits)
                        if result.endswith('.'):
                            result = result[:-1]
                        return result
                    
                    cleaned = original_col.apply(extract_valid_number)
                    numeric_series = pd.to_numeric(cleaned, errors='coerce')
                    
                    converted_count = numeric_series.notna().sum()
                    
                    if converted_count > 0 and converted_count > len(df) * 0.2:
                        mask = numeric_series.notna()
                        df.loc[mask, col] = numeric_series[mask]
                        if converted_count > len(df) * 0.8:
                            df[col] = pd.to_numeric(df[col], errors='ignore')
                            
                        changes[col] = {
                            'converted': int(converted_count),
                            'failed': int(len(df) - converted_count),
                            'preserved_original_values': int(len(df) - converted_count)
                        }
            except Exception as e:
                print(f"Warning: Failed to process column {col}: {str(e)}")
                pass
        
        return df, {'numeric_strings_cleaned': changes}
    
    @staticmethod
    def clean_weird_characters(df, columns=None):
        if columns is None:
            columns = df.select_dtypes(include=['object']).columns
        
        changes = {}
        for col in columns:
            if col not in df.columns:
                continue
            
            weird_pattern = r'[^\x00-\x7F]'
            before_weird = df[col].astype(str).str.contains(weird_pattern, na=False).sum()
            
            if before_weird > 0:
                df[col] = df[col].astype(str).str.replace(weird_pattern, '', regex=True)
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                df[col] = df[col].str.strip()
                
                changes[col] = {
                    'weird_removed': int(before_weird)
                }
        
        return df, {'weird_characters_cleaned': changes}
    
    @staticmethod
    def clean_leading_trailing_spaces(df, columns=None):
        if columns is None:
            columns = df.select_dtypes(include=['object']).columns
        
        changes = {}
        for col in columns:
            if col not in df.columns:
                continue
            
            before = df[col].astype(str).copy()
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            
            changed = (before != df[col]).sum()
            if changed > 0:
                changes[col] = int(changed)
        
        return df, {'spaces_cleaned': changes}
    
    @staticmethod
    def remove_duplicates(df):
        before = len(df)
        df = df.drop_duplicates(keep='first')
        after = len(df)
        return df, {'removed': int(before - after)}
    
    @staticmethod
    def remove_empty_columns(df):
        before = len(df.columns)
        df = df.dropna(axis=1, how='all')
        for col in df.columns:
            if df[col].dtype == 'object':
                if (df[col].str.strip() == '').all():
                    df = df.drop(columns=[col])
        after = len(df.columns)
        return df, {'removed': int(before - after)}
    
    @staticmethod
    def apply_all_cleaning(df, options=None, selected_columns=None):
        if options is None:
            options = {
                'clean_numeric_strings': True,
                'clean_weird_characters': True,
                'clean_spaces': True,
                'remove_duplicates': True,
                'remove_empty_columns': True,
                'fix_dates': True,
                'format_emails': True
            }
        
        clean_log = {
            'steps': [],
            'removed_duplicates': 0,
            'removed_empty_cols': 0,
            'numeric_strings_cleaned': {},
            'weird_characters_cleaned': {},
            'spaces_cleaned': {},
            'date_columns_fixed': 0,
            'emails_formatted': 0
        }
        
        df_clean = df.copy()
        
        if selected_columns:
            numeric_cols = [col for col in selected_columns if col in df_clean.columns and df_clean[col].dtype == 'object']
            text_cols = [col for col in selected_columns if col in df_clean.columns and df_clean[col].dtype == 'object']
            date_cols = [col for col in selected_columns if col in df_clean.columns and df_clean[col].dtype == 'object']
            email_cols = [col for col in selected_columns if col in df_clean.columns and df_clean[col].dtype == 'object']
        else:
            numeric_cols = df_clean.select_dtypes(include=['object']).columns.tolist()
            text_cols = df_clean.select_dtypes(include=['object']).columns.tolist()
            date_cols = df_clean.select_dtypes(include=['object']).columns.tolist()
            email_cols = df_clean.select_dtypes(include=['object']).columns.tolist()
        
        if options.get('clean_numeric_strings', True) and numeric_cols:
            df_clean, result = DataCleaner.clean_numeric_strings(df_clean, numeric_cols)
            if result.get('numeric_strings_cleaned'):
                clean_log['steps'].append('Cleaned numeric strings')
                clean_log['numeric_strings_cleaned'] = result['numeric_strings_cleaned']
        
        if options.get('clean_weird_characters', True) and text_cols:
            df_clean, result = DataCleaner.clean_weird_characters(df_clean, text_cols)
            if result.get('weird_characters_cleaned'):
                clean_log['steps'].append('Removed weird characters')
                clean_log['weird_characters_cleaned'] = result['weird_characters_cleaned']
        
        if options.get('clean_spaces', True) and text_cols:
            df_clean, result = DataCleaner.clean_leading_trailing_spaces(df_clean, text_cols)
            if result.get('spaces_cleaned'):
                clean_log['steps'].append('Cleaned spaces')
                clean_log['spaces_cleaned'] = result['spaces_cleaned']
        
        if options.get('format_emails', True):
            for col in df_clean.columns:
                if df_clean[col].dtype == 'object':
                    sample = df_clean[col].dropna().astype(str).head(20)
                    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                    partial_email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+$'
                    email_count = sample.str.match(email_pattern).sum()
                    partial_email_count = sample.str.match(partial_email_pattern).sum()
                    
                    if email_count > len(sample) * 0.2 or partial_email_count > len(sample) * 0.2:
                        if selected_columns is None or col in selected_columns:
                            df_clean, result = EmailFormatter.format_column(df_clean, col)
                            if result.get('changes', 0) > 0:
                                clean_log['steps'].append(f'Formatted emails in "{col}"')
                                clean_log['emails_formatted'] += result['changes']
        
        if options.get('remove_duplicates', True):
            df_clean, result = DataCleaner.remove_duplicates(df_clean)
            clean_log['steps'].append('Removed duplicates')
            clean_log['removed_duplicates'] = result.get('removed', 0)
        
        if options.get('remove_empty_columns', True):
            df_clean, result = DataCleaner.remove_empty_columns(df_clean)
            clean_log['steps'].append('Removed empty columns')
            clean_log['removed_empty_cols'] = result.get('removed', 0)
        
        if options.get('fix_dates', True) and date_cols:
            df_clean, result = DataCleaner.fix_dates(df_clean, date_cols)
            if result.get('date_columns_fixed', 0) > 0:
                clean_log['steps'].append('Fixed date columns')
                clean_log['date_columns_fixed'] = result['date_columns_fixed']
        
        return df_clean, clean_log
    
    @staticmethod
    def fix_dates(df, columns=None):
        if columns is None:
            columns = df.select_dtypes(include=['object']).columns
        fixed = 0
        for col in columns:
            if col not in df.columns:
                continue
            try:
                sample = df[col].dropna().head(10).astype(str)
                date_patterns = [
                    r'^\d{4}-\d{2}-\d{2}',
                    r'^\d{2}/\d{2}/\d{4}',
                    r'^\d{2}-\d{2}-\d{4}',
                    r'^\d{4}/\d{2}/\d{2}',
                ]
                is_date = False
                for pattern in date_patterns:
                    if sample.str.match(pattern).sum() > len(sample) * 0.8:
                        is_date = True
                        break
                
                if is_date:
                    continue
                
                numeric_test = pd.to_numeric(df[col], errors='coerce')
                if numeric_test.notna().sum() > len(df) * 0.8:
                    year_values = numeric_test[(numeric_test >= 1900) & (numeric_test <= 2100)]
                    if len(year_values) > len(df) * 0.5:
                        continue
                
                temp = pd.to_datetime(df[col], errors='coerce')
                if temp.notna().sum() > len(df) * 0.5:
                    df[col] = temp
                    fixed += 1
            except:
                pass
        return df, {'date_columns_fixed': int(fixed)}

# ===================== DATA FORMATTER CLASS =====================
class DataFormatter:
    @staticmethod
    def clean_headers(df):
        df.columns = [
            str(col).strip().replace('_', ' ').title() 
            for col in df.columns
        ]
        return df
    
    @staticmethod
    def detect_and_convert_types(df):
        for col in df.columns:
            if df[col].isnull().all():
                continue
            
            sample = df[col].dropna().head(10).astype(str)
            date_patterns = [
                r'^\d{4}-\d{2}-\d{2}',
                r'^\d{2}/\d{2}/\d{4}',
                r'^\d{2}-\d{2}-\d{4}',
                r'^\d{4}/\d{2}/\d{2}',
                r'^\d{2}\.\d{2}\.\d{4}',
                r'^\d{4}\.\d{2}\.\d{2}',
            ]
            is_date = False
            for pattern in date_patterns:
                if sample.str.match(pattern).sum() > len(sample) * 0.8:
                    is_date = True
                    break
            
            if is_date:
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    if df[col].dtype == 'datetime64[ns]':
                        continue
                except:
                    pass
            
            if df[col].dtype == 'object':
                sample = df[col].dropna().astype(str).head(20)
                bool_patterns = ['true', 'false', 'yes', 'no', 'y', 'n', 't', 'f', '1', '0']
                bool_count = sample.str.lower().str.strip().isin(bool_patterns).sum()
                
                if bool_count > len(sample) * 0.8:
                    def convert_to_bool(val):
                        if pd.isna(val) or val == '':
                            return ''
                        val_str = str(val).lower().strip()
                        if val_str in ['true', 'yes', 'y', 't', '1']:
                            return 'Yes'
                        elif val_str in ['false', 'no', 'n', 'f', '0']:
                            return 'No'
                        return val
                    
                    pattern_counts = {}
                    for val in sample.str.lower().str.strip():
                        if val in bool_patterns:
                            pattern_counts[val] = pattern_counts.get(val, 0) + 1
                    
                    if pattern_counts:
                        dominant = max(pattern_counts.items(), key=lambda x: x[1])
                        if dominant[0] in ['yes', 'no']:
                            df[col] = df[col].apply(lambda x: 'Yes' if str(x).lower().strip() in ['yes', 'y', 't', '1', 'true'] else 'No' if pd.notna(x) else '')
                        elif dominant[0] in ['true', 'false']:
                            df[col] = df[col].apply(lambda x: 'True' if str(x).lower().strip() in ['true', 'y', 't', '1', 'yes'] else 'False' if pd.notna(x) else '')
                        elif dominant[0] in ['y', 'n']:
                            df[col] = df[col].apply(lambda x: 'Y' if str(x).lower().strip() in ['y', 'yes', 't', '1', 'true'] else 'N' if pd.notna(x) else '')
                        elif dominant[0] in ['t', 'f']:
                            df[col] = df[col].apply(lambda x: 'T' if str(x).lower().strip() in ['t', 'true', 'y', '1', 'yes'] else 'F' if pd.notna(x) else '')
                        elif dominant[0] in ['1', '0']:
                            df[col] = df[col].apply(lambda x: '1' if str(x).lower().strip() in ['1', 'yes', 'true', 'y', 't'] else '0' if pd.notna(x) else '')
                        continue
            
            try:
                numeric_test = pd.to_numeric(df[col], errors='coerce')
                if numeric_test.notna().sum() > len(df) * 0.8:
                    year_values = numeric_test[(numeric_test >= 1900) & (numeric_test <= 2100)]
                    if len(year_values) > len(df) * 0.5:
                        df[col] = numeric_test
                        continue
                    
                    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                    sample_str = df[col].dropna().head(10).astype(str)
                    if sample_str.str.match(email_pattern).sum() > len(sample_str) * 0.5:
                        continue
                    
                    df[col] = pd.to_numeric(df[col], errors='ignore')
                    if df[col].dtype in ['int64', 'float64']:
                        continue
            except:
                pass
            
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('nan', '')
            df[col] = df[col].replace('None', '')
            
        return df
    
    @staticmethod
    def format_numbers(df):
        for col in df.columns:
            if df[col].dtype in ['int64', 'float64']:
                non_null = df[col].dropna()
                if len(non_null) > 0:
                    year_values = non_null[(non_null >= 1900) & (non_null <= 2100)]
                    if len(year_values) > len(non_null) * 0.5:
                        df[col] = df[col].apply(
                            lambda x: f"{int(x)}" if pd.notna(x) and x != '' and x is not None else ''
                        )
                        continue
                
                if df[col].dtype == 'int64':
                    df[col] = df[col].apply(
                        lambda x: f"{int(x):,}" if pd.notna(x) and x != '' and x is not None else ''
                    )
                else:
                    df[col] = df[col].apply(
                        lambda x: f"{float(x):,.2f}" if pd.notna(x) and x != '' and x is not None else ''
                    )
        return df
    
    @staticmethod
    def format_dates(df):
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                has_time = False
                for val in df[col].dropna().head(10):
                    if val.hour != 0 or val.minute != 0 or val.second != 0:
                        has_time = True
                        break
                
                if has_time:
                    df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M')
                else:
                    df[col] = df[col].dt.strftime('%Y-%m-%d')
        return df
    
    @staticmethod
    def handle_nulls(df):
        df = df.fillna('')
        df = df.replace([np.nan, pd.NA, None], '', regex=False)
        return df
    
    @staticmethod
    def apply_all_formattings(df):
        df = df.copy()
        df = DataFormatter.clean_headers(df)
        df = DataFormatter.detect_and_convert_types(df)
        df = DataFormatter.format_dates(df)
        df = DataFormatter.format_numbers(df)
        df = DataFormatter.handle_nulls(df)
        return df

    @staticmethod
    def convert_to_serializable(df):
        df_copy = df.copy()
        for col in df_copy.columns:
            if df_copy[col].dtype == 'datetime64[ns]':
                has_time = False
                for val in df_copy[col].dropna().head(10):
                    if val.hour != 0 or val.minute != 0 or val.second != 0:
                        has_time = True
                        break
                
                if has_time:
                    df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d %H:%M')
                else:
                    df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d')
        return df_copy

    @staticmethod
    def get_column_info(df):
        info = []
        for col in df.columns:
            dtype = str(df[col].dtype)
            non_null = df[col].ne('').sum() if df[col].dtype == 'object' else df[col].count()
            
            if 'bool' in dtype:
                col_type = 'Boolean'
            elif 'datetime' in dtype:
                col_type = 'Date/Time'
            elif 'int' in dtype:
                col_type = 'Number (Integer)'
            elif 'float' in dtype:
                col_type = 'Number (Float)'
            else:
                col_type = 'Text'
            
            info.append({
                'name': str(col),
                'type': col_type,
                'non_null': int(non_null)
            })
        return info
    
    @staticmethod
    def search_data(df, search_term, case_sensitive=False):
        if not search_term or search_term.strip() == '':
            return df
        
        search_term = search_term.strip()
        
        if len(df) > 10000:
            mask = pd.Series(False, index=df.index)
            for col in df.columns:
                if df[col].dtype == 'object':
                    if case_sensitive:
                        mask |= df[col].astype(str).str.contains(search_term, na=False, regex=False)
                    else:
                        mask |= df[col].astype(str).str.contains(search_term, na=False, case=False, regex=False)
                elif df[col].dtype == 'datetime64[ns]':
                    str_col = df[col].dt.strftime('%Y-%m-%d %H:%M')
                    if case_sensitive:
                        mask |= str_col.str.contains(search_term, na=False, regex=False)
                    else:
                        mask |= str_col.str.contains(search_term, na=False, case=False, regex=False)
            return df[mask]
        else:
            def row_contains(row):
                for val in row:
                    if pd.isna(val):
                        continue
                    if isinstance(val, (datetime, pd.Timestamp)):
                        val = val.strftime('%Y-%m-%d %H:%M')
                    else:
                        val = str(val)
                    if case_sensitive:
                        if search_term in val:
                            return True
                    else:
                        if search_term.lower() in val.lower():
                            return True
                return False
            
            return df[df.apply(row_contains, axis=1)]

# ===================== BULK UPDATER CLASS =====================
class BulkUpdater:
    @staticmethod
    def find_and_replace(df, search_col, search_term, replace_term, case_sensitive=False, single_row=None):
        if search_col not in df.columns:
            return df, {'error': f'Column "{search_col}" not found'}
        
        if single_row is not None:
            if single_row < 0 or single_row >= len(df):
                return df, {'error': f'Row {single_row} out of range'}
            
            cell_value = str(df.at[single_row, search_col]) if pd.notna(df.at[single_row, search_col]) else ''
            if case_sensitive:
                if search_term not in cell_value:
                    return df, {'affected_rows': 0, 'message': 'No match found in this cell'}
                new_value = cell_value.replace(search_term, replace_term)
            else:
                if search_term.lower() not in cell_value.lower():
                    return df, {'affected_rows': 0, 'message': 'No match found in this cell'}
                pattern = re.compile(re.escape(search_term), re.IGNORECASE)
                new_value = pattern.sub(replace_term, cell_value)
            
            df.at[single_row, search_col] = new_value
            
            return df, {
                'affected_rows': 1,
                'total_rows': int(len(df)),
                'column': str(search_col),
                'search_term': str(search_term),
                'replace_term': str(replace_term),
                'affected_rows_list': [int(single_row)],
                'single_row': True
            }
        else:
            if case_sensitive:
                mask = df[search_col].astype(str).str.contains(search_term, na=False, regex=False)
            else:
                mask = df[search_col].astype(str).str.contains(search_term, na=False, case=False, regex=False)
            
            affected_rows = mask.sum()
            
            if affected_rows == 0:
                return df, {'affected_rows': 0, 'message': 'No matches found'}
            
            if case_sensitive:
                df.loc[mask, search_col] = df.loc[mask, search_col].astype(str).str.replace(
                    search_term, replace_term, regex=False
                )
            else:
                df.loc[mask, search_col] = df.loc[mask, search_col].astype(str).str.replace(
                    search_term, replace_term, case=False, regex=False
                )
            
            return df, {
                'affected_rows': int(affected_rows),
                'total_rows': int(len(df)),
                'column': str(search_col),
                'search_term': str(search_term),
                'replace_term': str(replace_term),
                'affected_rows_list': [int(idx) for idx in mask[mask].index.tolist()],
                'single_row': False
            }
    
    @staticmethod
    def find_all_occurrences(df, search_term, case_sensitive=False):
        results = []
        search_term = search_term.strip()
        
        if not search_term:
            return results
        
        for col in df.columns:
            try:
                if df[col].dtype == 'datetime64[ns]':
                    str_col = df[col].dt.strftime('%Y-%m-%d %H:%M')
                    if case_sensitive:
                        mask = str_col.str.contains(search_term, na=False, regex=False)
                    else:
                        mask = str_col.str.contains(search_term, na=False, case=False, regex=False)
                else:
                    if df[col].dtype == 'object':
                        if case_sensitive:
                            mask = df[col].astype(str).str.contains(search_term, na=False, regex=False)
                        else:
                            mask = df[col].astype(str).str.contains(search_term, na=False, case=False, regex=False)
                    else:
                        continue
                
                if mask.any():
                    occurrences = []
                    for idx, row in df[mask].iterrows():
                        cell_value = row[col]
                        if isinstance(cell_value, (datetime, pd.Timestamp)):
                            cell_value = cell_value.strftime('%Y-%m-%d %H:%M')
                        else:
                            cell_value = str(cell_value) if pd.notna(cell_value) else ''
                        
                        if case_sensitive:
                            positions = [m.start() for m in re.finditer(re.escape(search_term), cell_value)]
                        else:
                            pattern = re.compile(re.escape(search_term), re.IGNORECASE)
                            positions = [m.start() for m in pattern.finditer(cell_value)]
                        
                        occurrences.append({
                            'row': int(idx),
                            'col': str(col),
                            'value': cell_value,
                            'positions': [int(p) for p in positions]
                        })
                    
                    results.append({
                        'column': str(col),
                        'count': int(mask.sum()),
                        'occurrences': occurrences
                    })
            except:
                pass
        
        return results

# ===================== HELPER FUNCTIONS =====================
def process_file(file):
    filename = file.filename.lower()
    
    try:
        if filename.endswith('.csv'):
            encodings = ['utf-8', 'latin1', 'cp1252']
            df = None
            for encoding in encodings:
                try:
                    sample = file.read(1024).decode(encoding)
                    file.seek(0)
                    
                    if ',' in sample and ';' not in sample:
                        delimiter = ','
                    elif ';' in sample:
                        delimiter = ';'
                    elif '\t' in sample:
                        delimiter = '\t'
                    else:
                        delimiter = ','
                    
                    df = pd.read_csv(file, encoding=encoding, delimiter=delimiter)
                    break
                except:
                    file.seek(0)
                    continue
            
            if df is None:
                raise ValueError("Could not decode CSV file")
                
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file, engine='openpyxl')
        else:
            raise ValueError("Unsupported file format")
        
        app.config['CURRENT_ORIGINAL_DF'] = df.copy()
        app.config['CLEANING_HISTORY'] = []
        
        formatted_df = DataFormatter.apply_all_formattings(df)
        return formatted_df, None
        
    except Exception as e:
        return None, str(e)

# ===================== ROUTES =====================

@app.route('/auth', methods=['POST'])
def authenticate():
    data = request.get_json()
    password = data.get('password', '')
    
    try:
        if check_password_hash(PASSWORD_HASH, password):
            session['authenticated'] = True
            session.permanent = True
            return jsonify({'success': True, 'message': 'Authentication successful'})
        else:
            return jsonify({'success': False, 'message': 'Invalid password'}), 401
    except Exception as e:
        return jsonify({'success': False, 'message': f'Authentication error: {str(e)}'}), 500

@app.route('/logout', methods=['POST'])
@login_required
def logout():
    session.pop('authenticated', None)
    return jsonify({'success': True, 'message': 'Logged out successfully'})

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    formatted_df, error = process_file(file)
    
    if error:
        return jsonify({'error': error}), 400
    
    serializable_df = DataFormatter.convert_to_serializable(formatted_df)
    column_info = DataFormatter.get_column_info(serializable_df)
    
    table_html = serializable_df.to_html(
        classes='formatted-table',
        index=False,
        escape=False
    )
    
    app.config['CURRENT_DF'] = formatted_df
    
    return jsonify({
        'success': True,
        'table_html': table_html,
        'columns': column_info,
        'rows': int(len(serializable_df)),
        'cols': int(len(serializable_df.columns)),
        'filename': file.filename
    })

# ===================== CLEANING ROUTES =====================

@app.route('/cleaning-report', methods=['POST'])
@login_required
def get_cleaning_report():
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        data = request.get_json() or {}
        selected_columns = data.get('columns', None)
        
        report = DataCleaner.get_cleaning_report(df, selected_columns)
        report = json.loads(json.dumps(report, default=str))
        return jsonify({'success': True, 'report': report})
    except Exception as e:
        return jsonify({'error': f'Failed to generate report: {str(e)}'}), 400

@app.route('/clean-data', methods=['POST'])
@login_required
def clean_data():
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        data = request.get_json() or {}
        options = data.get('options', {})
        selected_columns = data.get('columns', None)
        
        app.config['CLEANING_HISTORY'].append(df.copy())
        
        cleaned_df, clean_log = DataCleaner.apply_all_cleaning(df, options, selected_columns)
        
        formatted_df = DataFormatter.apply_all_formattings(cleaned_df)
        app.config['CURRENT_DF'] = formatted_df
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        clean_log = json.loads(json.dumps(clean_log, default=str))
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'clean_log': clean_log,
            'message': f"Data cleaned! {len(clean_log['steps'])} steps applied."
        })
    except Exception as e:
        return jsonify({'error': f'Cleaning failed: {str(e)}'}), 400

@app.route('/undo-cleaning', methods=['POST'])
@login_required
def undo_cleaning():
    if not app.config.get('CLEANING_HISTORY'):
        return jsonify({'error': 'Nothing to undo'}), 400
    
    try:
        df = app.config['CLEANING_HISTORY'].pop()
        formatted_df = DataFormatter.apply_all_formattings(df)
        app.config['CURRENT_DF'] = formatted_df
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'message': 'Cleaning undone successfully!'
        })
    except Exception as e:
        return jsonify({'error': f'Undo failed: {str(e)}'}), 400

# ===================== BULK UPDATE ROUTES =====================

@app.route('/bulk-find', methods=['POST'])
@login_required
def bulk_find():
    data = request.get_json()
    search_term = data.get('search_term', '')
    case_sensitive = data.get('case_sensitive', False)
    
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        results = BulkUpdater.find_all_occurrences(df, search_term, case_sensitive)
        results = json.loads(json.dumps(results, default=str))
        total_matches = sum(r['count'] for r in results)
        
        return jsonify({
            'success': True,
            'results': results,
            'total_matches': int(total_matches),
            'columns_affected': int(len(results))
        })
    except Exception as e:
        return jsonify({'error': f'Find failed: {str(e)}'}), 400

@app.route('/bulk-replace', methods=['POST'])
@login_required
def bulk_replace():
    data = request.get_json()
    search_col = data.get('column')
    search_term = data.get('search_term', '')
    replace_term = data.get('replace_term', '')
    case_sensitive = data.get('case_sensitive', False)
    single_row = data.get('single_row', None)
    
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        app.config['CLEANING_HISTORY'].append(df.copy())
        
        updated_df, result = BulkUpdater.find_and_replace(
            df, search_col, search_term, replace_term, case_sensitive, single_row
        )
        
        if 'error' in result:
            return jsonify({'error': result['error']}), 400
        
        app.config['CURRENT_DF'] = updated_df
        
        formatted_df = DataFormatter.apply_all_formattings(updated_df)
        app.config['CURRENT_DF'] = formatted_df
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        result = json.loads(json.dumps(result, default=str))
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'result': result
        })
    except Exception as e:
        return jsonify({'error': f'Replace failed: {str(e)}'}), 400

# ===================== DETECT DUPLICATES ROUTE =====================

@app.route('/detect-duplicates', methods=['GET'])
@login_required
def detect_duplicates():
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        dup_mask = df.duplicated(keep=False)
        duplicate_count = dup_mask.sum()
        duplicate_rows = df[dup_mask].head(10).to_dict('records')
        
        for row in duplicate_rows:
            for key, value in row.items():
                if pd.isna(value):
                    row[key] = ''
                elif isinstance(value, (datetime, pd.Timestamp)):
                    row[key] = value.strftime('%Y-%m-%d %H:%M')
                else:
                    row[key] = str(value)
        
        return jsonify({
            'success': True,
            'duplicate_count': int(duplicate_count),
            'duplicate_rows': duplicate_rows
        })
    except Exception as e:
        return jsonify({'error': f'Failed to detect duplicates: {str(e)}'}), 400

# ===================== SEARCH ROUTE =====================

@app.route('/search', methods=['POST'])
@login_required
def search_data():
    data = request.get_json()
    search_term = data.get('search', '')
    case_sensitive = data.get('case_sensitive', False)
    
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        result_df = DataFormatter.search_data(df, search_term, case_sensitive)
        serializable_df = DataFormatter.convert_to_serializable(result_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'search_term': search_term
        })
    except Exception as e:
        return jsonify({'error': f'Search failed: {str(e)}'}), 400

# ===================== EDIT ROUTES =====================

@app.route('/update-cell', methods=['POST'])
@login_required
def update_cell():
    data = request.get_json()
    row = data.get('row')
    col = data.get('col')
    value = data.get('value')
    
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        df.at[row, col] = value
        app.config['CURRENT_DF'] = df
        return jsonify({'success': True, 'message': 'Cell updated successfully'})
    except Exception as e:
        return jsonify({'error': f'Update failed: {str(e)}'}), 400

@app.route('/add-row', methods=['POST'])
@login_required
def add_row():
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        new_row = {col: '' for col in df.columns}
        new_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        app.config['CURRENT_DF'] = new_df
        return jsonify({'success': True, 'message': 'Row added successfully'})
    except Exception as e:
        return jsonify({'error': f'Failed to add row: {str(e)}'}), 400

@app.route('/delete-row', methods=['POST'])
@login_required
def delete_row():
    data = request.get_json()
    row_index = data.get('row')
    
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        df = df.drop(index=row_index).reset_index(drop=True)
        app.config['CURRENT_DF'] = df
        return jsonify({'success': True, 'message': 'Row deleted successfully'})
    except Exception as e:
        return jsonify({'error': f'Failed to delete row: {str(e)}'}), 400

# ===================== DOWNLOAD ROUTES =====================

@app.route('/download/<format>')
@login_required
def download(format):
    df = app.config.get('CURRENT_DF')
    if df is None:
        return jsonify({'error': 'No data to download'}), 400
    
    try:
        df_clean = df.copy()
        df_clean = df_clean.fillna('')
        df_clean = df_clean.replace([np.nan, pd.NA, None], '', regex=False)
        
        for col in df_clean.columns:
            if df_clean[col].dtype == 'datetime64[ns]':
                has_time = False
                for val in df_clean[col].dropna().head(10):
                    if val.hour != 0 or val.minute != 0 or val.second != 0:
                        has_time = True
                        break
                
                if has_time:
                    df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d %H:%M')
                else:
                    df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d')
        
        if format == 'csv':
            output = io.StringIO()
            df_clean.to_csv(output, index=False, na_rep='')
            output.seek(0)
            return send_file(
                io.BytesIO(output.getvalue().encode('utf-8-sig')),
                mimetype='text/csv',
                as_attachment=True,
                download_name='formatted_data.csv'
            )
        
        elif format == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_clean.to_excel(writer, index=False, sheet_name='Formatted Data', na_rep='')
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='formatted_data.xlsx'
            )
        
        return jsonify({'error': 'Invalid format'}), 400
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/reset', methods=['POST'])
@login_required
def reset_data():
    original_df = app.config.get('CURRENT_ORIGINAL_DF')
    if original_df is None:
        return jsonify({'error': 'No original data to reset to'}), 400
    
    try:
        formatted_df = DataFormatter.apply_all_formattings(original_df.copy())
        app.config['CURRENT_DF'] = formatted_df
        app.config['CLEANING_HISTORY'] = []
        return jsonify({'success': True, 'message': 'Data reset to original'})
    except Exception as e:
        return jsonify({'error': f'Reset failed: {str(e)}'}), 400

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)