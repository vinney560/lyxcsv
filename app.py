from flask import (
    Flask, render_template, request, send_file, jsonify, session,\
      redirect, stream_with_context, Response
)
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import check_password_hash, generate_password_hash
from flask_sqlalchemy import SQLAlchemy
import io, re, warnings, secrets, json, hashlib
from datetime import datetime
from zoneinfo import ZoneInfo
from decouple import config
from scipy import stats
import pandas as pd
import numpy as np
import openpyxl

warnings.filterwarnings("ignore", category=UserWarning)

app = Flask(__name__)

app.secret_key = config('SECRET_KEY', default='test_secret_eky')
app.config['MAX_CONTENT_LENGTH'] = int(config('MAX_CONTENT_LENGTH', default=100 * 1024 * 1024))
app.config['SQLALCHEMY_DATABASE_URI'] = config('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
    'pool_size': 10,
    'pool_recycle': 3600,
    'pool_pre_ping': True,
}

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.refresh_view = 'login'

# =============== GMT +3 ====================
NAIROBI_NOW = datetime.now(ZoneInfo('Africa/Nairobi'))
# ===================== USER MODEL =====================

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    email = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    last_login = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    is_active = db.Column(db.Boolean, default=True)
    
    uploads = db.relationship('UploadHistory', backref='user', lazy=True)
    processing_jobs = db.relationship('ProcessingJob', backref='user', lazy=True)
    saved_data = db.relationship('SavedCleanData', backref='user', lazy=True)
    active_session = db.relationship('ActiveSessionData', backref='user', lazy=True, uselist=False)
    
    def to_dict(self):
        return {
            'id': self.id,
            'username': self.username,
            'email': self.email,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'last_login': self.last_login.isoformat() if self.last_login else None,
            'is_active': self.is_active
        }

class UploadHistory(db.Model):
    __tablename__ = 'upload_history'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    filename = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.BigInteger)
    file_type = db.Column(db.String(50))
    original_headers = db.Column(db.Text)
    row_count = db.Column(db.Integer)
    column_count = db.Column(db.Integer)
    uploaded_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    status = db.Column(db.String(50), default='uploaded')
    session_id = db.Column(db.String(100))
    
    processing_jobs = db.relationship('ProcessingJob', backref='upload', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'filename': self.filename,
            'file_size': self.file_size,
            'file_type': self.file_type,
            'row_count': self.row_count,
            'column_count': self.column_count,
            'uploaded_at': self.uploaded_at.isoformat() if self.uploaded_at else None,
            'status': self.status
        }

class ProcessingJob(db.Model):
    __tablename__ = 'processing_jobs'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    upload_id = db.Column(db.Integer, db.ForeignKey('upload_history.id'), nullable=True)
    job_type = db.Column(db.String(50))
    status = db.Column(db.String(50), default='pending')
    started_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    completed_at = db.Column(db.DateTime)
    job_details = db.Column(db.Text)
    result_summary = db.Column(db.Text)
    input_data_hash = db.Column(db.String(64))
    output_data_hash = db.Column(db.String(64))
    session_id = db.Column(db.String(100))
    
    audit_logs = db.relationship('AuditLog', backref='job', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'job_type': self.job_type,
            'status': self.status,
            'started_at': self.started_at.isoformat() if self.started_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'job_details': json.loads(self.job_details) if self.job_details else {},
            'result_summary': json.loads(self.result_summary) if self.result_summary else {}
        }

class CleanedData(db.Model):
    __tablename__ = 'cleaned_data'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    upload_id = db.Column(db.Integer, db.ForeignKey('upload_history.id'), nullable=True)
    job_id = db.Column(db.Integer, db.ForeignKey('processing_jobs.id'), nullable=True)
    data_json = db.Column(db.Text)
    data_hash = db.Column(db.String(64))
    row_count = db.Column(db.Integer)
    column_count = db.Column(db.Integer)
    cleaning_steps = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    storage_path = db.Column(db.String(500))
    storage_type = db.Column(db.String(50), default='json')
    session_id = db.Column(db.String(100))
    
    def to_dict(self):
        return {
            'id': self.id,
            'row_count': self.row_count,
            'column_count': self.column_count,
            'cleaning_steps': json.loads(self.cleaning_steps) if self.cleaning_steps else [],
            'created_at': self.created_at.isoformat() if self.created_at else None
        }

class SavedCleanData(db.Model):
    __tablename__ = 'saved_clean_data'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    filename = db.Column(db.String(255), nullable=False)
    file_size = db.Column(db.BigInteger)
    file_type = db.Column(db.String(50), default='csv')
    data_json = db.Column(db.Text, nullable=False)
    data_hash = db.Column(db.String(64))
    row_count = db.Column(db.Integer)
    column_count = db.Column(db.Integer)
    cleaning_steps = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    session_id = db.Column(db.String(100))
    status = db.Column(db.String(50), default='saved')
    
    def to_dict(self):
        return {
            'id': self.id,
            'filename': self.filename,
            'file_type': self.file_type,
            'row_count': self.row_count,
            'column_count': self.column_count,
            'cleaning_steps': json.loads(self.cleaning_steps) if self.cleaning_steps else [],
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'status': self.status
        }

class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    job_id = db.Column(db.Integer, db.ForeignKey('processing_jobs.id'), nullable=True)
    action = db.Column(db.String(100), nullable=False)
    details = db.Column(db.Text)
    ip_address = db.Column(db.String(45))
    user_agent = db.Column(db.String(255))
    session_id = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=lambda: NAIROBI_NOW)
    
    def to_dict(self):
        return {
            'id': self.id,
            'action': self.action,
            'details': json.loads(self.details) if self.details else {},
            'created_at': self.created_at.isoformat() if self.created_at else None
        }

class ActiveSessionData(db.Model):
    __tablename__ = 'active_session_data'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, unique=True)
    current_df_json = db.Column(db.Text)
    original_df_json = db.Column(db.Text)
    cleaning_history_json = db.Column(db.Text)
    current_upload_id = db.Column(db.Integer)
    last_updated = db.Column(db.DateTime, default=lambda: NAIROBI_NOW, onupdate=lambda: NAIROBI_NOW)
    
    def to_dict(self):
        return {
            'id': self.id,
            'user_id': self.user_id,
            'current_upload_id': self.current_upload_id,
            'last_updated': self.last_updated.isoformat() if self.last_updated else None
        }

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# =========== Unique table IDs ================
def gen_unique_id(_tablename, max_attempts=100):
    for attempt in range(max_attempts):
        r_id = secrets.randbelow(900000) + 100000
        
        with db.session.begin_nested():
            existing = db.session.query(_tablename.id).filter_by(id=r_id).with_for_update().first()
            if not existing:
                return r_id
        
        if attempt < max_attempts - 1:
            import time
            time.sleep(0.01) 
    
    raise ValueError(f"Failed to generate unique ID after {max_attempts} attempts and fallback attempts")

# ===================== HELPER FUNCTIONS =====================

def get_or_create_session_id():
    if 'session_id' not in session:
        session['session_id'] = secrets.token_hex(16)
    return session['session_id']

def log_audit(action, details=None, job_id=None):
    try:
        audit = AuditLog(
            id=gen_unique_id(AuditLog),
            user_id=current_user.id if current_user.is_authenticated else None,
            job_id=job_id,
            action=action,
            details=json.dumps(details) if details else None,
            ip_address=request.remote_addr,
            user_agent=request.headers.get('User-Agent'),
            session_id=get_or_create_session_id()
        )
        db.session.add(audit)
        db.session.commit()
    except Exception as e:
        print(f"Audit log failed: {e}")
        db.session.rollback()

def validate_username(username):
    pattern = r'^[a-zA-Z0-9_ ]{3,20}$'
    return re.match(pattern, username) is not None

def validate_password(password):
    return 6 <= len(password) <= 14

def validate_email(email):
    if not email:
        return True
    pattern = r'^[a-zA-Z0-9._%+-]+@(gmail\.com|yahoo\.com)$'
    return re.match(pattern, email) is not None

def save_upload_record(filename, file_size, file_type, headers, rows, cols):
    try:
        upload = UploadHistory(
            id=gen_unique_id(UploadHistory),
            user_id=current_user.id,
            filename=filename,
            file_size=file_size,
            file_type=file_type,
            original_headers=json.dumps(headers) if headers else None,
            row_count=rows,
            column_count=cols,
            session_id=get_or_create_session_id()
        )
        db.session.add(upload)
        db.session.commit()
        return upload.id
    except Exception as e:
        print(f"Failed to save upload record: {e}")
        db.session.rollback()
        return None

def save_processing_job(upload_id, job_type, job_details=None):
    try:
        job = ProcessingJob(
            id=gen_unique_id(ProcessingJob),
            user_id=current_user.id,
            upload_id=upload_id,
            job_type=job_type,
            job_details=json.dumps(job_details) if job_details else None,
            session_id=get_or_create_session_id(),
            status='started'
        )
        db.session.add(job)
        db.session.commit()
        return job.id
    except Exception as e:
        print(f"Failed to save job record: {e}")
        db.session.rollback()
        return None

def update_job_status(job_id, status, result_summary=None):
    try:
        job = db.session.query(ProcessingJob).get(job_id)
        if job:
            job.status = status
            if status in ['completed', 'failed']:
                job.completed_at = NAIROBI_NOW
            if result_summary:
                job.result_summary = json.dumps(result_summary)
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to update job: {e}")
        db.session.rollback()
    return False

def save_cleaned_data(job_id, upload_id, df, cleaning_steps=None):
    try:
        data_json = df.to_json(orient='records', date_format='iso')
        data_hash = hashlib.sha256(data_json.encode()).hexdigest()
        
        cleaned = CleanedData(
            id=gen_unique_id(CleanedData),
            user_id=current_user.id,
            upload_id=upload_id,
            job_id=job_id,
            data_json=data_json,
            data_hash=data_hash,
            row_count=len(df),
            column_count=len(df.columns),
            cleaning_steps=json.dumps(cleaning_steps) if cleaning_steps else None,
            session_id=get_or_create_session_id()
        )
        db.session.add(cleaned)
        db.session.commit()
        return cleaned.id
    except Exception as e:
        print(f"Failed to save cleaned data: {e}")
        db.session.rollback()
        return None

def get_cleaned_data(cleaned_id):
    try:
        cleaned = db.session.query(CleanedData).get(cleaned_id)
        if cleaned and cleaned.data_json:
            df = pd.read_json(io.StringIO(cleaned.data_json), orient='records')
            return df
    except Exception as e:
        print(f"Failed to retrieve cleaned data: {e}")
    return None

def get_user_history_with_data(user_id=None):
    try:
        if user_id is None or user_id != current_user.id:
            return []
        
        uploads = UploadHistory.query.filter_by(user_id=user_id).order_by(UploadHistory.uploaded_at.desc()).limit(50).all()

        history = []
        for upload in uploads:
            upload_dict = upload.to_dict()
            cleaned_data = CleanedData.query.filter_by(upload_id=upload.id).order_by(CleanedData.created_at.desc()).first()
            if cleaned_data:
                upload_dict['has_data'] = True
                upload_dict['cleaned_id'] = cleaned_data.id
                upload_dict['cleaned_at'] = cleaned_data.created_at.isoformat() if cleaned_data.created_at else None
                upload_dict['cleaned_rows'] = cleaned_data.row_count
                upload_dict['cleaned_cols'] = cleaned_data.column_count
                upload_dict['status'] = 'saved'
            else:
                upload_dict['has_data'] = False
                upload_dict['status'] = upload.status or 'uploaded'
            history.append(upload_dict)
        return history
    except Exception as e:
        print(f"Failed to get history: {e}")
        return []

def save_saved_data_to_db(df, file_name):
    try:
        if not current_user.is_authenticated:
            return None
        
        data_json = df.to_json(orient='records', date_format='iso')
        data_hash = hashlib.sha256(data_json.encode()).hexdigest()
        
        saved_data = SavedCleanData(
            id=gen_unique_id(SavedCleanData),
            user_id=current_user.id,
            filename=f"{file_name}.csv",
            file_size=len(data_json.encode('utf-8')),
            file_type='csv',
            data_json=data_json,
            data_hash=data_hash,
            row_count=len(df),
            column_count=len(df.columns),
            cleaning_steps=json.dumps(['Final save']),
            status='saved'
        )
        db.session.add(saved_data)
        db.session.commit()
        
        log_audit('save_saved_data', {'filename': file_name, 'rows': len(df), 'saved_id': saved_data.id})
        return saved_data.id
    except Exception as e:
        db.session.rollback()
        print(f"Failed to save saved data: {e}")
        return None

def get_saved_data_list(user_id=None):
    try:
        if user_id:
            saved_data = SavedCleanData.query.filter_by(user_id=user_id).order_by(SavedCleanData.created_at.desc()).limit(50).all()
        else:
            session_id = get_or_create_session_id()
            saved_data = SavedCleanData.query.filter_by(session_id=session_id).order_by(SavedCleanData.created_at.desc()).limit(50).all()
        return [s.to_dict() for s in saved_data]
    except Exception as e:
        print(f"Failed to get saved data: {e}")
        return []

def load_saved_data_from_db(saved_id):
    try:
        saved_data = db.session.query(SavedCleanData).get(saved_id)
        if saved_data and saved_data.data_json:
            df = pd.read_json(io.StringIO(saved_data.data_json), orient='records')
            return df
    except Exception as e:
        print(f"Failed to load saved data: {e}")
    return None

def delete_saved_data_from_db(saved_id, user_id):
    try:
        saved_data = SavedCleanData.query.filter_by(id=saved_id, user_id=user_id).first()
        if saved_data:
            db.session.delete(saved_data)
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to delete saved data: {e}")
        db.session.rollback()
    return False

def clear_saved_data_for_user(user_id):
    try:
        if not user_id:
            return False
        SavedCleanData.query.filter_by(user_id=user_id).delete()
        db.session.commit()
        return True
    except Exception as e:
        print(f"Failed to clear saved data: {e}")
        db.session.rollback()
    return False

# ===================== ACTIVE SESSION DATA HELPERS =====================

def get_active_session_data():
    """Get or create active session data for current user"""
    try:
        if not current_user.is_authenticated:
            return None
        
        active_session = ActiveSessionData.query.filter_by(user_id=current_user.id).first()
        if not active_session:
            active_session = ActiveSessionData(
                id=gen_unique_id(ActiveSessionData),
                user_id=current_user.id
            )
            db.session.add(active_session)
            db.session.commit()
        return active_session
    except Exception as e:
        print(f"Failed to get active session data: {e}")
        db.session.rollback()
        return None

def save_current_df_to_db(df):
    """Save current dataframe to database"""
    try:
        active_session = get_active_session_data()
        if active_session and df is not None:
            active_session.current_df_json = df.to_json(orient='records', date_format='iso')
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to save current df to db: {e}")
        db.session.rollback()
    return False

def get_current_df_from_db():
    """Get current dataframe from database"""
    try:
        active_session = get_active_session_data()
        if active_session and active_session.current_df_json:
            df = pd.read_json(io.StringIO(active_session.current_df_json), orient='records')
            return df
    except Exception as e:
        print(f"Failed to get current df from db: {e}")
    return None

def save_original_df_to_db(df):
    """Save original dataframe to database"""
    try:
        active_session = get_active_session_data()
        if active_session and df is not None:
            active_session.original_df_json = df.to_json(orient='records', date_format='iso')
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to save original df to db: {e}")
        db.session.rollback()
    return False

def get_original_df_from_db():
    """Get original dataframe from database"""
    try:
        active_session = get_active_session_data()
        if active_session and active_session.original_df_json:
            df = pd.read_json(io.StringIO(active_session.original_df_json), orient='records')
            return df
    except Exception as e:
        print(f"Failed to get original df from db: {e}")
    return None

def save_cleaning_history_to_db(history):
    """Save cleaning history to database"""
    try:
        active_session = get_active_session_data()
        if active_session:
            # Convert DataFrames to JSON strings
            history_json = []
            for df in history:
                history_json.append(df.to_json(orient='records', date_format='iso'))
            active_session.cleaning_history_json = json.dumps(history_json)
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to save cleaning history to db: {e}")
        db.session.rollback()
    return False

def get_cleaning_history_from_db():
    """Get cleaning history from database"""
    try:
        active_session = get_active_session_data()
        if active_session and active_session.cleaning_history_json:
            history_json = json.loads(active_session.cleaning_history_json)
            history = []
            for df_json in history_json:
                df = pd.read_json(io.StringIO(df_json), orient='records')
                history.append(df)
            return history
    except Exception as e:
        print(f"Failed to get cleaning history from db: {e}")
    return []

def save_current_upload_id_to_db(upload_id):
    """Save current upload ID to database"""
    try:
        active_session = get_active_session_data()
        if active_session:
            active_session.current_upload_id = upload_id
            db.session.commit()
            return True
    except Exception as e:
        print(f"Failed to save current upload id to db: {e}")
        db.session.rollback()
    return False

def get_current_upload_id_from_db():
    """Get current upload ID from database"""
    try:
        active_session = get_active_session_data()
        if active_session:
            return active_session.current_upload_id
    except Exception as e:
        print(f"Failed to get current upload id from db: {e}")
    return None

# ===================== EMAIL FORMATTER =====================

class EmailFormatter:
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

# ===================== DATA CLEANER =====================

class DataCleaner:
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

# ===================== DATA FORMATTER =====================

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

# ===================== BULK UPDATER =====================

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

# ===================== FILE PROCESSING =====================

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
        elif filename.endswith(('.txt', '.text')):
            # Read plain text file
            encodings = ['utf-8', 'latin1', 'cp1252']
            text_content = None
            for encoding in encodings:
                try:
                    file.seek(0)
                    text_content = file.read().decode(encoding)
                    break
                except:
                    continue
            
            if text_content is None:
                raise ValueError("Could not decode text file")
            
            # Split into lines and create a DataFrame
            lines = text_content.splitlines()
            # Remove empty lines
            lines = [line.strip() for line in lines if line.strip()]
            
            if not lines:
                raise ValueError("Text file is empty")
            
            # Try to detect if it's delimited (comma, tab, semicolon, pipe)
            first_line = lines[0]
            delimiter = None
            
            if ',' in first_line:
                delimiter = ','
            elif '\t' in first_line:
                delimiter = '\t'
            elif ';' in first_line:
                delimiter = ';'
            elif '|' in first_line:
                delimiter = '|'
            
            if delimiter and len(lines) > 1:
                # Check if first line might be headers
                first_line_parts = len(first_line.split(delimiter))
                second_line_parts = len(lines[1].split(delimiter)) if len(lines) > 1 else 0
                
                if first_line_parts == second_line_parts:
                    # It's a delimited text file, parse as CSV
                    from io import StringIO
                    df = pd.read_csv(StringIO(text_content), delimiter=delimiter, encoding=encoding)
                else:
                    # Treat as single column
                    df = pd.DataFrame({'Text': lines})
            else:
                # Single column DataFrame
                df = pd.DataFrame({'Text': lines})
        else:
            raise ValueError("Unsupported file format")
        
        save_original_df_to_db(df.copy())
        save_cleaning_history_to_db([])
        
        formatted_df = DataFormatter.apply_all_formattings(df)
        
        try:
            upload_id = save_upload_record(
                filename=file.filename,
                file_size=file.content_length if hasattr(file, 'content_length') else None,
                file_type=filename.split('.')[-1],
                headers=df.columns.tolist(),
                rows=len(df),
                cols=len(df.columns)
            )
            save_current_upload_id_to_db(upload_id)
        except Exception as e:
            print(f"Failed to save upload to database: {e}")
        
        return formatted_df, None
        
    except Exception as e:
        return None, str(e)

# ===================== ROUTES =====================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=current_user.username)

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html')
    
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '')
    confirm_password = data.get('confirm_password', '')
    email = data.get('email', '').strip()
    
    if not username or not password or not confirm_password:
        return jsonify({'success': False, 'message': 'All fields are required'}), 400
    
    if not validate_username(username):
        return jsonify({'success': False, 'message': 'Username must be 3-20 characters (letters, numbers, underscores, spaces)'}), 400
    
    if not validate_password(password):
        return jsonify({'success': False, 'message': 'Password must be 6-14 characters'}), 400
    
    if password != confirm_password:
        return jsonify({'success': False, 'message': 'Passwords do not match'}), 400
    
    if email and not validate_email(email):
        return jsonify({'success': False, 'message': 'Email must be @gmail.com or @yahoo.com'}), 400
    
    if User.query.filter_by(username=username).first():
        return jsonify({'success': False, 'message': 'Username already taken'}), 400
    
    if email and User.query.filter_by(email=email).first():
        return jsonify({'success': False, 'message': 'Email already registered'}), 400
    
    try:
        user = User(
            id=gen_unique_id(User),
            username=username,
            password_hash=generate_password_hash(password),
            email=email if email else None
        )
        db.session.add(user)
        db.session.commit()
        
        login_user(user)
        log_audit('register', {'username': username, 'redirect': '/dashboard'})

        return jsonify({
                'success': True, 
                'message': 'Registration successful!',
                'username': user.username,
                'redirect': '/dashboard'
            }), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Registration failed: {str(e)}'}), 500

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')
    
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '')
    remember = data.get('remember', False)
    
    if not username or not password:
        return jsonify({'success': False, 'message': 'Username and password are required'}), 400
    
    # Check if user exists
    user = User.query.filter(
        (User.username == username) | (User.email == username)
    ).first()
    
    if not user:
        return jsonify({'success': False, 'message': 'Invalid username or password'}), 401
    
    # Verify password
    if not check_password_hash(user.password_hash, password):
        return jsonify({'success': False, 'message': 'Invalid username or password'}), 401
    
    # Login user
    login_user(user, remember=remember)
    log_audit('login', {'username': user.username})
    
    return jsonify({
        'success': True, 
        'message': 'Login successful!',
        'username': user.username,
        'redirect': '/dashboard'
    }), 200

@app.route('/api/check-auth', methods=['GET'])
def check_auth():
    """Check if user is authenticated"""
    if current_user.is_authenticated:
        return jsonify({
            'authenticated': True,
            'username': current_user.username,
            'user_id': current_user.id
        })
    else:
        return jsonify({
            'authenticated': False
        })
    
@app.route('/logout', methods=['POST'])
@login_required
def logout():
    log_audit('logout')
    logout_user()
    return jsonify({'success': True, 'message': 'Logged out successfully'})

@app.route('/txt-input')
@login_required
def txt_input():
    return render_template("get_text_file.html")

@app.route('/txt-file', methods=['POST'])
def txt_file():
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No JSON data provided'}), 400
        
        text = data.get('text', '').strip()
        
        if not text:
            return jsonify({'error': 'Text is empty'}), 400
        
        filename = data.get('filename', 'download.txt')
        if not filename.endswith('.txt'):
            filename += '.txt'
        
        if data.get('add_timestamp', False):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            name, ext = filename.rsplit('.', 1)
            filename = f"{name}_{timestamp}.{ext}"
        
        def generate():
            chunk_size = 8192
            text_bytes = text.encode('utf-8')
            
            for i in range(0, len(text_bytes), chunk_size):
                yield text_bytes[i:i+chunk_size]
        
        return Response(
            stream_with_context(generate()),
            mimetype='text/plain',
            headers={
                'Content-Disposition': f'attachment; filename={filename}',
                'Content-Type': 'text/plain; charset=utf-8'
            }
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

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
        log_audit('upload_failed', {'filename': file.filename, 'error': error})
        return jsonify({'error': error}), 400
    
    serializable_df = DataFormatter.convert_to_serializable(formatted_df)
    column_info = DataFormatter.get_column_info(serializable_df)
    
    table_html = serializable_df.to_html(
        classes='formatted-table',
        index=False,
        escape=False
    )
    
    save_current_df_to_db(formatted_df)
    
    log_audit('upload_success', {
        'filename': file.filename,
        'rows': len(serializable_df),
        'cols': len(serializable_df.columns)
    })
    
    return jsonify({
        'success': True,
        'table_html': table_html,
        'columns': column_info,
        'rows': int(len(serializable_df)),
        'cols': int(len(serializable_df.columns)),
        'filename': file.filename
    })

@app.route('/history', methods=['GET'])
@login_required
def get_history():
    user_id = current_user.id if current_user.is_authenticated else None
    history = get_user_history_with_data(user_id)
    return jsonify({'success': True, 'history': history})

@app.route('/load-history/<int:upload_id>', methods=['GET'])
@login_required
def load_history(upload_id):
    try:
        upload = UploadHistory.query.filter_by(id=upload_id, user_id=current_user.id).first()
        if not upload:
            return jsonify({'error': 'File not found or access denied'}), 404
        
        cleaned_data = CleanedData.query.filter_by(upload_id=upload_id).order_by(CleanedData.created_at.desc()).first()
        if not cleaned_data or not cleaned_data.data_json:
            return jsonify({'error': 'No data found for this upload'}), 404
        
        df = pd.read_json(io.StringIO(cleaned_data.data_json), orient='records')
        
        save_current_df_to_db(df)
        save_original_df_to_db(df.copy())
        save_cleaning_history_to_db([])
        save_current_upload_id_to_db(upload_id)
        
        serializable_df = DataFormatter.convert_to_serializable(df)
        column_info = DataFormatter.get_column_info(serializable_df)
        
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        log_audit('load_history', {'upload_id': upload_id, 'filename': upload.filename})
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'columns': column_info,
            'rows': int(len(serializable_df)),
            'cols': int(len(serializable_df.columns)),
            'filename': upload.filename,
            'upload_id': upload_id
        })
    except Exception as e:
        return jsonify({'error': f'Failed to load file: {str(e)}'}), 500

@app.route('/download-history/<int:upload_id>')
@login_required
def download_history(upload_id):
    try:
        upload = UploadHistory.query.filter_by(id=upload_id, user_id=current_user.id).first()
        if not upload:
            return jsonify({'error': 'File not found or access denied'}), 404
        
        cleaned_data = CleanedData.query.filter_by(upload_id=upload_id).order_by(CleanedData.created_at.desc()).first()
        if not cleaned_data or not cleaned_data.data_json:
            return jsonify({'error': 'No data found for this upload'}), 404
        
        df = pd.read_json(io.StringIO(cleaned_data.data_json), orient='records')
        
        output = io.StringIO()
        df.to_csv(output, index=False)
        output.seek(0)
        
        log_audit('download_history', {'upload_id': upload_id, 'filename': upload.filename})
        
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f"lyx_{upload.filename}"
        )
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/delete-history/<int:upload_id>', methods=['DELETE'])
@login_required
def delete_history(upload_id):
    try:
        upload = UploadHistory.query.filter_by(id=upload_id, user_id=current_user.id).first()
        if not upload:
            return jsonify({'error': 'File not found'}), 404
        
        CleanedData.query.filter_by(upload_id=upload_id).delete()
        ProcessingJob.query.filter_by(upload_id=upload_id).delete()
        db.session.delete(upload)
        db.session.commit()
        
        log_audit('delete_history', {'upload_id': upload_id})
        return jsonify({'success': True, 'message': 'Deleted successfully'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Delete failed: {str(e)}'}), 500

@app.route('/clear-history', methods=['POST'])
@login_required
def clear_history():
    try:
        ProcessingJob.query.filter_by(user_id=current_user.id).delete()
        CleanedData.query.filter_by(user_id=current_user.id).delete()
        UploadHistory.query.filter_by(user_id=current_user.id).delete()
        db.session.commit()
        log_audit('clear_history')
        return jsonify({'success': True, 'message': 'History cleared successfully'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Failed to clear history: {str(e)}'}), 500

@app.route('/detect-duplicates', methods=['GET'])
@login_required
def detect_duplicates():
    df = get_current_df_from_db()
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
        
        log_audit('detect_duplicates', {'duplicate_count': int(duplicate_count)})
        
        return jsonify({
            'success': True,
            'duplicate_count': int(duplicate_count),
            'duplicate_rows': duplicate_rows
        })
    except Exception as e:
        return jsonify({'error': f'Failed to detect duplicates: {str(e)}'}), 400

@app.route('/cleaning-report', methods=['POST'])
@login_required
def get_cleaning_report():
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        data = request.get_json() or {}
        selected_columns = data.get('columns', None)
        
        report = DataCleaner.get_cleaning_report(df, selected_columns)
        report = json.loads(json.dumps(report, default=str))
        
        log_audit('generate_report', {'columns': selected_columns})
        return jsonify({'success': True, 'report': report})
    except Exception as e:
        return jsonify({'error': f'Failed to generate report: {str(e)}'}), 400

@app.route('/clean-data', methods=['POST'])
@login_required
def clean_data():
    """Apply data cleaning operations to the current dataset.
    Retrieves current data from database, applies requested cleaning steps,
    saves the modified data back, and logs the operation for auditing.
    """
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        data = request.get_json() or {}
        options = data.get('options', {})
        selected_columns = data.get('columns', None)
        
        # Preserve current state for undo functionality
        cleaning_history = get_cleaning_history_from_db()
        cleaning_history.append(df.copy())
        save_cleaning_history_to_db(cleaning_history)
        
        cleaned_df, clean_log = DataCleaner.apply_all_cleaning(df, options, selected_columns)
        
        formatted_df = DataFormatter.apply_all_formattings(cleaned_df)
        save_current_df_to_db(formatted_df)
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        clean_log = json.loads(json.dumps(clean_log, default=str))
        
        upload_id = get_current_upload_id_from_db()
        if upload_id:
            job_id = save_processing_job(upload_id, 'cleaning', {
                'columns': selected_columns,
                'options': options
            })
            if job_id:
                cleaned_id = save_cleaned_data(job_id, upload_id, formatted_df, clean_log.get('steps', []))
                update_job_status(job_id, 'completed', clean_log)
                
                log_audit('clean_completed', {
                    'job_id': job_id,
                    'columns': selected_columns,
                    'steps': clean_log.get('steps', [])
                })
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'clean_log': clean_log,
            'message': f"Data cleaned! {len(clean_log['steps'])} steps applied."
        })
    except Exception as e:
        log_audit('clean_failed', {'error': str(e)})
        return jsonify({'error': f'Cleaning failed: {str(e)}'}), 400

@app.route('/undo-cleaning', methods=['POST'])
@login_required
def undo_cleaning():
    """Undo the most recent data cleaning operation.
    Reverts the dataset to its previous state from the cleaning history,
    updates the current data in the database, and logs the action.
    """
    cleaning_history = get_cleaning_history_from_db()
    if not cleaning_history:
        return jsonify({'error': 'Nothing to undo'}), 400
    
    try:
        df = cleaning_history.pop()
        save_cleaning_history_to_db(cleaning_history)
        
        formatted_df = DataFormatter.apply_all_formattings(df)
        save_current_df_to_db(formatted_df)
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        log_audit('undo_cleaning')
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'message': 'Cleaning undone successfully!'
        })
    except Exception as e:
        return jsonify({'error': f'Undo failed: {str(e)}'}), 400

@app.route('/bulk-find', methods=['POST'])
@login_required
def bulk_find():
    """Search for all occurrences of a term across the dataset.
    Finds matches in all columns, returns detailed results including
    which columns contain matches and how many occurrences exist.
    """
    data = request.get_json()
    search_term = data.get('search_term', '')
    case_sensitive = data.get('case_sensitive', False)
    
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        results = BulkUpdater.find_all_occurrences(df, search_term, case_sensitive)
        results = json.loads(json.dumps(results, default=str))
        total_matches = sum(r['count'] for r in results)
        
        log_audit('bulk_find', {
            'search_term': search_term,
            'matches': total_matches
        })
        
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
    """Find and replace text across multiple cells in the dataset.
    Can replace all occurrences or a single match, preserves history for undo,
    and logs the operation for audit purposes.
    """
    data = request.get_json()
    search_col = data.get('column')
    search_term = data.get('search_term', '')
    replace_term = data.get('replace_term', '')
    case_sensitive = data.get('case_sensitive', False)
    single_row = data.get('single_row', None)
    
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        # Preserve current state for undo functionality
        cleaning_history = get_cleaning_history_from_db()
        cleaning_history.append(df.copy())
        save_cleaning_history_to_db(cleaning_history)
        
        updated_df, result = BulkUpdater.find_and_replace(
            df, search_col, search_term, replace_term, case_sensitive, single_row
        )
        
        if 'error' in result:
            return jsonify({'error': result['error']}), 400
        
        save_current_df_to_db(updated_df)
        
        formatted_df = DataFormatter.apply_all_formattings(updated_df)
        save_current_df_to_db(formatted_df)
        
        serializable_df = DataFormatter.convert_to_serializable(formatted_df)
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        result = json.loads(json.dumps(result, default=str))
        
        log_audit('bulk_replace', {
            'column': search_col,
            'search_term': search_term,
            'replace_term': replace_term,
            'affected_rows': result.get('affected_rows', 0)
        })
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'rows': int(len(serializable_df)),
            'result': result
        })
    except Exception as e:
        return jsonify({'error': f'Replace failed: {str(e)}'}), 400

@app.route('/search', methods=['POST'])
@login_required
def search_data():
    """Search the entire dataset for a specific term.
    Returns filtered results matching the search query, with optional
    case-sensitive matching, formatted as HTML table for display.
    """
    data = request.get_json()
    search_term = data.get('search', '')
    case_sensitive = data.get('case_sensitive', False)
    
    df = get_current_df_from_db()
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

@app.route('/update-cell', methods=['POST'])
@login_required
def update_cell():
    """Update a single cell's value in the dataset.
    Modifies the specified row and column with the new value,
    persists changes to the database, and logs the edit.
    """
    data = request.get_json()
    row = data.get('row')
    col = data.get('col')
    value = data.get('value')
    
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        df.at[row, col] = value
        save_current_df_to_db(df)
        log_audit('cell_update', {'row': row, 'col': col})
        return jsonify({'success': True, 'message': 'Cell updated successfully'})
    except Exception as e:
        return jsonify({'error': f'Update failed: {str(e)}'}), 400

@app.route('/add-row', methods=['POST'])
@login_required
def add_row():
    """Append a new empty row to the end of the dataset.
    Creates a row with empty values for all columns, adds it to the dataframe,
    saves the updated data to the database, and logs the action.
    """
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        new_row = {col: '' for col in df.columns}
        new_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_current_df_to_db(new_df)
        log_audit('add_row')
        return jsonify({'success': True, 'message': 'Row added successfully'})
    except Exception as e:
        return jsonify({'error': f'Failed to add row: {str(e)}'}), 400

@app.route('/delete-row', methods=['POST'])
@login_required
def delete_row():
    """Remove a specific row from the dataset.
    Deletes the row at the specified index, resets the dataframe index,
    persists the changes to the database, and logs the deletion.
    """
    data = request.get_json()
    row_index = data.get('row')
    
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data loaded'}), 400
    
    try:
        df = df.drop(index=row_index).reset_index(drop=True)
        save_current_df_to_db(df)
        log_audit('delete_row', {'row': row_index})
        return jsonify({'success': True, 'message': 'Row deleted successfully'})
    except Exception as e:
        return jsonify({'error': f'Failed to delete row: {str(e)}'}), 400

@app.route('/download/<format>')
@login_required
def download(format):
    """Export the current dataset in the requested format.
    Supports CSV and Excel exports. Processes dates to proper string formats,
    cleans up null values, and sends the file as an attachment for download.
    """
    df = get_current_df_from_db()
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
        
        log_audit('download', {'format': format})
        
        if format == 'csv':
            output = io.StringIO()
            df_clean.to_csv(output, index=False, na_rep='')
            output.seek(0)
            return send_file(
                io.BytesIO(output.getvalue().encode('utf-8-sig')),
                mimetype='text/csv',
                as_attachment=True,
                download_name=f'lyx_formatted_data_{NAIROBI_NOW.strftime("%Y%m%d%H%M%S")}.csv'
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
                download_name=f'lyx_formatted_data_{NAIROBI_NOW.strftime("%Y%m%d%H%M%S")}.xlsx'
            )
        
        return jsonify({'error': 'Invalid format'}), 400
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/reset', methods=['POST'])
@login_required
def reset_data():
    """Reset the entire dataset back to its original uploaded state.
    Reverts all changes, cleaning operations, and edits by restoring the
    original dataframe from the database. Clears the undo history.
    """
    original_df = get_original_df_from_db()
    if original_df is None:
        return jsonify({'error': 'No original data to reset to'}), 400
    
    try:
        formatted_df = DataFormatter.apply_all_formattings(original_df.copy())
        save_current_df_to_db(formatted_df)
        save_cleaning_history_to_db([])
        log_audit('reset_data')
        return jsonify({'success': True, 'message': 'Data reset to original'})
    except Exception as e:
        return jsonify({'error': f'Reset failed: {str(e)}'}), 400

# ===================== SAVED DATA ROUTES =====================

@app.route('/history-page', methods=['GET'])
@login_required
def history_page():
    return render_template('history.html')

@app.route('/save-final', methods=['POST'])
@login_required
def save_final_data():
    """Save the current processed dataset as a finalized entry.
    Archives the current dataframe in the database with a user-provided filename,
    creates a permanent saved record that can be accessed later from the history.
    """
    df = get_current_df_from_db()
    if df is None:
        return jsonify({'error': 'No data to save'}), 400
    
    try:
        data = request.get_json() or {}
        file_name = data.get('filename', f'lyx_formatted_data_{NAIROBI_NOW.strftime("%Y%m%d%H%M%S")}')
        
        saved_id = save_saved_data_to_db(df, file_name)
        
        if saved_id:
            return jsonify({
                'success': True,
                'message': 'Data saved successfully!',
                'saved_id': saved_id
            })
        else:
            return jsonify({'error': 'Failed to save data'}), 500
            
    except Exception as e:
        log_audit('save_failed', {'error': str(e)})
        return jsonify({'error': f'Save failed: {str(e)}'}), 500

@app.route('/saved-data', methods=['GET'])
@login_required
def get_saved_data():
    user_id = current_user.id
    saved_data = get_saved_data_list(user_id)
    return jsonify({'success': True, 'saved_data': saved_data})

@app.route('/saved-data/load/<int:saved_id>', methods=['GET'])
@login_required
def load_saved_data(saved_id):
    """Load a previously saved dataset from the archive.
    Retrieves a saved dataset by its ID, verifies user ownership,
    sets it as the current active dataset in the database, and prepares
    it for viewing and editing in the application.
    """
    try:
        saved_data = SavedCleanData.query.filter_by(id=saved_id, user_id=current_user.id).first()
        if not saved_data:
            return jsonify({'error': 'File not found or access denied'}), 404
        
        df = load_saved_data_from_db(saved_id)
        if df is None:
            return jsonify({'error': 'No data found'}), 404
        
        save_current_df_to_db(df)
        save_original_df_to_db(df.copy())
        save_cleaning_history_to_db([])
        
        serializable_df = DataFormatter.convert_to_serializable(df)
        column_info = DataFormatter.get_column_info(serializable_df)
        
        table_html = serializable_df.to_html(
            classes='formatted-table',
            index=False,
            escape=False
        )
        
        log_audit('load_saved_data', {'saved_id': saved_id, 'filename': saved_data.filename})
        
        return jsonify({
            'success': True,
            'table_html': table_html,
            'columns': column_info,
            'rows': int(len(serializable_df)),
            'cols': int(len(serializable_df.columns)),
            'filename': saved_data.filename,
            'saved_id': saved_id
        })
    except Exception as e:
        return jsonify({'error': f'Failed to load file: {str(e)}'}), 500

@app.route('/saved-data/download/<int:saved_id>')
@login_required
def download_saved_data(saved_id):
    try:
        saved_data = SavedCleanData.query.filter_by(id=saved_id, user_id=current_user.id).first()
        if not saved_data:
            return jsonify({'error': 'File not found or access denied'}), 404
        
        df = load_saved_data_from_db(saved_id)
        if df is None:
            return jsonify({'error': 'No data found'}), 404
        
        output = io.StringIO()
        df.to_csv(output, index=False)
        output.seek(0)
        
        log_audit('download_saved_data', {'saved_id': saved_id, 'filename': saved_data.filename})
        
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f"lyx_{saved_data.filename}"
        )
    except Exception as e:
        return jsonify({'error': f'Download failed: {str(e)}'}), 500

@app.route('/saved-data/delete/<int:saved_id>', methods=['DELETE'])
@login_required
def delete_saved_data(saved_id):
    try:
        if delete_saved_data_from_db(saved_id, current_user.id):
            log_audit('delete_saved_data', {'saved_id': saved_id})
            return jsonify({'success': True, 'message': 'Deleted successfully'})
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Delete failed: {str(e)}'}), 500

@app.route('/saved-data/clear', methods=['POST'])
@login_required
def clear_saved_data():
    try:
        if clear_saved_data_for_user(current_user.id):
            log_audit('clear_saved_data')
            return jsonify({'success': True, 'message': 'Saved data cleared successfully'})
        else:
            return jsonify({'error': 'Failed to clear saved data'}), 500
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Clear failed: {str(e)}'}), 500

# ===================== CREATE TABLES =====================

with app.app_context():
    db.create_all()
    print("Database tables created successfully!")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)