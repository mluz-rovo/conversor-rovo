"""
logging_config.py — Logging estruturado para convertor ROVO

Implementa:
- Audit trail (quem, quando, o quê)
- Error tracking
- Performance monitoring
- Rastreabilidade completa

Para usar: importar no rovo_v2.py
"""

import logging
import json
from datetime import datetime
from pathlib import Path
import hashlib


class AuditLogger:
    """Logger estruturado para auditoria"""
    
    def __init__(self, log_dir="./logs"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # File handler (audit trail)
        self.audit_file = self.log_dir / "audit.log"
        self.error_file = self.log_dir / "errors.log"
        self.performance_file = self.log_dir / "performance.log"
        
        # Setup logging
        self.logger = logging.getLogger("ROVO")
        self.logger.setLevel(logging.DEBUG)
        
        # Format
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(name)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Audit file
        audit_handler = logging.FileHandler(self.audit_file)
        audit_handler.setLevel(logging.INFO)
        audit_handler.setFormatter(formatter)
        
        # Error file
        error_handler = logging.FileHandler(self.error_file)
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(formatter)
        
        # Console (para desenvolvimento)
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        
        self.logger.addHandler(audit_handler)
        self.logger.addHandler(error_handler)
        self.logger.addHandler(console_handler)
    
    def log_upload(self, user_email, filename, file_size, client):
        """Log: ficheiro foi uploadado"""
        log_entry = {
            "event": "FILE_UPLOAD",
            "timestamp": datetime.now().isoformat(),
            "user": user_email,
            "filename": filename,
            "file_size_bytes": file_size,
            "client": client,
        }
        self.logger.info(f"UPLOAD | {json.dumps(log_entry)}")
        return log_entry
    
    def log_parse_start(self, filename, client):
        """Log: começou parsing"""
        log_entry = {
            "event": "PARSE_START",
            "timestamp": datetime.now().isoformat(),
            "filename": filename,
            "client": client,
        }
        self.logger.info(f"PARSE_START | {json.dumps(log_entry)}")
        return log_entry
    
    def log_parse_success(self, filename, rows_extracted, models_found, colors_found):
        """Log: parsing foi bem-sucedido"""
        log_entry = {
            "event": "PARSE_SUCCESS",
            "timestamp": datetime.now().isoformat(),
            "filename": filename,
            "rows_extracted": rows_extracted,
            "models_count": len(models_found),
            "colors_count": len(colors_found),
            "models": list(models_found),
            "colors": list(colors_found),
        }
        self.logger.info(f"PARSE_SUCCESS | {json.dumps(log_entry)}")
        return log_entry
    
    def log_parse_error(self, filename, error_message, line_number=None):
        """Log: erro durante parsing"""
        log_entry = {
            "event": "PARSE_ERROR",
            "timestamp": datetime.now().isoformat(),
            "filename": filename,
            "error": error_message,
            "line": line_number,
        }
        self.logger.error(f"PARSE_ERROR | {json.dumps(log_entry)}")
        return log_entry
    
    def log_excel_generated(self, filename, rows_in_excel, file_size_bytes):
        """Log: Excel foi gerado"""
        log_entry = {
            "event": "EXCEL_GENERATED",
            "timestamp": datetime.now().isoformat(),
            "output_filename": filename,
            "rows": rows_in_excel,
            "file_size_bytes": file_size_bytes,
        }
        self.logger.info(f"EXCEL_GENERATED | {json.dumps(log_entry)}")
        return log_entry
    
    def log_download(self, user_email, filename, output_filename):
        """Log: user fez download"""
        log_entry = {
            "event": "DOWNLOAD",
            "timestamp": datetime.now().isoformat(),
            "user": user_email,
            "input_file": filename,
            "output_file": output_filename,
        }
        self.logger.info(f"DOWNLOAD | {json.dumps(log_entry)}")
        return log_entry
    
    def log_validation(self, user_email, filename, validation_status, comments=None):
        """Log: user validou dados antes de download"""
        log_entry = {
            "event": "USER_VALIDATION",
            "timestamp": datetime.now().isoformat(),
            "user": user_email,
            "filename": filename,
            "status": validation_status,  # "APPROVED" ou "REJECTED"
            "comments": comments,
        }
        self.logger.info(f"USER_VALIDATION | {json.dumps(log_entry)}")
        return log_entry
    
    def log_financial_control(self, user_email, rows_grouped, cpo_count, spo_count):
        """Log: Financial Control processing"""
        log_entry = {
            "event": "FINANCIAL_CONTROL",
            "timestamp": datetime.now().isoformat(),
            "user": user_email,
            "rows_processed": rows_grouped,
            "cpo_count": cpo_count,
            "spo_count": spo_count,
        }
        self.logger.info(f"FINANCIAL_CONTROL | {json.dumps(log_entry)}")
        return log_entry
    
    def get_audit_trail(self, filename=None, user_email=None):
        """Get audit trail for specific file or user"""
        entries = []
        
        try:
            with open(self.audit_file, 'r') as f:
                for line in f:
                    if filename and filename not in line:
                        continue
                    if user_email and user_email not in line:
                        continue
                    entries.append(line.strip())
        except FileNotFoundError:
            pass
        
        return entries
    
    def get_error_summary(self, hours=24):
        """Get error summary from last N hours"""
        errors = []
        
        try:
            with open(self.error_file, 'r') as f:
                for line in f:
                    # Parse timestamp and filter
                    # (simplified version)
                    errors.append(line.strip())
        except FileNotFoundError:
            pass
        
        return errors


# ============================================================================
# EXEMPLO DE USO (integrar em rovo_v2.py)
# ============================================================================

"""
# No topo do rovo_v2.py:
from logging_config import AuditLogger

# Inicializar
audit = AuditLogger()

# Quando user faz upload:
def handle_upload(uploaded_file, user_email, client):
    audit.log_upload(
        user_email=user_email,
        filename=uploaded_file.name,
        file_size=uploaded_file.size,
        client=client
    )
    
    # ... parsing ...
    
    # Se sucesso:
    audit.log_parse_success(
        filename=uploaded_file.name,
        rows_extracted=len(rows),
        models_found=set(r['model'] for r in rows),
        colors_found=set(r['color'] for r in rows)
    )
    
    # Se erro:
    audit.log_parse_error(
        filename=uploaded_file.name,
        error_message="Invalid format detected",
        line_number=42
    )
    
    # Após Excel gerado:
    audit.log_excel_generated(
        filename="IMPORT_Stussy.xlsx",
        rows_in_excel=len(rows),
        file_size_bytes=len(excel_bytes)
    )
    
    # Quando user faz download:
    audit.log_download(
        user_email=user_email,
        filename=uploaded_file.name,
        output_filename="IMPORT_Stussy.xlsx"
    )

# Para ver audit trail:
trail = audit.get_audit_trail(filename="order-001.xlsx")
for entry in trail:
    print(entry)

# Para ver erros:
errors = audit.get_error_summary(hours=24)
for error in errors:
    print(error)
"""


class PerformanceMonitor:
    """Monitor performance e alertas"""
    
    def __init__(self, log_file="./logs/performance.log"):
        self.log_file = log_file
        Path(self.log_file).parent.mkdir(exist_ok=True)
    
    def log_parse_time(self, filename, seconds, rows_processed):
        """Log tempo de parsing"""
        throughput = rows_processed / seconds if seconds > 0 else 0
        
        log_entry = {
            "event": "PARSE_PERFORMANCE",
            "filename": filename,
            "duration_seconds": round(seconds, 2),
            "rows_processed": rows_processed,
            "throughput_rows_per_sec": round(throughput, 2),
        }
        
        with open(self.log_file, 'a') as f:
            f.write(json.dumps(log_entry) + "\n")
        
        # Alert se muito lento
        if seconds > 30:
            print(f"⚠️ WARNING: Parsing levou {seconds:.1f}s (normal: <10s)")
    
    def log_file_size_warning(self, filename, file_size_mb):
        """Alert se ficheiro é muito grande"""
        if file_size_mb > 50:
            print(f"⚠️ WARNING: File {filename} é {file_size_mb}MB (recomendado: <50MB)")


# ============================================================================
# DASHBOARD/SUMMARY (pode ser usado em Streamlit tab)
# ============================================================================

def generate_audit_summary(audit_logger, days=7):
    """Gera summary do último N dias"""
    
    summary = {
        "period_days": days,
        "generated_at": datetime.now().isoformat(),
        "stats": {
            "total_uploads": 0,
            "total_downloads": 0,
            "total_errors": 0,
            "unique_users": set(),
            "clients_processed": set(),
        }
    }
    
    trail = audit_logger.get_audit_trail()
    
    for entry in trail:
        if "FILE_UPLOAD" in entry:
            summary["stats"]["total_uploads"] += 1
        if "DOWNLOAD" in entry:
            summary["stats"]["total_downloads"] += 1
        if "PARSE_ERROR" in entry:
            summary["stats"]["total_errors"] += 1
    
    # Convert sets to lists for JSON serialization
    summary["stats"]["unique_users"] = list(summary["stats"]["unique_users"])
    summary["stats"]["clients_processed"] = list(summary["stats"]["clients_processed"])
    
    return summary


if __name__ == "__main__":
    # Teste
    audit = AuditLogger()
    
    # Simular alguns eventos
    audit.log_upload("maria@company.com", "order-001.xlsx", 45000, "Stussy")
    audit.log_parse_success("order-001.xlsx", 50, {"SNW-001", "SNW-002"}, {"Navy", "Black"})
    audit.log_excel_generated("IMPORT_Stussy.xlsx", 50, 28000)
    audit.log_download("maria@company.com", "order-001.xlsx", "IMPORT_Stussy.xlsx")
    
    print("\n✅ Logging initialized successfully")
    print(f"📋 Audit log: {audit.audit_file}")
    print(f"❌ Error log: {audit.error_file}")
    print(f"⚡ Performance log: {audit.performance_file}")
