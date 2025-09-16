"""
Custom exceptions for ZERF Automation System
"""

class ZERFError(Exception):
    """Base exception for ZERF Automation System"""
    pass

class ConfigurationError(ZERFError):
    """Raised when configuration is invalid or missing"""
    pass

class SAPConnectionError(ZERFError):
    """Raised when SAP connection fails"""
    pass

class VBSScriptError(ZERFError):
    """Raised when VBS script execution fails"""
    pass

class FileNotFoundError(ZERFError):
    """Raised when expected files are not found"""
    pass

class DataProcessingError(ZERFError):
    """Raised when data processing fails"""
    pass

class SharePointError(ZERFError):
    """Raised when SharePoint operations fail"""
    pass

class AuthenticationError(ZERFError):
    """Raised when authentication fails"""
    pass

class ValidationError(ZERFError):
    """Raised when data validation fails"""
    pass

class TimeoutError(ZERFError):
    """Raised when operations timeout"""
    pass

class WorkflowError(ZERFError):
    """Raised when workflow execution fails"""
    
    def __init__(self, message: str, step: str = None, details: dict = None):
        super().__init__(message)
        self.step = step
        self.details = details or {}
    
    def __str__(self):
        if self.step:
            return f"Workflow failed at step '{self.step}': {super().__str__()}"
        return super().__str__()

class RetryableError(ZERFError):
    """Base class for errors that can be retried"""
    pass

class NetworkError(RetryableError):
    """Raised when network operations fail"""
    pass

class TemporaryFileError(RetryableError):
    """Raised when file operations fail temporarily"""
    pass