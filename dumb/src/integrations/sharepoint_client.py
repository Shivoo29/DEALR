"""
Modern SharePoint client using Microsoft Graph API
"""

import requests
import time
from pathlib import Path
from typing import Optional, Dict, Any
from urllib.parse import quote
from tenacity import retry, stop_after_attempt, wait_exponential

try:
    from msal import ConfidentialClientApplication, PublicClientApplication
    MSAL_AVAILABLE = True
except ImportError:
    MSAL_AVAILABLE = False

from ..utils.logger import get_logger
from ..utils.exceptions import SharePointError, AuthenticationError, NetworkError

logger = get_logger(__name__)

class SharePointClient:
    """Modern SharePoint client with Graph API integration"""
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.access_token = None
        self.token_expires_at = 0
        
        # SharePoint configuration
        self.site_url = config_manager.get_sharepoint_url()
        self.username = config_manager.get_sharepoint_username()
        self.password = config_manager.get_sharepoint_password()
        self.folder_path = config_manager.get_sharepoint_folder()
        
        # Extract tenant and site info from URL
        self._parse_sharepoint_url()
        
        if not MSAL_AVAILABLE:
            logger.warning("MSAL library not available - falling back to basic authentication")
    
    def _parse_sharepoint_url(self):
        """Parse SharePoint URL to extract tenant and site information"""
        try:
            if not self.site_url:
                self.tenant_id = None
                self.site_id = None
                return
            
            # Extract tenant from URL like: https://company.sharepoint.com/sites/sitename
            parts = self.site_url.replace('https://', '').split('/')
            if len(parts) >= 3 and 'sharepoint.com' in parts[0]:
                tenant_domain = parts[0].split('.')[0]
                self.tenant_id = f"{tenant_domain}.onmicrosoft.com"
                
                if 'sites' in parts:
                    site_index = parts.index('sites')
                    if len(parts) > site_index + 1:
                        self.site_name = parts[site_index + 1]
                    else:
                        self.site_name = 'root'
                else:
                    self.site_name = 'root'
            else:
                logger.warning(f"Could not parse SharePoint URL: {self.site_url}")
                self.tenant_id = None
                self.site_name = None
                
        except Exception as e:
            logger.error(f"Error parsing SharePoint URL: {e}")
            self.tenant_id = None
            self.site_name = None
    
    def _get_access_token(self) -> Optional[str]:
        """Get access token for Microsoft Graph API"""
        try:
            # Check if current token is still valid
            if self.access_token and time.time() < self.token_expires_at - 300:  # 5 min buffer
                return self.access_token
            
            if not MSAL_AVAILABLE:
                logger.error("MSAL library required for modern authentication")
                return self._fallback_authentication()
            
            # Use device code flow for interactive authentication
            app = PublicClientApplication(
                client_id="14d82eec-204b-4c2f-b7e8-296a70dab67e",  # Microsoft Graph PowerShell
                authority=f"https://login.microsoftonline.com/{self.tenant_id or 'common'}"
            )
            
            # Try to get token silently first
            accounts = app.get_accounts(username=self.username)
            if accounts:
                result = app.acquire_token_silent(
                    scopes=["https://graph.microsoft.com/.default"],
                    account=accounts[0]
                )
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    self.token_expires_at = time.time() + result.get("expires_in", 3600)
                    logger.info("✅ SharePoint token acquired silently")
                    return self.access_token
            
            # If silent acquisition fails, try username/password
            if self.username and self.password:
                result = app.acquire_token_by_username_password(
                    username=self.username,
                    password=self.password,
                    scopes=["https://graph.microsoft.com/.default"]
                )
                
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    self.token_expires_at = time.time() + result.get("expires_in", 3600)
                    logger.info("✅ SharePoint token acquired via credentials")
                    return self.access_token
                else:
                    error_msg = result.get("error_description", "Unknown authentication error")
                    logger.error(f"Authentication failed: {error_msg}")
                    raise AuthenticationError(f"SharePoint authentication failed: {error_msg}")
            
            # Fallback to device code flow
            logger.info("Initiating device code authentication...")
            flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])
            
            if "user_code" not in flow:
                raise AuthenticationError("Failed to create device code flow")
            
            print(flow["message"])
            logger.info("Please complete authentication in your browser")
            
            result = app.acquire_token_by_device_flow(flow)
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                self.token_expires_at = time.time() + result.get("expires_in", 3600)
                logger.info("✅ SharePoint token acquired via device code")
                return self.access_token
            else:
                error_msg = result.get("error_description", "Device code authentication failed")
                raise AuthenticationError(error_msg)
                
        except Exception as e:
            logger.error(f"Token acquisition failed: {e}")
            return self._fallback_authentication()
    
    def _fallback_authentication(self) -> Optional[str]:
        """Fallback authentication method using basic auth"""
        logger.warning("Using fallback authentication - this may not work with modern SharePoint")
        
        if not (self.username and self.password):
            logger.error("No credentials available for fallback authentication")
            return None
        
        try:
            # This is a simplified approach that may not work with all SharePoint configurations
            auth_url = f"https://login.microsoftonline.com/common/oauth2/token"
            
            data = {
                'grant_type': 'password',
                'client_id': '14d82eec-204b-4c2f-b7e8-296a70dab67e',
                'resource': 'https://graph.microsoft.com/',
                'username': self.username,
                'password': self.password,
                'scope': 'https://graph.microsoft.com/.default'
            }
            
            response = requests.post(auth_url, data=data, timeout=30)
            
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data.get('access_token')
                self.token_expires_at = time.time() + token_data.get('expires_in', 3600)
                logger.info("✅ Fallback authentication successful")
                return self.access_token
            else:
                logger.error(f"Fallback authentication failed: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"Fallback authentication error: {e}")
            return None
    
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=10),
        retry_error_callback=lambda retry_state: logger.warning(f"SharePoint upload retry {retry_state.attempt_number}/3")
    )
    def upload_file(self, file_path: Path, custom_filename: str = None) -> bool:
        """Upload file to SharePoint using Microsoft Graph API"""
        try:
            if not file_path.exists():
                raise SharePointError(f"File not found: {file_path}")
            
            logger.info(f"Starting SharePoint upload: {file_path.name}")
            
            # Get access token
            token = self._get_access_token()
            if not token:
                raise AuthenticationError("Failed to get SharePoint access token")
            
            # Get site ID
            site_id = self._get_site_id()
            if not site_id:
                raise SharePointError("Failed to get SharePoint site ID")
            
            # Upload file
            filename = custom_filename or file_path.name
            upload_url = self._build_upload_url(site_id, filename)
            
            success = self._upload_file_content(upload_url, file_path, token)
            
            if success:
                logger.info(f"✅ SharePoint upload successful: {filename}")
                return True
            else:
                raise SharePointError("File upload failed")
                
        except Exception as e:
            logger.error(f"SharePoint upload error: {e}")
            raise NetworkError(f"SharePoint upload failed: {e}")
    
    def _get_site_id(self) -> Optional[str]:
        """Get SharePoint site ID using Graph API"""
        try:
            token = self._get_access_token()
            if not token:
                return None
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Accept': 'application/json'
            }
            
            # Extract hostname and site path from URL
            url_parts = self.site_url.replace('https://', '').split('/')
            hostname = url_parts[0]
            site_path = '/'.join(url_parts[1:]) if len(url_parts) > 1 else ''
            
            # Get site ID
            if site_path:
                api_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
            else:
                api_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}"
            
            response = requests.get(api_url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                site_data = response.json()
                site_id = site_data.get('id')
                logger.debug(f"Site ID retrieved: {site_id}")
                return site_id
            else:
                logger.error(f"Failed to get site ID: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"Error getting site ID: {e}")
            return None
    
    def _build_upload_url(self, site_id: str, filename: str) -> str:
        """Build the upload URL for Microsoft Graph API"""
        # Encode folder path and filename
        encoded_folder = quote(self.folder_path)
        encoded_filename = quote(filename)
        
        # Build the upload URL
        upload_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
            f"{encoded_folder}/{encoded_filename}:/content"
        )
        
        return upload_url
    
    def _upload_file_content(self, upload_url: str, file_path: Path, token: str) -> bool:
        """Upload file content to SharePoint"""
        try:
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/octet-stream'
            }
            
            file_size = file_path.stat().st_size
            logger.info(f"Uploading {file_size} bytes...")
            
            # For large files (>4MB), use resumable upload
            if file_size > 4 * 1024 * 1024:  # 4MB
                return self._upload_large_file(upload_url, file_path, token)
            
            # Simple upload for smaller files
            with open(file_path, 'rb') as f:
                response = requests.put(
                    upload_url,
                    headers=headers,
                    data=f,
                    timeout=300  # 5 minutes for upload
                )
            
            if response.status_code in [200, 201]:
                logger.info("File upload completed successfully")
                return True
            else:
                logger.error(f"Upload failed: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"File upload error: {e}")
            return False
    
    def _upload_large_file(self, upload_url: str, file_path: Path, token: str) -> bool:
        """Upload large file using resumable upload session"""
        try:
            # Create upload session
            session_url = upload_url.replace('/content', '/createUploadSession')
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            session_data = {
                'item': {
                    '@microsoft.graph.conflictBehavior': 'replace'
                }
            }
            
            response = requests.post(session_url, headers=headers, json=session_data, timeout=30)
            
            if response.status_code != 200:
                logger.error(f"Failed to create upload session: {response.status_code}")
                return False
            
            upload_session = response.json()
            upload_url = upload_session.get('uploadUrl')
            
            if not upload_url:
                logger.error("No upload URL in session response")
                return False
            
            # Upload file in chunks
            chunk_size = 10 * 1024 * 1024  # 10MB chunks
            file_size = file_path.stat().st_size
            
            with open(file_path, 'rb') as f:
                bytes_uploaded = 0
                
                while bytes_uploaded < file_size:
                    chunk_start = bytes_uploaded
                    chunk_end = min(bytes_uploaded + chunk_size - 1, file_size - 1)
                    chunk_data = f.read(chunk_end - chunk_start + 1)
                    
                    chunk_headers = {
                        'Content-Length': str(len(chunk_data)),
                        'Content-Range': f'bytes {chunk_start}-{chunk_end}/{file_size}'
                    }
                    
                    chunk_response = requests.put(
                        upload_url,
                        headers=chunk_headers,
                        data=chunk_data,
                        timeout=300
                    )
                    
                    if chunk_response.status_code not in [202, 200, 201]:
                        logger.error(f"Chunk upload failed: {chunk_response.status_code}")
                        return False
                    
                    bytes_uploaded = chunk_end + 1
                    progress = (bytes_uploaded / file_size) * 100
                    logger.info(f"Upload progress: {progress:.1f}%")
            
            logger.info("Large file upload completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Large file upload error: {e}")
            return False
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        try:
            logger.info("Testing SharePoint connection...")
            
            if not self.site_url or not self.username:
                logger.error("SharePoint credentials not configured")
                return False
            
            # Test authentication
            token = self._get_access_token()
            if not token:
                logger.error("❌ SharePoint authentication failed")
                return False
            
            # Test site access
            site_id = self._get_site_id()
            if not site_id:
                logger.error("❌ SharePoint site access failed")
                return False
            
            # Test folder access
            if self.folder_path:
                folder_exists = self._check_folder_exists(site_id, token)
                if not folder_exists:
                    logger.warning(f"⚠️ SharePoint folder may not exist: {self.folder_path}")
                    # Don't fail the test - folder might be created automatically
            
            logger.info("✅ SharePoint connection test passed")
            return True
            
        except Exception as e:
            logger.error(f"SharePoint connection test failed: {e}")
            return False
    
    def _check_folder_exists(self, site_id: str, token: str) -> bool:
        """Check if the target folder exists"""
        try:
            headers = {
                'Authorization': f'Bearer {token}',
                'Accept': 'application/json'
            }
            
            encoded_folder = quote(self.folder_path)
            folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_folder}"
            
            response = requests.get(folder_url, headers=headers, timeout=30)
            return response.status_code == 200
            
        except Exception as e:
            logger.error(f"Error checking folder existence: {e}")
            return False
    
    def get_connection_info(self) -> Dict[str, Any]:
        """Get SharePoint connection information"""
        return {
            'site_url': self.site_url,
            'username': self.username,
            'folder_path': self.folder_path,
            'tenant_id': getattr(self, 'tenant_id', None),
            'site_name': getattr(self, 'site_name', None),
            'msal_available': MSAL_AVAILABLE,
            'token_valid': bool(self.access_token and time.time() < self.token_expires_at)
        }
    
    def create_folder_if_not_exists(self, folder_path: str) -> bool:
        """Create folder structure if it doesn't exist"""
        try:
            token = self._get_access_token()
            site_id = self._get_site_id()
            
            if not (token and site_id):
                return False
            
            # Check if folder exists
            if self._check_folder_exists(site_id, token):
                return True
            
            # Create folder
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            folder_data = {
                'name': folder_path.split('/')[-1],
                'folder': {}
            }
            
            parent_path = '/'.join(folder_path.split('/')[:-1])
            if parent_path:
                create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{quote(parent_path)}:/children"
            else:
                create_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
            
            response = requests.post(create_url, headers=headers, json=folder_data, timeout=30)
            
            if response.status_code in [200, 201]:
                logger.info(f"✅ Folder created: {folder_path}")
                return True
            else:
                logger.error(f"Failed to create folder: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"Error creating folder: {e}")
            return False