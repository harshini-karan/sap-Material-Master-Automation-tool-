"""
SAP Material Master Automation Tool
Main module for automating material master creation, update, and validation
"""

import os
import sys
import time
import logging
import configparser
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import pandas as pd

# Try to import SAP GUI scripting (Windows only)
try:
    import win32com.client
    SAP_GUI_AVAILABLE = True
except ImportError:
    SAP_GUI_AVAILABLE = False
    print("Warning: win32com not available. SAP GUI scripting will not work.")

# Try to import pyrfc for RFC connections
try:
    from pyrfc import Connection
    RFC_AVAILABLE = True
except ImportError:
    RFC_AVAILABLE = False
    print("Warning: pyrfc not available. RFC API functionality will not work.")


class MaterialMasterAutomation:
    """Main class for Material Master automation"""
    
    # Mandatory fields for material master creation
    MANDATORY_FIELDS = [
        'Material_Type',
        'Industry_Sector',
        'Description',
        'Base_Unit'
    ]
    
    # SAP return types for RFC success validation
    RFC_SUCCESS_TYPES = ['', 'S', 'W']  # Empty, Success, Warning
    
    def __init__(self, config_file: str = 'config.ini'):
        """
        Initialize the automation tool
        
        Args:
            config_file: Path to configuration file
        """
        self.config = self._load_config(config_file)
        self.logger = self._setup_logging()
        self.sap_session = None
        self.rfc_connection = None
        self.results = []
        
    def _load_config(self, config_file: str) -> configparser.ConfigParser:
        """Load configuration from file"""
        config = configparser.ConfigParser()
        
        # Use template if config file doesn't exist
        if not os.path.exists(config_file):
            config_file = 'config_template.ini'
            if not os.path.exists(config_file):
                print(f"Warning: Configuration file not found. Using defaults.")
                # Create default config
                config['SAP'] = {
                    'transaction_code': 'MM01',
                    'sap_language': 'EN'
                }
                config['Automation'] = {
                    'delay_between_actions': '0.5',
                    'max_retries': '3'
                }
                config['Logging'] = {
                    'log_level': 'INFO',
                    'log_file': 'logs/material_master.log'
                }
                return config
        
        config.read(config_file)
        return config
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging configuration"""
        log_level = self.config.get('Logging', 'log_level', fallback='INFO')
        log_file = self.config.get('Logging', 'log_file', fallback='logs/material_master.log')
        
        # Create logs directory if it doesn't exist
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        
        # Configure logging
        logging.basicConfig(
            level=getattr(logging, log_level),
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        
        return logging.getLogger(__name__)
    
    def connect_sap_gui(self) -> bool:
        """
        Connect to SAP GUI
        
        Returns:
            True if connection successful, False otherwise
        """
        if not SAP_GUI_AVAILABLE:
            self.logger.error("SAP GUI scripting not available (win32com not installed)")
            return False
        
        try:
            self.logger.info("Connecting to SAP GUI...")
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
            connection = application.Children(0)
            self.sap_session = connection.Children(0)
            self.logger.info("Successfully connected to SAP GUI")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect to SAP GUI: {str(e)}")
            return False
    
    def connect_rfc(self) -> bool:
        """
        Connect to SAP using RFC
        
        Returns:
            True if connection successful, False otherwise
        """
        if not RFC_AVAILABLE:
            self.logger.error("RFC not available (pyrfc not installed)")
            return False
        
        try:
            self.logger.info("Connecting to SAP via RFC...")
            rfc_config = {
                'ashost': self.config.get('RFC', 'ashost', fallback=''),
                'sysnr': self.config.get('RFC', 'sysnr', fallback='00'),
                'client': self.config.get('RFC', 'client', fallback='100'),
                'user': self.config.get('RFC', 'user', fallback=''),
                'passwd': self.config.get('RFC', 'passwd', fallback=''),
                'lang': self.config.get('RFC', 'lang', fallback='EN')
            }
            
            # Check if RFC config is complete
            if not all([rfc_config['ashost'], rfc_config['user'], rfc_config['passwd']]):
                self.logger.warning("RFC configuration incomplete")
                return False
            
            self.rfc_connection = Connection(**rfc_config)
            self.logger.info("Successfully connected to SAP via RFC")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect via RFC: {str(e)}")
            return False
    
    def validate_material_data(self, material_data: Dict) -> Tuple[bool, List[str]]:
        """
        Validate material master data
        
        Args:
            material_data: Dictionary containing material data
            
        Returns:
            Tuple of (is_valid, list_of_errors)
        """
        errors = []
        
        # Check mandatory fields
        for field in self.MANDATORY_FIELDS:
            if field not in material_data or not str(material_data[field]).strip():
                errors.append(f"Mandatory field '{field}' is missing or empty")
        
        # Validate material type
        valid_material_types = ['FERT', 'ROH', 'HALB', 'HAWA', 'VERP']
        if 'Material_Type' in material_data:
            mat_type = str(material_data['Material_Type']).upper()
            if mat_type and mat_type not in valid_material_types:
                errors.append(f"Invalid Material Type: {mat_type}")
        
        # Validate industry sector
        valid_sectors = ['M', 'C', 'P', 'A']
        if 'Industry_Sector' in material_data:
            sector = str(material_data['Industry_Sector']).upper()
            if sector and sector not in valid_sectors:
                errors.append(f"Invalid Industry Sector: {sector}")
        
        # Validate base unit
        if 'Base_Unit' in material_data:
            base_unit = str(material_data['Base_Unit']).strip()
            if base_unit and len(base_unit) > 3:
                errors.append(f"Base Unit too long: {base_unit}")
        
        # Validate price if provided
        if 'Price' in material_data and material_data['Price']:
            try:
                price = float(material_data['Price'])
                if price < 0:
                    errors.append("Price cannot be negative")
            except (ValueError, TypeError):
                errors.append(f"Invalid price value: {material_data['Price']}")
        
        is_valid = len(errors) == 0
        return is_valid, errors
    
    def read_input_file(self, file_path: str) -> pd.DataFrame:
        """
        Read input file (CSV or Excel)
        
        Args:
            file_path: Path to input file
            
        Returns:
            DataFrame with material data
        """
        self.logger.info(f"Reading input file: {file_path}")
        
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                raise ValueError(f"Unsupported file format: {file_path}")
            
            self.logger.info(f"Successfully read {len(df)} records from file")
            return df
        except Exception as e:
            self.logger.error(f"Error reading file: {str(e)}")
            raise
    
    def create_material_gui(self, material_data: Dict) -> Tuple[bool, str]:
        """
        Create material using SAP GUI scripting
        
        Args:
            material_data: Dictionary containing material data
            
        Returns:
            Tuple of (success, message)
        """
        if not self.sap_session:
            return False, "SAP GUI session not connected"
        
        try:
            delay = float(self.config.get('Automation', 'delay_between_actions', fallback='0.5'))
            transaction = self.config.get('SAP', 'transaction_code', fallback='MM01')
            
            # Start transaction
            self.logger.info(f"Starting transaction {transaction}")
            self.sap_session.StartTransaction(transaction)
            time.sleep(delay)
            
            # Fill in material type
            self.sap_session.findById("wnd[0]/usr/ctxtRMMG1-MBRSH").text = material_data.get('Industry_Sector', 'M')
            self.sap_session.findById("wnd[0]/usr/ctxtRMMG1-MTART").text = material_data.get('Material_Type', '')
            time.sleep(delay)
            
            # Press Enter
            self.sap_session.findById("wnd[0]").sendVKey(0)
            time.sleep(delay)
            
            # Fill in description
            self.sap_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpOVERVIEW/ssubSUBSCREEN_BODY:SAPLMGMM:2100/subSUB_VIEWSET:SAPLMGMM:2200/ctxtMAKT-MAKTX").text = material_data.get('Description', '')
            
            # Fill in base unit
            self.sap_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpOVERVIEW/ssubSUBSCREEN_BODY:SAPLMGMM:2100/subSUB_VIEWSET:SAPLMGMM:2200/ctxtMARA-MEINS").text = material_data.get('Base_Unit', '')
            
            # Fill in material group if provided
            if material_data.get('Material_Group'):
                self.sap_session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpOVERVIEW/ssubSUBSCREEN_BODY:SAPLMGMM:2100/subSUB_VIEWSET:SAPLMGMM:2200/ctxtMARA-MATKL").text = str(material_data['Material_Group'])
            
            time.sleep(delay)
            
            # Save (this is a simulation - actual field IDs may vary)
            self.sap_session.findById("wnd[0]").sendVKey(11)  # Ctrl+S
            time.sleep(delay * 2)
            
            # Get material number from status bar
            status_text = self.sap_session.findById("wnd[0]/sbar").text
            
            self.logger.info(f"Material created successfully: {status_text}")
            return True, status_text
            
        except Exception as e:
            error_msg = f"Error creating material via GUI: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def create_material_rfc(self, material_data: Dict) -> Tuple[bool, str]:
        """
        Create material using RFC API
        
        Args:
            material_data: Dictionary containing material data
            
        Returns:
            Tuple of (success, message)
        """
        if not self.rfc_connection:
            return False, "RFC connection not established"
        
        try:
            # Call BAPI_MATERIAL_SAVEDATA or similar
            # Note: This is a simplified example. Actual implementation requires
            # proper BAPI structure understanding
            
            result = self.rfc_connection.call(
                'BAPI_MATERIAL_SAVEDATA',
                HEADDATA={
                    'MATERIAL': material_data.get('Material_Number', ''),
                    'IND_SECTOR': material_data.get('Industry_Sector', 'M'),
                    'MATL_TYPE': material_data.get('Material_Type', ''),
                    'BASIC_VIEW': 'X'
                },
                CLIENTDATA={
                    'BASE_UOM': material_data.get('Base_Unit', '')
                },
                MATERIALDESCRIPTION={
                    'LANGU': 'EN',
                    'MATL_DESC': material_data.get('Description', '')
                }
            )
            
            if result.get('RETURN', {}).get('TYPE') in self.RFC_SUCCESS_TYPES:
                material_number = result.get('NUMBER', '')
                self.logger.info(f"Material created successfully via RFC: {material_number}")
                return True, f"Material {material_number} created successfully"
            else:
                error_msg = result.get('RETURN', {}).get('MESSAGE', 'Unknown error')
                self.logger.error(f"RFC error: {error_msg}")
                return False, error_msg
                
        except Exception as e:
            error_msg = f"Error creating material via RFC: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def process_materials(self, file_path: str, method: str = 'gui') -> Dict:
        """
        Process materials from input file
        
        Args:
            file_path: Path to input CSV or Excel file
            method: 'gui' for SAP GUI scripting or 'rfc' for RFC API
            
        Returns:
            Dictionary with processing results
        """
        self.logger.info(f"Starting material processing from {file_path} using {method} method")
        
        # Read input file
        try:
            df = self.read_input_file(file_path)
        except Exception as e:
            return {
                'status': 'error',
                'message': f"Failed to read input file: {str(e)}",
                'total': 0,
                'success': 0,
                'failed': 0,
                'results': []
            }
        
        # Connect to SAP
        if method == 'gui':
            if not self.connect_sap_gui():
                return {
                    'status': 'error',
                    'message': 'Failed to connect to SAP GUI',
                    'total': len(df),
                    'success': 0,
                    'failed': 0,
                    'results': []
                }
        elif method == 'rfc':
            if not self.connect_rfc():
                return {
                    'status': 'error',
                    'message': 'Failed to connect via RFC',
                    'total': len(df),
                    'success': 0,
                    'failed': 0,
                    'results': []
                }
        
        # Process each material
        results = []
        success_count = 0
        failed_count = 0
        
        for idx, row in df.iterrows():
            material_data = row.to_dict()
            record_num = idx + 1
            
            self.logger.info(f"Processing record {record_num}/{len(df)}")
            
            # Validate data
            is_valid, errors = self.validate_material_data(material_data)
            
            if not is_valid:
                result = {
                    'record': record_num,
                    'status': 'failed',
                    'message': f"Validation failed: {'; '.join(errors)}",
                    'data': material_data
                }
                self.logger.warning(f"Record {record_num} validation failed: {errors}")
                results.append(result)
                failed_count += 1
                continue
            
            # Create material
            if method == 'gui':
                success, message = self.create_material_gui(material_data)
            else:  # rfc
                success, message = self.create_material_rfc(material_data)
            
            result = {
                'record': record_num,
                'status': 'success' if success else 'failed',
                'message': message,
                'data': material_data
            }
            results.append(result)
            
            if success:
                success_count += 1
            else:
                failed_count += 1
        
        # Generate summary
        summary = {
            'status': 'completed',
            'total': len(df),
            'success': success_count,
            'failed': failed_count,
            'results': results,
            'timestamp': datetime.now().isoformat()
        }
        
        self.logger.info(f"Processing completed. Success: {success_count}, Failed: {failed_count}")
        
        # Save results to file
        self._save_results(summary)
        
        return summary
    
    def _save_results(self, summary: Dict):
        """Save processing results to file"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"logs/results_{timestamp}.csv"
        
        try:
            results_df = pd.DataFrame(summary['results'])
            results_df.to_csv(output_file, index=False)
            self.logger.info(f"Results saved to {output_file}")
        except Exception as e:
            self.logger.error(f"Failed to save results: {str(e)}")
    
    def disconnect(self):
        """Disconnect from SAP"""
        if self.sap_session:
            self.logger.info("Disconnecting from SAP GUI")
            self.sap_session = None
        
        if self.rfc_connection:
            self.logger.info("Closing RFC connection")
            self.rfc_connection.close()
            self.rfc_connection = None


def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(description='SAP Material Master Automation Tool')
    parser.add_argument('input_file', help='Path to input CSV or Excel file')
    parser.add_argument('--method', choices=['gui', 'rfc'], default='gui',
                       help='Method to use: gui (SAP GUI scripting) or rfc (RFC API)')
    parser.add_argument('--config', default='config.ini',
                       help='Path to configuration file')
    
    args = parser.parse_args()
    
    # Create automation instance
    automation = MaterialMasterAutomation(config_file=args.config)
    
    try:
        # Process materials
        summary = automation.process_materials(args.input_file, method=args.method)
        
        # Print summary
        print("\n" + "="*50)
        print("PROCESSING SUMMARY")
        print("="*50)
        print(f"Total Records: {summary['total']}")
        print(f"Successful: {summary['success']}")
        print(f"Failed: {summary['failed']}")
        print(f"Status: {summary['status']}")
        print("="*50)
        
        # Exit with appropriate code
        sys.exit(0 if summary['failed'] == 0 else 1)
        
    except KeyboardInterrupt:
        print("\nProcess interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)
    finally:
        automation.disconnect()


if __name__ == '__main__':
    main()
