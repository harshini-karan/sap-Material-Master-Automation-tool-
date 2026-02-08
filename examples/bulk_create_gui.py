"""
Example script: Bulk create materials from CSV file using SAP GUI
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from material_master import MaterialMasterAutomation


def main():
    """Example: Bulk create materials using SAP GUI scripting"""
    
    print("="*60)
    print("SAP Material Master - Bulk Creation Example (GUI)")
    print("="*60)
    
    # Path to input file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    input_file = os.path.join(project_root, 'sample_data', 'material_master_template.csv')
    
    # Create automation instance
    automation = MaterialMasterAutomation(config_file='../config.ini')
    
    try:
        # Process materials using GUI method
        print(f"\nProcessing materials from: {input_file}")
        print("Method: SAP GUI Scripting\n")
        
        summary = automation.process_materials(input_file, method='gui')
        
        # Display results
        print("\n" + "="*60)
        print("PROCESSING RESULTS")
        print("="*60)
        print(f"Total Records Processed: {summary['total']}")
        print(f"Successful: {summary['success']}")
        print(f"Failed: {summary['failed']}")
        print(f"Timestamp: {summary['timestamp']}")
        print("="*60)
        
        # Show individual results
        print("\nDetailed Results:")
        for result in summary['results']:
            status_icon = "✓" if result['status'] == 'success' else "✗"
            print(f"{status_icon} Record {result['record']}: {result['status'].upper()}")
            print(f"  Description: {result['data'].get('Description', 'N/A')}")
            print(f"  Message: {result['message']}")
            print()
        
    except FileNotFoundError:
        print(f"Error: Input file not found: {input_file}")
        print("Please ensure the file exists or update the path.")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)
    finally:
        automation.disconnect()


if __name__ == '__main__':
    main()
