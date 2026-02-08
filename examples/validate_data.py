"""
Example script: Validate material data before processing
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from material_master import MaterialMasterAutomation
import pandas as pd


def main():
    """Example: Validate material data from CSV"""
    
    print("="*60)
    print("SAP Material Master - Data Validation Example")
    print("="*60)
    
    # Path to input file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    input_file = os.path.join(project_root, 'sample_data', 'material_master_template.csv')
    
    # Create automation instance (no SAP connection needed for validation)
    automation = MaterialMasterAutomation()
    
    try:
        # Read input file
        print(f"\nReading data from: {input_file}\n")
        df = automation.read_input_file(input_file)
        
        print(f"Total records found: {len(df)}\n")
        print("Validating records...")
        print("-"*60)
        
        valid_count = 0
        invalid_count = 0
        
        # Validate each record
        for idx, row in df.iterrows():
            material_data = row.to_dict()
            record_num = idx + 1
            
            is_valid, errors = automation.validate_material_data(material_data)
            
            if is_valid:
                print(f"✓ Record {record_num}: VALID")
                print(f"  Description: {material_data.get('Description', 'N/A')}")
                print(f"  Type: {material_data.get('Material_Type', 'N/A')}")
                valid_count += 1
            else:
                print(f"✗ Record {record_num}: INVALID")
                print(f"  Description: {material_data.get('Description', 'N/A')}")
                print(f"  Errors:")
                for error in errors:
                    print(f"    - {error}")
                invalid_count += 1
            print()
        
        # Summary
        print("="*60)
        print("VALIDATION SUMMARY")
        print("="*60)
        print(f"Total Records: {len(df)}")
        print(f"Valid: {valid_count}")
        print(f"Invalid: {invalid_count}")
        print("="*60)
        
        if invalid_count > 0:
            print("\n⚠ Please fix validation errors before processing.")
            sys.exit(1)
        else:
            print("\n✓ All records are valid and ready for processing.")
            sys.exit(0)
        
    except FileNotFoundError:
        print(f"Error: Input file not found: {input_file}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
