import json
import os

class PayslipConfig:
    def __init__(self, config_file='payslip_config.json'):
        self.config_file = config_file
        self.config = self._load_config()

    def _load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except:
                return self._get_default_config()
        return self._get_default_config()

    def _get_default_config(self):
        return {
            'FIXED': {
                'earnings': {},
                'deductions': {}
            },
            'FTC': {
                'earnings': {},
                'deductions': {}
            },
            'payslip_month': None  # Add default month setting
        }

    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=4)

    def add_mapping(self, sheet_type, mapping_type, display_name, excel_header):
        """Add a new mapping for earnings or deductions.
        display_name and excel_header can be strings (single column) or lists/tuples of two (double column).
        """
        if sheet_type not in self.config:
            self.config[sheet_type] = {'earnings': {}, 'deductions': {}}
        # Accept comma-separated string for double columns (for backward compatibility)
        if isinstance(excel_header, str) and ',' in excel_header:
            excel_header = [h.strip() for h in excel_header.split(',', 1)]
        if isinstance(display_name, str) and ',' in display_name:
            display_name = [d.strip() for d in display_name.split(',', 1)]
        self.config[sheet_type][mapping_type][str(display_name) if isinstance(display_name, list) else display_name] = excel_header
        self.save_config()

    def remove_mapping(self, sheet_type, mapping_type, display_name):
        """Remove a mapping"""
        if sheet_type in self.config and mapping_type in self.config[sheet_type]:
            # Try both string and list/tuple keys
            self.config[sheet_type][mapping_type].pop(display_name, None)
            if isinstance(display_name, list):
                self.config[sheet_type][mapping_type].pop(str(display_name), None)
            self.save_config()

    def get_mappings(self, sheet_type):
        """Get all mappings for a sheet type"""
        return self.config.get(sheet_type, {'earnings': {}, 'deductions': {}})

    def set_payslip_month(self, month):
        """Set the payslip month"""
        self.config['payslip_month'] = month
        self.save_config()

    def get_payslip_month(self):
        """Get the current payslip month"""
        return self.config.get('payslip_month', None)
