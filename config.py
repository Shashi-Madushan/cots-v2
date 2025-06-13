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
            }
        }

    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=4)

    def add_mapping(self, sheet_type, mapping_type, display_name, excel_header):
        """Add a new mapping for earnings or deductions"""
        if sheet_type not in self.config:
            self.config[sheet_type] = {'earnings': {}, 'deductions': {}}
        self.config[sheet_type][mapping_type][display_name] = excel_header
        self.save_config()

    def remove_mapping(self, sheet_type, mapping_type, display_name):
        """Remove a mapping"""
        if sheet_type in self.config and mapping_type in self.config[sheet_type]:
            self.config[sheet_type][mapping_type].pop(display_name, None)
            self.save_config()

    def get_mappings(self, sheet_type):
        """Get all mappings for a sheet type"""
        return self.config.get(sheet_type, {'earnings': {}, 'deductions': {}})
