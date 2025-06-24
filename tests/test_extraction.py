import unittest
import os
from extract_transactions_simple import extract_transactions_simple

class TestPDFExtraction(unittest.TestCase):
    def setUp(self):
        self.sample_pdf = "tests/sample_data/test_file.pdf"
    
    def test_basic_extraction(self):
        """Test basic PDF extraction functionality."""
        if os.path.exists(self.sample_pdf):
            result = extract_transactions_simple(self.sample_pdf)
            self.assertIsNotNone(result)
            self.assertGreater(len(result), 0)
    
    def test_column_cleaning(self):
        """Test that numeric columns are properly cleaned."""
        # Add specific tests for data cleaning logic
        pass

if __name__ == '__main__':
    unittest.main()
