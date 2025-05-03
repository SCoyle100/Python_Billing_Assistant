import os
import sys
import unittest
from unittest.mock import patch, MagicMock

# Add parent directory to path for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.pdf_utils import combine_pdfs, combine_vendor_pdfs

class TestPdfUtils(unittest.TestCase):
    
    @patch('utils.pdf_utils.PdfMerger')
    @patch('utils.pdf_utils.os.makedirs')
    @patch('utils.pdf_utils.logging')
    def test_combine_pdfs(self, mock_logging, mock_makedirs, mock_pdf_merger):
        # Setup mock PdfMerger
        merger_instance = MagicMock()
        mock_pdf_merger.return_value = merger_instance
        
        # Test with multiple PDF files
        pdf_files = ['/path/to/pdf1.pdf', '/path/to/pdf2.pdf', '/path/to/pdf3.pdf']
        output_path = '/path/to/output.pdf'
        
        result = combine_pdfs(pdf_files, output_path)
        
        # Verify expected calls
        self.assertEqual(result, output_path)
        mock_makedirs.assert_called_once()
        self.assertEqual(merger_instance.append.call_count, 3)
        merger_instance.write.assert_called_once_with(output_path)
        merger_instance.close.assert_called_once()
        
        # Test with one PDF file
        pdf_files = ['/path/to/pdf1.pdf']
        result = combine_pdfs(pdf_files, output_path)
        
        # Should return the single PDF path without combining
        self.assertEqual(result, pdf_files[0])
        
        # Test with empty list
        pdf_files = []
        result = combine_pdfs(pdf_files, output_path)
        
        # Should return None
        self.assertIsNone(result)
    
    @patch('utils.pdf_utils.combine_pdfs')
    @patch('utils.pdf_utils.os.path.exists')
    @patch('utils.pdf_utils.logging')
    def test_combine_vendor_pdfs(self, mock_logging, mock_path_exists, mock_combine_pdfs):
        # Setup mocks
        mock_path_exists.return_value = True
        mock_combine_pdfs.return_value = '/directory/Combined_Vendor.pdf'
        
        # Test with multiple vendor PDFs
        directory = '/directory'
        vendor_name = 'Matrix Media'
        vendor_map = {
            'file1.pdf': 'Matrix Media',
            'file2.pdf': 'Matrix Media',
            'file3.pdf': 'Capitol Media',
            'file4.pdf': 'Matrix Media'
        }
        
        result = combine_vendor_pdfs(directory, vendor_name, vendor_map)
        
        # Should call combine_pdfs with the correct files
        expected_files = [
            '/directory/file1.pdf',
            '/directory/file2.pdf',
            '/directory/file4.pdf'
        ]
        expected_files.sort()  # Files should be sorted alphabetically
        
        mock_combine_pdfs.assert_called_once()
        called_args = mock_combine_pdfs.call_args[0]
        self.assertEqual(sorted(called_args[0]), sorted(expected_files))
        self.assertEqual(called_args[1], '/directory/Combined_Matrix_Media.pdf')
        
        # Test with only one vendor PDF
        vendor_map = {
            'file1.pdf': 'Matrix Media',
            'file2.pdf': 'Capitol Media',
        }
        
        mock_combine_pdfs.reset_mock()
        result = combine_vendor_pdfs(directory, vendor_name, vendor_map)
        
        # Should not call combine_pdfs for a single PDF
        mock_combine_pdfs.assert_not_called()
        
        # Test with no vendor PDFs
        vendor_map = {
            'file1.pdf': 'Capitol Media',
            'file2.pdf': 'Capitol Media',
        }
        
        result = combine_vendor_pdfs(directory, vendor_name, vendor_map)
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()