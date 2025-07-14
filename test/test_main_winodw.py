import unittest
import os
import sys

# Get the absolute path of the current file
current_file_path = os.path.abspath(__file__)
# Get the project root directory
project_root = os.path.abspath(os.path.join(current_file_path, '..', '..', '..'))
# Add project root to sys.path
sys.path.insert(0, project_root)

from unittest.mock import MagicMock, patch
from PySide6.QtWidgets import QApplication



# Import the class to test
from src.gui.main_window import Worker, Ui_Form

class MainWindowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Initialize QApplication once for all tests
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        # Create a mock Ui_Form instance
        self.form = Ui_Form()
        self.form.worker = Worker()

    @patch('src.gui.main_window.image_to_excel')
    @patch('src.gui.main_window.os.path.exists')
    @patch('src.gui.main_window.os.makedirs')
    def test_temp_store_photo_inputs(self, mock_makedirs, mock_exists, mock_image_to_excel):
        """Test storing photo inputs functionality"""
        # Setup mock behaviors
        mock_exists.return_value = True  # Assume paths exist
        
        # Create test photo paths
        test_photos = [
            os.path.join(project_root, "test_photo1.jpg"),
            os.path.join(project_root, "test_photo2.jpg")
        ]
        
        # Call the method (assuming it takes photo paths as input)
        self.form.temp_store_photo_inputs(test_photos)
        
        # Verify the photos were processed
        self.assertEqual(mock_image_to_excel.call_count, len(test_photos))
        
        # Verify each photo was processed with the correct path
        for photo in test_photos:
            mock_image_to_excel.assert_any_call(
                photo,
                save_folder=os.path.join(project_root, "src", "data", "input", "manual")
            )
        
        # Verify directories were checked/created
        mock_makedirs.assert_called_once_with(
            os.path.join(project_root, "src", "data", "input", "manual"),
            exist_ok=True
        )

    @classmethod
    def tearDownClass(cls):
        # Clean up QApplication
        if cls.app:
            cls.app.quit()

if __name__ == '__main__':
    unittest.main()