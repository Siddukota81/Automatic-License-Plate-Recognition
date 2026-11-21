import os
import unittest
import cv2
from main import recognize_license_plate, run_ocr


class TestLicensePlateRecognition(unittest.TestCase):

    def test_car1_detection_and_recognition(self):
        """Verify that car1.JPG correctly detects the license plate and OCR extracts 'HR 26 DA 2330'."""
        self.assertTrue(os.path.exists("car1.JPG"), "car1.JPG must exist in the workspace")

        result = recognize_license_plate(image_path="car1.JPG", show_gui=False)

        # 1. Bounding box check
        self.assertIsNotNone(result['bbox'], "A 4-corner bounding box must be detected")
        x, y, w, h = result['bbox']
        self.assertTrue(60 <= x <= 85, f"Expected x ~ 72, got {x}")
        self.assertTrue(150 <= y <= 180, f"Expected y ~ 167, got {y}")
        self.assertTrue(130 <= w <= 160, f"Expected w ~ 147, got {w}")
        self.assertTrue(40 <= h <= 60, f"Expected h ~ 49, got {h}")

        # 2. License plate text check
        detected_text = result['text']
        self.assertIn("HR", detected_text)
        self.assertIn("26", detected_text)
        self.assertIn("2330", detected_text)

        # 3. Output files check
        self.assertTrue(os.path.exists("output_plate.png"), "output_plate.png must be saved")
        self.assertTrue(os.path.exists("output_result.png"), "output_result.png must be saved")
        self.assertGreater(os.path.getsize("output_plate.png"), 0)
        self.assertGreater(os.path.getsize("output_result.png"), 0)

    def test_run_ocr_direct(self):
        """Verify OCR function directly on the extracted plate."""
        plate = cv2.imread("output_plate.png", cv2.IMREAD_GRAYSCALE)
        self.assertIsNotNone(plate, "output_plate.png must be readable")
        text = run_ocr(plate)
        self.assertIn("HR 26", text)

    def test_missing_image_exception(self):
        """Verify clean exception when input image does not exist."""
        with self.assertRaises(FileNotFoundError):
            recognize_license_plate(image_path="non_existent_image.jpg", show_gui=False)


if __name__ == "__main__":
    unittest.main(verbosity=2)
