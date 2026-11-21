# Automatic Number Plate Recognition (ANPR) System

This project is part of the **30 Days AI Challenge (Day 20)** covering computer vision and optical character recognition (OCR) for vehicle license plates.

---

## 📌 Project Overview

The project extracts vehicle license plates from car photographs using OpenCV image processing and recognizes the characters using Tesseract OCR.

### Pipeline Workflow:
1. **Image Ingestion**: Loads the input vehicle image (`car1.JPG`).
2. **Grayscale Conversion**: Converts the 3-channel BGR image to grayscale.
3. **Canny Edge Detection**: Detects edge gradients with intensity thresholds `(170, 200)`.
4. **Contour Extraction**: Identifies closed contours and sorts them by area.
5. **ROI Approximation**: Uses `approxPolyDP` to identify 4-corner rectangular contours corresponding to the license plate.
6. **Thresholding & Filtering**: Binarizes the region of interest (`THRESH_BINARY`) and removes high-frequency noise using a `bilateralFilter`.
7. **OCR Character Recognition**: Extracts alphanumeric plate text using Tesseract OCR (with Page Segmentation Mode `--psm 7`).
8. **Annotation & Export**: Draws a bounding box around the plate, overlays the recognized text, and saves the output to disk.

---

## 🚀 Setup & Installation

### 1. Requirements
Install the required Python packages:
```bash
pip install -r requirements.txt
```

### 2. OCR Engine
- **Embedded Engine**: Includes local language model files in `tessdata/eng.traineddata` powered by `tesserocr` for out-of-the-box execution without system configuration.
- **System Tesseract (Optional)**: If you have installed Tesseract OCR system-wide (e.g. `C:\Program Files\Tesseract-OCR\tesseract.exe`), `main.py` detects and uses it automatically.

---

## 💻 How to Run

### Interactive GUI Mode
Runs the pipeline and displays interactive OpenCV windows for each step:
```bash
python main.py
```
*(Press any key to close the OpenCV display windows).*

### Run with a Custom Image
```bash
python main.py car2.JPG
```

### Headless / Automated Mode (No GUI)
Runs without opening GUI popup windows and saves output images directly to disk:
```bash
python main.py --headless
```

---

## 🧪 Testing

Run the test suite using `pytest`:
```bash
pytest test_main.py -v
```

All tests verify:
- Accurate plate contour and bounding box localization on `car1.JPG`.
- OCR extraction matching expected plate number (`HR 26 DA 2330`).
- Export of `output_plate.png` and `output_result.png`.
- Error handling for missing files.

---

## 📁 Project Structure

```
NLP_Project/
├── main.py              # Main license plate recognition script
├── test_main.py         # Automated pytest test suite
├── requirements.txt     # Python dependencies
├── README.md            # Project documentation
├── car1.JPG             # Sample test image 1 (HR 26 DA 2330)
├── car2.JPG             # Sample test image 2
├── car3.JPG             # Sample test image 3
├── car4.JPG             # Sample test image 4
├── tessdata/            # English OCR trained data
│   └── eng.traineddata
├── output_plate.png     # Cropped & preprocessed license plate
└── output_result.png    # Annotated output image
```
