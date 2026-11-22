# Andhra Pradesh License Plate Recognition and Regional Information Extraction

An Automatic License Plate Recognition (ALPR) system that combines Computer Vision, Optical Character Recognition (OCR), and NLP-based text processing to detect vehicle registration numbers and extract associated Andhra Pradesh registration information.

## Project Overview

This project extends an existing Automatic License Plate Recognition system by introducing an additional text-processing and regional information extraction module for Andhra Pradesh vehicle registrations.

The system processes vehicle images, detects license plates, extracts registration text using OCR, and identifies registration components such as state code, RTO code, registration series, and vehicle number.

## Features

### Existing Computer Vision Features
- Vehicle image processing using OpenCV.
- Grayscale conversion and image preprocessing.
- Canny edge detection.
- Contour-based license plate localization.
- License plate image extraction.
- Image annotation and output generation.

### OCR Features
- Text extraction using Tesseract OCR.
- OCR preprocessing and text recognition.
- Handling of OCR recognition errors.

### NLP and Information Extraction
- Registration text normalization.
- Tokenization and registration component extraction.
- State code identification.
- RTO code extraction.
- Registration series and number extraction.
- Registration format validation.
- Context-aware OCR correction.

### Andhra Pradesh Regional Information Extraction
- Andhra Pradesh registration code identification.
- RTO information lookup using a structured CSV dataset.
- Historical and current registration mapping distinction.
- Statewide registration series handling.
- Mapping verification status.

## System Architecture

Vehicle Image  
↓  
OpenCV Image Processing  
↓  
License Plate Detection  
↓  
Tesseract OCR  
↓  
NLP Text Processing  
↓  
Registration Entity Extraction  
↓  
AP Registration Validation  
↓  
AP RTO Dataset Lookup  
↓  
Structured Registration Information

## Technologies Used

- Python
- OpenCV
- Tesseract OCR
- pytesseract
- NumPy
- NLP text-processing techniques
- Regular Expressions
- Pandas (if used by the implementation)
- Pytest

## Project Structure

```text
NLP_Project/
│
├── main.py
├── nlp_processor.py
├── rto_lookup.py
│
├── data/
│   └── ap_rto_codes.csv
│
├── tests/
│   ├── test_main.py
│   ├── test_nlp_processor.py
│   └── test_rto_lookup.py
│
├── documentation/
│   ├── License_Plate_Recognition_Presentation.pptx
│   ├── License_Plate_Recognition_Project_Report.docx
│   ├── License_Plate_Recognition_Project_Report.pdf
│   └── images/
│
├── car1.JPG
├── output_plate.png
├── output_result.png
├── tessdata/
├── requirements.txt
├── README.md
└── .gitignore
```

The actual project structure may include additional required files.

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Siddukota81/Automatic-License-Plate-Recognition.git
cd Automatic-License-Plate-Recognition
```

### 2. Create a Virtual Environment

```bash
python -m venv venv
```

Activate it on Windows:

```bash
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Install Tesseract OCR

Install Tesseract OCR and configure its executable path if it is not automatically detected.

## Execution

Run the application:

```bash
python main.py
```

```bash
python main.py --headless
```

Follow the application's input instructions to process a vehicle image.

## Testing

Run the automated test suite:

```bash
pytest -v
```

## Project Deliverables

The academic deliverables are available in the `documentation/` directory:

- [Project Presentation](documentation/License_Plate_Recognition_Presentation.pptx)
- [Project Report (DOCX)](documentation/License_Plate_Recognition_Project_Report.docx)
- [Project Report (PDF)](documentation/License_Plate_Recognition_Project_Report.pdf)

## Dataset

The Andhra Pradesh RTO registration information is maintained in:

`data/ap_rto_codes.csv`

The dataset distinguishes historical registration associations from current registration information wherever verified records are available.

Registration codes identify registration authorities or series; they do not establish a vehicle's current physical location or owner's residential address.

## Limitations

- OCR accuracy depends on image quality, lighting, angle, and plate visibility.
- Uncertain OCR characters may require manual verification.
- Regional lookup is limited to the available verified dataset.
- Historical registration codes should not automatically be interpreted as current RTO office assignments.
- The system does not access private vehicle-owner information or live vehicle tracking data.

## Future Enhancements

- Improved license plate detection using deep learning.
- More robust OCR correction.
- Expanded verified regional registration datasets.
- Graphical user interface.
- Improved evaluation using larger vehicle image datasets.

## Author

Developed as an academic project integrating Computer Vision, OCR, and NLP-based information extraction.
