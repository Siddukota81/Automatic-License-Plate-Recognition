import os
import sys
from pathlib import Path
import cv2
import numpy as np

# Import NLP normalization and Andhra Pradesh RTO extraction modules
from nlp_processor import process_registration_text, format_registration_summary, ProcessedRegistration
from rto_lookup import lookup_rto, RTOLookupResult

# Configure OCR: support pytesseract if tesseract.exe is present, with seamless fallback to tesserocr
tess_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'

try:
    import pytesseract
    # Check common locations for tesseract.exe
    candidate_paths = [
        tess_cmd,
        'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe',
        os.path.expandvars('%LOCALAPPDATA%/Programs/Tesseract-OCR/tesseract.exe'),
    ]
    for cp in candidate_paths:
        if os.path.exists(cp):
            pytesseract.pytesseract.tesseract_cmd = cp
            break
except ImportError:
    pytesseract = None


def run_ocr(image_array):
    """
    Perform OCR on the cropped license plate image array.
    Tries pytesseract first (with PSM 7 for single-line text).
    If tesseract.exe is not found, seamlessly falls back to tesserocr using local tessdata.
    """
    text_result = ""

    # 1. Try pytesseract if available and executable exists
    if pytesseract is not None:
        try:
            # PSM 7: Treat the image as a single text line
            text_result = pytesseract.image_to_string(image_array, config='--psm 7').strip()
            if text_result:
                return text_result
            # Fallback to default PSM
            text_result = pytesseract.image_to_string(image_array).strip()
            if text_result:
                return text_result
        except Exception:
            # Executable missing or invocation error, proceed to fallback
            pass

    # 2. Seamless fallback: tesserocr (uses embedded C++ Tesseract engine)
    try:
        import tesserocr
        from PIL import Image

        pil_img = Image.fromarray(image_array) if isinstance(image_array, np.ndarray) else image_array
        script_dir = os.path.dirname(os.path.abspath(__file__))
        tessdata_path = os.path.join(script_dir, 'tessdata')

        for psm in [7, 6, 8, 3]:
            with tesserocr.PyTessBaseAPI(path=tessdata_path, lang='eng', psm=psm) as api:
                api.SetImage(pil_img)
                res = api.GetUTF8Text().strip()
                if res:
                    return res
    except Exception as e:
        print(f"[OCR Warning] Fallback OCR encountered: {e}")

    return text_result


def recognize_license_plate(image_path='car1.JPG', show_gui=True):
    """
    Detects and recognizes the license plate from the given car image.
    Returns:
        dict: {
            'text': recognized_text,
            'bbox': (x, y, w, h),
            'plate_image': cropped_and_filtered_plate,
            'annotated_image': image_with_bbox_and_text
        }
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Input image not found: '{image_path}'. Ensure car image is in the project folder.")

    # 1. Read the image file
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError(f"Could not read image file: '{image_path}'")

    if show_gui:
        cv2.imshow("Original", image)

    # 2. Convert to Grayscale Image
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    if show_gui:
        cv2.imshow("gray_image", gray_image)

    # 3. Canny Edge Detection
    canny_edge = cv2.Canny(gray_image, 170, 200)
    if show_gui:
        cv2.imshow("canny_edge", canny_edge)

    # 4. Find contours based on Edges
    contours, _ = cv2.findContours(canny_edge.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)[:150]

    # Initialize license Plate contour and x,y,w,h coordinates
    contour_with_license_plate = None
    license_plate = None
    x, y, w, h = None, None, None, None

    # 5. Find the contour with 4 potential corners and create ROI around it
    for contour in contours:
        # Find Perimeter of contour and it should be a closed contour
        perimeter = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.01 * perimeter, True)
        print("approx", len(approx))
        if len(approx) == 4:  # see whether it is a Rectangle
            contour_with_license_plate = approx
            x, y, w, h = cv2.boundingRect(contour)
            license_plate = gray_image[y:y + h, x:x + w]
            break

    if license_plate is None:
        print("[Warning] No 4-corner contour found for license plate.")
        empty_nlp = process_registration_text("")
        return {
            'text': '',
            'bbox': None,
            'plate_image': None,
            'annotated_image': image,
            'nlp_result': empty_nlp,
            'rto_result': None,
            'formatted_output': format_registration_summary(empty_nlp, None),
        }

    # 6. Thresholding
    _, license_plate = cv2.threshold(license_plate, 127, 255, cv2.THRESH_BINARY)
    if show_gui:
        cv2.imshow("plate", license_plate)

    # 7. Removing Noise from the detected image, before sending to Tesseract
    license_plate = cv2.bilateralFilter(license_plate, 11, 17, 17)

    # 8. Text Recognition
    text = run_ocr(license_plate)
    print("License Plate OCR Raw Output:", text)

    # 9. NLP Text Processing & Andhra Pradesh Regional Information Extraction
    nlp_result = process_registration_text(text)
    rto_result = None
    if nlp_result.entities.state_code:
        rto_result = lookup_rto(nlp_result.entities.state_code, nlp_result.entities.rto_code)

    formatted_summary = format_registration_summary(nlp_result, rto_result)
    print("\n" + formatted_summary + "\n")

    # 10. Draw License Plate rectangle and write the Text
    annotated_image = image.copy()
    annotated_image = cv2.rectangle(annotated_image, (x, y), (x + w, y + h), (0, 0, 255), 3)

    # Keep text coordinates inside visible bounds (display normalized/corrected string if available)
    display_text = nlp_result.corrected_text if nlp_result.corrected_text else text
    text_x = max(10, x - 50)
    text_y = max(30, y - 10)
    annotated_image = cv2.putText(
        annotated_image, display_text, (text_x, text_y),
        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2, cv2.LINE_AA
    )

    # 11. Save results to disk
    cv2.imwrite("output_plate.png", license_plate)
    cv2.imwrite("output_result.png", annotated_image)
    print("[Success] Saved 'output_plate.png' and 'output_result.png'")

    if show_gui:
        cv2.imshow("license_plate", annotated_image)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    return {
        'text': text,
        'bbox': (x, y, w, h),
        'plate_image': license_plate,
        'annotated_image': annotated_image,
        'nlp_result': nlp_result,
        'rto_result': rto_result,
        'formatted_output': formatted_summary,
    }


if __name__ == "__main__":
    # Support command line args: python main.py [image_file] [--headless]
    img_file = 'car1.JPG'
    headless = os.environ.get('HEADLESS', '0') == '1'

    for arg in sys.argv[1:]:
        if arg in ('--headless', '-h', '--no-gui'):
            headless = True
        elif not arg.startswith('-'):
            img_file = arg

    recognize_license_plate(image_path=img_file, show_gui=not headless)
