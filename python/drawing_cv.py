#!/usr/bin/env python3
"""
PMC Civil AI Agent — Drawing Computer Vision Module
Extracts scale, dimensions, text from civil drawings
"""
import cv2
import numpy as np
import base64
import json
import sys
from PIL import Image
import io

def decode_image(b64_str):
    """Decode base64 image to numpy array"""
    img_bytes = base64.b64decode(b64_str)
    img_array = np.frombuffer(img_bytes, dtype=np.uint8)
    img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    return img

def detect_scale_bar(img):
    """
    Detect graphical scale bar in drawing
    Returns: pixels_per_meter or None
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    h, w = gray.shape
    
    # Scale bars usually in bottom portion of drawing
    bottom_region = gray[int(h*0.7):, :]
    
    # Find horizontal lines (scale bars are horizontal)
    edges = cv2.Canny(bottom_region, 50, 150)
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=50, minLineLength=50, maxLineGap=5)
    
    scale_candidates = []
    if lines is not None:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            # Horizontal lines only
            if abs(y2 - y1) < 3:
                length_px = abs(x2 - x1)
                if 50 < length_px < w * 0.4:  # Reasonable scale bar width
                    scale_candidates.append(length_px)
    
    return scale_candidates

def extract_dimension_lines(img):
    """Extract dimension lines and their pixel lengths"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Threshold to get black lines on white background
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    
    # Detect horizontal and vertical lines
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    
    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel)
    
    # Find contours
    h_contours, _ = cv2.findContours(horizontal_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    v_contours, _ = cv2.findContours(vertical_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    h_lengths = []
    v_lengths = []
    
    for cnt in h_contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 30:  # Minimum length
            h_lengths.append({'x': x, 'y': y, 'length_px': w})
    
    for cnt in v_contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if h > 30:
            v_lengths.append({'x': x, 'y': y, 'length_px': h})
    
    return {'horizontal': h_lengths[:300], 'vertical': v_lengths[:300]}

def detect_rooms_and_spaces(img):
    """Detect closed regions (rooms, spaces) in floor plan"""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    
    # Find contours (closed regions)
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    
    spaces = []
    img_area = img.shape[0] * img.shape[1]
    
    for cnt in contours:
        area = cv2.contourArea(cnt)
        # Filter: meaningful area (not too small, not full image)
        if img_area * 0.001 < area < img_area * 0.5:
            x, y, w, h = cv2.boundingRect(cnt)
            perimeter = cv2.arcLength(cnt, True)
            # Compactness check (rooms are roughly rectangular)
            compactness = 4 * np.pi * area / (perimeter ** 2) if perimeter > 0 else 0
            if compactness > 0.3:  # Reasonably rectangular
                spaces.append({
                    'x': int(x), 'y': int(y),
                    'width_px': int(w), 'height_px': int(h),
                    'area_px': int(area),
                    'compactness': round(compactness, 3)
                })
    
    # Sort by area descending
    spaces.sort(key=lambda s: s['area_px'], reverse=True)
    return spaces[:15]

def analyze_drawing_cv(b64_image, image_type='image/png'):
    """
    Main CV analysis function
    Returns dict with extracted geometric data
    """
    try:
        img = decode_image(b64_image)
        h, w = img.shape[:2]
        
        result = {
            'image_dimensions': {'width_px': w, 'height_px': h},
            'scale_bar_candidates_px': detect_scale_bar(img),
            'dimension_lines': extract_dimension_lines(img),
            'detected_spaces': detect_rooms_and_spaces(img),
            'image_analysis': {
                'aspect_ratio': round(w/h, 3),
                'is_landscape': w > h,
                'estimated_type': 'floor_plan' if w > h else 'section_drawing'
            }
        }
        
        # Calculate pixel density hints
        if result['scale_bar_candidates_px']:
            common_scales = [1, 2, 5, 10, 20, 50, 100, 200, 500]  # meters
            avg_scale_bar_px = np.mean(result['scale_bar_candidates_px'])
            hints = []
            for scale_m in common_scales:
                px_per_m = avg_scale_bar_px / scale_m
                hints.append(f"If scale bar = {scale_m}m → 1px = {1/px_per_m:.4f}m ({px_per_m:.1f}px/m)")
            result['scale_interpretation_hints'] = hints[:5]
        
        return result
    except Exception as e:
        return {'error': str(e), 'traceback': str(e)}

if __name__ == '__main__':
    # Test with command line: python3 drawing_cv.py <base64_file>
    if len(sys.argv) > 1:
        with open(sys.argv[1]) as f:
            b64 = f.read().strip()
        result = analyze_drawing_cv(b64)
        print(json.dumps(result, indent=2))
    else:
        print('Usage: python3 drawing_cv.py <base64_image_file>')
