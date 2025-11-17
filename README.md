# Camera Game Case Scanner

A Python application that scans physical video game cases using a laptop camera, processes the image through the OCR.Space API, detects the console system, and saves the results into a clean, alphabetized Excel sheet. The tool supports continuous scanning sessions, title editing, retry options, and automatic duplicate checking.

---

## About the Developer

My name is **Bryan Saucedo-Mondragon**, an aspiring Electrical Engineer with a strong interest in applying software tools to real-world problems. I created this project to automate cataloging my game collection while learning more about OCR, APIs, and data-handling workflows.

This is an **ongoing project** that I continue to expand as I improve my skills in Python, automation, and computer vision.

---

## Features

- 📸 Capture game case images using OpenCV  
- 🔤 Extract text via the **OCR.Space API**  
- 🎮 Detect console system (PS3, Xbox 360, Wii, etc.)  
- ✏️ Retry or edit extracted titles  
- 🔁 Continuous scanning loop for multiple items  
- 📄 Export to Excel with sorting & duplicate prevention  
- 🧹 Automatic title cleaning (removes extra tokens, subtitles, etc.)  
- ⚠️ Robust error handling and missing-case warnings  

---

## Installation

```bash
pip install opencv-python openpyxl requests
