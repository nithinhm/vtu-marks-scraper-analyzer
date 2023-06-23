# VTU Marks Scraper App 2023

The contents of this repository include an automated tool designed to retrieve and organize student marks data from the VTU (Visvesvaraya Technological University) website in an excel file and also generate result analysis for the specified branch, as well as a graphical representation of subject performance in the form of an image file. The repository provides both the source code and a pre-compiled executable (GUI) for ease of use.

Additionally there is now another tool that collects data of revaluation marks (source code and app included).

## Features

- Collects marks data for all students of a branch from the VTU website.
- Supports automation using Selenium WebDriver.
- Handles the USN entry and CAPTCHA verification automatically.
- Processes the collected data to generate an Excel file.
- Provides an easy-to-use graphical-user interface (GUI) through Tkinter.
- Handles common errors. If the number of retries crosses the set limit, the data collected till that point will be saved.

## Prerequisites

Before running the application, ensure that you have the following prerequisites:

- Python 3.x
- Chrome WebDriver
- Pytesseract

Make sure to download and configure the Chrome WebDriver according to your system.

Install [Pytesseract (for WIndows)](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe) in the default directory.

## Usage

1. Clone the repository to your local machine:

```
git clone https://github.com/nithinhm/2023-VTU-marks-scraper-app.git
```

2. Navigate to the project directory:

```
cd 2023-VTU-marks-scraper-app
```

3. Install the required dependencies:

```
pip install -r requirements.txt
```

4. Run the application using Python (and similarly for the other app):

```
python "vtu marks scraper app by nithinhm - GUI.py"
```

5. Enter the required information, such as the college code, branch code, batch year etc.

6. Click Verify to verify the input. If the entered data is valid, you can click Collect. Next, sit back and relax while the tool collects the marks data for the specified students and processes it.

7. Once the process is complete, you will find the processed marks data in an Excel file within the appropriate subfolder in the same directory.

8. If you wish, you can continue collecting data for other branches or simply quit.

## Executable

Alternatively, you can download and use the pre-compiled executable files (both kinds) present in [this folder](https://drive.google.com/drive/folders/1OrhIpXU_E2krhoOlCQMNalobZo_RIoXX?usp=sharing). The executables are compiled using PyInstaller and do not require Python installation. But you still need to install Pytesseract in the default directory.

Run the executable(s) by double-clicking it.

## License

This project is licensed under the [GNU General Public License v2.0](LICENSE).

## Contribution

Contributions to this project are welcome. Feel free to open issues and submit pull requests to suggest improvements or fix any bugs.

Please ensure that your contributions align with the project's coding style and conventions.

## Acknowledgements

This project was inspired by the need to automate the laborious procedure of collection of marks data from the VTU website. Special thanks to the open-source community for their valuable resources and tools. And a big shoutout to [Suhas P K](https://github.com/suhaspk) for his constant feedback, suggestions, and code-breaking!

## Contact

For any inquiries or suggestions, please contact the project maintainer:

Prof. Nithin H M

Email: nithinmanju111@gmail.com

Feel free to reach out with any questions or feedback you may have.