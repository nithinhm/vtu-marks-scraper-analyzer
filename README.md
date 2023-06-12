# VTU Marks Scraper App 2023

The contents of this repository include an automated tool designed to retrieve and organize student marks data from the VTU (Visvesvaraya Technological University) website in an excel file and which also generates a result analysis for the specified branch, as well as a graphical representation of subject performance in the form of an image file. The repository provides both the source code and a pre-compiled executable for ease of use.

## Features

- Collects marks data for all students of a branch from the VTU website
- Supports automation using Selenium WebDriver
- Handles the USN entry and CAPTCHA verification automatically
- Processes the collected data to generate an Excel file
- Provides an easy-to-use command-line interface (CLI)
- Handles common errors. If the number of retries crosses the set limit, the data collected till that point will be saved.

## Prerequisites

Before running the application, ensure that you have the following prerequisites:

- Python 3.x
- Chrome WebDriver
- Pytesseract

Make sure to download and configure the Chrome WebDriver according to your system.

Install Pytesseract (for WIndows) using the [installer](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe) in the default directory.

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

4. Run the application using Python:

```
python "vtu marks scraper app by nithinhm.py"
```

5. Follow the prompts and enter the required information, such as the college code, branch code, batch year etc.

6. Sit back and relax while the tool collects the marks data for all students and generates the Excel file.

7. Once the process is complete, you will find the processed marks data in an Excel file named `20{batch} {branch} semester {number} {first_USN} to {last_USN} VTU results.xlsx` within the appropriate subfolder in the same directory.

8. If you wish, you can continue collecting data for other branches or simply quit.

## Executable

Alternatively, you can use the pre-compiled executable file located in the `dist` subfolder of this repository. The executable is compiled using PyInstaller and does not require Python installation. But you still need to install Pytesseract in the default directory.

Run the executable by double-clicking it or using the command line.

## License

This project is licensed under the [MIT License](LICENSE).

## Contribution

Contributions to this project are welcome. Feel free to open issues and submit pull requests to suggest improvements or fix any bugs.

Please ensure that your contributions align with the project's coding style and conventions.

## Acknowledgements

This project was inspired by the need to automate the laborious procedure of collection of marks data from the VTU website. Special thanks to the open-source community for their valuable resources and tools.

## Contact

For any inquiries or suggestions, please contact the project maintainer:

Prof. Nithin H M

Email: nithin.manju@amceducation.in

Feel free to reach out with any questions or feedback you may have.
