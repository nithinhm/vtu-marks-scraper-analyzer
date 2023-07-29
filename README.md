# VTU Marks Scraper and Analyzer

**Updated 28/07/2023**

The contents of this repository include 2 apps:

1. **Scraper and Analyzer App**

    This main app contains the following two options:

    - **Marks Scraper**: An automated tool designed to retrieve student marks data from the VTU (Visvesvaraya Technological University) website and store it in a CSV file. This option is intended for data scraping and gathering marks information.

    - **Marks Analyer**: A tool that generates result analysis based on the data collected by the scraper and saves it in an Excel file. It also provides a graphical representation of subject performance in the form of an image file. This option is focused on analyzing the collected marks data and presenting insights.

2. **Revaluation Marks Scraper App**

    An automated tool that collects data of revaluation marks similar to the above scraper.

The repository provides both the source code and pre-compiled executables (GUI) of both apps for ease of use.

## Features

- Collects marks data for all students of a branch from the VTU website.
- Supports automation using Selenium WebDriver.
- Handles the USN entry and CAPTCHA verification automatically.
- Processes the collected data to generate an Excel file.
- Provides easy-to-use graphical-user interface (GUI) apps through Tkinter.
- Handles common errors. If the number of retries crosses the set limit, the data collected till that point will be saved.

## Prerequisites

Before running the application, ensure that you have the following prerequisites:

- Python 3.x
- Chrome WebDriver
- Pytesseract

Make sure to download and configure the Chrome WebDriver according to your system.

Install [Pytesseract (for Windows)](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe) in the default directory.

## Usage

1. Clone the repository to your local machine:

```
git clone https://github.com/nithinhm/vtu-marks-scraper-analyzer.git
```

2. Navigate to the project directory:

```
cd vtu-marks-scraper-analyzer
```

3. Install the required dependencies:

```
pip install -r requirements.txt
```

4. Run the main app using Python:

```
python "scraper_analyzer_app_nithinhm.py"
```
5. Click on the `Marks Scraper` button. The scraper app will open.

6. Enter the required information, such as the college code, branch code, batch year etc.

7. Click "Verify" to verify the input. If the entered data is valid, you can click "Collect".

8. Next, sit back and relax while the tool collects the marks data for the specified students. (You could also abort the collection process by clicking "Abort." Data collected till that point, if any, will be saved.)

9. Once the process is complete, you will find the collected marks data in a CSV file within the appropriate subfolder in the same directory.

10. If you wish, you can continue collecting data for other branches or simply quit. (You could also collect data in chunks for the same branch and use the analyzer app to combine all the chunks to generate the final analysis report.)

11. After you have collected the data for all the students from a branch, close the scraper window. The main window reappears. Click the `Marks Analyzer` button. The analyzer app will open.

12. Click on "Browse" and select the saved CSV file(s) that contain the marks data.

13. Click on "Analyze" and do as instructed. You will get the analysis report for the branch in the selected folder.

## Executable

Alternatively, you can download and use the pre-compiled executable file present in [this folder](https://drive.google.com/drive/folders/1OrhIpXU_E2krhoOlCQMNalobZo_RIoXX?usp=sharing). The executable is compiled using PyInstaller and does not require Python installation. But you still need to install Pytesseract in the default directory.

Run the executable by double-clicking it.

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