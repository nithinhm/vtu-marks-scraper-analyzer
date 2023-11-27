# VTU Marks Scraper and Analyzer

This app is designed to retrieve student marks data from the VTU (Visvesvaraya Technological University) website and analyze it. It includes the following options:

- **Marks Scraper**: An automated tool that retrieves student marks data from the VTU website and stores it in a CSV file. Regular and arrear results will be stored in seperate files. Depending on the URL link that is entered in the scraper, the tool will decide whether to collect revaluation marks or regular/arrear marks.

- **Marks Analyer**: A tool that generates result analysis based on the data collected by the scraper and saves it in an Excel file. It also provides a graphical representation of subject performance in the form of an image file. This tool can update the old marks with the new revaluation marks (if any) and provide an updated analysis.

The repository provides both the source code and a pre-compiled executable (GUI) for ease of use.

## Features

- Collects marks data for all students of a branch from the VTU website.
- Supports automation using Selenium WebDriver.
- **Handles the USN entry and CAPTCHA verification automatically.**
- Processes the collected data to generate an Excel file.
- Provides an easy-to-use graphical-user interface (GUI) app through Tkinter.
- Handles common errors. While collecting data, if the number of retries crosses the set limit due to any errors, the data collected till that point will be saved.

## Prerequisites

Before running the application, ensure that you have the following prerequisites:

- Python 3.x
- Google Chrome
- Tesseract

Install [Tesseract (for Windows)](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe) in the default directory.

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
python "main.py"
```
5. When the GUI window opens, click on the `Marks Scraper` button. The scraper app will open.

6. Enter the required information, such as the first USN, last USN, current semester, etc.

7. Click "Verify" to verify the input. If the entered data is valid, you can click "Collect". You will be asked whether you want to look at the automation process or not.

8. Next, sit back and relax while the tool collects the marks data for the specified students. (You could also abort the collection process by clicking "Abort." Data collected till that point, if any, will be saved.)

9. Once the process is complete, you will find the collected marks data in a raw CSV file(s) within the appropriate subfolder(s) in the same directory. (Regular and arrear marks data will be in seperate folders.)

10. If you wish, you can continue to gather data for other branches or simply stop. (You can also gather data in segments for the same branch and use the analyzer app to merge all the segments to create the final analysis report.)

11. After you have collected the data for all the students from a branch, close the scraper window. The main window reappears. Click the `Marks Analyzer` button. The analyzer app will open.

12. Select "Browse" and choose the saved raw CSV files that hold the marks data of regular marks only. If you've used the scraper to gather revaluation marks, you can include those files here along with the original CSV containing the old marks.

13. Click on "Analyze" and do as instructed. This will generate an Excel file with the marks and their analysis (check the different sheets in the Excel workbook) in the selected folder.

## Important Note

Please avoid running the scraper on multiple systems simultaneously using the same internet IP address. There's a risk that your IP could be blocked from accessing the specific result link forever due to excessive requests in a short period of time. Although necessary measures are taken in the app to avoid IP block, one must ensure to use the tool responsibly to prevent any disruptions in access.

## Executable

Alternatively, you can download and utilize the pre-compiled executable file available in [this folder](https://drive.google.com/drive/folders/1OrhIpXU_E2krhoOlCQMNalobZo_RIoXX?usp=sharing). The executable, compiled with PyInstaller, eliminates the need for Python installation. However, ensure Tesseract is installed in the default directory.

To run the executable, simply double-click it. A command prompt window may appear briefly; wait for some time. Subsequently, the GUI window will open. Note that closing the command prompt will also close the GUI window, so exercise caution.

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