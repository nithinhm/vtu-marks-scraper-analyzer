# VTU Marks Scraper and Analyzer

This app is designed to retrieve student marks data from the VTU (Visvesvaraya Technological University) website and analyze it. It includes the following options:

- **Marks Scraper**: An automated tool that retrieves student marks data from the VTU website and stores it in a CSV file. Regular and arrear results will be stored in seperate files. Depending on the URL link that is entered in the scraper, the tool will decide whether to collect revaluation marks or regular/arrear marks.

- **Marks Analyer**: A tool that generates result analysis based on the data collected by the scraper and saves it in an Excel file. It also provides a graphical representation of subject performance in the form of an image file. This tool can update the old marks with the new revaluation marks (if any) and provide an updated analysis.

The repository provides both the source code and a packaged release for ease of use.

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

<p align="center">
    <img src="images\main.png" width="400">
</p>

5. When the GUI window opens, click on the `Marks Scraper` button. The scraper app will open.

<p align="center">
    <img src="images\scraper.png" width="400">
</p>

6. Enter the required information, such as the first USN, last USN, current semester, etc.

7. Click "Verify" to verify the input. If the entered data is valid, you can click "Collect". You will be asked whether you want to look at the automation process or not.

8. Next, sit back and relax while the tool collects the marks data for the specified students. (You could also abort the collection process by clicking "Abort." Data collected till that point, if any, will be saved.)

9. Once the process is complete, you will be asked to select a folder in which to save the collected marks data. Once the folder is selected, the data will be saved in it as a raw CSV file(s) within the appropriate subfolder(s) (regular and arrear marks data will be in seperate folders). Make sure to not change the names of the downloaded files.

10. If you wish, you can continue to gather data for other branches or simply stop. (You can also gather data in segments for the same branch and use the analyzer app to merge all the segments to create the final analysis report.)

<p align="center">
    <img src="images\analyzer.png" width="400">
</p>

11. After you have collected the data for all the students from a branch, close the scraper window. The main window reappears. Click the `Marks Analyzer` button. The analyzer app will open.

12. Select "Browse" and choose the saved raw CSV files that hold the marks data of regular marks, arrear marks, and revaluation marks (if any). Selecting files containing regular marks is compulsory.

13. Click on "Analyze" and do as instructed. This will generate an Excel file with the marks and their analysis (check the different sheets in the Excel workbook) in the selected folder.

## Important Note

Please avoid running the scraper on multiple systems simultaneously using the same internet IP address. There's a risk that your IP could be blocked from accessing the specific result link forever due to excessive requests in a short period of time. Although necessary measures are taken in the app to avoid IP block, one must ensure to use the tool responsibly to prevent any disruptions in access.

## Packaged Release

Alternatively, you can download and utilize the packaged version of this project available in the [Releases](https://github.com/nithinhm/vtu-marks-scraper-analyzer/releases) section of this repositary. This package does not require manual installation of Python. More on how to use the packaged version is available in the description of the release.

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
