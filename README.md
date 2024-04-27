# Version Excel

## Part I: Introduction and Installation

### Version controlling `.XLSX` and `.XLSM` Excel workbooks

This library aims to provide a robust versioning solution for Excel workbooks in the Office Open XML format (Workbooks with the `.xls` are not supported). This will provide human readable diffs that give context on the changes made between versions with a separate VBA code file for versioning independent of the parent workbook. This is implemented via a set of git hooks which convert workbooks to a `.yml` extension prior to a commit containing all XML data needed to recreate the workbook alongside any VBA code. Many workbooks will see reduced commit and repo size using this approach when compared to a binary format.

### Building the versioning executable
It's recommended to build the versioning code into an executable for ease of deployment.  This executable can either be stored within the repo to be versioned or separately on the local system. This step can be skipped by using an executable from the releases section.  To proceed with building locally, first install the necessary dependncies via the package installer
```
pip3 install -r requirements.txt
```

Next we should run the `build_executable.bat` script to create an executable using PyInstaller. This will create a single file with no external dependencies for deployment. 
```
./build_executable.bat
```

### Modifying the `.gitignore` file
Within the repo to be version controlled, we should update the `.gitignore` file located in the root directory of the repo to exclude Excel files from Git. The `.yml` files output by the script will be tracked by Git and so we no longer need to track the changes to these binary files. These lines should be added to the file:
```
# Excel Files
*.xlsx
*.xlsm
```

### Installing the Git Hooks
Within the repo to be version controlled add the `pre-commit` and `post-checkout` files into `.git/hooks` folder. These files can be found in the `hook_templates` folder. By default these hooks expect to find the `version_xlsx` executable in the root directory of the repo. This can modified to provide a different path to the executable by modifying the `./version_xlsx` path as shown in the example below:
```
#!/bin/sh

if (./version_xlsx convert_to_yml); then
  echo "pre-commit success"
  git add .
  exit 0
else
  echo "pre-commit failed"
  exit 1
fi
```

### Setting up the configuration file
Finally, copy the `version_sheet_settings.yml` file into the same folder as the `version_xlsx` executable (typically the root directory of the repo being version controlled). This file provides a set of options to configure how the library operates on workbooks.

- `enabled` A global switch to disable or enable the versioning library. Defaults to `True`
- `convert_xlsx` Optionally enable or disable processing of workbooks with a `.xlsx` extension. Defaults to `True`
- `convert_xlsm` Optionally enable or disable processing of workbooks with a `.xlsm` extension. Defaults to `True`
- `convert_vba_separate_file` Outputs VBA code from `.XLSM` workbooks into a separate file. Defaults to `True`
- `enable_logging` Enables the logging of script operations and execution time into a logfile. Defaults to `True`
- `logfile` Specifies a path to write the log file if `enable_logging` is enabled.
- `exclude_directories` A list of directories that should not be scanned when versioning workbooks. By default `.git` should be included to improve scan performance.

```yml
options: 
  enabled: True
  convert_xlsx: True
  convert_xlsm: True
  convert_vba_separate_file: True
  enable_logging: True
  logfile: '../version_log.txt'
exclude_directories:
  - "New folder"
  - ".git"
  - "utils"
  - "sample_sheets"
```

## Part II: Usage and Output
Proceed to edit and update Excel Workbooks within your repo, always ensure that you save and exit the workbook before committing your changes to source control.  You can commit changes as usual through either Git, GitHub Desktop or another application. Note that GitHub desktop does not currently allow an empty commit so other changes must be present on the branch to create a commit from the interface. From terminal an empty commit can be created as shown below:
```
git commit --allow-empty -m "message"
```

If an Excel workbook is locked for editing when the script requires access, an error message will be displayed and the commit will fail. Simply close the workbook and retry the commit to proceed. 

The changes that have been made to the workbook will be visible as a human readable diff within GitHub, Bitbucket or a local diff tool:
![image](https://github.com/nd4321/version_excel/assets/16249888/9de9b6df-420a-4a3d-bff5-c8b05b5cd9fe)

Similarly the VBA code is visible both within the `.xlsm.yml` file and within a separate `.vba` file (provided `convert_vba_separate_file` is enabled within the configuration file)
```vba
vba: 
  filename: "Module2"
  code: |
    Attribute VB_Name = "Module2"
    Sub Macro1()
    Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Macro1 Macro

        Range("E12:N12").Select
        Selection.Font.Bold = False
        Range("P12").Select
        Selection.Font.Bold = False
        Range("H3").Select
        ActiveCell.FormulaR1C1 = "test"
        Range("F3").Select
    End Sub
```

Locally you should see all three of these files within your repo. Note that the `.xlsm` or `.xlsx` file is not stored within the hosted repository, but is created within the local repo when necessary via the `post-checkout` hook.  
![image](https://github.com/nd4321/version_excel/assets/16249888/3e8943bc-8cfa-499f-abe2-b65e36a39b17)

## Part III: Performance

The time taken to convert been the Excel Workbook and versioned YML file is minimal, but varies primarily based on the size of the input files. If the conversion operation takes to long, consider moving some files to an excluded directory. Below we can see the output from the log file and some sample execution times for conversion operations:
```
Sun Apr 21 14:38:44 2024 | convert_to_yml | Success | Execution time: 0.039 seconds | .\sample.xlsx
Sun Apr 21 14:38:44 2024 | convert_to_yml | Success | Execution time: 0.123 seconds | .\test_sheet.xlsx
Sun Apr 21 14:48:50 2024 | convert_to_excel | Success | Execution time: 0.027 seconds | .\sample.xlsx.yml
Sun Apr 21 14:48:51 2024 | convert_to_excel | Success | Execution time: 0.39 seconds | .\test_sheet.xlsx.yml
```

In the majority of cases, the size of commits will be smaller after the workbook has been converted to versioned YML.  This is even true for many complex cases such as images being added to the workbook.  A greater number of edits to a workbook increases the likelihood that the YML versioning will result in space savings. This is caused by a high number of edits disrupting the compression algorithm used on the binary files resulting in a significantly different output from the previous version. Below we can see some testing data from a set of sample workbooks:

| Workbook Name                                   | Commit Excel (Bytes) | Commit YML (Bytes) | Space Savings | Notes |
| ----------------------------------------------- | -------------------- | ------------------ | ------------- | ----- |
| DatasetLabelingWorkbook                         | 30,179               | 12,787             | 57.63%        |       |
| ddp-6-caltrans-dataset-metadata-template v3-1-1 | 3,262,692            | 1,613,138          | 50.56%        | Image |
| IFIDefaultGridFactors2021v3-1-unfccc            | 49,684               | 8,234              | 83.43%        |       |
| IMFInvestmentandCapitalStockDataset2021         | 4,839,880            | 6,520,781          | \-34.73%      |       |
| local-seo-citation-building-template            | 13,928               | 10,020             | 28.06%        |       |
| oecd-monthly-capital-flow-dataset               | 1,970,948            | 3,008,095          | \-52.62%      |       |
| sampledataworkorders                            | 231,560              | 58,642             | 74.68%        |       |
| Search-Engine-Optimization-Certificate-2020     | 24,471               | 1,041              | 95.75%        |       |
| SuperStoreUS-2015                               | 410,393              | 568,024            | \-38.41%      |       |
| tableaudataset202009                            | 576,450              | 13,535             | 97.65%        |       |
| TheSEOChecklistTemplate-V2-0                    | 81,833               | 131,900            | \-61.18%      |       |
| UKEnergyinBrief2022dataset                      | 324,837              | 222,880            | 31.39%        |

Additional Testing Data (Testing Round 4)

| Workbook Name                                           | Commit Excel (Bytes) | Commit YML (Bytes) | Space Savings | Notes |
| ------------------------------------------------------- | -------------------- | ------------------ | ------------- | ----- |
| 15577298                                                | 17,976               | 5,599              | 68.85%        |       |
| 2021-Distribution-Rates-Database                        | 216,179              | 86,962             | 59.77%        |       |
| 7c93adb5-en                                             | 36,111               | 39,721             | \-10.00%      |       |
| aisc-shapes-database-v15.0                              | 1,217,588            | 886,183            | 27.22%        |       |
| budget-2022-fiscal-plan-economic-outlook-tables-2022-25 | 160,365              | 43,765             | 72.71%        |       |
| cgfs60-datasets                                         | 165,008              | 15,805             | 90.42%        |       |
| database-of-treatment-effects-excel-2610552205          | 243,976              | 346,142            | \-41.88%      |       |
| exsituap10                                              | 218,953              | 30,464             | 86.09%        |       |
| FID-Data-Use-Matrix                                     | 13,525               | 639                | 95.28%        |       |
| JSTOR-BusinessIVCollection-2024-04-25                   | 19,760               | 30,515             | \-54.43%      |       |
| MeasuringBiodiversityDatabase                           | 178,791              | 54,108             | 69.74%        |       |
| newlevy                                                 | 8,993                | 259                | 97.12%        |


## Demonstration 

https://github.com/nd4321/version_excel/assets/16249888/d6fb382d-07df-4283-9e98-5b98df4f511c



