# Code Overview

This repository contains code for processing revenue details and generating settlement reports.

## Prerequisites

Before running the code, ensure that you have the following libraries installed:

- reformat
- refund
- openpyxl
- os
- formula
- scratchu(this is additional function)

## Usage

1. Clone the repository to your local machine.
2. Ensure that you have the necessary input files in the correct directories:
   - `营收详情导出.xlsx` in the `ask` folder.
   - `2022上學期南村結算表.xlsx` in the main directory.
3. Update the `data_directory` variable in the code to specify the correct file paths.
4. Run the code using a Python interpreter.
5. The output will be saved in the `2022上學期南村結算表.xlsx` file with the processed data.

## Code Explanation

The code performs the following steps:

1. Imports the required libraries.
2. Downloads an XLS file using the `scratchu` module.
3. Transfers the downloaded file to an Excel file using the `scratchu` module.
4. Imports necessary data from Excel files.
5. Processes the data and performs calculations.
6. Updates and saves the output in the `2022上學期南村結算表.xlsx` file.

Please refer to the code comments for more detailed explanations of each step.

## License

This code is released under the [MIT License](LICENSE).
