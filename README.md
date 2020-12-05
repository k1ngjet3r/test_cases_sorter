# test_cases_sorter

To make our life easier by avoiding sort the 1000+ test cases manually

## Setup
* Python with version 3.7+
* Visual Studio Code (optional)
* Openpyxl installed

## How to use
1. Open the MY_22_sorter.py using VS Code
2. Drug the test case and last week result file in the form of .xlsx to the VS Code
3. Set the input file name, output file name, and last week result file name in the testing = Tc_sorter('input_name', 'output_name', 'last week result')
4. Execute the script and wait for the output file generated

## Version History

* V1.2
    * Rename the TC_sorter_class.py to MY_22_sorter.py
    * Introducing the MY_23_sortor.py for sorting the MY-23-related cases
    * Adding formater method to the class for the output format
    * Create a new sheet names difficult_case that contain the list of failed cases from last week
    * Appending the name of the tester who executed the cases, also the result from previous week


* V1.1
    * Using python class for the main structure
    * Update the matching process using the regular expression
    * Determine the User, sign-status, and connection instead of function
    * Isolate the bench-only cases

* V1.0
    * Tried fixing the "air" and "pair" matching issue by subsituting the sign in the sentence with space and than spliting the sentence at space and store it in a list for matching process

* V0.0
    * Initial Release
    * Focusing on the function determination
    * Matching the keyword by slicing the sentence with the length of the keyword and match the keyword
    * Read and write the .xlsx file using the Openpyxl python module
    * Known Issue: "air" and "pair" issue