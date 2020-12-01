# test_cases_sorter

To make our life easier by avoiding sort the 1000+ test cases manually 

## Version History

* V1.2
    * Rename the TC_sorter_class.py to MY_22_sorter.py
    * Introducing the MY_23_sortor.py for sorting the MY-23-related cases

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