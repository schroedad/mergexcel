Create a python script that will combine all Excel files in a specified folder into a single file as follows:

Assumptions:
1. All excel files (with suffix .xlsx) in the folder specified will be included.
2. The first file processed dictates the number and name of all the worksheets (up to 100).  Subsequent files typically have the same number of worksheets and names, but if any worksheets are missing (by name comparison), skip those worksheets but process the rest of the worksheets in the file. If there are extra worksheets in subsequent files that do not match the name of the first one processed, skip them as well.  Except, if no worksheets match, or if only the codebook worksheets match, then stop with an error.
3. The first n number of worksheets may or may not be static codebooks. If there are > 0 codebooks, all Excel files will have the codebooks but we should only include the codebooks from the first excel file processed. 

General Approach
1. First, ask the user a couple of questions on the command line. The first question will be "Specify the folder containing excel files [] :"  where the current folder is listed as the default. Support relative paths like ~/files. The second will ask "How many codebook worksheets are there [1]?" The default is 1 for this question, but 0 or other positive integers are valid responses. Ensure an integer response. 0 means there are no codebooks.
2. Create a list of all Excel files in the specified folder for looping
3. Open the first file in the list and use its contents to create a new Excel file that will be the combined Excel file. If the user specified that there are codebooks, copy the entire first worksheet as the codebook from this first file into the garget combined version (including the name of the worksheet). Then, copy everything from subsequent worksheets, including the name of each worksheet and the header row of each into subsequent worksheets on the target combined Excel file.
4. Iterate through the rest of the files. For each, if the user said the files contain a codebook, skip the first worksheet. Otherwise don't skip it. Then copy rows 2 - n from each worksheet into the corresponding worksheet of our target combined file. We start on row 2 because we only need one copy of the header row in our target file.

More details
1. Call the python command mergexcel
2. Use the openpyxl library
3. Try to limit use of libraries beyond openpyxl
4. Provide comments within the code consistent with open source projects and with enough detail for those newer to Python and openpyxl to understand
5. When run, show a brief overview of what the program does to the user. E.g. "This will merge all like Excel files in a specified directory into a single target Excel file." 
6. Prompt the user to answer the two questions.
7. The command prompts and output should be styled well. The Laravel installer can be referenced as an example for styling and layout. Be sure to use good spacing for readability. For example, put a blank line between the questions and the output progress.
8. As the program iterates through and combines files, show a progress bar in the terminal with what is happening at each step to the right of the bar. This should use ASCII art and the progress bar should take up about 20% of a standard terminal width with text about what is happening to the right such as "Opening (filename)" and "Copying worksheet 3 of x".  Be specific. Replace existing text so the terminal doesn't scroll vertically. Also clear through the end of the line after each update so we don't see text artifacts from the previous message.  Reference how the Laravel installer does this as an example of good design. Be sure the progress bar includes the entire process.
9. For the codeboook worksheets, copy with formatting to preserve layout. For the rest of the worksheets, Make the header row on the combined file bold and shaded a light grey. This can be done at the end. Also auto-set the width for all the columns in the combined file at the end. Be sure to account for this in the progress bar and provide feedback to the user about doing this because formatting can take time. Avoid bugs in code related to using formatting with openpyxl. 
10. For files where no worksheets match the initial file, or where only the codebook(s) match but not others, stop with an error message. If at least some match, keep track of which worksheets are missing and which are extra for each file. This will be displayed in the summary.
11. When complete, show a summary of what was done including number of files, number of rows, etc. For each file, indicate the number of missing and/or extra worksheets.
12. If errors, still save the combined file as-is for debugging purposes.
13. Use emoji or ASCII art in the terminal where helpful but in a subtle way.
14. Make as performant as possible with readable code.
15. Document the code well with comments.
16. Also create a readme.md file and any other assets needed such as requirements.txt so that this could be uploaded as an open-source utility to Github.