CAP776 MINOR PROJECT 1 - MY DATA, MY STORY
Himanshu Gupta | 12602516 | MCA, D1P2635

FILES
12602516.xlsx          Original activity workbook, unchanged
12602516.py            Python program and reusable functions
12602516.docx          Filled report in the supplied template
12602516_results.txt   Output saved by the Python program

RUN IN WINDOWS POWERSHELL
1. Keep the Python file and workbook in the same folder.
2. Open PowerShell in that folder. For the Desktop copy, enter:
   cd "C:\Users\Hiii\Desktop\12602516"
3. Run:
   python 12602516.py

Python 3.14.7 and openpyxl 3.1.5 are already available on this computer.
On another computer, if openpyxl is missing, install it once with:
   python -m pip install openpyxl

The program prints the results and replaces 12602516_results.txt with
the latest output. It does not change the workbook or upload anything.
The Word report is a saved report, so rerunning Python does not update it.
If activity data or period dates change, update the report to match.

FINAL RECORDING PERIOD
17 August to 21 September 2026, inclusive: 36 expected days.
The student confirmed that the teacher specified 17 August as the start,
superseding the 13 August date printed in the earlier lecture/template.
The workbook has 34 valid days in this period. 17 and 18 August are missing.
22 and 23 September remain in the workbook but are outside this period.
No missing activity entries have been invented.
DCI = 34/36 x 100 = 94.44%. PAI = 317.48.

CODE SAMJHNE KE LIYE
read_data(): Excel ke Daily Log ki row 6 se entries padhta hai. Blank
template rows skip karta hai. Invalid values, duplicate dates aur period
ke bahar ki entries alag karta hai. Total tracked time aur free time
minute columns se calculate hote hain; class count minutes mein nahi judta.

average(): Loop se values jodkar valid days ki count se divide karta hai.

activity_indices(): Required 9 indices nikalta hai. PAI mein report
template ke weights hain; DCI ka weight 0.05 hai. ABI alag report hota hai.

correlation(): Pearson r basic loops aur arithmetic se nikalta hai.
Positive r ka matlab dono values saath badhne ka pattern hai. Isse
cause-and-effect prove nahi hota. Constant values ke liye r undefined hai.

main(): Functions call karke results screen aur text file mein likhta hai.
Missing workbook, wrong sheet aur empty valid data par error batata hai.

SCORING
Feeling: Excellent=5, Good=4, Neutral=3, Low=2, Stressed=1.
Satisfaction: Very Satisfied=5, Satisfied=4, Neutral=3,
Unsatisfied=2, Very Unsatisfied=1.
Energy: High=3, Medium=2, Low=1.
EI is the average of the three scores without rescaling energy.
The template labels EI /5, but these prescribed scores give a maximum
of 13/3 (about 4.33). PAI is a weighted index, not a percentage.

BEFORE USING THE REPORT
Read the personal findings and improvement plan and revise them to
reflect your own views. Signature and signature date are left blank.
Lecture 9, page 32 states: "Do not copy or use LLMs (we will figure out)".
This code and report were prepared with AI assistance. Use them in line
with your course rules; sign the declaration only if it is accurate.

REFERENCES
CAP776-Lecture-9-15.09.2026 openpyxl+MP1.pdf: period and rubric.
CAP776-Lecture-11-22.09.2026 MP1+Exceptions.pdf: revised expectations.
CAP776-MinorProject1Report.docx: report layout and index formulas.
OpenPyXL_Step_by_Step_Classroom.pdf: workbook reading and Python examples.
