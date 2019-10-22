# excel-combine
Python code to clean and combine multiple excel files into one data.

Suppose you have 100 excel files that you want to combine into one data. Doing it manually would cost a lot of your time.
I've experienced this so many times in the past and would love to automatize the process.

This is the generic code that you might use to clean and combine multiple excel files into one excel file. 
Much better than creating VBA code for each project you stumbled upon.

The files that you want to combine might come from several format that is not suitable for the final excel file you want.
Most of the time, you only want the main data of the file, not the cosmetics row or column.

First, you create a stand alone excel file containing only the header of the file you want to combine. 
Write the header on the very first row. In this case, the header are: "No", "Nama", "Judul", "Semester", and "Progress". 

#### Excel file with header information of final result you want:
![Image of header file](https://hendriyono.files.wordpress.com/2019/10/header-example-e1571735106670.png)

#### Save the file and put them in folder structure like below:
![Image of folder structure](https://hendriyono.files.wordpress.com/2019/10/folder-structure-e1571735118967.png)

The "Combine" folder is where you put all of your excel file which you want to combine.

Then, we clean the excel file by stating unwanted part such as: top row, left side column, bottom row, right side column, and blank column.

#### Example of unwanted part of the file:

![image of excel cleaning setup](https://hendriyono.files.wordpress.com/2019/10/excel-cleaning-final-e1571732454294.png)

#### The cleaned file:

![image of cleaned excel](https://hendriyono.files.wordpress.com/2019/10/cleaned-file-e1571732706425.png)

#### Final result (combination of 2 excel files):

![image of final file](https://hendriyono.files.wordpress.com/2019/10/combined-file-e1571733649282.png)

The code will guide you in defining the unwanted part. Just press start!

You will need to have numpy and openpyxl packages installed. Check the requirements.txt file. 
If you find difficulties in setting up this code, please contact me via email: hendri.rach at gmail.com or via instagram: (at)hendriyono
