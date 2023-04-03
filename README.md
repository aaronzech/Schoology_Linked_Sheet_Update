# Schoology_Linked_Sheet_Update


configure the gspread service account on Google Cloud Console.

<H1>Prerequisites</h1>
Have the following Python libraries installed.
<ul><li>gspread</li>
<li>tkinter</li>
<li>pandas</li>

<li>df2gspread</li></ul>

  ```sh
  python -m pip install df2gspread
  ```
  ```sh
  python -m pip install gspread
  ```
 <H1>Directions</h1>
 
 1) Download the Parent Access Code CSV file from Schoology
 
 ![](https://github.com/aaronzech/images/blob/main/Screenshot_241.png)
 
 2) Run the python script
 3) The program will look for the most recent 'parent-code-export' file in the Downloads directory.
 4) Open the Google Sheet
 
 ![](https://github.com/aaronzech/images/blob/main/Screenshot_242.png)

 5) Format the Buidling ID column to a number format. Optional run the macro to do this.
 
 ![](https://github.com/aaronzech/images/blob/main/Screenshot_243.png)
