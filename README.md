# A4-Final_project
We have made an Excel file with the costplan for the Duplex files loadbaering walls.

## How the tool works:
In the main.py file have we made a script, to help the Contractors, the client and the workers. 

The main py file are making an excel file wich contains the: area of the exterior wallparts, it shows the prise for the loadbearing wall pr m2., and the prise for each loadbearing wallparts in building.

The output looks like this. 

![Our Cost-plan](https://github.com/AnjaHolmquist/A4-Final_project/blob/main/the%20costplan.png)

## But here is how we get to this output.
First of all we have to open the main.py file, as you can se now,

it imports ifcopenshell, the model, Pandas (This is the program that makes the excel file in the buttom of the script), and then the json file (The json file is the molio price data file frol learn)

The first loop sorts out the prise for the in-situ loadbaering concrete wall in the json file.


The sekend loop makes a list for the Roof level, afterworts it sorts out the exterialwalls on level Roof and gives us the materials and the area in m2


Now we have made some Values we want the list to contain so we can see it in the excel file.

So the first thing is the level, then we want it to tell us the informations about each wall exterial wall on the level, and then list it in a column in the excel sheet.

The structure of the excel sheet are defines in line 125, 
the first part is column A wich is defind with describing text we have made.

The sekend part is column B here are the Roof level defined with each wall on the level and the informations about it.

column C here are the 2. level defined.

column D here are the 1. level defined.

In line 127 will Pandas print the output for you in an excel file, to make it work on your computer you have to make the parth to were you want it to be saved.
![Line 127](https://github.com/AnjaHolmquist/A4-Final_project/blob/main/Line%20127.png)

Here you can see a short video about this use case.
![Video](https://github.com/AnjaHolmquist/A4-Final_project/blob/main/Group11-A4_Final_Project.zip)

# If we had more time
We wanted to make this tool more userfrendly, so maybe there could be a kind af drop down meny so the contractor can select the building part he/here wants to know about.
it could also be a HTML file where the Client can send an link and the contractor could send the answer back to the Client.
