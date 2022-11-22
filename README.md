# A4-Final_project
###### We have made an Excel file with the costplan for the Duplex files loadbaering walls.

## How the tool works:
###### In the main.py file have we made a script, to help the Contractors, the client and the workers. 
###### The main py file are making an excel file wich contains the: area of the exterior wallparts, it shows the prise for the loadbearing wall pr m2., and the prise for each loadbearing wallparts in building.
###### The output looks like this. 
![Our Cost-plan](https://github.com/AnjaHolmquist/A4-Final_project/blob/main/the%20costplan.png)

## But here is how we get to this output.
###### First of all we have to open the main.py file, as you can se now,
###### it imports ifcopenshell, the model, Pandas (This is the program that makes the excel file in the buttom of the script), and then the json file (The json file is the molio price data file frol learn)
###### the first loop sorts out the prise for the in-situ loadbaering concrete wall in the json file.
###### the sekend loop makes a list for the Roof level, afterworts it sorts out the exterialwalls on level Roof and gives us the materials and the area in m2
###### now we have made some Values we want the list to contain so we can see it in the excel file.
###### so the first thing is the value 
