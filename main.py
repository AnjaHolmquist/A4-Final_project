#11034 - Advanced BIM E21
#Team 11
#------------------------------------
#Rene Bjørnholt Christensen - s180524
#Anja Holmquist Egebo - s181535
#Vahid Soleimaniazar - s180096
#------------------------------------
#Assignent 05 - BIM Tool
#------------------------------------

#first we need to import hte ifcopenshell and the data set
import ifcopenshell
model = ifcopenshell.open('model/duplex.ifc')

#then we import Pandas so we can make a excel file with the external wall data for the contractors.
import pandas as pd

#this loop are going to import the price of the loadbearing exterior wall, the price comes from the molio JSON file from learn.
import json

materialName = "pladsstÃ¸bt beton" #in-situ concrete for the load bearing part of the exterior wall
materialPrice = 0
with open('data.json') as fp:
    data = json.load(fp)
    for x in data:
       if not isinstance(x, str) and x["name"] and materialName in x["name"]:
         materialPrice = x["price"]
         break

#to make the excel sheet look nice and clear are we making a list for each level we are starting from the top.
#this is the list for the roof.
roof_walls = []
# this for loop will give us the informations for the external walls in level roof.
for ifcwall_element in model.by_type('IfcWall'):
    # if IfcWall.is_a(''):
        for relContainedInSpatialStructure in ifcwall_element.ContainedInStructure:
           if relContainedInSpatialStructure.RelatingStructure.Name == "Roof":
                if (ifcwall_element.Name).find("Exterior") > 0:
                    print(ifcwall_element.Name)
                    print(relContainedInSpatialStructure.RelatingStructure.Name)
                    for definition in ifcwall_element.IsDefinedBy:
                        if definition.is_a('IfcRelDefinesByProperties'):
                            property_set = definition.RelatingPropertyDefinition			
                            # Sort by the name of the propertySet
                            if property_set.Name == 'PSet_Revit_Dimensions':
                                # Look through properties and find Volume
                                for property in property_set.HasProperties:
                                    # Replace Volume with other property e.g. Area
                                    if property.Name == "Area":
                                        print(property.NominalValue.wrappedValue)
                                        value = ifcwall_element.Name
                                        value1 = relContainedInSpatialStructure.RelatingStructure.Name
                                        value2 = property.NominalValue.wrappedValue
                                        value9 = materialPrice
                                        value10 = property.NominalValue.wrappedValue * materialPrice
                                        #Now all the necessary information can be added to the list.
                                        roof_walls.append(value)
                                        roof_walls.append(value2)
                                        roof_walls.append(value9)
                                        roof_walls.append(value10)

#this is the list for the 1. floor.
level2_walls = []
# this for loop will give us the informations for the external walls in level 2.
for ifcwall_element in model.by_type('IfcWall'):
    # if IfcWall.is_a(''):
        for relContainedInSpatialStructure in ifcwall_element.ContainedInStructure:
           if relContainedInSpatialStructure.RelatingStructure.Name == "Level 2":
                if (ifcwall_element.Name).find("Exterior") > 0:
                    print(ifcwall_element.Name)
                    print(relContainedInSpatialStructure.RelatingStructure.Name)
                    for definition in ifcwall_element.IsDefinedBy:
                        if definition.is_a('IfcRelDefinesByProperties'):
                            property_set = definition.RelatingPropertyDefinition			
                            # Sort by the name of the propertySet
                            if property_set.Name == 'PSet_Revit_Dimensions':
                                # Look through properties and find Volume
                                for property in property_set.HasProperties:
                                    # Replace Volume with other property e.g. Area
                                    if property.Name == "Area":
                                        print(property.NominalValue.wrappedValue)
                                        value3 = ifcwall_element.Name
                                        value4 = relContainedInSpatialStructure.RelatingStructure.Name
                                        value5 = property.NominalValue.wrappedValue
                                        value11 = materialPrice
                                        value12 = property.NominalValue.wrappedValue * materialPrice
                                        #Now all the necessary information can be added to the list.
                                        level2_walls.append(value3)
                                        level2_walls.append(value5)
                                        level2_walls.append(value11)
                                        level2_walls.append(value12)

#this is the list for the ground floor.
level1_walls = []
# this for loop will give us the informations for the external walls in level 1.
for ifcwall_element in model.by_type('IfcWall'):
    # if IfcWall.is_a(''):
        for relContainedInSpatialStructure in ifcwall_element.ContainedInStructure:
           if relContainedInSpatialStructure.RelatingStructure.Name == "Level 1":
                if (ifcwall_element.Name).find("Exterior") > 0:
                    print(ifcwall_element.Name)
                    print(relContainedInSpatialStructure.RelatingStructure.Name)
                    for definition in ifcwall_element.IsDefinedBy:
                        if definition.is_a('IfcRelDefinesByProperties'):
                            property_set = definition.RelatingPropertyDefinition			
                            # Sort by the name of the propertySet
                            if property_set.Name == 'PSet_Revit_Dimensions':
                                # Look through properties and find Volume
                                for property in property_set.HasProperties:
                                    # Replace Volume with other property e.g. Area
                                    if property.Name == "Area":
                                        print(property.NominalValue.wrappedValue)
                                        value6 = ifcwall_element.Name
                                        value7 = relContainedInSpatialStructure.RelatingStructure.Name
                                        value8 = property.NominalValue.wrappedValue
                                        value13 = materialPrice
                                        value14 = property.NominalValue.wrappedValue * materialPrice
                                        #Now all the necessary information can be added to the list.
                                        level1_walls.append(value6)
                                        level1_walls.append(value8)
                                        level1_walls.append(value13)
                                        level1_walls.append(value14)


df = pd.DataFrame({"Elevation":["Walltype, Material, Wall id","Area [m2]","Concrete price pr. m2 [DKKR]","Concrete price for the wall [DKKR]","Walltype, Material, Wall id","Area [m2]","Concrete price pr. m2 [DKKR]","Concrete price for the wall [DKKR]","Walltype, Material, Wall id","Area [m2]","Concrete price pr. m2 [DKKR]","Concrete price for the wall [DKKR]","Walltype, Material, Wall id","Area [m2]","Concrete price pr. m2 [DKKR]","Concrete price for the wall [DKKR]"], value1: roof_walls, value4: level2_walls, value7: level1_walls,})
print (df)
df.to_excel(r'C:\Code\Group11-A4_Final_Project\output\Costplan_for_exterior_walls.xlsx', sheet_name='Load bearing walls', index = False)

#Thanks to Tim, Ann-Britt, Martyna and Piotr for the help to understand how to make a script. ;)