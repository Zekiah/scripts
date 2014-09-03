import arcpy

arcpy.env.overwriteOutput = 1
workspace = "D:\Temp"
arcpy.env.workspace = workspace

input1 = r"D:\Path_To...\Area_of_Interest.shp"
input2 = r"D:\Path_To...\Infrastructure_dataset.shp"
output1 = r"D:\Path_To...\Area_of_Interest_Projected.shp"
output2 = r"D:\Path_To...\Infrastructure_dataset_Projected.shp"


#sr = arcpy.SpatialReference(26918)
arcpy.Project_management(input1, output1, arcpy.SpatialReference(26918))
arcpy.Project_management(input2, output2, arcpy.SpatialReference(26918))


arcpy.Near_analysis(output1, output2, "", "LOCATION", "ANGLE")
field1 = "NEAR_DIST"
field2 = "NEAR_ANGLE"
cursor = arcpy.SearchCursor(output1)
for row in cursor:
    distance = (row.getValue(field1))
    angle = (row.getValue(field2))
distance = (distance / 1609.34)

arcpy.Delete_management(output1)
arcpy.Delete_management(output2)
print "The distance value is ",distance, " miles."
print "The angle value is ",angle, " degrees."

distance = str(round(distance,2))
angle = round(angle,2)
direction = ""

if angle < 0:
    angle = angle +360
    print angle
else:
    angle = angle

if angle >= 0 and angle < 22.5:
    direction = "east"
elif angle >= 22.5 and angle < 67.5:
    direction = "northeast"
elif angle >= 67.5 and angle < 112.5:
    direction = "north"
elif angle >= 112.5 and angle < 157.5:
    direction = "northwest"
elif angle >= 157.5 and angle < 202.5:
    direction = "west"
elif angle >= 202.5 and angle < 247.5:
    direction = "southwest"
elif angle >= 247.5 and angle < 292.5:
    direction = "south"
elif angle >= 292.5 and angle < 337.5:
    direction = "southeast"
elif angle >= 337.5 and angle < 360:
    direction = "east"
print "The direction value returned is "+direction+ "."   
dist_ang = "The nearest piece of infrastructure is approximately "+ distance+ " miles\n"+direction+" of the area of interest."
print dist_ang
    
        
