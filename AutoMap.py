from timeit import default_timer as timer
start = timer()
import arcpy
import arcpy.mapping


print 'success'
end = timer()
print "Completed in : " + str(end - start) + "s"

#Check if file is valid for arc
def isValid(inPath):
    return arcpy.Exists(inPath)


#Arc starts here
arcpy.env.workspace =  "C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Maintained\CP_Automate.gdb"
#arcpy.env.workspace =  "Database Connection/C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Maintained\CP_Automate.gdb"
mxd = arcpy.mapping.MapDocument(r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Maps\20170429_DGMCoverage.mxd")
intrusive_table = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Weekly\Intrusive_Results_Master_0524.csv"
anomaly_table = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Weekly\Anomaly_Table.xls"


arcpy.mapping.ListDataFrames(mxd)[0].name
arcpy.mapping.ListDataFrames(mxd)[1].name
mxd.activeDataFrame

arcpy.RefreshTOC()

for lyr in arcpy.mapping.ListLayers(mxd):
    print lyr.name
try:
    #Make XY Here
    x_coords = "Easting"
    y_coords = "Northing"
    out_Layer = "Intrusive_layer1"
    saved_Layer = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Output\Intrusive1.lyr"
        
    # Set the spatial reference
    spRef = r"Coordinate Systems\Projected Coordinate Systems\Utm\Nad 1983\NAD 1983 (2011) UTM Zone 19N.prj"

    # Make the XY event layer...
    arcpy.MakeXYEventLayer_management(intrusive_table, x_coords, y_coords, out_Layer, spRef)

    # Save to a layer file
    arcpy.SaveToLayerFile_management(out_Layer, saved_Layer)

except Exception as err:
    print(err.args[0])
#Access from ListLayers index works!
mrs = arcpy.mapping.ListLayers(mxd)[8]
#mrs = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Maintained\CP_Automate.gdb\Water_101316_Acreage_Clip"
#Clip starts here
#in_features = saved_Layer
#in_f Needs Definition Query attached...
in_f = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Output\Intrusive1.lyr"
clip_f = mrs
out_f = r"C:\Users\alogan\Desktop\Arcpy Automate Maps\Data\Output\Intrusive_clip1.shp"
xy_tolerance = ""

#NOTE - Clip seems to produce faulty output if input is SHP and clip is .GDB database
# Execute Clip
arcpy.Clip_analysis(in_f, clip_f, out_f)










"""
lyrs = arcpy.mapping.ListLayers(mxd)
#Lists layers
for lyr in lyrs:     
    try:         
        print lyr.name,lyr.dataSource + '\n'
    except:         
        pass
"""
