
# -*- coding: utf-8 -*-	
"""	
CBYD tickets into WebLayer auto update	
"""	

#import arcpy	
import arcpy	
import os, sys	
from arcgis import GIS	

**#xxx denotes user must change	

#set workplace and overwrite feature	
arcpy.env.workspace = "xxx"	
arcpy.env.overwriteOutput = True	
workspace = "xxx"	

#variables	
in_table = "xxx"	
Output_Table = "CBYD_Tickets_All"	
Refined_Table = "CBYD_Refined"	

#convert excel sheet to dbf table	
arcpy.conversion.ExcelToTable(in_table, Output_Table, "All Tickets", "")	

#add numeric fields to X,Y cordinates to avoid null geometry	
arcpy.management.AddField(Output_Table, "LongNum", "FLOAT")	
arcpy.management.AddField(Output_Table, "LatNum", "FLOAT")	

#select features 30 days before today and export to table	
arcpy.conversion.TableToTable(Output_Table, workspace, Refined_Table, "Date >= CURRENT_DATE() - 30")	

#fill in fields long at lat numerical	
arcpy.management.CalculateField(Refined_Table, "LongNum", "!Longitude!", "PYTHON3")	
arcpy.management.CalculateField(Refined_Table, "LatNum", "!Latitude!", "PYTHON3")	

#x,y points to shapefile	
arcpy.management.XYTableToPoint(Refined_Table, "CBYD" , "LongNum", "LatNum", "")	

#add layer to arcgis pro project	
xy_layer = "xxx"	
aprx = arcpy.mp.ArcGISProject("xxx")	
map = aprx.listMaps()[0]	
map.addDataFromPath(xy_layer)	
aprx.save()	

### Start setting variables	
# Set the path to the project	
prjPath = r"xxx"	

# Update the following variables to match:	
# Feature service/SD name in arcgis.com, user/password of the owner account	
sd_fs_name = "CBYD"	
portal = "xxx" # Can also reference a local portal	
user = "xxx"	
password = "xxx"	

# Set sharing options	
shrOrg = True	
shrEveryone = False	
shrGroups = ""	

### End setting variables	

# Local paths to create temporary content	
relPath = os.path.dirname(prjPath)	
sddraft = os.path.join(relPath, "WebUpdate.sddraft")	
sd = os.path.join(relPath, "WebUpdate.sd")	

# Create a new SDDraft and stage to SD	
print("Creating SD file")	
arcpy.env.overwriteOutput = True	
prj = arcpy.mp.ArcGISProject(prjPath)	
mp = prj.listMaps()[0]	
arcpy.mp.CreateWebLayerSDDraft(mp, sddraft, sd_fs_name, "MY_HOSTED_SERVICES", "FEATURE_ACCESS","", True, True)	
arcpy.StageService_server(sddraft, sd)	

print("Connecting to {}".format(portal))	
gis = GIS(portal, user, password)	

# Find the SD, update it, publish /w overwrite and set sharing and metadata	
print("Search for original SD on portal…")	
sdItem = gis.content.search("{} AND owner:{}".format(sd_fs_name, user), item_type="Service Definition")[0]	
print("Found SD: {}, ID: {} n Uploading and overwriting…".format(sdItem.title, sdItem.id))	
sdItem.update(data=sd)	
print("Overwriting existing feature service…")	
fs = sdItem.publish(overwrite=True)	

if shrOrg or shrEveryone or shrGroups:	
  print("Setting sharing options…")	
  fs.share(org=shrOrg, everyone=shrEveryone, groups=shrGroups)	

print("Finished updating: {} – ID: {}".format(fs.title, fs.id))	

#remove previous cbyd layer	
map.removeLayer(map.listLayers()[0])	
aprx.save()
