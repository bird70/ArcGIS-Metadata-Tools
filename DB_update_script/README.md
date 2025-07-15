# DB Metadata Editing script for ArcGIS GeoDatabase 

When you frequently create and export datasets for delivery, it can be tedious to keep track of metadata updates. This script has been designed to be used in ArcGIS Pro:

It reads the contents of a metadata template file (XML) that has been pre-populated with values that you want to be used (example provided). Those metadata
are then used to update the items in your File Geodatabase or SDE.

Use with care: CONTENTS WILL BE OVERWRITTEN - THERE'S NO UNDO. (as a safeguard, the previous contents are exported to a temp directory, but those would need to
be re-imported manually). 

TEST WITH A DEV DATABASE BEFORE USING IN YOUR REAL GEODATABASE
