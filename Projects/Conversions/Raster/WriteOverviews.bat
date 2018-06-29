mkdir "oviews_restored"
for %%x in (*.tif) do mr_file -i1 -o1 -H1 -Y8 -E -D72 -S256 -Ka %%x ./"oviews_restored"/%%x REM (loops through all existing .tif files, writing new rasters to 'oviews_restored')
