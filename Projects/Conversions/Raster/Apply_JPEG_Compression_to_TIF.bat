mkdir "Oviews"
for %%a in (*.tif) do mr_file -T -E -C j -Q 5 -S256 -Ka %%a ./"OViews"/%%a
