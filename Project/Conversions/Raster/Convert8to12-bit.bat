mkdir "8bit"
for %%a in (*.tif) do mr_file -o1 -T -Y8 -m0 -M4095 -E -D0 -S256 -Ka %%a ./"8bit"/%%a
