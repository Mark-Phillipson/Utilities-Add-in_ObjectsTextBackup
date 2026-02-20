SELECT tmpFileNamesEntered.FileName
FROM tmpFileNamesEntered
GROUP BY tmpFileNamesEntered.FileName
HAVING (((tmpFileNamesEntered.FileName) Is Not Null))
ORDER BY tmpFileNamesEntered.FileName;

