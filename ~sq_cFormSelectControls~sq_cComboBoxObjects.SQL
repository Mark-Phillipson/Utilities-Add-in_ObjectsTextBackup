SELECT tmpControls.ControlType, tblControlCodes.Control_Type
FROM tmpControls INNER JOIN tblControlCodes ON tmpControls.ControlType = tblControlCodes.Control_Code
GROUP BY tmpControls.ControlType, tblControlCodes.Control_Type;

