SELECT tmpControls.id, tmpControls.Name, tmpControls.Parent, tblControlCodes.Control_Type AS Type, tmpControls.Controltype
FROM tmpControls INNER JOIN tblControlCodes ON tmpControls.ControlType = tblControlCodes.Control_Code
ORDER BY tmpControls.Name;

