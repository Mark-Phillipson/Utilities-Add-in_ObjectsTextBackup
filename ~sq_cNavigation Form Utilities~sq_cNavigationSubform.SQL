SELECT SwitchboardItems_AddIn.Argument, SwitchboardItems_AddIn.MenuHelpText
FROM SwitchboardItems_AddIn
WHERE (((SwitchboardItems_AddIn.Argument) Is Not Null) AND ((SwitchboardItems_AddIn.MenuHelpText) Is Not Null))
ORDER BY SwitchboardItems_AddIn.Argument;

