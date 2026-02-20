SELECT [Object Properties].PropertyName, IIf([Object Properties].FormApplicable,"Yes","no") AS Expr1, [Object Properties].ReportApplicable, [Object Properties].TextBoxApplicable, [Object Properties].LabelApplicable, [Object Properties].ComboBoxApplicable, [Object Properties].ListboxApplicable, [Object Properties].LineApplicable, [Object Properties].CommandButtonApplicable, [Object Properties].ToggleButtonApplicable, [Object Properties].CheckboxApplicable, [Object Properties].SubFormorReportApplicable
FROM [Object Properties]
WHERE ((([Object Properties].PropertyName) Like "*Lef*") AND (([Object Properties].MakeAvailable)=True))
ORDER BY [Object Properties].PropertyName;

