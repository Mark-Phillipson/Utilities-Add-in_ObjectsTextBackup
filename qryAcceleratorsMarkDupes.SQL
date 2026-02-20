UPDATE tmpAccelerators SET tmpAccelerators.Duplicated = True
WHERE (((tmpAccelerators.Accelerator) In (SELECT [Accelerator] FROM [tmpAccelerators] As Tmp GROUP BY [Accelerator],[TabPageName] HAVING Count(*)>1  And [TabPageName] = [tmpAccelerators].[TabPageName])) AND ((Len([tmpAccelerators]![Accelerator]))=1));

