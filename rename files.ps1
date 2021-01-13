

Get-ChildItem -Filter "*The.Big.Bang.Theory*" -Recurse | Rename-Item -NewName {$_.name -replace 'The.Big.Bang.Theory','TBBT' }

Get-ChildItem -Filter "*The Big Bang Theory*" -Recurse | Rename-Item -NewName {$_.name -replace 'The Big Bang Theory','TBBT' }