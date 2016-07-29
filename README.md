# AngularExcelExporter

Simple javascript "class" to perform dato to excel

### Example
```
  var fe = new FileExporter("Nombre del proyecto");
  fe.createExcelFile(
        [
          [['Name', 'Job'],['Andrew C. Oliver', 'Strategic Developer']], // sheet 1
          [['Id', 'Note'], ['1', 'Hello word!'], [2, 'It's works!']] // sheet 2
        ], ['Persons', 'Notes']);
  fe.download();
```
Javascript download data to Excel file

##### Note: When I made this, I need it in AngularJS, so It's work very well there

