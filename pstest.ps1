#import a csv from Sql and name it the Name of the table 

#Put the name table (exclude .cvs)asdfasdfasdfasdf
$tableName = "Forestry"

$apiFolder = "C:\Source\repos\NaturalResources.API\CN.Reporting\Excel\${tableName}\"

if(!(Test-Path $apiFolder)){

New-Item -Path $apiFolder -ItemType "directory"

}

$classFilePath = "${apiFolder}${tableName}Config.cs"

#type the path where the csv file is here
$rows = Import-Csv "${tableName}.csv"
#type the file path to the table here. 



if(Test-Path $classFilePath ){

Remove-Item $classFilePath 

}

New-Item $classFilePath -ItemType "file" -Value "using CN.Core.Utilities;
using CN.DataLayer.Reporting;
using System;
using System.Collections.Generic;

namespace CN.Reporting.Excel.${tableName}
{ 
public class ${tableName}Config
{
 public List<ReportColumnInfoModel> GetColumnConfig(int dataCallYear)
 {   
 var rtrn = new List<ReportColumnInfoModel>();
 
 "


ForEach ($row in $rows){
$header=$row.HeaderText
$description =$row.Description
$alignment = $row.Alignment
$dataType = $row.DataType
$group=$row.Group

$lineArr= @()


$lineArr+= ("HeaderText="""+$row.HeaderText+""" ")

$lineArr+= if($row.Column){ ("Column="+$row.Column)}  else {"Column=1"}

$lineArr+= if($row.Width) {("Width="+$row.Width)} else {"Width=20"} 
$lineArr+= if((-Not [string]::IsNullOrEmpty($group))-and ($group -ne 'NULL')){("Group="""+$row.Group+""" ")}else{("Group=""""")}

$lineArr+=if(($row.QuestionId)-and ($row.QuestionId -ne 'NULL')){("QuestionId="+$row.QuestionId)}else  {"QuestionId=-1"} 


$lineArr+=if($row.Height){("Height="+$row.Height)}else{"Height=15"}  

$lineArr+=if([string]::IsNullOrEmpty($row.Alignment)){"Alignment=""LEFT"""}else{("Alignment="""+$row.Alignment +""" ")}
$lineArr+=if(([string]::IsNullOrEmpty($row.DataType))){"DataType=""STRING"""} else {("DataType="""+$row.DataType+""" ")}
$lineArr+= ("Description="""+$row.Description +""" ")
$lineArr+= if($row.Visible-eq 0){"Visible=false"} else {"Visible=true"}
$lineArr+= if($row.TotalField -eq 0){"TotalField=false"} else {"TotalField=true"}
$lineArr+= if($row.ItemInTable-eq 0){"ItemInTable=false"} else {"ItemInTable=true"}
$lineArr+= if($row.WrapText -eq 0){"WrapText=false"} else {"WrapText=true"}



$joinedArr=($lineArr -join ", ")
$line = ("rtrn.Add(new ReportColumnInfoModel{"+$joinedArr+"});")



$line | Add-Content "${classFilePath}"
}


Add-Content -Path ${classFilePath} -Value "return rtrn;

}
}
}"


$classFilePath = "${apiFolder}${tableName}.cs"





if(Test-Path $classFilePath ){

Remove-Item $classFilePath 

}

New-Item $classFilePath -ItemType "file" -Value "using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using CN.DataLayer;
using CN.DataLayer.Helpers;
using CN.DataLayer.Interfaces;
using CN.DataLayer.Reporting;
using CN.DataLayer.ReportingModels;
using CN.DataLayer.Utilities;
using CN.Reporting.Helpers;
using CN.Reporting.Interfaces;
using CN.Reporting.Models;
using SpreadsheetLight;
using CN.Core.Utilities;
using CN.DataLayer.Models;
using CN.Reporting.Excel.EmrSikesAct;
using CN.DataLayer.Repository;

namespace CN.Reporting.Excel.${tableName}
{
    public class ${tableName} : BaseExcelReport
    {
    
    }
    }
    
    
    "
    





