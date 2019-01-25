#directions
#Create a folder for the sql query export
#Just go ahead and shove this file in that newly created folder 
#run SQL query 
#right click on results and select "Save Results as"
#Name the .cvs file the same name you are going to call the folder in Natural Resources (example: EmrSikesAct.csv)



#Put the name table here (exclude .cvs)
$tableName = "InrmpPoamActivev2OandETab"

#Add the path to your current NaturalResources.API\CN.Reporting\Excel

$excelFolder= "C:\Source\repos\NaturalResources.API\CN.Reporting\Excel\"


$newFolder = "${excelFolder}${tableName}\"


#save 

#right click on this file
#select "Run With Powershell"

#To check if it worked, Open Visual studio and go to Natural Resources.APT\CN.Reporting\Excel\ 
#Select the one button to "Show All Files"
#If you the file, add it to the solution and take a look and tell me if you want it better or different or whatever. 

if(!(Test-Path $newFolder)){

New-Item -Path $newFolder -ItemType "directory"

}

$configClassPath = "${newFolder}${tableName}Config.cs"

#type the path where the csv file is here
$rows = Import-Csv "${tableName}.csv"
#type the file path to the table here. 

$questions = Import-Csv "$InrmpPoamActivev2Questions.csv"



if(Test-Path $configClassPath ){

Remove-Item $configClassPath 

}

New-Item $configClassPath -ItemType "file" -Value "using CN.Core.Utilities;
using CN.DataLayer.Reporting;
using System;
using System.Collections.Generic;

namespace CN.Reporting.Excel.${tableName}
{ 
public class ${tableName}Config
{
 public List<ReportColumnInfoModel> GetColumnConfig()
 {   
 var rtrn = new List<ReportColumnInfoModel>();
 
 "


ForEach ($row in $rows){
$header=$row.HeaderText
$description =$row.Description
$alignment = $row.Alignment
$dataType = $row.DataType
$group=$row.Group

if ([string]::IsNullOrEmpty($header) -or ($header -eq "NULL")){
ForEach($question in $questions){

if ($question.Description -eq $row.Description)


{

$header=($question.Number+". "+$question.QuestionText)

}

}

}
$lineArr= @()


$lineArr+= ("HeaderText=""${header}""")
$lineArr+=if(($row.QuestionId)-and ($row.QuestionId -ne 'NULL')){("QuestionId="+$row.QuestionId)}else  {"QuestionId=-1"} 
$lineArr+= if($row.Column){ ("Column="+$row.Column)}  else {"Column=1"}

$lineArr+= if($row.Width) {("Width="+$row.Width)} else {"Width=20"} 
$lineArr+=if([string]::IsNullOrEmpty($row.Alignment)){"Alignment=""LEFT"""}else{("Alignment="""+$row.Alignment +""" ")}




$lineArr+= if((-Not [string]::IsNullOrEmpty($group))-and ($group -ne 'NULL')){("Group="""+$row.Group+""" ")}else{("Group=""""")}




$lineArr+=if($row.Height){("Height="+$row.Height)}else{"Height=15"}  


$lineArr+=if(([string]::IsNullOrEmpty($row.DataType))){"DataType=""STRING"""} else {("DataType="""+$row.DataType+""" ")}
$lineArr+= ("Description="""+$row.Description +""" ")

$lineArr+= if($row.ShowComment -eq 0){"ShowComment=false"} else {"ShowComment=true"}
$lineArr+= if($row.Visible-eq 0){"Visible=false"} else {"Visible=true"}
$lineArr+= if($row.TotalField -eq 0){"TotalField=false"} else {"TotalField=true"}
$lineArr+= if($row.ItemInTable-eq 0){"ItemInTable=false"} else {"ItemInTable=true"}
$lineArr+= if($row.WrapText -eq 0){"WrapText=false"} else {"WrapText=true"}



$joinedArr=($lineArr -join ", ")
$line = ("rtrn.Add(new ReportColumnInfoModel{"+$joinedArr+"});")



$line | Add-Content "${configClassPath}"
}


Add-Content -Path ${configClassPath} -Value "return rtrn;

}
}
}"


$classFilePath = "${newFolder}${tableName}.cs"




#if(Test-Path $classFilePath ){

#Remove-Item $classFilePath 

#}




#Copy-Item -Path "${excelFolder}${tableName}.cs" -Destination "${newFolder}"

