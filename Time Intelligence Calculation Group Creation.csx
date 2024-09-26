#r "Microsoft.VisualBasic"
using Microsoft.VisualBasic;
//
// CHANGELOG:
// '2021-05-01 / B.Agullo / 
// '2021-05-17 / B.Agullo / added affected measure table
// '2021-06-19 / B.Agullo / data label measures
// '2021-07-10 / B.Agullo / added flag expression to avoid breaking already special format strings
// '2021-09-23 / B.Agullo / added code to prompt for parameters (code credit to Daniel Otykier) 
// '2021-09-27 / B.Agullo / added code for general name 
// '2022-10-11 / B.Agullo / added MMT and MWT calc item groups
// '2024-09-26 / back1ply / revamed measures, fixed bug
//
// by Bernat Agulló
// twitter: @AgulloBernat
// www.esbrina-ba.com/blog
//
// REFERENCE: 
// Check out https://www.esbrina-ba.com/time-intelligence-the-smart-way/ where this script is introduced
// 
// FEATURED: 
// this script featured in GuyInACube https://youtu.be/_j0iTUo2HT0
//
// THANKS:
// shout out to Johnny Winter for the base script and SQLBI for daxpatterns.com

//select the measures that you want to be affected by the calculation group
//before running the script. 
//measure names can also be included in the following array (no need to select them) 
string[] preSelectedMeasures = {}; //include measure names in double quotes, like: {"Profit","Total Cost"};

//AT LEAST ONE MEASURE HAS TO BE AFFECTED!, 
//either by selecting it or typing its name in the preSelectedMeasures Variable

//
// ----- do not modify script below this line -----
//

string affectedMeasures = "{";

int i = 0; 

for (i=0;i<preSelectedMeasures.GetLength(0);i++){
  
    if(affectedMeasures == "{") {
    affectedMeasures = affectedMeasures + "\"" + preSelectedMeasures[i] + "\"";
    }else{
        affectedMeasures = affectedMeasures + ",\"" + preSelectedMeasures[i] + "\"" ;
    }; 
    
};


if (Selected.Measures.Count != 0) {
    
    foreach(var m in Selected.Measures) {
        if(affectedMeasures == "{") {
        affectedMeasures = affectedMeasures + "\"" + m.Name + "\"";
        }else{
            affectedMeasures = affectedMeasures + ",\"" + m.Name + "\"" ;
        };
    };  
};

//check that by either method at least one measure is affected
if(affectedMeasures == "{") { 
    Error("No measures affected by calc group"); 
    return; 
};

string calcGroupName; 
string columnName; 

if(Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group")) {
    calcGroupName = Model.CalculationGroups.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Name;
    
}else {
    calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
}; 

if(calcGroupName == "") return;


if(Model.CalculationGroups.Any(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group")) {
    columnName = Model.Tables.Where(cg => cg.GetAnnotation("@AgulloBernat") == "Time Intel Calc Group").First().Columns.First().Name;
    
}else {
    columnName = Interaction.InputBox("Provide a name for your Calc Group Column", "Calc Group Column Name", calcGroupName, 740, 400);
}; 

if(columnName == "") return;

string affectedMeasuresTableName; 

if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    affectedMeasuresTableName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Name;

} else {
    affectedMeasuresTableName = Interaction.InputBox("Provide a name for affected measures table", "Affected Measures Table Name", calcGroupName  + " Affected Measures", 740, 400);

};

if(affectedMeasuresTableName == "") return;

string affectedMeasuresColumnName;

if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    affectedMeasuresColumnName = Model.Tables.Where(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table").First().Columns.First().Name;

} else {
    affectedMeasuresColumnName = Interaction.InputBox("Provide a name for affected measures column", "Affected Measures Table Column Name", "Measure", 740, 400);

};

if(affectedMeasuresColumnName == "") return;

string labelAsValueMeasureName = "Label as Value Measure"; 
string labelAsFormatStringMeasureName = "Label as format string"; 


 // '2021-09-24 / B.Agullo / model object selection prompts! 
var factTable = SelectTable(label: "Select your fact table");
if(factTable == null) return;

var factTableDateColumn = SelectColumn(factTable.Columns, label: "Select the main date column");
if(factTableDateColumn == null) return;

Table dateTableCandidate = null;

if(Model.Tables.Any
    (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table" 
        || x.Name == "Date" 
        || x.Name == "Calendar")){
            dateTableCandidate = Model.Tables.Where
                (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table" 
                    || x.Name == "Date" 
                    || x.Name == "Calendar").First();

};

var dateTable = 
    SelectTable(
        label: "Select your date table",
        preselect:dateTableCandidate);

if(dateTable == null) {
    Error("You just aborted the script"); 
    return;
} else {
    dateTable.SetAnnotation("@AgulloBernat","Time Intel Date Table");
}; 


Column dateTableDateColumnCandidate = null; 

if(dateTable.Columns.Any
            (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date")){
    dateTableDateColumnCandidate = dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Date Column" || x.Name == "Date").First();
};

var dateTableDateColumn = 
    SelectColumn(
        dateTable.Columns, 
        label: "Select the date column",
        preselect: dateTableDateColumnCandidate);

if(dateTableDateColumn == null) {
    Error("You just aborted the script"); 
    return;
} else { 
    dateTableDateColumn.SetAnnotation("@AgulloBernat","Time Intel Date Table Date Column"); 
}; 

Column dateTableYearColumnCandidate = null;
if(dateTable.Columns.Any(x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year")){
    dateTableYearColumnCandidate = dateTable.Columns.Where
        (x => x.GetAnnotation("@AgulloBernat") == "Time Intel Date Table Year Column" || x.Name == "Year").First();
};

var dateTableYearColumn = 
    SelectColumn(
        dateTable.Columns, 
        label: "Select the year column", 
        preselect:dateTableYearColumnCandidate);

if(dateTableYearColumn == null) {
    Error("You just aborted the script"); 
    return;
} else {
    dateTableYearColumn.SetAnnotation("@AgulloBernat","Time Intel Date Table Year Column"); 
};


//these names are for internal use only, so no need to be super-fancy, better stick to daxpatterns.com model
string ShowValueForDatesMeasureName = "_ShowValueForDates";
string dateWithTransactionsColumnName = "DateWithTransactions";

// '2021-09-24 / B.Agullo / I put the names back to variables so I don't have to touch the script
string factTableName = factTable.Name;
string factTableDateColumnName = factTableDateColumn.Name;
string dateTableName = dateTable.Name;
string dateTableDateColumnName = dateTableDateColumn.Name;
string dateTableYearColumnName = dateTableYearColumn.Name; 

// '2021-09-24 / B.Agullo / this is for internal use only so better leave it as is 
string flagExpression = "UNICHAR( 8204 )"; 

string calcItemProtection = "<CODE>"; //default value if user has selected no measures
string calcItemFormatProtection = "<CODE>"; //default value if user has selected no measures

// check if there's already an affected measure table
if(Model.Tables.Any(t => t.GetAnnotation("@AgulloBernat") == "Time Intel Affected Measures Table")) {
    //modifying an existing calculated table is not risk-free
    Info("Make sure to include measure names to the table " + affectedMeasuresTableName);
} else { 
    // create calculated table containing all names of affected measures
    // this is why you need to enable 
    if(affectedMeasures != "{") { 
        
        affectedMeasures = affectedMeasures + "}";
        
        string affectedMeasureTableExpression = 
            "SELECTCOLUMNS(" + affectedMeasures + ",\"" + affectedMeasuresColumnName + "\",[Value])";

        var affectedMeasureTable = 
            Model.AddCalculatedTable(affectedMeasuresTableName,affectedMeasureTableExpression);
        
        affectedMeasureTable.FormatDax(); 
        affectedMeasureTable.Description = 
            "Measures affected by " + calcGroupName + " calculation group." ;
        
        affectedMeasureTable.SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table"); 
       
        // this causes error
        // affectedMeasureTable.Columns[affectedMeasuresColumnName].SetAnnotation("@AgulloBernat","Time Intel Affected Measures Table Column");

        affectedMeasureTable.IsHidden = true;     
        
    };
};

//if there were selected or preselected measures, prepare protection code for expression and formatstring
string affectedMeasuresValues = "VALUES('" + affectedMeasuresTableName + "'[" + affectedMeasuresColumnName + "])";

calcItemProtection = 
    "SWITCH(" + 
    "   TRUE()," + 
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
    "   <CODE> ," + 
    "   ISSELECTEDMEASURE([" + labelAsValueMeasureName + "])," + 
    "   <LABELCODE> ," + 
    "   SELECTEDMEASURE() " + 
    ")";
    
    
calcItemFormatProtection = 
    "SWITCH(" + 
    "   TRUE() ," + 
    "   SELECTEDMEASURENAME() IN " + affectedMeasuresValues + "," + 
    "   <CODE> ," + 
    "   ISSELECTEDMEASURE([" + labelAsFormatStringMeasureName + "])," + 
    "   <LABELCODEFORMATSTRING> ," +
    "   SELECTEDMEASUREFORMATSTRING() " + 
    ")";

    
string dateColumnWithTable = "'" + dateTableName + "'[" + dateTableDateColumnName + "]"; 
string yearColumnWithTable = "'" + dateTableName + "'[" + dateTableYearColumnName + "]"; 
string factDateColumnWithTable = "'" + factTableName + "'[" + factTableDateColumnName + "]";
string dateWithTransactionsColumnWithTable = "'" + dateTableName + "'[" + dateWithTransactionsColumnName + "]";
string calcGroupColumnWithTable = "'" + calcGroupName + "'[" + columnName + "]";

//check to see if a table with this name already exists
//if it doesn't exist, create a calculation group with this name
if (!Model.Tables.Contains(calcGroupName)) {
  var cg = Model.AddCalculationGroup(calcGroupName);
  cg.Description = "Calculation group for time intelligence. Availability of data is taken from " + factTableName + ".";
  cg.SetAnnotation("@AgulloBernat","Time Intel Calc Group"); 
};

//set variable for the calc group
Table calcGroup = Model.Tables[calcGroupName];

//if table already exists, make sure it is a Calculation Group type
if (calcGroup.SourceType.ToString() != "CalculationGroup") {
  Error("Table exists in Model but is not a Calculation Group. Rename the existing table or choose an alternative name for your Calculation Group.");
  return;
};

//adds the two measures that will be used for label as value, label as format string 
var labelAsValueMeasure = calcGroup.AddMeasure(labelAsValueMeasureName,"");
labelAsValueMeasure.Description = "Use this measure to show the year evaluated in tables"; 

var labelAsFormatStringMeasure = calcGroup.AddMeasure(labelAsFormatStringMeasureName,"0");
labelAsFormatStringMeasure.Description = "Use this measure to show the year evaluated in charts"; 

//by default the calc group has a column called Name. If this column is still called Name change this in line with specified variable
if (calcGroup.Columns.Contains("Name")) {
  calcGroup.Columns["Name"].Name = columnName;

};

calcGroup.Columns[columnName].Description = "Select value(s) from this column to apply time intelligence calculations.";
calcGroup.Columns[columnName].SetAnnotation("@AgulloBernat","Time Intel Calc Group Column"); 


//Only create them if not in place yet (reruns)
if(!Model.Tables[dateTableName].Columns.Any(C => C.GetAnnotation("@AgulloBernat") == "Date with Data Column")){
    string DateWithTransactionsCalculatedColumnExpression = 
        dateColumnWithTable + " <= MAX ( " + factDateColumnWithTable + " )";

    Column dateWithTransactionsColumn = dateTable.AddCalculatedColumn(dateWithTransactionsColumnName, DateWithTransactionsCalculatedColumnExpression);
    dateWithTransactionsColumn.SetAnnotation("@AgulloBernat","Date with Data Column");
};

if(!Model.Tables[dateTableName].Measures.Any(M => M.Name == ShowValueForDatesMeasureName)) {
    string ShowValueForDatesMeasureExpression = 
        "VAR LastDateWithData = " + 
        "    CALCULATE ( " + 
        "        MAX (  " + factDateColumnWithTable + " ), " + 
        "        REMOVEFILTERS () " +
        "    )" +
        "VAR FirstDateVisible = " +
        "    MIN ( " + dateColumnWithTable + " ) " + 
        "VAR Result = " +  
        "    FirstDateVisible <= LastDateWithData " +
        "RETURN " + 
        "    Result ";

    var ShowValueForDatesMeasure = dateTable.AddMeasure(ShowValueForDatesMeasureName, ShowValueForDatesMeasureExpression); 

    ShowValueForDatesMeasure.FormatDax();
};


// Defining expressions and format strings for each calculation item

string defFormatString = "SELECTEDMEASUREFORMATSTRING()";

string pctFormatString = 
"IF(" + 
"\n FIND( "+ flagExpression + ", SELECTEDMEASUREFORMATSTRING(), 1, -1 ) <> -1," + 
"\n SELECTEDMEASUREFORMATSTRING()," + 
"\n \"#,##0.# %\"" + 
"\n)";

// Define calculation item expressions

// Basic period calculations
string YTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        DATESYTD ( " + dateColumnWithTable + " ) " +
    "    ) " +
    ")";

string QTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        DATESQTD ( " + dateColumnWithTable + " ) " +
    "    ) " +
    ")";

string MTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        DATESMTD ( " + dateColumnWithTable + " ) " +
    "    ) " +
    ")";

// Previous periods
string PY =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, YEAR ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

string PQ =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, QUARTER ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

string PM =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, MONTH ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

// Period over period changes
string YOY =
    "VAR __ValueCurrentPeriod = SELECTEDMEASURE() " +
    "VAR __ValuePreviousPeriod = " + PY + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string YOYpct =
    "DIVIDE ( " +
    "    " + YOY + ", " +
    "    (" + PY + ") " +
    ")";

string QOQ =
    "VAR __ValueCurrentPeriod = SELECTEDMEASURE() " +
    "VAR __ValuePreviousPeriod = " + PQ + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string QOQpct =
    "DIVIDE ( " +
    "    " + QOQ + ", " +
    "    (" + PQ + ") " +
    ")";

string MOM =
    "VAR __ValueCurrentPeriod = SELECTEDMEASURE() " +
    "VAR __ValuePreviousPeriod = " + PM + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string MOMpct =
    "DIVIDE ( " +
    "    " + MOM + ", " +
    "    (" + PM + ") " +
    ")";

// Previous period to date
string PYTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        " + YTD + ", " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, YEAR ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

string PQTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        " + QTD + ", " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, QUARTER ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

string PMTD =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        " + MTD + ", " +
    "        CALCULATETABLE ( " +
    "            DATEADD ( " + dateColumnWithTable + ", -1, MONTH ), " +
    "            " + dateWithTransactionsColumnWithTable + " = TRUE " +
    "        ) " +
    "    ) " +
    ")";

// Period over period to date changes
string YOYTD =
    "VAR __ValueCurrentPeriod = " + YTD + " " +
    "VAR __ValuePreviousPeriod = " + PYTD + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string YOYTDpct =
    "DIVIDE ( " +
    "    " + YOYTD + ", " +
    "    (" + PYTD + ") " +
    ")";

string QOQTD =
    "VAR __ValueCurrentPeriod = " + QTD + " " +
    "VAR __ValuePreviousPeriod = " + PQTD + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string QOQTDpct =
    "DIVIDE ( " +
    "    " + QOQTD + ", " +
    "    (" + PQTD + ") " +
    ")";

string MOMTD =
    "VAR __ValueCurrentPeriod = " + MTD + " " +
    "VAR __ValuePreviousPeriod = " + PMTD + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string MOMTDpct =
    "DIVIDE ( " +
    "    " + MOMTD + ", " +
    "    (" + PMTD + ") " +
    ")";

// Parallel period calculations
string PYC =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        PARALLELPERIOD ( " + dateColumnWithTable + ", -1, YEAR ) " +
    "    ) " +
    ")";

string PQC =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        PARALLELPERIOD ( " + dateColumnWithTable + ", -1, QUARTER ) " +
    "    ) " +
    ")";

string PMC =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        PARALLELPERIOD ( " + dateColumnWithTable + ", -1, MONTH ) " +
    "    ) " +
    ")";

// Period over parallel period calculations
string YTDOPY =
    "VAR __ValueCurrentPeriod = " + YTD + " " +
    "VAR __ValuePreviousPeriod = " + PYC + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string YTDOPYpct =
    "DIVIDE ( " +
    "    " + YTDOPY + ", " +
    "    (" + PYC + ") " +
    ")";

string QTDOPQ =
    "VAR __ValueCurrentPeriod = " + QTD + " " +
    "VAR __ValuePreviousPeriod = " + PQC + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string QTDOPQpct =
    "DIVIDE ( " +
    "    " + QTDOPQ + ", " +
    "    (" + PQC + ") " +
    ")";

string MTDOPM =
    "VAR __ValueCurrentPeriod = " + MTD + " " +
    "VAR __ValuePreviousPeriod = " + PMC + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        NOT ISBLANK ( __ValueCurrentPeriod ) && NOT ISBLANK ( __ValuePreviousPeriod ), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string MTDOPMpct =
    "DIVIDE ( " +
    "    " + MTDOPM + ", " +
    "    (" + PMC + ") " +
    ")";

// Moving Annual Total
string MAT =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        SELECTEDMEASURE(), " +
    "        DATESINPERIOD ( " +
    "            " + dateColumnWithTable + ", " +
    "            MAX ( " + dateColumnWithTable + " ), " +
    "            -1, " +
    "            YEAR " +
    "        ) " +
    "    ) " +
    ")";

string PYMAT =
    "IF ( " +
    "    [" + ShowValueForDatesMeasureName + "], " +
    "    CALCULATE ( " +
    "        " + MAT + ", " +
    "        DATEADD ( " + dateColumnWithTable + ", -1, YEAR ) " +
    "    ) " +
    ")";

string MATG =
    "VAR __ValueCurrentPeriod = " + MAT + " " +
    "VAR __ValuePreviousPeriod = " + PYMAT + " " +
    "VAR __Result = " +
    "    IF ( " +
    "        ISBLANK ( __ValueCurrentPeriod ) || ISBLANK ( __ValuePreviousPeriod ), " +
    "        BLANK(), " +
    "        __ValueCurrentPeriod - __ValuePreviousPeriod " +
    "    ) " +
    "RETURN " +
    "    __Result";

string MATGpct =
    "DIVIDE ( " +
    "    " + MATG + ", " +
    "    (" + PYMAT + ") " +
    ")";

// Define the calcItems array
string[ , ] calcItems = 
{
    {"YTD",         YTD,        defFormatString,    "Year to date",                 "\"YTD\""},
    {"QTD",         QTD,        defFormatString,    "Quarter to date",              "\"QTD\""},
    {"MTD",         MTD,        defFormatString,    "Month to date",                "\"MTD\""},
    {"PY",          PY,         defFormatString,    "Previous Year",                "\"PY\""},
    {"YOY",         YOY,        defFormatString,    "Year over year",               "\"YOY\""},
    {"YOY %",       YOYpct,     pctFormatString,    "Year over year %",             "\"YOY %\""},
    {"PQ",          PQ,         defFormatString,    "Previous Quarter",             "\"PQ\""},
    {"QOQ",         QOQ,        defFormatString,    "Quarter over quarter",         "\"QOQ\""},
    {"QOQ %",       QOQpct,     pctFormatString,    "Quarter over quarter %",       "\"QOQ %\""},
    {"PM",          PM,         defFormatString,    "Previous Month",               "\"PM\""},
    {"MOM",         MOM,        defFormatString,    "Month over month",             "\"MOM\""},
    {"MOM %",       MOMpct,     pctFormatString,    "Month over month %",           "\"MOM %\""},
    {"PYTD",        PYTD,       defFormatString,    "Previous Year to date",        "\"PYTD\""},
    {"YOYTD",       YOYTD,      defFormatString,    "Year over year to date",       "\"YOYTD\""},
    {"YOYTD %",     YOYTDpct,   pctFormatString,    "Year over year to date %",     "\"YOYTD %\""},
    {"PQTD",        PQTD,       defFormatString,    "Previous Quarter to date",     "\"PQTD\""},
    {"QOQTD",       QOQTD,      defFormatString,    "Quarter over quarter to date", "\"QOQTD\""},
    {"QOQTD %",     QOQTDpct,   pctFormatString,    "Quarter over quarter to date %","\"QOQTD %\""},
    {"PMTD",        PMTD,       defFormatString,    "Previous Month to date",       "\"PMTD\""},
    {"MOMTD",       MOMTD,      defFormatString,    "Month over month to date",     "\"MOMTD\""},
    {"MOMTD %",     MOMTDpct,   pctFormatString,    "Month over month to date %",   "\"MOMTD %\""},
    {"PYC",         PYC,        defFormatString,    "Parallel Period Year",         "\"PYC\""},
    {"YTDOPY",      YTDOPY,     defFormatString,    "YTD over Parallel Year",       "\"YTDOPY\""},
    {"YTDOPY %",    YTDOPYpct,  pctFormatString,    "YTD over Parallel Year %",     "\"YTDOPY %\""},
    {"PQC",         PQC,        defFormatString,    "Parallel Period Quarter",      "\"PQC\""},
    {"QTDOPQ",      QTDOPQ,     defFormatString,    "QTD over Parallel Quarter",    "\"QTDOPQ\""},
    {"QTDOPQ %",    QTDOPQpct,  pctFormatString,    "QTD over Parallel Quarter %",  "\"QTDOPQ %\""},
    {"PMC",         PMC,        defFormatString,    "Parallel Period Month",        "\"PMC\""},
    {"MTDOPM",      MTDOPM,     defFormatString,    "MTD over Parallel Month",      "\"MTDOPM\""},
    {"MTDOPM %",    MTDOPMpct,  pctFormatString,    "MTD over Parallel Month %",    "\"MTDOPM %\""},
    {"MAT",         MAT,        defFormatString,    "Moving Annual Total",          "\"MAT\""},
    {"PYMAT",       PYMAT,      defFormatString,    "Previous Year MAT",            "\"PYMAT\""},
    {"MATG",        MATG,       defFormatString,    "MAT Growth",                   "\"MATG\""},
    {"MATG %",      MATGpct,    pctFormatString,    "MAT Growth %",                 "\"MATG %\""}
};

int j = 0;

// Create calculation items for each calculation with format string and description
foreach(var cg in Model.CalculationGroups) {
    if (cg.Name == calcGroupName) {
        for (j = 0; j < calcItems.GetLength(0); j++) {
            
            string itemName = calcItems[j,0];
            
            string itemExpression = calcItemProtection.Replace("<CODE>",calcItems[j,1]);
            itemExpression = itemExpression.Replace("<LABELCODE>",calcItems[j,4]); 
            
            string itemFormatExpression = calcItemFormatProtection.Replace("<CODE>",calcItems[j,2]);
            itemFormatExpression = itemFormatExpression.Replace("<LABELCODEFORMATSTRING>","\"\"\"\" & " + calcItems[j,4] + " & \"\"\"\"");
            
            string itemDescription = calcItems[j,3];
            
            if (!cg.CalculationItems.Contains(itemName)) {
                var nCalcItem = cg.AddCalculationItem(itemName, itemExpression);
                nCalcItem.FormatStringExpression = itemFormatExpression;
                nCalcItem.FormatDax();
                nCalcItem.Ordinal = j; 
                nCalcItem.Description = itemDescription;
                
            };

        };

        
    };
};
