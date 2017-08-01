import java.io.FileOutputStream

import org.apache.poi.hssf.util.CellReference
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

import scala.collection.mutable.ArrayBuffer

/**
  * Created by CCROWE on 8/1/2017.
  */

object Main {
  val dataPath = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Scoe\\SCoE Facility List as of 13 June 2017.xlsx";
  def main(args: Array[String]): Unit = {
    val getFacilities:GetAddedFacilities = new GetAddedFacilities;
    var data = new FindFacilityData;
    var facilities:ArrayBuffer[Facility] = ArrayBuffer[Facility]();
    var inserter:InsertDataIntoWorkbook = new InsertDataIntoWorkbook;
    for(facilityName <- getFacilities.facilities){
      val facility:Facility = data.Find(facilityName);
      if(facility != null){
        facilities += facility;
        println(facility.Designator);
        inserter.Insert(facility);
      }else{
        println("Null! " + facility.toString);
      }
    }
    inserter.Close();
  }
}

class InsertDataIntoWorkbook{
  val workbook:XSSFWorkbook = new XSSFWorkbook();
  workbook.createSheet("Datas")
  val sheet:XSSFSheet = workbook.getSheet("Datas");
  var row:Int = 0;
  def Insert(facility:Facility): Unit ={
    row += 1;
    WriteToCell(facility.Designator,"B" + row.toString);
    WriteToCell(facility.NewDescription,"C" + row.toString);
    WriteToCell(facility.OriginalDescription,"D" + row.toString);
    WriteToCell(facility.PrimaryConstructionMaterial,"E" + row.toString);
    WriteToCell(facility.DetailField,"F" + row.toString);
    WriteToCell(facility.LookupType,"G" + row.toString);
    WriteToCell(facility.LookupToNoun,"H" + row.toString);
    WriteToCell(facility.LookupToStandard,"I" + row.toString);
    WriteToCell(facility.LookupToMasterPlanningCategory,"J" + row.toString);
    WriteToCell(facility.ProponentRecommendation,"K" + row.toString);
    WriteToCell(facility.ProponentComments,"L" + row.toString);
    WriteToCell(facility.DesignAgentRecommendation,"M" + row.toString);
    WriteToCell(facility.DesignAgentComments,"N" + row.toString);
    WriteToCell(facility.PrimaryProponent,"O" + row.toString);
    WriteToCell(facility.SecondaryProponent,"P" + row.toString);
    WriteToCell(facility.VettingDate,"Q" + row.toString);
  }
  def Close(): Unit ={
    var fileOut = new FileOutputStream("data.xlsx");
    workbook.write(fileOut);
    fileOut.close();
  }
  def WriteToCell(value:String,address:String): Unit ={
    var cellRef: CellReference = new CellReference(address);
    var row = sheet.getRow(cellRef.getRow);
    if(row == null){
      sheet.createRow(cellRef.getRow);
      row = sheet.getRow(cellRef.getRow);
    }
    row.createCell(cellRef.getCol);
    var cell = row.getCell(cellRef.getCol);
    cell.setCellValue(value);
    val cellVal = cell.getStringCellValue;
    println(cellVal);
  }
}

class Facility(val Designator:String,val NewDescription:String,val OriginalDescription:String,val PrimaryConstructionMaterial:String,val DetailField:String,
               val LookupType:String,val LookupToNoun:String,val LookupToStandard:String,val LookupToMasterPlanningCategory:String, val ProponentRecommendation:String,
               val ProponentComments:String,val DesignAgentRecommendation:String,val DesignAgentComments:String,val PrimaryProponent:String,
               val SecondaryProponent:String, val VettingDate:String){
def Print(): Unit ={
  println();
  if(this.Designator != ""){
    println(this.Designator);
  }
  if(this.NewDescription != ""){
    println(this.NewDescription);
  }
  if(this.OriginalDescription != ""){
    println(this.OriginalDescription);
  }
  if(this.PrimaryConstructionMaterial != ""){
    println(this.PrimaryConstructionMaterial);
  }
  if(this.DetailField != ""){
    println(this.DetailField);
  }
  if(this.LookupType != ""){
    println(this.LookupType);
  }
  if(this.LookupToNoun != ""){
    println(this.LookupToNoun);
  }
  if(this.LookupToStandard != ""){
    println(this.LookupToStandard);
  }
  if(this.LookupToMasterPlanningCategory != ""){
    println(this.LookupToMasterPlanningCategory);
  }
  if(this.ProponentRecommendation != ""){
    println(this.ProponentRecommendation);
  }
  if(this.ProponentComments != ""){
    println(this.ProponentComments);
  }
  if(this.DesignAgentComments != ""){
    println(this.DesignAgentComments)
  }
  if(this.DesignAgentRecommendation != ""){
    println(this.DesignAgentRecommendation);
  }
  if(this.DesignAgentComments != ""){
    println(this.DesignAgentComments);
  }
  if(this.PrimaryProponent != ""){
    println(this.PrimaryProponent);
  }
  println();
  println();
}
}

class FindFacilityData{
  val workbook:XSSFWorkbook = new XSSFWorkbook(OPCPackage.open(Main.dataPath));
  def Find(name:String): Facility ={
    var iterator = workbook.sheetIterator();
    while(iterator.hasNext) {
      var sheet:XSSFSheet = workbook.getSheetAt(workbook.getSheetIndex(iterator.next().getSheetName));
      for(i <- 3.to(sheet.getLastRowNum)){
        val facilityName:String = CellStringGetter.GetString(sheet,"B" + i.toString)
        if(facilityName == name){
          val facility:Facility = new Facility(CellStringGetter.GetString(sheet,"B" + i.toString),
            CellStringGetter.GetString(sheet,"C" + i.toString),
            CellStringGetter.GetString(sheet,"D" + i.toString),
            CellStringGetter.GetString(sheet,"E" + i.toString),
            CellStringGetter.GetString(sheet,"F" + i.toString),
            CellStringGetter.GetString(sheet,"G" + i.toString),
            CellStringGetter.GetString(sheet,"H" + i.toString),
            CellStringGetter.GetString(sheet,"I" + i.toString),
            CellStringGetter.GetString(sheet,"J" + i.toString),
            CellStringGetter.GetString(sheet,"K" + i.toString),
            CellStringGetter.GetString(sheet,"L" + i.toString),
            CellStringGetter.GetString(sheet,"M" + i.toString),
            CellStringGetter.GetString(sheet,"N" + i.toString),
            CellStringGetter.GetString(sheet,"O" + i.toString),
            CellStringGetter.GetString(sheet,"P" + i.toString),
            CellStringGetter.GetString(sheet,"Q" + i.toString));
          return facility;
        }
      }
    }
    return null;
  }
}

class GetAddedFacilities{
  val changesPath = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Scoe\\SCoE_Changes.xlsx"
  val workbook:XSSFWorkbook = new XSSFWorkbook(OPCPackage.open(changesPath));
  val sheet:XSSFSheet = workbook.getSheet("Added_to_SCoE");
  var facilities:ArrayBuffer[String] = ArrayBuffer[String]();
  for(i <- 5.to(43)){
    var cellRef:CellReference = new CellReference("A" + i.toString);
    var row = sheet.getRow(cellRef.getRow());
    var cell = row.getCell(cellRef.getCol);
    val formatter = new DataFormatter
    val formattedCellValue = formatter.formatCellValue(cell)
    if(formattedCellValue != null){
      facilities += formattedCellValue;
    }
  }
  workbook.close();
}


object CellStringGetter {
  def GetString(workbook: XSSFSheet, address: String): String = {
    var cellRef: CellReference = new CellReference(address);
    var row = workbook.getRow(cellRef.getRow);
    if(row == null){
      return "";
    }
    var cell = row.getCell(cellRef.getCol);
    if (cell != null) {
      val formatter = new DataFormatter
      val formattedCellValue = formatter.formatCellValue(cell)
      return formattedCellValue;
    } else {
      return "";
    }
  }
}