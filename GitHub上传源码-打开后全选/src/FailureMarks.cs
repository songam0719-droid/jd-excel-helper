using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace JdHelper {
 // The status travels with the workbook. Red formatting alone never means "skip".
 // Including the waybill prevents a corrected/replaced number inheriting an old skip.
 public static class FailureMarks {
  static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const string Header="京东查单状态",Prefix="不可读｜";
  public static bool IsMarked(string value,string bill){return (value??"").StartsWith(Prefix+bill.ToUpperInvariant()+"｜",StringComparison.Ordinal);}
  static XDocument Read(ZipArchive zip,string name){using(var stream=zip.GetEntry(name).Open())return XDocument.Load(stream,LoadOptions.PreserveWhitespace);}
  static void Save(ZipArchive zip,string name,XDocument doc){var old=zip.GetEntry(name);if(old!=null)old.Delete();using(var stream=zip.CreateEntry(name,CompressionLevel.Optimal).Open())doc.Save(stream,SaveOptions.DisableFormatting);}
  static List<string> Strings(ZipArchive zip){return zip.GetEntry("xl/sharedStrings.xml")==null?new List<string>():Read(zip,"xl/sharedStrings.xml").Descendants(N+"si").Select(s=>String.Concat(s.Descendants(N+"t").Select(t=>t.Value))).ToList();}
  static int Number(string address){int result=0;foreach(char c in address){if(c<'A'||c>'Z')break;result=result*26+c-'A'+1;}return result;}
  static string Column(int number){string result="";while(number>0){number--;result=(char)('A'+number%26)+result;number/=26;}return result;}
  public static string StatusColumn(XDocument doc,List<string> strings,bool create){
   var data=doc.Root.Element(N+"sheetData");var header=data.Elements(N+"row").FirstOrDefault(r=>(int)r.Attribute("r")==1);
   var matches=header==null?new List<XElement>():header.Elements(N+"c").Where(c=>WorkbookStore.CellValue(c,strings)==Header).ToList();
   if(matches.Count>1)throw new FatalError("表格存在多个“京东查单状态”列，未猜测写入位置");
   if(matches.Count==1){if(matches[0].Element(N+"f")!=null)throw new FatalError("京东查单状态表头不能是公式");return Column(Number((string)matches[0].Attribute("r")));}
   if(!create)return null;
   int last=6;
   foreach(var cell in data.Descendants(N+"c"))if(cell.Element(N+"v")!=null||cell.Element(N+"is")!=null||cell.Element(N+"f")!=null)last=Math.Max(last,Number((string)cell.Attribute("r")));
   foreach(var merge in doc.Descendants(N+"mergeCell"))last=Math.Max(last,Number(((string)merge.Attribute("ref")).Split(':').Last()));
   if(last>=16384)throw new FatalError("没有空列可写入京东查单状态");
   string col=Column(last+1);if(header==null){header=new XElement(N+"row",new XAttribute("r",1));data.AddFirst(header);}Set(header,col+"1",Header);
   var cols=doc.Root.Element(N+"cols");if(cols==null){cols=new XElement(N+"cols");data.AddBeforeSelf(cols);}
   // Split an existing empty-column width range if the appended column lies in it.
   foreach(var item in cols.Elements(N+"col").ToList()){
    int min=(int)item.Attribute("min"),max=(int)item.Attribute("max");if(min>last+1||max<last+1)continue;
    if(min<last+1){var before=new XElement(item);before.SetAttributeValue("max",last);item.AddBeforeSelf(before);}
    if(max>last+1){var after=new XElement(item);after.SetAttributeValue("min",last+2);item.AddAfterSelf(after);}item.Remove();
   }
   cols.Add(new XElement(N+"col",new XAttribute("min",last+1),new XAttribute("max",last+1),new XAttribute("width",52),new XAttribute("customWidth",1)));
   cols.ReplaceNodes(cols.Elements().OrderBy(c=>(int)c.Attribute("min")).ToList());
   var dimension=doc.Root.Element(N+"dimension");if(dimension!=null){int bottom=data.Elements(N+"row").Max(r=>(int)r.Attribute("r"));dimension.SetAttributeValue("ref","A1:"+col+bottom);}
   return col;
  }
  static void Set(XElement row,string address,string value){
   var cell=row.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")==address);
   if(cell==null){cell=new XElement(N+"c",new XAttribute("r",address));var next=row.Elements(N+"c").FirstOrDefault(c=>Number((string)c.Attribute("r"))>Number(address));if(next!=null)next.AddBeforeSelf(cell);else row.Add(cell);}
   if(cell.Element(N+"f")!=null)throw new FatalError(address+" 含公式，未覆盖状态");
   cell.Elements(N+"v").Remove();cell.Elements(N+"is").Remove();cell.SetAttributeValue("t","inlineStr");cell.Add(new XElement(N+"is",new XElement(N+"t",value)));
  }
  public static void Clear(ZipArchive zip,XDocument doc,int rowNumber){
   var strings=Strings(zip);string col=StatusColumn(doc,strings,false);if(col==null)return;
   var row=doc.Root.Element(N+"sheetData").Elements(N+"row").Single(r=>(int)r.Attribute("r")==rowNumber);
   var cell=row.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")==col+rowNumber);
   if(WorkbookStore.CellValue(cell,strings).StartsWith(Prefix,StringComparison.Ordinal))Set(row,col+rowNumber,"");
  }
  public static void Mark(ZipArchive zip,XDocument doc,int rowNumber,string bill,string reason){
   var strings=Strings(zip);string col=StatusColumn(doc,strings,true);
   var row=doc.Root.Element(N+"sheetData").Elements(N+"row").Single(r=>(int)r.Attribute("r")==rowNumber);
   var cell=row.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")==col+rowNumber);string existing=WorkbookStore.CellValue(cell,strings);
   if(existing.Length>0&&!existing.StartsWith(Prefix,StringComparison.Ordinal))throw new FatalError("京东查单状态列已有其他内容，未覆盖");
   reason=(reason??"读取失败").Replace('\r',' ').Replace('\n',' ');if(reason.Length>220)reason=reason.Substring(0,220);
   Set(row,col+rowNumber,Prefix+bill+"｜"+reason);
   string formula="LEFT($"+col+"2,LEN($F2)+5)=\""+Prefix+"\"&$F2&\"｜\"";
   int lastRow=doc.Root.Element(N+"sheetData").Elements(N+"row").Max(r=>(int)r.Attribute("r"));
   string range="F2:F"+lastRow+" "+col+"2:"+col+lastRow;
   var existingRule=doc.Descendants(N+"cfRule").FirstOrDefault(r=>r.Elements(N+"formula").Any(f=>f.Value==formula));
   if(existingRule!=null){existingRule.Parent.SetAttributeValue("sqref",range);return;}
   int dxf=RedStyle(zip);
   foreach(var rule in doc.Descendants(N+"cfRule")){int priority=(int?)rule.Attribute("priority")??1;rule.SetAttributeValue("priority",priority+1);}
   var cf=new XElement(N+"conditionalFormatting",new XAttribute("sqref",range),new XElement(N+"cfRule",new XAttribute("type","expression"),new XAttribute("dxfId",dxf),new XAttribute("priority",1),new XAttribute("stopIfTrue",1),new XElement(N+"formula",formula)));
   var before=doc.Root.Elements().FirstOrDefault(e=>new[]{"dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags","drawing","legacyDrawing","legacyDrawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"}.Contains(e.Name.LocalName));
   if(before!=null)before.AddBeforeSelf(cf);else doc.Root.Add(cf);
  }
  static int RedStyle(ZipArchive zip){
   const string name="xl/styles.xml";XDocument doc;
   if(zip.GetEntry(name)!=null)doc=Read(zip,name);
   else{
    doc=XDocument.Parse("<styleSheet xmlns='"+N+"'><fonts count='1'><font><sz val='11'/><name val='Calibri'/></font></fonts><fills count='2'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill></fills><borders count='1'><border/></borders><cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs><cellXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/></cellXfs></styleSheet>");
    var rel=Read(zip,"xl/_rels/workbook.xml.rels");XNamespace rn=rel.Root.Name.Namespace;rel.Root.Add(new XElement(rn+"Relationship",new XAttribute("Id","jdStyles"+Guid.NewGuid().ToString("N")),new XAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),new XAttribute("Target","styles.xml")));Save(zip,"xl/_rels/workbook.xml.rels",rel);
    var types=Read(zip,"[Content_Types].xml");XNamespace tn=types.Root.Name.Namespace;types.Root.Add(new XElement(tn+"Override",new XAttribute("PartName","/xl/styles.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));Save(zip,"[Content_Types].xml",types);
   }
   var dxfs=doc.Root.Element(N+"dxfs");if(dxfs==null){dxfs=new XElement(N+"dxfs");var next=doc.Root.Elements().FirstOrDefault(e=>new[]{"tableStyles","colors","extLst"}.Contains(e.Name.LocalName));if(next!=null)next.AddBeforeSelf(dxfs);else doc.Root.Add(dxfs);}
   int index=dxfs.Elements(N+"dxf").Count();dxfs.Add(new XElement(N+"dxf",new XElement(N+"font",new XElement(N+"color",new XAttribute("rgb","FF9C0006"))),new XElement(N+"fill",new XElement(N+"patternFill",new XAttribute("patternType","solid"),new XElement(N+"fgColor",new XAttribute("rgb","FFFFC7CE")),new XElement(N+"bgColor",new XAttribute("indexed",64))))));dxfs.SetAttributeValue("count",index+1);Save(zip,name,doc);return index;
  }
 }
}
