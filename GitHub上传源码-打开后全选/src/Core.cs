using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace JdHelper {
 public class StopRequested : Exception { public StopRequested(string reason):base(reason){} }
 public class FatalError : Exception { public FatalError(string m):base(m){} }
 public class UnreadableOrder : Exception { public UnreadableOrder(string message):base(message){} }
 public class NoWaybill : UnreadableOrder { public NoWaybill():base("暂无运单信息"){} }
 public class Receiver {
  public string Name, Phone, Address, Waybill;
  public void Validate(string expected) {
   if(Waybill!=expected) throw new Exception("详情页运单号不一致，未写入");
   if(String.IsNullOrWhiteSpace(Name)||Name.Length>60||Name.Contains("*")||!Regex.IsMatch(Phone??"",@"^1\d{10}$")||String.IsNullOrWhiteSpace(Address)||Address.Length<5||Address.Contains("*"))
    throw new Exception("收件字段不完整或仍被隐藏，未写入");
  }
 }
 public class OrderRow { public int Row; public string Waybill,Name,Phone,Address; public bool Unreadable; public bool Complete {get{return !String.IsNullOrWhiteSpace(Name)&&!Name.Contains("*")&&Regex.IsMatch(Phone??"",@"^1\d{10}$")&&!String.IsNullOrWhiteSpace(Address)&&!Address.Contains("*");}} }
 public class Record { public int Row,Attempts; public string Waybill,Status,Error,Time; public Receiver Result; }
 public class Ledger { public int Version=1; public string Workbook,Sheet,LastRunStatus,LastStopReason,LastRunError,LastRunTime; public Dictionary<string,Record> Records=new Dictionary<string,Record>(); }
 public class RunSummary { public int Completed,Skipped,Failed,Attempted,SkippedUnreadable,Unreadable,Pending; public int Remaining=-1; public bool Stopped,AllComplete,AuxFilesCleaned; public string Backup,Error,Workbook,CleanupWarning,StopReason; }
 public class RunControl {
  readonly object gate=new object();volatile bool stopped;string stopReason;
  public volatile bool Paused;
  public bool Stop{get{return stopped;}}
  public string StopReason{get{return stopReason;}}
  public bool RequestStop(string reason){lock(gate){if(stopped)return false;stopReason=String.IsNullOrWhiteSpace(reason)?"未注明来源的停止请求":reason;stopped=true;Paused=false;return true;}}
  public void Check(){if(Stop)throw new StopRequested(StopReason); while(Paused){if(Stop)throw new StopRequested(StopReason);Thread.Sleep(100);} }
  public void Wait(int ms){for(int elapsed=0;elapsed<ms;elapsed+=100){Check();Thread.Sleep(Math.Min(100,ms-elapsed));}}
 }
 public static class Diagnostics {
  public static string Error(Exception e){
   string message=String.IsNullOrWhiteSpace(e.Message)?"操作失败，系统未返回文字说明":e.Message.Trim();
   string result=message+" ["+e.GetType().Name+", 0x"+e.HResult.ToString("X8")+"]";
   if(e.InnerException!=null)result+="；内部错误："+Error(e.InnerException);return result;
  }
  public static string NewRunLog(string folder){Directory.CreateDirectory(folder);string path=System.IO.Path.Combine(folder,"运行_"+DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")+"_"+Guid.NewGuid().ToString("N").Substring(0,6)+".log");File.WriteAllText(path,"京东查单工具 1.4"+Environment.NewLine,new UTF8Encoding(false));return path;}
  public static void AppendLog(string path,string text){File.AppendAllText(path,DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+" "+text+Environment.NewLine,new UTF8Encoding(false));}
 }
 public static class Json {
  public static string Encode(object o){return new JavaScriptSerializer().Serialize(o);}
  public static T Decode<T>(string s){return new JavaScriptSerializer().Deserialize<T>(s);}
  public static void Save(string path,object value){
   var temp=path+".tmp."+Guid.NewGuid().ToString("N");
   File.WriteAllText(temp,Encode(value),new UTF8Encoding(false));
   try {if(File.Exists(path))File.Replace(temp,path,null);else File.Move(temp,path);}finally{if(File.Exists(temp))File.Delete(temp);}
  }
 }
 public class WorkbookStore {
  static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  public readonly string Path,Sheet;
  string entry,expectedHash;
  public WorkbookStore(string path,string sheet){Path=System.IO.Path.GetFullPath(path);Sheet=sheet;if(!Path.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase))throw new FatalError("请选择 .xlsx 文件");LoadRows();}
  public static string Hash(byte[] bytes){using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-","");}
  static XDocument Xml(ZipArchive zip,string name){var e=zip.GetEntry(name);if(e==null)throw new FatalError("表格内部缺少 "+name);using(var s=e.Open())return XDocument.Load(s,LoadOptions.PreserveWhitespace);}
  public List<OrderRow> LoadRows(){
   byte[] bytes=File.ReadAllBytes(Path);expectedHash=Hash(bytes);
   using(var mem=new MemoryStream(bytes))using(var zip=new ZipArchive(mem,ZipArchiveMode.Read)){
    var wb=Xml(zip,"xl/workbook.xml");
    var sheet=wb.Descendants(N+"sheet").SingleOrDefault(s=>(string)s.Attribute("name")==Sheet);
    if(sheet==null)throw new FatalError("没有找到工作表："+Sheet);
    XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    string rid=(string)sheet.Attribute(R+"id");
    var rel=Xml(zip,"xl/_rels/workbook.xml.rels").Root.Elements().Single(s=>(string)s.Attribute("Id")==rid);
    string target=(string)rel.Attribute("Target");
    entry=new Uri(new Uri("http://local/xl/workbook.xml"),target).AbsolutePath.TrimStart('/');
    List<string> ss=new List<string>();
    if(zip.GetEntry("xl/sharedStrings.xml")!=null)ss=Xml(zip,"xl/sharedStrings.xml").Descendants(N+"si").Select(s=>String.Concat(s.Descendants(N+"t").Select(t=>t.Value))).ToList();
    var doc=Xml(zip,entry);
    string statusColumn=FailureMarks.StatusColumn(doc,ss,false);
    if(doc.Descendants(N+"sheetProtection").Any())throw new FatalError("工作表受保护，请先在表格软件中解除保护后再运行");
    var rows=new List<OrderRow>();
    foreach(var r in doc.Descendants(N+"sheetData").Elements(N+"row")){
     int row=(int)r.Attribute("r");if(row<2)continue;
     Func<string,string> val=col=>CellValue(r.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")==col+row),ss);
     string bill=(val("F")??"").Trim();if(bill.Length==0)continue;
     var f=r.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")=="F"+row);
     if(f!=null&&f.Element(N+"f")!=null)throw new FatalError("F"+row+" 是公式，请将运单列转换为文本值后运行");
     if(!bill.StartsWith("JD",StringComparison.OrdinalIgnoreCase))continue;
     rows.Add(new OrderRow{Row=row,Waybill=bill.ToUpperInvariant(),Name=val("B"),Phone=val("C"),Address=val("D"),Unreadable=statusColumn!=null&&FailureMarks.IsMarked(val(statusColumn),bill)});
    }
    return rows;
   }
  }
  internal static string CellValue(XElement c,List<string> ss){
   if(c==null)return "";string type=(string)c.Attribute("t");string value=(string)c.Element(N+"v")??"";
   if(type=="s"){int i;if(!Int32.TryParse(value,out i)||i<0||i>=ss.Count)throw new FatalError("表格共享文本索引损坏");return ss[i];}
   if(type=="inlineStr")return String.Concat(c.Descendants(N+"t").Select(t=>t.Value));return value;
  }
  public void CheckWritable(){using(var s=new FileStream(Path,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){} }
  public string Backup(){
   CheckWritable();var folder=System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Path),"京东查单备份");Directory.CreateDirectory(folder);
   var dest=System.IO.Path.Combine(folder,System.IO.Path.GetFileNameWithoutExtension(Path)+"_"+DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")+"_"+Guid.NewGuid().ToString("N").Substring(0,6)+".xlsx");
   File.Copy(Path,dest,false);return dest;
  }
  static bool MaskMatches(string mask,string actual){
   if(String.IsNullOrWhiteSpace(mask))return true;
   if(mask.Contains("*")) {
    int a=mask.IndexOf('*'),b=mask.LastIndexOf('*');
    return actual.StartsWith(mask.Substring(0,a),StringComparison.Ordinal)&&actual.EndsWith(mask.Substring(b+1),StringComparison.Ordinal);
   }
   return mask.Trim()==actual.Trim();
  }
  public void Write(OrderRow order,Receiver result){
   result.Validate(order.Waybill);
   if(!MaskMatches(order.Name,result.Name)||!MaskMatches(order.Phone,result.Phone))throw new Exception("查询结果与表格已有姓名或电话不匹配，未覆盖");
   if(!String.IsNullOrWhiteSpace(order.Address)&&!order.Address.Contains("*")&&order.Address.Trim()!=result.Address.Trim())throw new Exception("表格已有地址与查询结果不同，未覆盖");
   CheckWritable();var original=File.ReadAllBytes(Path);
   if(Hash(original)!=expectedHash)throw new FatalError("表格被其他程序修改，已停止。请关闭表格后重新开始，以最新内容续跑");
   string temp=Path+".jdtmp."+Guid.NewGuid().ToString("N")+".xlsx";
   try{
    File.WriteAllBytes(temp,original);
    using(var zip=ZipFile.Open(temp,ZipArchiveMode.Update)){
     var doc=Xml(zip,entry);var row=doc.Descendants(N+"sheetData").Elements(N+"row").Single(r=>(int)r.Attribute("r")==order.Row);
     foreach(var pair in new[]{new[]{"B",result.Name},new[]{"C",result.Phone},new[]{"D",result.Address}}){
      string address=pair[0]+order.Row;var cell=row.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")==address);
      if(cell==null){cell=new XElement(N+"c",new XAttribute("r",address));var next=row.Elements(N+"c").FirstOrDefault(c=>Column((string)c.Attribute("r"))>Column(address));if(next!=null)next.AddBeforeSelf(cell);else row.Add(cell);}
      if(cell.Element(N+"f")!=null)throw new FatalError(address+" 含有公式，已停止，避免覆盖");
      cell.Elements(N+"v").Remove();cell.Elements(N+"is").Remove();cell.SetAttributeValue("t","inlineStr");
      cell.Add(new XElement(N+"is",new XElement(N+"t",new XAttribute(XNamespace.Xml+"space","preserve"),pair[1])));
     }
     FailureMarks.Clear(zip,doc,order.Row);
     zip.GetEntry(entry).Delete();var updated=zip.CreateEntry(entry,CompressionLevel.Optimal);using(var stream=updated.Open())doc.Save(stream,SaveOptions.DisableFormatting);
    }
    // Confirm that only the three intended values changed before replacing the original.
    var check=new WorkbookStore(temp,Sheet);var written=check.LoadRows().Single(r=>r.Row==order.Row);
    if(written.Waybill!=order.Waybill||written.Name!=result.Name||written.Phone!=result.Phone||written.Address!=result.Address)throw new FatalError("写入验证失败，原文件未替换");
    CheckWritable();if(Hash(File.ReadAllBytes(Path))!=expectedHash)throw new FatalError("保存前检测到外部修改，已停止");
    File.Replace(temp,Path,null);expectedHash=Hash(File.ReadAllBytes(Path));
   }finally{if(File.Exists(temp))File.Delete(temp);}
  }
  public void MarkUnreadable(OrderRow order,string reason){
   CheckWritable();var original=File.ReadAllBytes(Path);
   if(Hash(original)!=expectedHash)throw new FatalError("表格被其他程序修改，未写入不可读标记，请关闭表格后重新开始");
   string temp=Path+".jdtmp."+Guid.NewGuid().ToString("N")+".xlsx";
   try{
    File.WriteAllBytes(temp,original);
    using(var zip=ZipFile.Open(temp,ZipArchiveMode.Update)){
     var doc=Xml(zip,entry);FailureMarks.Mark(zip,doc,order.Row,order.Waybill,reason);
     zip.GetEntry(entry).Delete();using(var stream=zip.CreateEntry(entry,CompressionLevel.Optimal).Open())doc.Save(stream,SaveOptions.DisableFormatting);
    }
    var check=new WorkbookStore(temp,Sheet).LoadRows().Single(r=>r.Row==order.Row);
    if(check.Waybill!=order.Waybill||!check.Unreadable||check.Name!=order.Name||check.Phone!=order.Phone||check.Address!=order.Address)throw new FatalError("不可读标记校验失败，原文件未替换");
    CheckWritable();if(Hash(File.ReadAllBytes(Path))!=expectedHash)throw new FatalError("保存标记前检测到外部修改，已停止");
    File.Replace(temp,Path,null);expectedHash=Hash(File.ReadAllBytes(Path));
   }finally{if(File.Exists(temp))File.Delete(temp);}
  }
  static int Column(string address){int n=0;foreach(char c in address){if(c<'A'||c>'Z')break;n=n*26+c-'A'+1;}return n;}
 }
 public interface ILookup { Receiver Query(string waybill); }
 public class BatchRunner {
  public Action<string> Log=delegate{};
  public RunSummary Run(string path,string sheet,int limit,int retries,ILookup lookup,RunControl control,bool retryUnreadable=false){
   var summary=new RunSummary{Workbook=path};
   string statePath=path+".京东查单记录.json";
   string lockPath=path+".京东查单运行锁";bool removeLock=false;
   using(var mutexFile=new FileStream(lockPath,FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None)){
    var book=new WorkbookStore(path,sheet);var rows=book.LoadRows();book.CheckWritable();
    var ledger=File.Exists(statePath)?Json.Decode<Ledger>(File.ReadAllText(statePath,Encoding.UTF8)):new Ledger{Workbook=System.IO.Path.GetFullPath(path),Sheet=sheet};
    if(ledger==null||ledger.Version!=1||ledger.Records==null||ledger.Sheet!=sheet)throw new FatalError("续跑记录格式或工作表不匹配，请保留记录并联系维护人员");
    summary.Backup=book.Backup();Log("已备份："+summary.Backup);Log("读取到 "+rows.Count+" 条运单；本次最多 "+(limit==0?"全部":limit.ToString())+" 条待补全记录");
    ledger.LastRunStatus="运行中";ledger.LastStopReason=null;ledger.LastRunError=null;ledger.LastRunTime=DateTime.Now.ToString("s");
    Json.Save(statePath,ledger);summary.Remaining=rows.Count(r=>!r.Complete);
    try{
     foreach(var order in rows){
      control.Check();if(order.Complete){summary.Skipped++;continue;}
      if(order.Unreadable&&!retryUnreadable){summary.SkippedUnreadable++;continue;}
      if(limit>0&&summary.Attempted>=limit)break;summary.Attempted++;
      string key=sheet+"|"+order.Row+"|"+order.Waybill;Record record;
      if(!ledger.Records.TryGetValue(key,out record))record=new Record{Row=order.Row,Waybill=order.Waybill};ledger.Records[key]=record;
      string label="第 "+order.Row+" 行 / 尾号 "+order.Waybill.Substring(Math.Max(0,order.Waybill.Length-6));
      bool success=false;
      for(int attempt=0;attempt<=retries;attempt++){
       control.Check();record.Attempts++;record.Time=DateTime.Now.ToString("s");
       try{
        Receiver result=!order.Unreadable&&record.Status=="待写入"&&record.Result!=null?record.Result:null;
        if(result==null){Log(label+"：查询（"+(attempt+1)+"/"+(retries+1)+"）");record.Status="查询中";Json.Save(statePath,ledger);result=lookup.Query(order.Waybill);result.Validate(order.Waybill);record.Result=result;record.Status="待写入";Json.Save(statePath,ledger);}
        control.Check();book.Write(order,result);record.Status="已完成";record.Error=null;record.Time=DateTime.Now.ToString("s");Json.Save(statePath,ledger);summary.Completed++;summary.Remaining--;success=true;Log(label+"：已核对并写入 B、C、D 列");break;
       }catch(StopRequested){throw;}
       catch(FatalError){throw;}
       catch(IOException e){throw new FatalError("保存文件失败，已停止；请关闭 Excel/WPS 后续跑。"+e.Message);}
       catch(UnauthorizedAccessException e){throw new FatalError("没有文件写入权限："+e.Message);}
       catch(Exception e){record.Status="失败";record.Error=Diagnostics.Error(e);record.Result=null;record.Time=DateTime.Now.ToString("s");Json.Save(statePath,ledger);Log(label+"："+record.Error);if(e is UnreadableOrder)break;if(attempt<retries)control.Wait(800);}
      }
      if(!success){control.Check();book.MarkUnreadable(order,record.Error);record.Status="不可读";Json.Save(statePath,ledger);summary.Failed++;Log(label+"：已标红并写入不可读原因，下次默认跳过；继续下一单");}
      control.Wait(250);
     }
     // Re-read the saved workbook: reaching a batch limit is not whole-sheet completion.
     control.Check();var finalRows=book.LoadRows();summary.Remaining=finalRows.Count(r=>!r.Complete);
     summary.Unreadable=finalRows.Count(r=>!r.Complete&&r.Unreadable);summary.Pending=finalRows.Count(r=>!r.Complete&&!r.Unreadable);
     summary.AllComplete=summary.Failed==0&&finalRows.Count>0&&summary.Remaining==0;
     SaveRunState(statePath,ledger,summary.AllComplete?"全部完成":"本批结束",null,null);
     if(summary.AllComplete){
      control.Check();
      try{
       // Remove the checkpoint while still holding the workbook lock, so a new run's
       // checkpoint cannot be accidentally deleted by this run.
       File.Delete(statePath);removeLock=true;
      }catch(IOException e){summary.CleanupWarning=e.Message;}
      catch(UnauthorizedAccessException e){summary.CleanupWarning=e.Message;}
     }
    }catch(StopRequested e){summary.Stopped=true;summary.StopReason=e.Message;SaveRunState(statePath,ledger,"已停止",summary.StopReason,null);Log("已停止，来源："+summary.StopReason+"；已完成记录已保存，下次从未完成记录继续");}
    catch(Exception e){summary.Error=Diagnostics.Error(e);SaveRunState(statePath,ledger,"异常停止",null,summary.Error);Log("已停止："+summary.Error);}
   }
   if(removeLock){
    // Reopen with delete sharing but no read/write sharing. This permits our delete
    // while excluding new runners. If another run used the gap to save a checkpoint,
    // preserve its lock as well. Never remove another run's checkpoint.
    try{
     using(var cleanupGuard=new FileStream(lockPath,FileMode.Open,FileAccess.ReadWrite,FileShare.Delete)){
      if(File.Exists(statePath))summary.CleanupWarning="检测到新的续跑记录，已保留运行锁";
      else{File.Delete(lockPath);summary.AuxFilesCleaned=true;}
     }
    }
    catch(FileNotFoundException){if(!File.Exists(statePath))summary.AuxFilesCleaned=true;else summary.CleanupWarning="检测到新的续跑记录，已保留";}
    catch(IOException e){summary.CleanupWarning=e.Message;}
    catch(UnauthorizedAccessException e){summary.CleanupWarning=e.Message;}
   }
   if(summary.AuxFilesCleaned)Log("整表运单均已填完整，已自动删除续跑记录和运行锁；原表及备份保留");
   else if(summary.CleanupWarning!=null)Log("表格已完成，但辅助文件清理未完成；未删除的文件已保留："+summary.CleanupWarning);
   else Log("续跑记录和运行锁已保留"+(summary.Remaining>0?"，还有 "+summary.Remaining+" 条运单未填完整":""));
   Log("本次完成 "+summary.Completed+" 条，新标记不可读 "+summary.Failed+" 条，跳过完整记录 "+summary.Skipped+" 条，跳过已标记不可读 "+summary.SkippedUnreadable+" 条");
   if(!summary.Stopped&&summary.Error==null)Log("待补全且未标记 "+summary.Pending+" 条；不可读 "+summary.Unreadable+" 条（仅勾选“重查不可读单”才重查）");return summary;
  }
  void SaveRunState(string path,Ledger ledger,string status,string reason,string error){
   ledger.LastRunStatus=status;ledger.LastStopReason=reason;ledger.LastRunError=error;ledger.LastRunTime=DateTime.Now.ToString("s");
   try{Json.Save(path,ledger);}catch(IOException e){Log("停止/结束状态未能写入续跑记录："+Diagnostics.Error(e));}
   catch(UnauthorizedAccessException e){Log("停止/结束状态未能写入续跑记录："+Diagnostics.Error(e));}
  }
 }
}
