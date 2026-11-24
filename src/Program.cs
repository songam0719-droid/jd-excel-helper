using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace JdHelper {
 static class Program {
  [STAThread] static int Main(string[] args){
   Native.Dpi();Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
   if(args.Length==0){using(var singleton=new Mutex(false,"Local\\JdExcelHelper_v1")){if(!singleton.WaitOne(0)){MessageBox.Show("工具已经在运行，请使用已打开的窗口。");return 2;}Application.Run(new MainForm());return 0;}}
   string report=Arg(args,"--report",Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"运行结果.json"));
   Action<string> log=s=>{File.AppendAllText(report+".log",DateTime.Now.ToString("HH:mm:ss")+" "+s+Environment.NewLine,Encoding.UTF8);};
   try{
    var control=new RunControl();var lookup=new JdLookup(control){Log=log};
    if(args[0]=="--self-test"){Json.Save(report,SelfTests.Run(Path.GetDirectoryName(Path.GetFullPath(report))));return 0;}
    if(args[0]=="--inspect"){var b=new WorkbookStore(args[1],Arg(args,"--sheet","Sheet1"));Json.Save(report,b.LoadRows().Select(r=>new{r.Row,r.Waybill,r.Complete}).ToArray());return 0;}
    if(args[0]=="--query"){Json.Save(report,lookup.Query(args[1]));return 0;}
    if(args[0]=="--run"){
     int limit=Int32.Parse(Arg(args,"--limit","3")),retry=Int32.Parse(Arg(args,"--retry","2"));
     if(limit<0||retry<0||retry>5)throw new Exception("数量或重试参数不合法");
     var result=new BatchRunner{Log=log}.Run(args[1],Arg(args,"--sheet","Sheet1"),limit,retry,lookup,control,Array.IndexOf(args,"--retry-unreadable")>=0);Json.Save(report,result);return result.Error==null?0:1;
    }
    throw new Exception("未知命令参数");
   }catch(Exception e){Json.Save(report,new{Error=Diagnostics.Error(e),Exception=e.GetType().Name});log("停止："+Diagnostics.Error(e));return 1;}
  }
  static string Arg(string[] a,string key,string fallback){int i=Array.IndexOf(a,key);return i>=0&&i+1<a.Length?a[i+1]:fallback;}
 }
 public static class Hotkeys {
  public const uint PauseModifiers=0x4000,StopModifiers=0x4003;
  public static bool IsPause(int id,long data){return id==8&&((data>>16)&0xFFFF)==0x77&&(data&0xFFFF)==0;}
  public static bool IsStop(int id,long data){return id==9&&((data>>16)&0xFFFF)==0x78&&(data&0xFFFF)==3;}
 }
 public class MainForm:Form {
  readonly TextBox file=new TextBox(),sheet=new TextBox(),log=new TextBox();
  readonly NumericUpDown limit=new NumericUpDown(),retry=new NumericUpDown();
  readonly CheckBox retryUnreadable=new CheckBox();
  readonly Button start=new Button(),pause=new Button(),stop=new Button(),browse=new Button();
  readonly Button openLog=new Button();readonly Label status=new Label();RunControl control;bool running,logFailed;string runLogPath;
  public MainForm(){
   Text="京东查单回填工具 1.4";ClientSize=new Size(1020,550);MinimumSize=new Size(1020,550);Font=new Font("Microsoft YaHei UI",10);StartPosition=FormStartPosition.CenterScreen;
   var layout=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(16),ColumnCount=1,RowCount=6};
   layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));Controls.Add(layout);
   layout.Controls.Add(new Label{Text="微信中打开京东“查快递”页，关闭待处理的 Excel/WPS；每单保存，启动先备份。\nF 列读取运单，B/C/D 列回填；不可读单标红并记录原因，再次开始默认跳过。",AutoSize=true},0,0);
   var row1=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2};row1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));row1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,110));file.Dock=DockStyle.Fill;browse.Text="选择表格";browse.Dock=DockStyle.Fill;row1.Controls.Add(file);row1.Controls.Add(browse);layout.Controls.Add(row1,0,1);
   var settings=new FlowLayoutPanel{Dock=DockStyle.Fill};settings.Controls.Add(new Label{Text="工作表",AutoSize=true,Margin=new Padding(3,7,3,0)});sheet.Text="Sheet1";sheet.Width=90;settings.Controls.Add(sheet);settings.Controls.Add(new Label{Text="本次数量（0=全部）",AutoSize=true,Margin=new Padding(14,7,3,0)});limit.Maximum=100000;limit.Value=3;limit.Width=75;settings.Controls.Add(limit);settings.Controls.Add(new Label{Text="每单额外重试",AutoSize=true,Margin=new Padding(14,7,3,0)});retry.Maximum=5;retry.Value=1;retry.Width=50;settings.Controls.Add(retry);retryUnreadable.Text="重查不可读单";retryUnreadable.AutoSize=true;retryUnreadable.Margin=new Padding(16,5,0,0);settings.Controls.Add(retryUnreadable);layout.Controls.Add(settings,0,2);
   var tips=new ToolTip();tips.SetToolTip(retry,"仅作用于当前单的临时失败。1 表示首次查询加 1 次重试。异常单数量没有上限；取消或无信息直接标记。");tips.SetToolTip(retryUnreadable,"默认不勾选：跳过已标记不可读的单。勾选后重新查询；成功会清除标记和红色提示。");
   var buttons=new FlowLayoutPanel{Dock=DockStyle.Fill};start.Text="开始 / 断点续跑";start.Width=165;pause.Text="暂停（F8）";stop.Text="停止（Ctrl+Alt+F9）";foreach(var b in new[]{start,pause,stop}){b.Height=36;if(b!=start)b.Width=b==stop?200:135;buttons.Controls.Add(b);}openLog.Text="查看日志";openLog.Width=110;openLog.Height=36;openLog.Enabled=false;buttons.Controls.Add(openLog);pause.Enabled=false;stop.Enabled=false;layout.Controls.Add(buttons,0,3);
   status.Text="准备就绪。运行期间请让京东窗口保持可见，并暂时不要操作鼠标。";status.AutoSize=true;layout.Controls.Add(status,0,4);
   log.Multiline=true;log.ReadOnly=true;log.ScrollBars=ScrollBars.Vertical;log.Dock=DockStyle.Fill;layout.Controls.Add(log,0,5);
   browse.Click+=delegate{using(var d=new OpenFileDialog{Filter="Excel 工作簿|*.xlsx",CheckFileExists=true})if(d.ShowDialog(this)==DialogResult.OK)file.Text=d.FileName;};
   start.Click+=delegate{StartRun();};pause.Click+=delegate{PauseRun();};stop.Click+=delegate{StopRun("点击停止按钮");};
   openLog.Click+=delegate{if(!String.IsNullOrEmpty(runLogPath)&&File.Exists(runLogPath))Process.Start(new ProcessStartInfo(runLogPath){UseShellExecute=true});};
   FormClosing+=delegate(object sender,FormClosingEventArgs e){if(running){StopRun("关闭工具窗口（"+e.CloseReason+"）");e.Cancel=true;Append("正在停止，保存完成后可关闭窗口。");}};
  }
  protected override void OnHandleCreated(EventArgs e){base.OnHandleCreated(e);bool a=Native.RegisterHotKey(Handle,8,Hotkeys.PauseModifiers,0x77),b=Native.RegisterHotKey(Handle,9,Hotkeys.StopModifiers,0x78);if(!a||!b)Append("F8 或 Ctrl+Alt+F9 可能被其他程序占用，可用窗口中的暂停/停止按钮。");}
  protected override void OnHandleDestroyed(EventArgs e){Native.UnregisterHotKey(Handle,8);Native.UnregisterHotKey(Handle,9);base.OnHandleDestroyed(e);}
  protected override void WndProc(ref Message m){if(m.Msg==0x0312){if(Hotkeys.IsPause(m.WParam.ToInt32(),m.LParam.ToInt64()))PauseRun();if(Hotkeys.IsStop(m.WParam.ToInt32(),m.LParam.ToInt64()))StopRun("快捷键 Ctrl+Alt+F9");}base.WndProc(ref m);}
  void Append(string s){if(InvokeRequired){BeginInvoke(new Action<string>(Append),s);return;}log.AppendText(DateTime.Now.ToString("HH:mm:ss")+"  "+s+Environment.NewLine);if(!String.IsNullOrEmpty(runLogPath)&&!logFailed){try{Diagnostics.AppendLog(runLogPath,s);}catch(Exception e){logFailed=true;log.AppendText("日志文件写入失败："+Diagnostics.Error(e)+Environment.NewLine);}}}
  void PauseRun(){if(!running)return;control.Paused=!control.Paused;pause.Text=control.Paused?"继续（F8）":"暂停（F8）";status.Text=control.Paused?"已请求暂停；再次按 F8 继续。":"正在运行…";Append(control.Paused?"收到暂停请求":"收到继续请求");}
  void StopRun(string source){if(!running)return;if(control.RequestStop(source))Append("收到停止请求，来源："+source);status.Text="正在停止，已完成的数据会保留…";}
  void StartRun(){
   if(!File.Exists(file.Text)){MessageBox.Show(this,"请先选择有效的 .xlsx 表格。");return;}
   string path=file.Text,tab=sheet.Text.Trim();int count=(int)limit.Value,tries=(int)retry.Value;bool recheck=retryUnreadable.Checked;
   runLogPath=null;logFailed=false;openLog.Enabled=false;
   try{runLogPath=Diagnostics.NewRunLog(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"JdExcelHelper","Logs"));openLog.Enabled=true;Append("本次日志："+runLogPath);}
   catch(Exception e){Append("未能创建日志文件："+Diagnostics.Error(e));}
   running=true;control=new RunControl();start.Enabled=browse.Enabled=file.Enabled=sheet.Enabled=limit.Enabled=retry.Enabled=retryUnreadable.Enabled=false;pause.Enabled=stop.Enabled=true;status.Text="正在运行…";WindowState=FormWindowState.Minimized;
   var thread=new Thread(delegate(){
    RunSummary result;
    try{result=new BatchRunner{Log=Append}.Run(path,tab,count,tries,new JdLookup(control){Log=Append},control,recheck);}
    catch(Exception e){result=new RunSummary{Error=Diagnostics.Error(e)};Append("停止："+result.Error);}
    BeginInvoke(new Action(delegate{
     running=false;start.Enabled=browse.Enabled=file.Enabled=sheet.Enabled=limit.Enabled=retry.Enabled=retryUnreadable.Enabled=true;pause.Enabled=stop.Enabled=false;pause.Text="暂停（F8）";WindowState=FormWindowState.Normal;
     status.Text=result.Error!=null?"已停止："+result.Error:
      result.Stopped?"已停止："+result.StopReason+"；本次完成 "+result.Completed+" 条，续跑文件已保留。":
      result.AuxFilesCleaned?"整表已完成，两个辅助文件已自动删除；可以打开原表检查。":
      result.CleanupWarning!=null?"整表已完成；辅助文件清理未完成，请查看日志。":
      "本次完成 "+result.Completed+" 条；待补全 "+result.Pending+" 条，不可读 "+result.Unreadable+" 条（下次默认跳过）；续跑文件已保留。";
    }));
   });thread.IsBackground=true;thread.SetApartmentState(ApartmentState.MTA);thread.Start();
  }
 }
}
