using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

namespace JdHelper {
 public static class Native {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int n);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
  [DllImport("user32.dll")] static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr extra);
  [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT p);
  [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flags);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint processId);
  static string Owner(IntPtr h){uint pid;GetWindowThreadProcessId(h,out pid);try{return Process.GetProcessById((int)pid).ProcessName+"/"+h.ToInt64();}catch{return "unknown/"+h.ToInt64();}}
  [DllImport("user32.dll",SetLastError=true)] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
  [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int width,int height,uint flags);
  [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr h,int id,uint mod,uint key);
  [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr h,int id);
  [StructLayout(LayoutKind.Sequential)] struct POINT { public int x,y; public POINT(int a,int b){x=a;y=b;} }
  public static void Dpi(){try{if(SetProcessDpiAwarenessContext(new IntPtr(-4))||Marshal.GetLastWin32Error()==5)return;if(SetProcessDpiAwarenessContext(new IntPtr(-3))||Marshal.GetLastWin32Error()==5)return;}catch(EntryPointNotFoundException){}SetProcessDPIAware();}
  public static void Click(IntPtr h,AutomationElement target,AutomationElement root,RunControl c){
   c.Check();
   if(IsIconic(h))ShowWindow(h,9);
   if(GetForegroundWindow()!=h){SetForegroundWindow(h);for(int i=0;i<6&&GetForegroundWindow()!=h;i++)c.Wait(100);c.Wait(100);}
   // Fetch bounds after foreground activation and pause/resume, not from an old snapshot.
   var current=target.Current;var r=current.BoundingRectangle;var window=root.Current.BoundingRectangle;
   if(current.IsOffscreen||r.IsEmpty||r.Width<1||r.Height<1||!window.Contains(new Point(r.Left+r.Width/2,r.Top+r.Height/2)))throw new Exception("目标位置已变化，等待重新定位");
   int x=(int)(r.Left+r.Width/2),y=(int)(r.Top+r.Height/2);
   if(GetForegroundWindow()!=h||GetAncestor(WindowFromPoint(new POINT(x,y)),2)!=h)throw new FatalError("京东窗口被其他窗口遮挡，请将小程序放到前台后继续（目标="+Owner(h)+"，前台="+Owner(GetForegroundWindow())+"，点击位置="+Owner(GetAncestor(WindowFromPoint(new POINT(x,y)),2))+"）");
   c.Check();SetCursorPos(x,y);Thread.Sleep(80);mouse_event(2,0,0,0,UIntPtr.Zero);Thread.Sleep(80);mouse_event(4,0,0,0,UIntPtr.Zero);c.Wait(120);
  }
  public static void ScrollReceipt(IntPtr h,Rect viewport,bool down,RunControl c){
   c.Check();SetForegroundWindow(h);c.Wait(150);int x=(int)(viewport.Left+viewport.Width*.83),y=(int)(viewport.Top+viewport.Height*.55);
   if(GetForegroundWindow()!=h||GetAncestor(WindowFromPoint(new POINT(x,y)),2)!=h)throw new FatalError("京东窗口被遮挡，未执行滚动");
   SetCursorPos(x,y);Thread.Sleep(80);mouse_event(0x0800,0,0,unchecked((uint)(down?-360:360)),UIntPtr.Zero);c.Wait(400);
  }
  public static bool EnlargeToScreen(IntPtr h,Rect current){
   var work=System.Windows.Forms.Screen.FromHandle(h).WorkingArea;
   if(current.Height>=work.Height-4)return false;
   int width=(int)Math.Min(current.Width,work.Width),x=(int)Math.Max(work.Left,Math.Min(current.Left,work.Right-width));
   return SetWindowPos(h,IntPtr.Zero,x,work.Top,width,work.Height,0x0004|0x0010);
  }
 }
 public class UiNode { public AutomationElement Element;public string Text,Value;public ControlType Type;public Rect Box;public int Index;public bool Offscreen; }
 public class UiSnapshot {
  public AutomationElement Root;public IntPtr Handle;public Rect Box;public List<UiNode> Nodes,AllNodes;
  public List<UiNode> Texts{get{return Nodes.Where(n=>n.Type==ControlType.Text&&!String.IsNullOrWhiteSpace(n.Text)).ToList();}}
  public bool Has(string text){return Nodes.Any(n=>text.StartsWith("JD")?n.Text.Trim()==text:UiRules.IsLabel(n.Text,text));}
  public UiNode Find(string text){return Nodes.FirstOrDefault(n=>UiRules.IsLabel(n.Text,text));}
  public bool Missing{get{return Nodes.Any(n=>UiRules.Label(n.Text).Contains("暂无运单信息"));}}
  public bool Receipt{get{return Has("存根信息")&&Nodes.Any(n=>UiRules.FieldLabel(n.Text,"运单号"));}}
 }
 public static class UiRecovery {
  public static bool IsStale(Exception e){return e is ElementNotAvailableException||(e is COMException&&e.HResult==unchecked((int)0x80040201));}
  public static T Read<T>(Func<T> read,RunControl control,Action<string> log){
   for(int attempt=1;attempt<=4;attempt++){
    control.Check();try{return read();}
    catch(Exception e){if(!IsStale(e))throw;if(attempt==4)throw new Exception("页面控件连续刷新，等待后仍未稳定",e);log("页面控件已刷新，重新读取（"+attempt+"/3）");control.Wait(300);}
   }throw new InvalidOperationException("页面读取未完成");
  }
 }
 // A visible control may not yet respond to input. Retry only while the source
 // state still applies, and inspect a fresh snapshot immediately before each click.
 public static class ClickFeedback {
  public static T Run<T>(Func<T> read,Func<T,bool> done,Func<T,bool> canClick,Action<T> click,
   RunControl control,Action<string> log,string label,int timeoutMs,int intervalMs,int maxClicks,
   Func<long> now=null,Action<int> wait=null,int initialDelayMs=0){
   var timer=Stopwatch.StartNew();if(now==null)now=()=>timer.ElapsedMilliseconds;if(wait==null)wait=control.Wait;
   long start=now(),nextClick=start+initialDelayMs;int clicks=0;
   while(now()-start<timeoutMs){
    control.Check();var state=read();if(done(state))return state;
    if(clicks<maxClicks&&now()>=nextClick&&canClick(state)){
     control.Check();state=read();if(done(state))return state;
     if(canClick(state)){
      if(clicks>0)log(label+"上次点击尚未生效，已重新核对页面，重试点击（"+(clicks+1)+"/"+maxClicks+"）");
      click(state);clicks++;nextClick=now()+intervalMs;
     }
    }
    wait(200);
   }
   // Include a response that arrived at the deadline, but never click after it.
   control.Check();var last=read();if(done(last))return last;
   throw new Exception(label+"在 "+(timeoutMs/1000)+" 秒内未生效（已点击 "+clicks+" 次）");
  }
 }
 public class JdLookup:ILookup {
  readonly RunControl control;public Action<string> Log=delegate{};
  static readonly Regex Phone=new Regex(@"^(?:1\d{10}|1\d{2}\*{4}\d{4})$");
  public JdLookup(RunControl c){control=c;}
  public UiSnapshot Snap(){
   return UiRecovery.Read(Capture,control,Log);
  }
  UiSnapshot Capture(){
   control.Check();var candidates=new List<AutomationElement>();
   foreach(AutomationElement el in AutomationElement.RootElement.FindAll(TreeScope.Children,Condition.TrueCondition)){
    try{if(el.Current.Name.Contains("京东快递")&&Process.GetProcessById(el.Current.ProcessId).ProcessName.Equals("WeChatAppEx",StringComparison.OrdinalIgnoreCase))candidates.Add(el);}catch(ElementNotAvailableException){}
   }
   if(candidates.Count!=1)throw new FatalError(candidates.Count==0?"请登录微信并打开京东快递小程序的“查快递”页":"打开了多个京东快递窗口，请只保留一个");
   var root=candidates[0];var h=new IntPtr(root.Current.NativeWindowHandle);
   if(Native.IsIconic(h)){Native.ShowWindow(h,9);control.Wait(350);}
   var list=new List<UiNode>();int i=0;
   var cache=new CacheRequest();cache.TreeScope=TreeScope.Element;
   foreach(var property in new[]{AutomationElement.NameProperty,AutomationElement.ControlTypeProperty,AutomationElement.BoundingRectangleProperty,AutomationElement.IsOffscreenProperty,ValuePattern.ValueProperty})cache.Add(property);
   AutomationElementCollection descendants;using(cache.Activate())descendants=root.FindAll(TreeScope.Descendants,Condition.TrueCondition);
   foreach(AutomationElement el in descendants){
    try{
     var p=el.Cached;int index=i++;var rect=p.BoundingRectangle;if(rect.IsEmpty||rect.Width<1||rect.Height<1)continue;
     object raw=el.GetCachedPropertyValue(ValuePattern.ValueProperty,true);string value=raw as string??"";
     list.Add(new UiNode{Element=el,Text=p.Name??"",Value=value??"",Type=p.ControlType,Box=rect,Index=index,Offscreen=p.IsOffscreen});
    }catch(ElementNotAvailableException){}
   }
   var active=list.LastOrDefault(n=>n.Type==ControlType.Document&&(n.Text=="Page-Frame"||n.Value.StartsWith("https://kd.jd.com/")));
   var visible=active==null?list:list.Where(n=>n.Index>=active.Index).ToList();
   var s=new UiSnapshot{Root=root,Handle=h,Box=root.Current.BoundingRectangle,Nodes=visible,AllNodes=list};
   if(s.Nodes.Any(n=>n.Text.Contains("验证码")||n.Text.Contains("安全验证")||n.Text.Contains("访问过于频繁")))throw new FatalError("小程序要求验证或限制访问，已停止，请人工处理后继续");
   return s;
  }
  UiSnapshot Wait(Func<UiSnapshot,bool> ready,string message,int seconds){
   var sw=Stopwatch.StartNew();while(sw.Elapsed.TotalSeconds<seconds){control.Check();var s=Snap();if(s.Missing)throw new NoWaybill();if(ready(s))return s;control.Wait(200);}throw new Exception(message+"（等待超时）");
  }
  void ClickTarget(Func<UiSnapshot,UiNode> selector,string label,Func<UiSnapshot,bool> alreadyDone=null){
   bool enlarged=false;double lastY=Double.NaN;int unchanged=0;
   for(int attempt=0;attempt<8;attempt++){
    try{
    control.Check();var s=Snap();if(alreadyDone!=null&&alreadyDone(s))return;var n=selector(s);if(n==null){control.Wait(450);continue;}
    if(UiRules.Visible(s,n)){Native.Click(s.Handle,n.Element,s.Root,control);return;}
    if(!Double.IsNaN(lastY)&&Math.Abs(n.Box.Top-lastY)<2)unchanged++;else unchanged=0;lastY=n.Box.Top;
    object pattern;bool scrolled=false;
    if(unchanged<2){
     try{
      if(n.Element.TryGetCurrentPattern(ScrollItemPattern.Pattern,out pattern)){((ScrollItemPattern)pattern).ScrollIntoView();scrolled=true;}
      else{
       var parent=TreeWalker.ControlViewWalker.GetParent(n.Element);
       for(int depth=0;parent!=null&&depth<8;depth++,parent=TreeWalker.ControlViewWalker.GetParent(parent)){
        if(parent.Equals(s.Root))break;
        if(parent.TryGetCurrentPattern(ScrollPattern.Pattern,out pattern)&&((ScrollPattern)pattern).Current.VerticallyScrollable){((ScrollPattern)pattern).Scroll(ScrollAmount.NoAmount,n.Box.Top>s.Box.Top+s.Box.Height/2?ScrollAmount.LargeIncrement:ScrollAmount.LargeDecrement);scrolled=true;break;}
       }
      }
     }catch(InvalidOperationException){}
     if(!scrolled&&s.Receipt){
      var doc=s.Nodes.FirstOrDefault(p=>p.Type==ControlType.Document&&p.Value.StartsWith("https://kd.jd.com/"));var viewport=doc==null?s.Box:Rect.Intersect(doc.Box,s.Box);
      Native.ScrollReceipt(s.Handle,viewport,n.Box.Top>viewport.Top+viewport.Height/2,control);scrolled=true;
     }
    }
    if(scrolled){Log(label+"不在可见区域，滚动后重新定位");control.Wait(400);continue;}
    if(!enlarged&&Native.EnlargeToScreen(s.Handle,s.Box)){enlarged=true;Log("窗口高度不足，已增高至当前屏幕可用高度");control.Wait(600);continue;}
    throw new FatalError(label+"仍不在可见区域，请调整窗口后续跑；未猜测点击位置");
    }catch(Exception e){if(!UiRecovery.IsStale(e))throw;Log(label+"的控件已刷新，等待后重新定位（"+(attempt+1)+"/8）："+Diagnostics.Error(e));control.Wait(400);}
   }
   throw new Exception("等待后仍未能唯一定位"+label+"，未继续点击");
  }
  static UiNode Input(UiSnapshot s){return UiRules.Input(s);}
  static bool SearchPage(UiSnapshot s){return Input(s)!=null&&!s.Has("电子存根")&&!s.Receipt;}
  void Back(UiSnapshot s){
   string before=PageKind(s);
   ClickFeedback.Run(Snap,p=>PageKind(p)!="loading"&&PageKind(p)!=before,p=>PageKind(p)==before,
    p=>ClickTarget(UiRules.Back,"返回按钮",q=>PageKind(q)!=before),control,Log,"返回按钮",8000,1400,4);
  }
  static string PageKind(UiSnapshot s){return s.Receipt?"receipt":SearchPage(s)?"search":s.Missing?"missing":s.Texts.Any(n=>Regex.IsMatch(n.Text.Trim(),@"^JD[A-Z0-9]{8,30}$"))?"detail":"loading";}
  UiSnapshot ToSearch(){
   var timer=Stopwatch.StartNew();int backs=0;
   while(timer.Elapsed.TotalSeconds<25&&backs<6){var s=Snap();if(SearchPage(s))return s;if(PageKind(s)=="loading"){control.Wait(250);continue;}Back(s);backs++;}
   throw new FatalError("未能回到查询页，请手动回到“查快递”页");
  }
  public Receiver Query(string bill){
   try{return QueryCore(bill);}
   catch(UnreadableOrder){
    // The next task starts from the search page, including when this is the last row.
    Log("当前单不可读，返回查询页");
    try{ToSearch();}catch(StopRequested){throw;}catch(Exception e){Log("返回查询页未完成，下一单会再次恢复："+Diagnostics.Error(e));}
    throw;
   }
  }
  Receiver QueryCore(string bill){
   var queryTimer=Stopwatch.StartNew();
   if(!Regex.IsMatch(bill??"",@"^JD[A-Z0-9]{8,30}$"))throw new UnreadableOrder("运单号格式不正确");
   Log("返回查询页并输入运单号");var s=ToSearch();Log("当前京东窗口："+(int)s.Box.Width+" × "+(int)s.Box.Height+" 像素");ClickTarget(Input,"运单输入框");s=Snap();var edit=Input(s);if(edit==null)throw new Exception("运单输入框不可用");
   object vp;if(!edit.Element.TryGetCurrentPattern(ValuePattern.Pattern,out vp)||((ValuePattern)vp).Current.IsReadOnly)throw new FatalError("当前微信版本没有开放输入控件，需适配此版本");
   ((ValuePattern)vp).SetValue(bill);s=Snap();
   if(Input(s)==null||Input(s).Value.Trim()!=bill)throw new Exception("运单号输入后核对失败");
   ClickTarget(UiRules.Query,"查询按钮");control.Wait(250);
   Func<UiSnapshot> checkedRead=()=>{var state=Snap();if(state.Missing)throw new NoWaybill();return state;};
   s=ClickFeedback.Run(checkedRead,p=>p.Receipt||UiRules.SameBill(p,bill),p=>SearchPage(p)&&Input(p).Value.Trim()==bill,
    p=>ClickTarget(UiRules.Query,"查询按钮",q=>q.Receipt||UiRules.SameBill(q,bill)),control,Log,"查询结果加载",15000,2000,3,initialDelayMs:1800);
   if(!s.Has("电子存根")&&!s.Receipt){
    if(UiRules.Cancelled(s,bill))throw new UnreadableOrder("运单已取消，没有可读取的电子存根");
    Log("打开查询结果的运单详情");
    Func<UiSnapshot,bool> detailReady=p=>p.Receipt||(!SearchPage(p)&&UiRules.SameBill(p,bill));
    s=ClickFeedback.Run(checkedRead,detailReady,p=>SearchPage(p)&&UiRules.SameBill(p,bill)&&UiRules.ResultCard(p,bill)!=null,
     p=>ClickTarget(q=>SearchPage(q)?UiRules.ResultCard(q,bill):null,"当前运单卡片",detailReady),control,Log,"打开运单详情",15000,1500,5);
    s=WaitForReceiptEntry(bill);
   }
   if(!s.Receipt){
    Log("打开电子存根");
    s=ClickFeedback.Run(checkedRead,p=>ReceiptReady(p,bill),p=>!p.Receipt&&UiRules.SameBill(p,bill)&&p.Has("电子存根"),
     p=>ClickTarget(q=>!q.Receipt&&UiRules.SameBill(q,bill)?q.Find("电子存根"):null,"电子存根",q=>q.Receipt),control,Log,"电子存根入口",20000,1800,5);
   }
   if(!ReceiptReady(s,bill))s=Wait(p=>ReceiptReady(p,bill),"电子存根字段没有加载完整",12);
   var result=ReadReceipt(s,bill);
   if(result.Phone.Contains("*")){
    s=ClickFeedback.Run(checkedRead,p=>PhoneRevealed(p,bill),p=>ReceiptReady(p,bill)&&ReadReceipt(p,bill).Phone.Contains("*"),
     p=>ClickTarget(q=>{
      if(!ReceiptReady(q,bill))return null;var current=ReadReceipt(q,bill);if(!current.Phone.Contains("*"))return null;
      var phone=q.Texts.LastOrDefault(n=>n.Text.Trim()==current.Phone);return UiRules.Eye(q,phone);
     },"收件电话的小眼睛",q=>PhoneRevealed(q,bill)),control,Log,"收件电话的小眼睛",15000,1800,6);
    result=ReadReceipt(s,bill);
   }
   result.Validate(bill);Log("已核对详情运单号并取得完整收件字段；查询耗时 "+queryTimer.Elapsed.TotalSeconds.ToString("F1")+" 秒");return result;
  }
  static bool ReceiptReady(UiSnapshot s,string bill){try{return s.Receipt&&ReadReceipt(s,bill)!=null;}catch{return false;}}
  UiSnapshot WaitForReceiptEntry(string bill){
   var timer=Stopwatch.StartNew();UiSnapshot s=null;
   do{
    control.Check();s=Snap();if(s.Missing)throw new NoWaybill();
    if(s.Receipt||(UiRules.SameBill(s,bill)&&s.Has("电子存根")))return s;
    if(UiRules.Cancelled(s,bill))throw new UnreadableOrder("运单已取消，没有可读取的电子存根");
    control.Wait(200);
   }while(timer.Elapsed.TotalSeconds<6);
   if(s!=null&&UiRules.SameBill(s,bill)&&Input(s)==null)throw new UnreadableOrder("详情页在 6 秒内未提供电子存根入口；可勾选重查不可读单再试");
   throw new Exception("运单详情在 6 秒内没有加载完成");
  }
  bool PhoneRevealed(UiSnapshot s,string bill){try{return Regex.IsMatch(ReadReceipt(s,bill).Phone,@"^1\d{10}$");}catch{return false;}}
  public static Receiver ReadReceipt(UiSnapshot s,string bill){
   var texts=s.Texts;int start=0,end=texts.FindIndex(n=>new[]{"商家单号","平台单号","运单号"}.Any(label=>UiRules.FieldLabel(n.Text,label)));
   if(!s.Has("存根信息")||end<=start)throw new Exception("电子存根字段结构无法识别");
   var receipt=texts.Skip(start).Take(end-start).ToList();var phones=receipt.Select((n,i)=>new{Node=n,I=i}).Where(x=>Phone.IsMatch(x.Node.Text.Trim())).ToList();
   if(phones.Count!=2)throw new Exception("未能唯一识别寄件和收件电话，未写入");
   int index=phones[1].I;if(index<1||index+1>=receipt.Count)throw new Exception("收件字段不完整");
   var bills=new List<string>();for(int i=0;i<texts.Count;i++){
    if(UiRules.IsLabel(texts[i].Text,"运单号")){if(i+1<texts.Count)bills.Add(texts[i+1].Text.Trim());}
    else{var m=Regex.Match(texts[i].Text,@"^\s*运\s*单\s*号\s*[:：]\s*(JD[A-Z0-9]+)\s*$");if(m.Success)bills.Add(m.Groups[1].Value);}
   }
   if(bills.Count!=1||bills[0]!=bill)throw new Exception("电子存根运单号与当前任务不匹配");
   return new Receiver{Name=receipt[index-1].Text.Trim(),Phone=receipt[index].Text.Trim(),Address=receipt[index+1].Text.Trim(),Waybill=bill};
  }
 }
}
