using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Automation;

namespace JdHelper {
 public static class LayoutTests {
  const string Bill="JDVA00000000001";
  static void Assert(bool ok,string message){if(!ok)throw new Exception("LAYOUT TEST FAILED: "+message);}
  static UiNode N(string text,ControlType type,double x,double y,double w,double h){return new UiNode{Text=text,Value="",Type=type,Box=new Rect(x,y,w,h)};}
  static UiSnapshot S(params UiNode[] nodes){for(int i=0;i<nodes.Length;i++)nodes[i].Index=i;return new UiSnapshot{Box=new Rect(0,0,620,1170),Nodes=nodes.ToList(),AllNodes=nodes.ToList()};}
  static UiSnapshot Scale(UiSnapshot s,double scale,double x,double y){
   Func<Rect,Rect> t=r=>new Rect(r.X*scale+x,r.Y*scale+y,r.Width*scale,r.Height*scale);
   var nodes=s.Nodes.Select(n=>new UiNode{Text=n.Text,Value=n.Value,Type=n.Type,Index=n.Index,Box=t(n.Box),Offscreen=n.Offscreen}).ToArray();var copy=S(nodes);copy.Box=t(s.Box);return copy;
  }
  static UiSnapshot Card(){return S(
   N("",ControlType.Edit,84,40,288,53),N("",ControlType.Edit,84,40,288,53),
   N("查\n询",ControlType.Text,520,54,50,26),N("查询",ControlType.Text,388,686,42,25),
   N(Bill,ControlType.Text,171,278,155,22),N("新增标签",ControlType.Text,40,310,65,20),
   N("发件城市",ControlType.Text,105,340,82,31),N("寄件人",ControlType.Text,114,378,64,25),
   N("运输中",ControlType.Text,294,348,64,25),N("收件城市",ControlType.Text,434,340,82,31),
   N("任意长度姓名",ControlType.Text,434,378,130,25),N("备注",ControlType.Text,202,580,42,25));}
  static UiSnapshot Receipt(){return S(
   N("存根\n信息",ControlType.ListItem,60,110,110,70),
   N("寄件人",ControlType.Text,100,300,80,31),N("189****3365",ControlType.Text,205,300,145,31),
   N("",ControlType.Group,366,297,37,37),N("",ControlType.Group,412,297,37,37),
   N("测试寄件地址",ControlType.Text,100,342,460,55),
   N("收件测试人",ControlType.Text,100,546,100,31),N("136****8300",ControlType.Text,205,546,145,31),
   N("",ControlType.Group,366,543,37,37),N("",ControlType.Group,412,543,37,37),
   N("测试省测试市测试区完整地址",ControlType.Text,100,588,460,55),
   N("商家单号 ：",ControlType.Text,60,700,110,25),N("example",ControlType.Text,180,700,150,25),
   N("运 单 号：",ControlType.Text,60,840,100,25),N(Bill,ControlType.Text,180,840,180,25));}
  public static List<string> Run(){
   var passed=new List<string>();int cases=0;
   foreach(double scale in new[]{1.0,1.25,1.5,1.75,2.0})foreach(int offset in new[]{-1600,0,1800})foreach(double height in new[]{760.0,900.0,1170.0}){
    var baseCard=Card();baseCard.Box=new Rect(0,0,620,height);var card=Scale(baseCard,scale,offset,100);
    Assert(UiRules.Input(card)!=null,"scaled/nested input");Assert(UiRules.Query(card)!=null&&UiRules.Query(card).Text=="查\n询","query must not use footer link");
    Assert(UiRules.ResultCard(card,Bill)!=null&&UiRules.ResultCard(card,Bill).Text=="收件城市","card position across scale/height");
    var receipt=Scale(Receipt(),scale,offset,100);var phone=receipt.Texts.Single(n=>n.Text=="136****8300");var eye=UiRules.Eye(receipt,phone);
    Assert(eye!=null&&Math.Abs(eye.Box.Left-(366*scale+offset))<.01,"scaled eye");
    Assert(JdLookup.ReadReceipt(receipt,Bill).Address=="测试省测试市测试区完整地址","receipt text independent of scale");cases++;
   }
   passed.Add(cases+" simulated scale/position/height layouts select correct input, card, eye and receipt");
   var shuffled=Card();shuffled.Nodes.Reverse();Assert(UiRules.ResultCard(shuffled,Bill).Text=="收件城市","enumeration reorder");passed.Add("card selection tolerates extra labels and reordered accessibility nodes");
   var wrong=Card();Assert(UiRules.ResultCard(wrong,"JDVA00000000002")==null,"near-matching bill must fail");wrong.Nodes.Add(N(Bill,ControlType.Text,171,800,155,22));Assert(UiRules.ResultCard(wrong,Bill)==null,"duplicate card must fail");passed.Add("near-matching and duplicate waybills are rejected");
   var ambiguous=Card();ambiguous.Nodes.Add(N("未知按钮",ControlType.Text,360,340,60,31));Assert(UiRules.ResultCard(ambiguous,Bill)==null,"ambiguous card position must fail");passed.Add("ambiguous card target is rejected");
   foreach(double font in new[]{.9,1.0,1.1})foreach(int length in new[]{2,3,4,6}){
    var phone=N("136****8300",ControlType.Text,100+length*25,546,145*font,31*font);double gap=16*font;
    var a=N("",ControlType.Group,phone.Box.Right+gap,phone.Box.Top-3*font,37*font,37*font);var b=N("",ControlType.Group,a.Box.Right+9*font,a.Box.Top,a.Box.Width,a.Box.Height);
    var r=S(phone,a,b);Assert(UiRules.Eye(r,phone)==a,"font/name length eye adaptation");
   }passed.Add("eye tolerates varied name lengths and 10 percent font size changes");
   var single=Receipt();var ph=single.Texts.Single(n=>n.Text=="136****8300");single.Nodes.RemoveAll(n=>n.Type==ControlType.Group&&Math.Abs(n.Box.Left-366)<1&&n.Box.Top>500);Assert(UiRules.Eye(single,ph)==null,"single call button must not be clicked");passed.Add("missing eye cannot fall back to the call button");
   var named=S(N("136****8300",ControlType.Text,205,546,145,31),N("显示手机号",ControlType.Button,366,543,37,37));Assert(UiRules.Eye(named,named.Nodes[0])==named.Nodes[1],"named eye");passed.Add("explicit show-phone control is recognized without fuzzy customer data");
   var rcp=Receipt();rcp.Nodes.Last().Text="JDVA00000000002";bool refused=false;try{JdLookup.ReadReceipt(rcp,Bill);}catch{refused=true;}Assert(refused,"receipt waybill exact check");passed.Add("receipt still rejects one-character waybill mismatch");
   rcp=Receipt();rcp.Nodes.RemoveAt(rcp.Nodes.Count-1);rcp.Nodes.Last().Text="运 单 号 ： "+Bill;Assert(JdLookup.ReadReceipt(rcp,Bill).Waybill==Bill,"inline bill field");passed.Add("whitespace, colon and inline field-label variants are recognized");
   rcp=Receipt();rcp.Nodes.Last().Text="JDVA 00000000001";refused=false;try{JdLookup.ReadReceipt(rcp,Bill);}catch{refused=true;}Assert(refused,"identifier whitespace cannot be guessed");passed.Add("waybill data is never normalized with fuzzy label rules");
   var viewport=S(N("目标",ControlType.Text,100,100,80,25));viewport.Box=new Rect(0,0,620,110);Assert(!UiRules.Visible(viewport,viewport.Nodes[0]),"clipped center");viewport.Box=new Rect(0,0,620,1170);Assert(UiRules.Visible(viewport,viewport.Nodes[0]),"visible center");passed.Add("partially clipped controls wait for visibility rather than blind clicking");
   var cancelled=Card();cancelled.Nodes.Add(N("已取消",ControlType.Text,300,450,80,25));Assert(UiRules.Cancelled(cancelled,Bill),"cancelled current order");Assert(!UiRules.Cancelled(cancelled,"JDVA00000000002"),"other cancelled bill");cancelled.Nodes.Add(N("电子存根",ControlType.Text,420,280,100,25));Assert(!UiRules.Cancelled(cancelled,Bill),"cancelled with readable receipt must still be queried");passed.Add("cancelled detection requires exact bill and absent receipt; unrelated cancellations cannot mark the current order");
   return passed;
  }
 }
}
