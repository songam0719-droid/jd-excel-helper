using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;

namespace JdHelper {
 // Tolerances apply only to control labels and layout, never to customer data or waybills.
 public static class UiRules {
  public static string Label(string text){return Regex.Replace(text??"",@"[\s\u200B\uFEFF]","").TrimEnd(':','：','>','›','»','→');}
  public static bool IsLabel(string text,string expected){return Label(text)==Label(expected);}
  public static bool FieldLabel(string text,string label){return Regex.IsMatch(text??"",@"^\s*"+String.Join(@"\s*",label.Select(c=>Regex.Escape(c.ToString())))+@"\s*[:：]")||IsLabel(text,label);}
  public static bool SameBill(UiSnapshot s,string bill){var bills=s.Texts.Where(n=>Regex.IsMatch(n.Text.Trim(),@"^JD[A-Z0-9]{8,30}$")).ToList();return bills.Count==1&&bills[0].Text.Trim()==bill;}
  public static bool Cancelled(UiSnapshot s,string bill){return SameBill(s,bill)&&!s.Has("电子存根")&&!s.Receipt&&s.Texts.Any(n=>IsLabel(n.Text,"已取消")||Label(n.Text).StartsWith("您已取消寄件"));}
  static double X(UiNode n){return n.Box.Left+n.Box.Width/2;}
  static double Y(UiNode n){return n.Box.Top+n.Box.Height/2;}
  static bool SameBox(UiNode a,UiNode b){double t=Math.Max(2,Math.Min(a.Box.Height,b.Box.Height)*.2);return Math.Abs(X(a)-X(b))<t&&Math.Abs(Y(a)-Y(b))<t;}
  static List<UiNode> DistinctBoxes(IEnumerable<UiNode> nodes){var result=new List<UiNode>();foreach(var n in nodes)if(!result.Any(p=>SameBox(n,p)))result.Add(n);return result;}
  public static bool Visible(UiSnapshot s,UiNode n){
   if(n==null||n.Offscreen||n.Box.IsEmpty)return false;var area=Rect.Intersect(s.Box,n.Box);
   return !area.IsEmpty&&area.Width>=Math.Min(n.Box.Width,8)&&area.Height>=Math.Min(n.Box.Height,8)&&area.Contains(new Point(X(n),Y(n)));
  }
  public static UiNode Input(UiSnapshot s){
   var options=DistinctBoxes(s.Nodes.Where(n=>n.Type==ControlType.Edit&&!n.Offscreen&&n.Box.Top<s.Box.Top+Math.Min(s.Box.Height*.4,s.Box.Width*.35)&&n.Box.Width>s.Box.Width*.2).OrderByDescending(n=>n.Index));
   return options.Count==1?options[0]:null;
  }
  public static UiNode Query(UiSnapshot s){
   var input=Input(s);if(input==null)return null;
   var options=DistinctBoxes(s.Nodes.Where(n=>IsLabel(n.Text,"查询")&&!n.Offscreen&&n.Box.Left>=input.Box.Right-input.Box.Height&&Math.Abs(Y(n)-Y(input))<input.Box.Height*1.8));
   return options.Count==1?options[0]:null;
  }
  public static UiNode ResultCard(UiSnapshot s,string bill){
   var bills=s.Texts.Where(n=>n.Text.Trim()==bill).ToList();if(bills.Count!=1)return null;var anchor=bills[0];double line=anchor.Box.Height;
   double bottom=anchor.Box.Bottom+line*8;
   var next=s.Texts.Where(n=>Regex.IsMatch(n.Text.Trim(),@"^JD[A-Z0-9]{8,30}$")&&n.Box.Top>anchor.Box.Bottom).OrderBy(n=>n.Box.Top).FirstOrDefault();
   if(next!=null)bottom=Math.Min(bottom,next.Box.Top);
   var rows=s.Texts.Where(n=>n.Box.Top>=anchor.Box.Bottom+line*.35&&n.Box.Bottom<bottom&&n.Text.Trim().Length>=2&&n.Text.Trim().Length<=24&&!Regex.IsMatch(n.Text,@"分享|备注|客服|评价|优惠|有礼|预计|下单|查询|JD[A-Z0-9]" )).ToList();
   // Locate the matched card's two city/name columns geometrically. Extra text nodes
   // or changed enumeration order do not shift the target as with a fixed +5 index.
   var left=rows.Where(n=>X(n)<s.Box.Left+s.Box.Width*.43).ToList();
   var right=rows.Where(n=>X(n)>s.Box.Left+s.Box.Width*.58&&X(n)<s.Box.Right-s.Box.Width*.02&&left.Any(l=>Math.Abs(Y(l)-Y(n))<=Math.Max(l.Box.Height,n.Box.Height)*.55)).OrderBy(n=>n.Box.Top).ToList();
   if(right.Count==0)return null;
   var top=DistinctBoxes(right.Where(n=>Math.Abs(Y(n)-Y(right[0]))<right[0].Box.Height*.55));
   return top.Count==1?top[0]:null;
  }
  public static UiNode Eye(UiSnapshot s,UiNode phone){
   if(phone==null)return null;double line=phone.Box.Height;
   var icons=DistinctBoxes(s.Nodes.Where(n=>(n.Type==ControlType.Group||n.Type==ControlType.Button||n.Type==ControlType.Image)&&n.Box.Width>=line*.5&&n.Box.Width<=line*2&&n.Box.Height>=line*.5&&n.Box.Height<=line*2&&n.Box.Left>=phone.Box.Right-line*.2&&n.Box.Left-phone.Box.Right<line*5&&Math.Abs(Y(n)-Y(phone))<line*.65).OrderBy(n=>n.Box.Left).ThenBy(n=>n.Type==ControlType.Group?0:1));
   var named=icons.Where(n=>new[]{"显示手机号","查看手机号","显示号码","查看号码","显示电话"}.Any(t=>IsLabel(n.Text,t))).ToList();
   if(named.Count==1)return named[0];
   // The observed unlabeled controls are an eye followed by a call button. Require
   // the full pair; a lone call button must never be mistaken for the eye.
   if(icons.Count!=2||!String.IsNullOrWhiteSpace(icons[0].Text)||!String.IsNullOrWhiteSpace(icons[1].Text))return null;
   var a=icons[0];var b=icons[1];
   if(a.Box.Left-phone.Box.Right>line*2.2||b.Box.Left-a.Box.Right<0||b.Box.Left-a.Box.Right>line*1.4||Math.Abs(a.Box.Height-b.Box.Height)>line*.4)return null;
   return a;
  }
  public static UiNode Back(UiSnapshot s){
   return s.AllNodes.Where(n=>!n.Offscreen&&(n.Type==ControlType.Group||n.Type==ControlType.Image||n.Type==ControlType.Button)&&n.Box.Left>=s.Box.Left&&n.Box.Left<s.Box.Left+s.Box.Width*.18&&n.Box.Top>=s.Box.Top&&n.Box.Top<s.Box.Top+Math.Min(s.Box.Height*.22,s.Box.Width*.22)&&n.Box.Width<s.Box.Width*.16&&n.Box.Height<s.Box.Width*.16).OrderByDescending(n=>n.Index).FirstOrDefault();
  }
 }
}
