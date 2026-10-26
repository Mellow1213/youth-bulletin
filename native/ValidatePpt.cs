using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
class ValidatePpt {
 static bool Overflow(dynamic s) {
  dynamic r=s.TextFrame2.TextRange;
  return (double)r.BoundTop+(double)r.BoundHeight>(double)s.Top+(double)s.Height+0.8 || (double)r.BoundLeft+(double)r.BoundWidth>(double)s.Left+(double)s.Width+0.8;
 }
 static string Compact(string text) {
  var result=new StringBuilder();
  foreach(string raw in Regex.Split(text,"[\\r\\n\\v]+")) {
   string line=raw.Trim();if(line.Length==0)continue;
   if(result.Length>0)result.Append(Regex.IsMatch(line,@"^\d+\s*(?:[.．:)]|\s)")?"\r":" ");
   result.Append(line);
  }
  return result.ToString();
 }
 [STAThread] static int Main(string[] args){
  dynamic app=null,deck=null; bool quit=false,changed=false,compacted=false; double fontSize=11,referenceFontSize=12; var issues=new List<string>();
  try {
   app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));quit=app.Presentations.Count==0;
   deck=app.Presentations.Open(Path.GetFullPath(args[0]),0,0,0);
   for(int i=1;i<=deck.Slides.Count;i++){
    dynamic slide=deck.Slides.Item(i);
    for(int j=1;j<=slide.Shapes.Count;j++){
     dynamic s=slide.Shapes.Item(j);int id=s.Id;
     if(!((i==1&&id==128)||(i==2&&(id==3||id==11||id==97||id==2200))))continue;
     if(s.HasTextFrame==0||s.TextFrame.HasText==0)continue;
     if(i==2&&id==11&&Overflow(s)) {
      referenceFontSize=(double)s.TextFrame2.TextRange.Font.Size;
      while(Overflow(s)&&referenceFontSize>1){referenceFontSize=Math.Max(1,referenceFontSize-0.25);s.TextFrame2.TextRange.Font.Size=(float)referenceFontSize;changed=true;}
     }
     if(i==2&&id==3&&Overflow(s)) {
      string original=s.TextFrame2.TextRange.Text, cleaned=Compact(original);
      if(cleaned!=original.TrimEnd('\r')) {s.TextFrame2.TextRange.Text=cleaned;s.TextFrame2.TextRange.Font.Size=11;compacted=true;changed=true;}
      // Keep the original typeface, color, box geometry and proportional line spacing.
      while(Overflow(s)&&fontSize>1) {fontSize=Math.Max(1,fontSize-0.25);s.TextFrame2.TextRange.Font.Size=(float)fontSize;s.TextFrame.TextRange.ParagraphFormat.LineRuleAfter=0;s.TextFrame.TextRange.ParagraphFormat.SpaceAfter=(float)fontSize;changed=true;}
     }
     if(Overflow(s))issues.Add(i==1?"광고 내용이 정해진 영역을 넘습니다. 내용을 줄이거나 줄바꿈을 조정하세요.":id==3?"본문을 PowerPoint의 최소 글자 크기까지 줄였지만 영역에 들어가지 않습니다.":"설교 제목 또는 성경 구절 제목이 정해진 영역을 넘습니다.");
    }
   }
   if(issues.Count==0&&changed)deck.Save();
   if(issues.Count==0&&args.Length>2)deck.SaveAs(Path.GetFullPath(args[2]),32);
   File.WriteAllText(args[1],new JavaScriptSerializer().Serialize(new{ok=issues.Count==0,issues=issues,fitting=new{compacted=compacted,fontSize=fontSize,referenceFontSize=referenceFontSize}}),new UTF8Encoding(false));return 0;
  }catch(Exception e){
   File.WriteAllText(args[1],new JavaScriptSerializer().Serialize(new{ok=false,issues=new[]{"PowerPoint 출력 검증에 실패했습니다: "+e.Message}}),new UTF8Encoding(false));return 1;
  }finally{
   // Only the successful path above saves. Discard temporary edits on failure without a hidden prompt.
   if(deck!=null){try{deck.Saved=-1;deck.Close();}catch{}}
   if(app!=null&&quit){try{app.Quit();}catch{}}
  }
 }
}




