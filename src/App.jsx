import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

var F  = "'Segoe UI','Arial Hebrew',Arial,sans-serif";
var BL = "#6366F1", GR = "#10B981", RD = "#EF4444", GO = "#F59E0B";
var SH = "0 1px 4px rgba(0,0,0,0.08)";

function fmt(n){ return "\u20AA"+Number(n||0).toLocaleString("he-IL",{maximumFractionDigits:0}); }
function uid(){ return Date.now()+Math.floor(Math.random()*99999); }
function mAmt(e){ var a=(parseFloat(e.amount)||0)*(parseFloat(e.qty)||1); return e.freq==="monthly"?a:e.freq==="quarterly"?a/3:a/12; }
function ce(tag,props){ var args=[tag,props||null]; for(var i=2;i<arguments.length;i++)args.push(arguments[i]); return React.createElement.apply(React,args); }

var CATS = [
  {id:"housing",l:"דיור",i:"🏠"},{id:"utilities",l:"שירותים",i:"💡"},
  {id:"insurance",l:"ביטוחים",i:"🛡️"},{id:"car",l:"רכב",i:"🚗"},
  {id:"kids",l:"ילדים",i:"👶"},{id:"food",l:"מזון",i:"🛒"},
  {id:"health",l:"בריאות",i:"⚕️"},{id:"savings",l:"חסכונות",i:"💰"},
  {id:"other",l:"אחר",i:"📦"},
];
var MONTHS = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני","יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"];
var SHORT  = ["ינ׳","פב׳","מרץ","אפ׳","מאי","יוני","יול׳","אוג׳","ספ׳","אוק׳","נוב׳","דצ׳"];
var NOW_M  = new Date().getMonth();
var NOW_Y  = new Date().getFullYear();
var NOW_D  = new Date().getDate();
var SK = "bgt_v2";

var ACCS = [
  {id:"a1",name:"ברקת",emoji:"👩",color:"#6366F1",grad:"linear-gradient(135deg,#6366F1,#8B5CF6)",balance:0},
  {id:"a2",name:"תומר",emoji:"👨",color:"#0EA5E9",grad:"linear-gradient(135deg,#0EA5E9,#38BDF8)",balance:0},
  {id:"a3",name:"משותף",emoji:"🏠",color:"#10B981",grad:"linear-gradient(135deg,#10B981,#34D399)",balance:0},
];

function loadData(){ try{ var r=localStorage.getItem(SK); return r?JSON.parse(r):null; }catch(e){ return null; } }
function saveData(d){ try{ localStorage.setItem(SK,JSON.stringify(d)); }catch(e){} }
function doExport(accs,exp,inc,fut,sav,mon){
  var wb=XLSX.utils.book_new();
  var s1=[["חשבון","הכנסות/חודש","הוצ׳ קבועות","פער","יתרה"]];
  accs.forEach(function(a){
    var fx=(exp[a.id]||[]).reduce(function(s,e){return s+mAmt(e);},0);
    var ic=(inc[a.id]||[]).reduce(function(s,r){return s+(parseFloat(r.amount)||0);},0);
    s1.push([a.name,Math.round(ic),Math.round(fx),Math.round(ic-fx),parseFloat(a.balance)||0]);
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s1),"סיכום");
  var s2=[["חשבון","שם","קטגוריה","סכום","תדירות","חודשי"]];
  accs.forEach(function(a){(exp[a.id]||[]).forEach(function(e){var c=CATS.find(function(x){return x.id===e.category;})||{l:""};s2.push([a.name,e.name,c.l,e.amount,e.freq,Math.round(mAmt(e))]);});});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s2),"הוצאות קבועות");
  var s3=[["חשבון","שם","סכום"]];
  accs.forEach(function(a){(inc[a.id]||[]).forEach(function(r){s3.push([a.name,r.name,r.amount]);});});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s3),"הכנסות");
  var s4=[["שנה","חודש","שם","חשבון","סכום","חוזרת"]];
  (fut||[]).forEach(function(f){var a=accs.find(function(x){return x.id===f.account;})||{name:""};s4.push([f.year||NOW_Y,MONTHS[f.month],f.name,a.name,f.amount,f.recurring?"כן":""]);});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s4),"תכנון עתידי");
  var s5=[["חשבון","שם","נוכחי","יעד","הפקדה"]];
  (sav||[]).forEach(function(sv){var a=accs.find(function(x){return x.id===sv.accId;})||{name:""};s5.push([a.name,sv.name,sv.current,sv.target,sv.monthly]);});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s5),"חסכונות");
  var b64=XLSX.write(wb,{bookType:"xlsx",type:"base64"});
  var a2=document.createElement("a");
  a2.href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"+b64;
  a2.download="תקציב_ביתי.xlsx"; document.body.appendChild(a2); a2.click(); document.body.removeChild(a2);
}

function buildDefaults(){
  var exp={a1:[],a2:[],a3:[]}, inc={a1:[],a2:[],a3:[]};
  function E(id,n,cat,amt,freq){ exp[id].push({id:uid(),name:n,category:cat,amount:amt,qty:1,freq:freq||"monthly",note:""}); }
  function I(id,n,amt){ inc[id].push({id:uid(),name:n,amount:amt,type:"salary",note:""}); }
  I("a1","משכורת ברקת",15000);
  E("a1","פנסיה","savings",800); E("a1","קרן השתלמות","savings",400);
  E("a1","ביטוח חיים","insurance",180); E("a1","ביטוח בריאות","insurance",220);
  E("a1","קופת מכבי","health",200); E("a1","כרטיס אשראי","other",2000);
  I("a2","משכורת תומר",18000);
  E("a2","פנסיה","savings",900); E("a2","קרן השתלמות","savings",450);
  E("a2","ביטוח חיים","insurance",200); E("a2","ביטוח בריאות","insurance",180);
  E("a2","קופת מכבי","health",200); E("a2","כרטיס אשראי","other",2500);
  E("a3","משכנתא","housing",6500); E("a3","ארנונה","housing",850);
  E("a3","ועד בית","housing",400); E("a3","חשמל","utilities",500);
  E("a3","מים","utilities",150); E("a3","גז","utilities",80);
  E("a3","אינטרנט","utilities",120); E("a3","טלפון ברקת","utilities",100);
  E("a3","טלפון תומר","utilities",100); E("a3","ביטוח דירה","insurance",180);
  E("a3","צהרונים","kids",1800); E("a3","חוגים ילדים","kids",600);
  E("a3","ליסינג","car",2200); E("a3","ביטוח רכב","car",350);
  E("a3","דלק","car",600); E("a3","סופרמרקט","food",3500);
  var fut=[
    {id:uid(),name:"רישיון רכב",month:1,year:NOW_Y,amount:500,account:"a3",category:"car",recurring:true},
    {id:uid(),name:"טסט",month:3,year:NOW_Y,amount:300,account:"a3",category:"car",recurring:true},
    {id:uid(),name:"ביטוח רכב שנתי",month:4,year:NOW_Y,amount:4200,account:"a3",category:"insurance",recurring:true},
    {id:uid(),name:"קייטנת קיץ",month:6,year:NOW_Y,amount:4500,account:"a3",category:"kids",recurring:true},
    {id:uid(),name:"ציוד לבית ספר",month:8,year:NOW_Y,amount:1200,account:"a3",category:"kids",recurring:true},
    {id:uid(),name:"חגים",month:9,year:NOW_Y,amount:1500,account:"a3",category:"other",recurring:false},
    {id:uid(),name:"טיול משפחתי",month:7,year:NOW_Y,amount:8000,account:"a3",category:"other",recurring:false}
  ];
  var sav=[
    {id:uid(),accId:"a1",name:"חיסכון בנקאי",current:15000,target:50000,monthly:500},
    {id:uid(),accId:"a1",name:"קרן השתלמות ברקת",current:40000,target:120000,monthly:400},
    {id:uid(),accId:"a2",name:"קרן השתלמות תומר",current:55000,target:150000,monthly:450},
    {id:uid(),accId:"a2",name:"תיק השקעות",current:80000,target:300000,monthly:1000}
  ];
  return {exp:exp,inc:inc,fut:fut,sav:sav};
}

// ── CASH FLOW ENGINE — running balance ───────────────────────────────────────
function computeFlow(acc, exp, inc, mon, fut, year){
  var fixE=(exp[acc.id]||[]).reduce(function(s,e){return s+mAmt(e);},0);
  var fixI=(inc[acc.id]||[]).reduce(function(s,r){return s+(parseFloat(r.amount)||0);},0);
  var bal=parseFloat(acc.balance)||0;
  var futList=fut||[];
  return Array.from({length:12},function(_,mo){
    var key=year+"-"+mo;
    var vd=((mon[acc.id]||{})[key])||{items:[]};
    var varE=(vd.items||[]).reduce(function(s,it){return s+(parseFloat(it.amount)||0);},0);
    var planned=futList.filter(function(f){return f.account===acc.id&&(f.year||NOW_Y)===year&&f.month===mo;});
    var oneT=planned.reduce(function(s,f){return s+(parseFloat(f.amount)||0);},0);
    var net=fixI-fixE-varE-oneT;
    var open=bal; bal=bal+net;
    var hasData=varE>0||oneT>0;
    return {mo:mo,key:key,open:open,fixI:fixI,fixE:fixE,varE:varE,vd:vd,oneT:oneT,planned:planned,net:net,close:bal,hasData:hasData,past:year<NOW_Y||(year===NOW_Y&&mo<NOW_M),isNow:year===NOW_Y&&mo===NOW_M};
  });
}

// ── CONFIRM ───────────────────────────────────────────────────────────────────
function Confirm(props){
  return ce("div",{style:{position:"fixed",inset:0,zIndex:2000,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:16},onClick:function(e){if(e.target===e.currentTarget)props.onCancel();}},
    ce("div",{style:{background:"#fff",borderRadius:20,padding:28,maxWidth:300,width:"100%",textAlign:"center",fontFamily:F,boxShadow:"0 20px 50px rgba(0,0,0,0.2)"}},
      ce("div",{style:{fontSize:34,marginBottom:10}},"🗑️"),
      ce("div",{style:{fontSize:16,fontWeight:800,marginBottom:6}},"למחוק?"),
      ce("div",{style:{fontSize:13,color:"#666",marginBottom:22,lineHeight:1.7}},'בטוחה שאת רוצה למחוק "'+(props.name||"זה")+'"?'),
      ce("div",{style:{display:"flex",gap:10,justifyContent:"center"}},
        ce("button",{style:{padding:"9px 20px",borderRadius:10,border:"1.5px solid #ddd",background:"#fff",cursor:"pointer",fontFamily:F,fontWeight:700},onClick:props.onCancel},"ביטול"),
        ce("button",{style:{padding:"9px 20px",borderRadius:10,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontWeight:700},onClick:props.onConfirm},"מחקי ✓")
      )
    )
  );
}

// ── SUMMARY ───────────────────────────────────────────────────────────────────
function SummaryTab(props){
  var accs=props.accs,exp=props.exp,inc=props.inc,mon=props.mon,fut=props.fut||[];
  var simM=useState(NOW_M), simY=useState(NOW_Y);
  var rows=accs.map(function(a){
    var fixE=(exp[a.id]||[]).reduce(function(s,e){return s+mAmt(e);},0);
    var fixI=(inc[a.id]||[]).reduce(function(s,r){return s+(parseFloat(r.amount)||0);},0);
    return {a:a,fixE:fixE,fixI:fixI,gap:fixI-fixE};
  });
  var totI=rows.reduce(function(s,r){return s+r.fixI;},0);
  var totE=rows.reduce(function(s,r){return s+r.fixE;},0);
  // annual chart
  var rb=accs.reduce(function(s,a){return s+(parseFloat(a.balance)||0);},0);
  var chart=Array.from({length:12},function(_,mi){
    var out=accs.reduce(function(s,a){var f=computeFlow(a,exp,inc,mon,fut,NOW_Y)[mi];return s+(f?f.fixE+f.varE+f.oneT:0);},0);
    var i2=accs.reduce(function(s,a){var f=computeFlow(a,exp,inc,mon,fut,NOW_Y)[mi];return s+(f?f.fixI:0);},0);
    rb+=(i2-out); return {mi:mi,bal:rb};
  });
  var maxA=Math.max.apply(null,chart.map(function(m){return Math.abs(m.bal);}));
  if(!maxA)maxA=1;
  // per-account projection
  var mEl=simY[0]===NOW_Y?simM[0]+1:simY[0]>NOW_Y?(simY[0]-NOW_Y)*12+simM[0]+1:0;
  var missing=NOW_D>=15&&accs.some(function(a){var m=(mon[a.id]||{})[NOW_Y+"-"+NOW_M];return !(m&&(m.items||[]).length>0);});
  var upAlert=fut.filter(function(f){return (f.year||NOW_Y)===NOW_Y&&(f.month===NOW_M||f.month===NOW_M+1);});
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    missing ? ce("div",{style:{background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:12,padding:"10px 14px",marginBottom:14,display:"flex",gap:10,alignItems:"center"}},
      ce("span",{style:{fontSize:18}},"⏰"),
      ce("div",null,
        ce("div",{style:{fontSize:12,fontWeight:700,color:"#92400E"}},"לא הוזנו הוצ׳ משתנות ל-"+MONTHS[NOW_M]),
        ce("div",{style:{fontSize:11,color:"#B45309"}},"כבר ה-"+NOW_D+" — זכרי להזין")
      )
    ) : null,
    upAlert.length>0 ? ce("div",{style:{background:"#FEF3C7",border:"1px solid #FDE68A",borderRadius:12,padding:"10px 14px",marginBottom:14}},
      ce("div",{style:{fontSize:12,fontWeight:700,color:"#92400E",marginBottom:6}},"🔔 הוצאות קרובות"),
      ce("div",{style:{display:"flex",gap:6,flexWrap:"wrap"}},
        upAlert.map(function(f){
          var cat=CATS.find(function(c){return c.id===f.category;})||{i:"📦"};
          return ce("div",{key:f.id,style:{background:"#fff",borderRadius:8,padding:"4px 10px",fontSize:11,border:"1px solid #FDE68A"}},cat.i+" "+f.name+" — "+MONTHS[f.month]+" "+fmt(f.amount));
        })
      )
    ) : null,
    // KPI
    ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:14}},
      [{l:"הכנסות/חודש",v:fmt(Math.round(totI)),c:GR},
       {l:"הוצ׳ קבועות",v:fmt(Math.round(totE)),c:RD},
       {l:"פער חודשי",v:fmt(Math.round(totI-totE)),c:(totI-totE)>=0?GR:RD}
      ].map(function(k){
        return ce("div",{key:k.l,style:{background:"#fff",borderRadius:12,padding:"12px 8px",textAlign:"center",boxShadow:SH}},
          ce("div",{style:{fontSize:9,color:"#666",fontWeight:700,textTransform:"uppercase",marginBottom:4}},k.l),
          ce("div",{style:{fontSize:14,fontWeight:900,color:k.c}},k.v)
        );
      })
    ),
    // Income vs Expenses
    ce("div",{style:{background:"#fff",borderRadius:14,boxShadow:SH,overflow:"hidden",marginBottom:14}},
      ce("div",{style:{padding:"12px 16px",fontWeight:800,fontSize:13,borderBottom:"1px solid #eee"}},"📊 הכנסות vs. הוצאות"),
      rows.map(function(r,i){
        return ce("div",{key:r.a.id,style:{padding:"12px 16px",borderBottom:i<rows.length-1?"1px solid #f5f5f5":"none"}},
          ce("div",{style:{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}},
            ce("span",{style:{fontWeight:700}},r.a.emoji+" "+r.a.name),
            ce("span",{style:{fontWeight:800,color:r.gap>=0?GR:RD,padding:"2px 10px",borderRadius:20,background:r.gap>=0?"#ECFDF5":"#FEF2F2"}},(r.gap>=0?"+":"")+fmt(Math.round(r.gap))+"/חודש")
          ),
          ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6}},
            [{l:"הכנסות",v:fmt(Math.round(r.fixI)),c:GR},{l:"קבועות",v:fmt(Math.round(r.fixE)),c:RD}].map(function(item){
              return ce("div",{key:item.l,style:{background:"#f8f9fa",borderRadius:8,padding:"6px 8px",textAlign:"center"}},
                ce("div",{style:{fontSize:9,color:"#666",marginBottom:2}},item.l),
                ce("div",{style:{fontSize:13,fontWeight:700,color:item.c}},item.v)
              );
            })
          )
        );
      })
    ),
    // Annual chart
    ce("div",{style:{background:"#fff",borderRadius:14,padding:"14px 16px",marginBottom:14,boxShadow:SH}},
      ce("div",{style:{fontSize:13,fontWeight:800,marginBottom:10}},"📈 תזרים שנתי "+NOW_Y+" — יתרה מצטברת"),
      ce("div",{style:{display:"flex",alignItems:"flex-end",gap:2,height:80}},
        chart.map(function(m){
          var isN=m.mi===NOW_M,h=Math.max(Math.round((Math.abs(m.bal)/maxA)*72),4),pos=m.bal>=0;
          return ce("div",{key:m.mi,style:{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:1}},
            ce("div",{style:{width:"100%",height:h+"px",borderRadius:"4px 4px 0 0",background:pos?GR+"CC":RD+"CC",border:isN?"2px solid "+BL:"none",boxSizing:"border-box"},title:SHORT[m.mi]+": "+fmt(Math.round(m.bal))}),
            ce("div",{style:{fontSize:8,color:isN?BL:"#aaa",fontWeight:isN?700:400}},SHORT[m.mi])
          );
        })
      )
    ),
    // Per-account projection
    ce("div",{style:{fontSize:11,fontWeight:700,color:"#666",marginBottom:8,textTransform:"uppercase"}},"יתרה צפויה עד "+MONTHS[simM[0]]+" "+simY[0]),
    ce("div",{style:{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}},
      ce("select",{value:simM[0],onChange:function(e){simM[1](+e.target.value);},style:{fontFamily:F,fontSize:12,padding:"5px 8px",borderRadius:8,border:"1px solid #ddd",direction:"rtl"}},
        MONTHS.map(function(m,i){return ce("option",{key:i,value:i},m);})),
      ce("select",{value:simY[0],onChange:function(e){simY[1](+e.target.value);},style:{fontFamily:F,fontSize:12,padding:"5px 8px",borderRadius:8,border:"1px solid #ddd"}},
        [NOW_Y-1,NOW_Y,NOW_Y+1,NOW_Y+2].map(function(y){return ce("option",{key:y,value:y},y);}))
    ),
    ce("div",{style:{display:"flex",flexDirection:"column",gap:10}},
      accs.map(function(a){
        var fixE=(exp[a.id]||[]).reduce(function(s,e){return s+mAmt(e);},0);
        var fixI=(inc[a.id]||[]).reduce(function(s,r){return s+(parseFloat(r.amount)||0);},0);
        var bal=fixI-fixE, open=parseFloat(a.balance)||0, proj=open+bal*mEl;
        return ce("div",{key:a.id,style:{background:"#fff",borderRadius:14,overflow:"hidden",boxShadow:SH}},
          ce("div",{style:{background:a.grad,padding:"12px 16px",color:"#fff"}},
            ce("div",{style:{fontSize:12,opacity:.8,marginBottom:2}},a.emoji+" "+a.name),
            ce("div",{style:{fontSize:20,fontWeight:900}},(bal>=0?"+":"")+fmt(Math.round(bal))+"/חודש")
          ),
          ce("div",{style:{padding:"10px 16px",display:"flex",gap:14,flexWrap:"wrap"}},
            [["פתיחה",open,"#666"],["יתרה ב"+MONTHS[simM[0]],proj,proj>=0?GR:RD]].map(function(r){
              return ce("div",{key:r[0]},
                ce("div",{style:{fontSize:9,color:"#aaa",fontWeight:600,textTransform:"uppercase",marginBottom:2}},r[0]),
                ce("div",{style:{fontSize:14,fontWeight:800,color:r[2]}},fmt(Math.round(r[1])))
              );
            })
          )
        );
      })
    )
  );
}

// ── YEAR PLANNER ──────────────────────────────────────────────────────────────
function YearTab(props){
  var accs=props.accs,exp=props.exp,inc=props.inc,mon=props.mon,setMon=props.setMon,fut=props.fut||[],setFut=props.setFut;
  var yr=useState(NOW_Y), detail=useState(null), addF=useState(null), delT=useState(null);
  function addPlanned(data){setFut(fut.concat([Object.assign({id:uid()},data)]));addF[1](null);}
  function delPlanned(id){setFut(fut.filter(function(f){return f.id!==id;}));delT[1](null);}
  function togRec(id){setFut(fut.map(function(f){return f.id===id?Object.assign({},f,{recurring:!f.recurring}):f;}));}
  var fixedM=Object.keys(exp).reduce(function(s,aId){return s+(exp[aId]||[]).reduce(function(s2,e){return s2+mAmt(e);},0);},0);
  var totI=Object.keys(inc).reduce(function(s,aId){return s+(inc[aId]||[]).reduce(function(s2,r){return s2+(parseFloat(r.amount)||0);},0);},0);
  var totOpen=accs.reduce(function(s,a){return s+(parseFloat(a.balance)||0);},0);
  var rb=totOpen;
  var planned=fut.filter(function(f){return (f.year||NOW_Y)===yr[0];});
  var mData=Array.from({length:12},function(_,mi){
    var varE=accs.reduce(function(s,a){var vd=((mon[a.id]||{})[yr[0]+"-"+mi])||{items:[]};return s+(vd.items||[]).reduce(function(s2,it){return s2+(parseFloat(it.amount)||0);},0);},0);
    var items=planned.filter(function(f){return f.month===mi;});
    var autos=fut.filter(function(f){return f.recurring&&(f.year||NOW_Y)===yr[0]-1&&f.month===mi;});
    var allIt=items.concat(autos);
    var oneT=allIt.reduce(function(s,f){return s+(parseFloat(f.amount)||0);},0);
    var totO=fixedM+varE+oneT, net=totI-totO, open=rb; rb+=net;
    var isPast=yr[0]<NOW_Y||(yr[0]===NOW_Y&&mi<NOW_M);
    var isNow=yr[0]===NOW_Y&&mi===NOW_M;
    var hasVar=varE>0||oneT>0;
    var st=isPast?(hasVar?"entered":"missing"):isNow?(hasVar?"entered":"pending"):"future";
    return {mi:mi,varE:varE,oneT:oneT,items:allIt,totO:totO,net:net,open:open,close:rb,st:st,isPast:isPast,isNow:isNow};
  });
  var stCol={entered:GR,missing:RD,pending:GO,future:"#aaa"};
  var stLbl={entered:"הוזן",missing:"חסר",pending:"ממתין",future:""};
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    ce("div",{style:{display:"flex",alignItems:"center",gap:10,marginBottom:14}},
      ce("button",{onClick:function(){yr[1](function(y){return y-1;});},style:{width:36,height:36,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:18,fontFamily:F}},"\u2039"),
      ce("div",{style:{flex:1,textAlign:"center",fontSize:22,fontWeight:900}},yr[0],yr[0]===NOW_Y?ce("span",{style:{fontSize:10,color:BL,marginRight:6,background:"#EEF2FF",padding:"2px 8px",borderRadius:10}}," שנה נוכחית"):null),
      ce("button",{onClick:function(){yr[1](function(y){return y+1;});},style:{width:36,height:36,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:18,fontFamily:F}},"\u203a")
    ),
    ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:14}},
      [{l:"הכנסה/חודש",v:fmt(Math.round(totI)),c:GR},{l:"קבועות/חודש",v:fmt(Math.round(fixedM)),c:BL},
       {l:"סה\"כ הוצ׳",v:fmt(Math.round(mData.reduce(function(s,m){return s+m.totO;},0))),c:RD},{l:"ממוצע/חודש",v:fmt(Math.round(mData.reduce(function(s,m){return s+m.totO;},0)/12)),c:GO}
      ].map(function(item){
        return ce("div",{key:item.l,style:{background:"#fff",border:"1px solid "+item.c+"22",borderRadius:12,padding:"10px 8px",textAlign:"center",boxShadow:SH}},
          ce("div",{style:{fontSize:8,color:"#aaa",fontWeight:700,textTransform:"uppercase",marginBottom:3}},item.l),
          ce("div",{style:{fontSize:13,fontWeight:900,color:item.c}},item.v)
        );
      })
    ),
    ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}},
      mData.map(function(d){
        var col=stCol[d.st];
        return ce("div",{key:d.mi,
          onClick:function(){detail[1](d);},
          style:{background:d.st==="missing"?"#FFF5F5":d.st==="entered"?"#F0FDF4":"#fff",borderRadius:14,padding:"11px 12px 9px",cursor:"pointer",border:"1.5px solid "+(d.isNow?BL:(d.varE>0||d.oneT>0)?"#FDE68A":"#eee"),boxShadow:d.isNow?"0 0 0 3px "+BL+"22":SH,opacity:d.isPast?0.75:1,transition:"all .18s",position:"relative"},
          onMouseEnter:function(ev){ev.currentTarget.style.transform="translateY(-2px)";ev.currentTarget.style.boxShadow="0 6px 16px rgba(0,0,0,0.12)";},
          onMouseLeave:function(ev){ev.currentTarget.style.transform="translateY(0)";ev.currentTarget.style.boxShadow=d.isNow?"0 0 0 3px "+BL+"22":SH;}},
          d.st!=="future"?ce("div",{style:{position:"absolute",top:7,left:9,display:"flex",alignItems:"center",gap:3}},
            ce("div",{style:{width:6,height:6,borderRadius:"50%",background:col}}),
            ce("span",{style:{fontSize:8,color:col,fontWeight:700}},stLbl[d.st])
          ):null,
          d.isNow?ce("div",{style:{position:"absolute",top:7,right:9,fontSize:8,fontWeight:800,color:BL}},"עכשיו"):null,
          ce("div",{style:{fontSize:11,fontWeight:700,marginBottom:12}},SHORT[d.mi]+" "+yr[0]),
          ce("div",{style:{fontSize:12,fontWeight:800}},fmt(Math.round(d.totO))),
          ce("div",{style:{fontSize:10,fontWeight:700,color:d.net>=0?GR:RD,marginTop:1}},(d.net>=0?"↑":"")+fmt(Math.round(d.net))),
          !d.varE&&!d.oneT&&d.isPast?ce("div",{style:{fontSize:8,color:"#aaa",marginTop:1}},"קבועות בלבד ✓"):null,
          d.oneT>0?ce("div",{style:{fontSize:9,color:GO,fontWeight:700,marginTop:1}},"⚡ "+fmt(Math.round(d.oneT))+" חד\"פ"):null,
          ce("div",{style:{fontSize:9,color:"#aaa",marginTop:3,borderTop:"1px dashed #eee",paddingTop:3}},
            "יתרה: ",ce("span",{style:{fontWeight:700,color:d.close>=0?GR:RD}},fmt(Math.round(d.close)))
          ),
          ce("button",{onClick:function(e){e.stopPropagation();addF[1]({month:d.mi,year:yr[0]});},style:{marginTop:5,width:"100%",padding:"3px",borderRadius:6,border:"1.5px dashed #ddd",background:"transparent",color:"#aaa",cursor:"pointer",fontFamily:F,fontSize:10,fontWeight:700}},"+ חד\"פ")
        );
      })
    ),
    ce("div",{style:{display:"flex",gap:12,marginTop:10,fontSize:9,color:"#aaa",flexWrap:"wrap"}},
      [{c:GR+"55",l:"הכנסות"},{c:BL+"55",l:"קבועות"},{c:GO+"99",l:"משתנות"}].map(function(x){
        return ce("span",{key:x.l,style:{display:"flex",alignItems:"center",gap:4}},
          ce("span",{style:{width:9,height:9,borderRadius:2,background:x.c,display:"inline-block"}}),x.l);
      })
    ),
    // Month detail modal
    detail[0]?ce("div",{style:{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.4)",display:"flex",alignItems:"center",justifyContent:"center",padding:16},onClick:function(e){if(e.target===e.currentTarget)detail[1](null);}},
      ce("div",{style:{background:"#fff",borderRadius:20,padding:22,width:"100%",maxWidth:440,maxHeight:"90vh",overflowY:"auto",fontFamily:F,direction:"rtl",boxShadow:"0 25px 60px rgba(0,0,0,0.2)"}},
        ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}},
          ce("span",{style:{fontSize:15,fontWeight:800}},MONTHS[detail[0].mi]+" "+yr[0]+" — פירוט"),
          ce("button",{onClick:function(){detail[1](null);},style:{width:28,height:28,borderRadius:"50%",border:"1px solid #ddd",background:"#f5f5f5",cursor:"pointer",fontSize:14,fontFamily:F}},"✕")
        ),
        ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:16}},
          [{l:"הכנסות",v:fmt(Math.round(totI)),c:GR},{l:"הוצ׳ סה\"כ",v:fmt(Math.round(detail[0].totO)),c:RD},
           {l:"פער",v:(detail[0].net>=0?"+":"")+fmt(Math.round(detail[0].net)),c:detail[0].net>=0?GR:RD},{l:"יתרת סגירה",v:fmt(Math.round(detail[0].close)),c:detail[0].close>=0?GR:RD}
          ].map(function(item){
            return ce("div",{key:item.l,style:{background:item.c+"0D",border:"1px solid "+item.c+"22",borderRadius:10,padding:"10px 8px",textAlign:"center"}},
              ce("div",{style:{fontSize:9,color:"#666",fontWeight:700,textTransform:"uppercase",marginBottom:3}},item.l),
              ce("div",{style:{fontSize:16,fontWeight:900,color:item.c}},item.v)
            );
          })
        ),
        accs.map(function(a){
          var row=computeFlow(a,exp,inc,mon,fut,yr[0])[detail[0].mi];
          if(!row)return null;
          var vi=(mon[a.id]||{})[yr[0]+"-"+detail[0].mi]||{items:[]};
          return ce("div",{key:a.id,style:{background:a.color+"08",border:"1px solid "+a.color+"22",borderRadius:12,padding:"12px 14px",marginBottom:10}},
            ce("div",{style:{fontSize:13,fontWeight:800,color:a.color,marginBottom:8}},a.emoji+" "+a.name),
            ce("div",{style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}},
              [["פתיחה",row.open,"#666"],["הכנסות",row.fixI,GR],["קבועות",row.fixE,BL],["משתנות",row.varE,GO],["פער",row.net,row.net>=0?GR:RD],["סגירה",row.close,row.close>=0?GR:RD]].map(function(r){
                return ce("div",{key:r[0],style:{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:"rgba(255,255,255,0.7)",borderRadius:6}},
                  ce("span",{style:{color:"#666",fontSize:11}},r[0]),
                  ce("span",{style:{fontWeight:700,color:r[2],fontSize:12}},fmt(Math.round(r[1])))
                );
              })
            ),
            (vi.items||[]).length>0?ce("div",{style:{background:"rgba(255,255,255,0.8)",borderRadius:8,padding:"8px 10px"}},
              ce("div",{style:{fontSize:10,color:"#666",fontWeight:700,marginBottom:4}},"פירוט משתנות"),
              (vi.items||[]).map(function(it,i){
                return ce("div",{key:i,style:{display:"flex",justifyContent:"space-between",fontSize:11,color:"#666",marginBottom:2}},
                  ce("span",null,it.name),ce("span",{style:{fontWeight:700}},fmt(it.amount)));
              })
            ):null
          );
        }),
        ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,marginTop:4}},
          ce("div",{style:{fontSize:12,fontWeight:700,color:"#666"}},"⚡ חד\"פ מתוכנן"),
          ce("button",{onClick:function(){var d2=detail[0];detail[1](null);addF[1]({month:d2.mi,year:yr[0]});},style:{padding:"5px 12px",borderRadius:8,border:"none",cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:11,background:BL,color:"#fff"}},"+ הוסף")
        ),
        (detail[0].items||[]).length===0?ce("div",{style:{fontSize:12,color:"#aaa",textAlign:"center",padding:"12px 0"}},"אין חד\"פ"):null,
        (detail[0].items||[]).map(function(f){
          var cat=CATS.find(function(c){return c.id===f.category;})||{i:"📦",l:""};
          var a2=accs.find(function(a){return a.id===f.account;})||{emoji:"",name:""};
          var isAuto=String(f.id).indexOf("auto_")===0;
          return ce("div",{key:f.id,style:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 10px",background:"#fff",borderRadius:9,marginBottom:6,border:"1px solid #eee"}},
            ce("div",{style:{flex:1}},
              ce("div",{style:{fontWeight:700,fontSize:12}},f.name),
              ce("div",{style:{fontSize:10,color:"#aaa",marginTop:1}},cat.i+" "+cat.l+" · "+a2.emoji+" "+a2.name+(f.recurring?" · 🔄 חוזרת":""))
            ),
            ce("div",{style:{display:"flex",alignItems:"center",gap:6}},
              ce("span",{style:{fontSize:12,fontWeight:700,color:GO}},fmt(parseFloat(f.amount)||0)),
              !isAuto?ce("button",{onClick:function(){togRec(f.id);},style:{padding:"3px 6px",borderRadius:6,border:"1px solid "+(f.recurring?BL:"#ddd"),background:f.recurring?BL+"12":"transparent",color:f.recurring?BL:"#aaa",cursor:"pointer",fontFamily:F,fontSize:10}},"🔄"):null,
              !isAuto?ce("button",{onClick:function(){delT[1](f);},style:{padding:"5px 10px",borderRadius:7,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontSize:11,fontWeight:700}},"🗑"):null
            )
          );
        })
      )
    ):null,
    addF[0]?ce("div",{style:{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.4)",display:"flex",alignItems:"center",justifyContent:"center",padding:16},onClick:function(e){if(e.target===e.currentTarget)addF[1](null);}},
      ce("div",{style:{background:"#fff",borderRadius:20,padding:20,width:"100%",maxWidth:420,fontFamily:F,direction:"rtl",boxShadow:"0 25px 60px rgba(0,0,0,0.2)"}},
        ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}},
          ce("span",{style:{fontSize:15,fontWeight:800}},"+ חד\"פ — "+MONTHS[addF[0].month]+" "+addF[0].year),
          ce("button",{onClick:function(){addF[1](null);},style:{width:28,height:28,borderRadius:"50%",border:"1px solid #ddd",background:"#f5f5f5",cursor:"pointer",fontSize:14,fontFamily:F}},"✕")
        ),
        ce(AddFutureForm,{accs:accs,month:addF[0].month,year:addF[0].year,onSave:addPlanned,onClose:function(){addF[1](null);}})
      )
    ):null,
    delT[0]?ce(Confirm,{name:delT[0].name,onConfirm:function(){delPlanned(delT[0].id);},onCancel:function(){delT[1](null);}}):null
  );
}
function AddFutureForm(props){
  var f=useState({name:"",category:"car",amount:"",account:props.accs[0].id,recurring:false});
  function s(k,v){f[1](Object.assign({},f[0],{[k]:v}));}
  var sI={fontFamily:F,fontSize:13,padding:"9px 12px",borderRadius:9,border:"1.5px solid #ddd",width:"100%",direction:"rtl",boxSizing:"border-box"};
  var sc=["ביטוח רכב","רישיון רכב","טסט","קייטנה","ביטוח חיים","תחזוקה","ביגוד","מנוי שנתי"];
  return ce("div",null,
    ce("div",{style:{marginBottom:12}},
      ce("div",{style:{fontSize:11,fontWeight:700,color:"#666",marginBottom:5}},"שם"),
      ce("input",{value:f[0].name,onChange:function(e){s("name",e.target.value);},placeholder:"לדוג׳ ביטוח רכב",style:sI}),
      ce("div",{style:{display:"flex",gap:5,flexWrap:"wrap",marginTop:7}},
        sc.map(function(sh){return ce("button",{key:sh,onClick:function(){s("name",sh);},style:{padding:"3px 10px",borderRadius:14,border:"1.5px solid "+(f[0].name===sh?BL:"#ddd"),background:f[0].name===sh?BL+"10":"transparent",color:f[0].name===sh?BL:"#666",cursor:"pointer",fontFamily:F,fontSize:10,fontWeight:600}},sh);})
      )
    ),
    ce("div",{style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}},
      ce("div",null,ce("div",{style:{fontSize:11,fontWeight:700,color:"#666",marginBottom:5}},"קטגוריה"),
        ce("select",{value:f[0].category,onChange:function(e){s("category",e.target.value);},style:sI},
          CATS.map(function(c){return ce("option",{key:c.id,value:c.id},c.i+" "+c.l);}))),
      ce("div",null,ce("div",{style:{fontSize:11,fontWeight:700,color:"#666",marginBottom:5}},"חשבון"),
        ce("select",{value:f[0].account,onChange:function(e){s("account",e.target.value);},style:sI},
          props.accs.map(function(a){return ce("option",{key:a.id,value:a.id},a.emoji+" "+a.name);})))
    ),
    ce("div",{style:{marginBottom:12}},
      ce("div",{style:{fontSize:11,fontWeight:700,color:"#666",marginBottom:5}},"סכום ₪"),
      ce("input",{type:"number",value:f[0].amount,onChange:function(e){s("amount",e.target.value);},placeholder:"0",style:sI})
    ),
    ce("label",{style:{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",marginBottom:18,padding:"9px 12px",borderRadius:10,background:f[0].recurring?"#EEF2FF":"transparent",border:"1.5px solid "+(f[0].recurring?BL:"#ddd")}},
      ce("input",{type:"checkbox",checked:f[0].recurring,onChange:function(e){s("recurring",e.target.checked);},style:{width:15,height:15}}),
      ce("span",{style:{fontWeight:600,color:f[0].recurring?BL:"#333"}},"🔄 חוזרת כל שנה")
    ),
    ce("div",{style:{display:"flex",gap:10,justifyContent:"flex-end"}},
      ce("button",{onClick:props.onClose,style:{padding:"10px 16px",borderRadius:10,border:"1.5px solid #ddd",background:"#fff",color:"#666",cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:13}},"ביטול"),
      ce("button",{onClick:function(){if(f[0].name.trim())props.onSave({name:f[0].name,category:f[0].category,amount:parseFloat(f[0].amount)||0,account:f[0].account,month:props.month,year:props.year,recurring:f[0].recurring});},style:{padding:"10px 22px",borderRadius:10,border:"none",cursor:"pointer",fontFamily:F,fontWeight:800,fontSize:13,background:f[0].name.trim()?"linear-gradient(135deg,#6366F1,#8B5CF6)":"#ccc",color:"#fff"}},"הוסף ✓")
    )
  );
}

// ── MONTHLY ENTRY ─────────────────────────────────────────────────────────────
function MonthTab(props){
  var accs=props.accs,exp=props.exp,inc=props.inc,mon=props.mon,setMon=props.setMon,fut=props.fut||[];
  var selA=useState(accs[0].id), selM=useState(NOW_M), selY=useState(NOW_Y), delT=useState(null);
  var acc=accs.find(function(a){return a.id===selA[0];})||accs[0];
  var key=selY[0]+"-"+selM[0];
  var vd=((mon[selA[0]]||{})[key])||{items:[]};
  var flow=computeFlow(acc,exp,inc,mon,fut,selY[0])[selM[0]];
  var fixTot=(exp[selA[0]]||[]).reduce(function(s,e){return s+mAmt(e);},0);
  function setVD(d){var nm=Object.assign({},mon);var na=Object.assign({},nm[selA[0]]||{});na[key]=d;nm[selA[0]]=na;setMon(nm);}
  function addIt(name){setVD({items:(vd.items||[]).concat([{id:uid(),name:name||"",amount:0}])});}
  function updIt(id,f,v){setVD({items:(vd.items||[]).map(function(it){return it.id===id?Object.assign({},it,{[f]:v}):it;})});}
  function doDelIt(id){setVD({items:(vd.items||[]).filter(function(it){return it.id!==id;})});delT[1](null);}
  var varT=(vd.items||[]).reduce(function(s,it){return s+(parseFloat(it.amount)||0);},0);
  var sc=["מכולת","דלק","פארם","קפה","מסעדה","בגדים","ניקוי","אחר"];
  var sI={fontFamily:F,fontSize:13,padding:"7px 10px",borderRadius:8,border:"1px solid #ddd",direction:"rtl"};
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    // Month nav
    ce("div",{style:{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}},
      ce("button",{onClick:function(){if(selM[0]===0){selM[1](11);selY[1](function(y){return y-1;});}else selM[1](function(m){return m-1;});},style:{width:36,height:36,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:18,fontFamily:F}},"\u2039"),
      ce("div",{style:{flex:1,textAlign:"center",fontSize:17,fontWeight:800}},MONTHS[selM[0]]+" "+selY[0]),
      ce("button",{onClick:function(){if(selM[0]===11){selM[1](0);selY[1](function(y){return y+1;});}else selM[1](function(m){return m+1;});},style:{width:36,height:36,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:18,fontFamily:F}},"\u203a"),
      ce("div",{style:{display:"flex",gap:6,flexWrap:"wrap"}},
        accs.map(function(a){
          return ce("button",{key:a.id,onClick:function(){selA[1](a.id);},style:{padding:"6px 14px",borderRadius:20,border:"1.5px solid "+(selA[0]===a.id?a.color:"#ddd"),background:selA[0]===a.id?a.grad:"#fff",color:selA[0]===a.id?"#fff":"#666",fontFamily:F,fontWeight:700,fontSize:12,cursor:"pointer"}},a.emoji+" "+a.name);
        })
      )
    ),
    // Flow summary bar
    flow?ce("div",{style:{background:acc.grad,borderRadius:14,padding:"14px 18px",marginBottom:14,color:"#fff"}},
      ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}},
        ce("div",{style:{fontSize:12,fontWeight:600,opacity:.8}},MONTHS[selM[0]]+" "+selY[0]+" — "+acc.name),
        !flow.hasData?ce("span",{style:{fontSize:10,opacity:.6,background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"2px 8px"}},"קבועות בלבד"):null
      ),
      ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:flow.varE>0?8:0}},
        [["פתיחה",flow.open],["הכנסות",flow.fixI],["קבועות",flow.fixE],["סגירה",flow.close]].map(function(r){
          return ce("div",{key:r[0],style:{textAlign:"center"}},
            ce("div",{style:{fontSize:9,opacity:.7,fontWeight:700,textTransform:"uppercase",marginBottom:2}},r[0]),
            ce("div",{style:{fontSize:13,fontWeight:800}},fmt(Math.round(r[1])))
          );
        })
      ),
      flow.varE>0?ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,paddingTop:8,borderTop:"1px solid rgba(255,255,255,0.2)"}},
        [["משתנות",flow.varE],["פער",flow.net],["ירידה",flow.varE]].slice(0,2).map(function(r){
          return ce("div",{key:r[0],style:{textAlign:"center"}},
            ce("div",{style:{fontSize:9,opacity:.7,fontWeight:700,textTransform:"uppercase",marginBottom:2}},r[0]),
            ce("div",{style:{fontSize:13,fontWeight:800,color:r[0]==="פער"&&r[1]<0?"#FCA5A5":r[0]==="פער"&&r[1]>=0?"#6EE7B7":"#fff"}},fmt(Math.round(r[1])))
          );
        })
      ):null
    ):null,
    // Empty month info
    flow&&!flow.hasData?ce("div",{style:{background:"#EEF2FF",border:"1px solid #C7D2FE",borderRadius:12,padding:"10px 14px",marginBottom:12,display:"flex",gap:8}},
      ce("span",{style:{fontSize:14}},"ℹ️"),
      ce("div",{style:{fontSize:11,color:"#4338CA",lineHeight:1.7}},
        ce("strong",null,"חודש ריק")," — היתרה מחושבת לפי קבועות ועוברת לחודש הבא אוטומטית."
      )
    ):null,
    // Fixed expenses (read-only)
    ce("div",{style:{background:"#fff",borderRadius:14,padding:"12px 14px",marginBottom:12,boxShadow:SH}},
      ce("div",{style:{fontSize:13,fontWeight:800,color:BL,marginBottom:8}},"📋 הוצ׳ קבועות — אוטומטי"),
      (exp[selA[0]]||[]).length===0?ce("div",{style:{fontSize:12,color:"#aaa"}},"אין הוצ׳ קבועות"):
      (exp[selA[0]]||[]).map(function(e){
        var cat=CATS.find(function(c){return c.id===e.category;})||{i:"📦"};
        return ce("div",{key:e.id,style:{display:"flex",justifyContent:"space-between",fontSize:12,padding:"4px 0",borderBottom:"1px solid #f5f5f5"}},
          ce("span",{style:{color:"#666"}},cat.i+" "+e.name),
          ce("span",{style:{fontWeight:700,color:BL}},fmt(Math.round(mAmt(e))))
        );
      }),
      ce("div",{style:{display:"flex",justifyContent:"space-between",fontSize:13,fontWeight:800,marginTop:8,paddingTop:8,borderTop:"2px solid "+BL+"22"}},
        ce("span",null,"סה\"כ"),ce("span",{style:{color:BL}},fmt(Math.round(fixTot)))
      )
    ),
    // Variable expenses
    ce("div",{style:{background:"#FFFBEB",borderRadius:14,border:"1.5px solid #FDE68A",padding:"12px 14px",marginBottom:12}},
      ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}},
        ce("div",{style:{fontSize:13,fontWeight:800,color:"#92400E"}},"📦 הוצ׳ משתנות"),
        ce("button",{onClick:function(){addIt("");},style:{padding:"6px 14px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:12,background:"#D97706",color:"#fff"}},"+ הוסף")
      ),
      ce("div",{style:{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}},
        sc.map(function(s){return ce("button",{key:s,onClick:function(){addIt(s);},style:{padding:"4px 10px",borderRadius:14,border:"1px solid #FDE68A",background:"#FEF3C7",color:"#92400E",cursor:"pointer",fontFamily:F,fontSize:11,fontWeight:600}},s);})
      ),
      (vd.items||[]).length===0?ce("div",{style:{textAlign:"center",color:"#B45309",fontSize:12,padding:"8px 0"}},'לחצי קיצור דרך או "+ הוסף"'):null,
      (vd.items||[]).map(function(it){
        return ce("div",{key:it.id,style:{background:"#fff",borderRadius:12,padding:"11px 13px",marginBottom:8,border:"1.5px solid #FDE68A"}},
          ce("div",{style:{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap"}},
            ce("input",{value:it.name,onChange:function(e){updIt(it.id,"name",e.target.value);},placeholder:"שם הוצאה",style:Object.assign({},sI,{flex:"1 1 100px",minWidth:0,background:"#f8f9fa"})}),
            ce("button",{onClick:function(){delT[1](it);},style:{padding:"7px 14px",borderRadius:9,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontSize:13,fontWeight:700,flexShrink:0}},"🗑 מחק")
          ),
          ce("div",{style:{display:"flex",alignItems:"center",gap:8}},
            ce("span",{style:{fontSize:13,color:"#92400E",fontWeight:700,flexShrink:0}},"₪"),
            ce("input",{type:"number",value:it.amount||"",onChange:function(e){updIt(it.id,"amount",parseFloat(e.target.value)||0);},placeholder:"0",style:Object.assign({},sI,{flex:1,fontSize:17,fontWeight:900,color:"#92400E",border:"1.5px solid #FDE68A",minWidth:0})}),
            (it.amount||0)>0?ce("span",{style:{fontWeight:900,color:"#92400E",background:"#FEF3C7",padding:"5px 12px",borderRadius:8,flexShrink:0}},fmt(parseFloat(it.amount)||0)):null
          )
        );
      }),
      (vd.items||[]).length>0?ce("div",{style:{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:13,marginTop:8,paddingTop:8,borderTop:"2px solid #FDE68A",color:"#92400E"}},
        ce("span",null,"סה\"כ משתנות"),ce("span",null,fmt(Math.round(varT)))
      ):null
    ),
    delT[0]?ce(Confirm,{name:delT[0].name||"הוצאה זו",onConfirm:function(){doDelIt(delT[0].id);},onCancel:function(){delT[1](null);}}):null
  );
}

// ── EXPENSES ──────────────────────────────────────────────────────────────────
function ExpTab(props){
  var accs=props.accs,exp=props.exp,setExp=props.setExp;
  var selA=useState(accs[0].id), q=useState(""), delT=useState(null);
  var acc=accs.find(function(a){return a.id===selA[0];})||accs[0];
  var rows=exp[selA[0]]||[];
  var ql=q[0].toLowerCase();
  var filtered=ql?rows.filter(function(e){var cat=CATS.find(function(c){return c.id===e.category;})||{l:""};return (e.name||"").toLowerCase().includes(ql)||cat.l.toLowerCase().includes(ql);}):rows;
  var total=rows.reduce(function(s,e){return s+mAmt(e);},0);
  function add(){var ne=Object.assign({},exp);ne[selA[0]]=(ne[selA[0]]||[]).concat([{id:uid(),name:"הוצאה חדשה",category:"housing",amount:0,qty:1,freq:"monthly",note:""}]);setExp(ne);}
  function del(id){var ne=Object.assign({},exp);ne[selA[0]]=(ne[selA[0]]||[]).filter(function(e){return e.id!==id;});setExp(ne);delT[1](null);}
  function upd(id,f,v){var ne=Object.assign({},exp);ne[selA[0]]=(ne[selA[0]]||[]).map(function(e){return e.id===id?Object.assign({},e,{[f]:v}):e;});setExp(ne);}
  var sI={fontFamily:F,fontSize:13,padding:"6px 8px",borderRadius:8,border:"1px solid #ddd",direction:"rtl",background:"#fff",width:"100%",boxSizing:"border-box"};
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    ce("div",{style:{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}},
      accs.map(function(a){return ce("button",{key:a.id,onClick:function(){selA[1](a.id);q[1]("");},style:{padding:"7px 16px",borderRadius:20,border:"1.5px solid "+(selA[0]===a.id?a.color:"#ddd"),background:selA[0]===a.id?a.grad:"#fff",color:selA[0]===a.id?"#fff":"#666",fontFamily:F,fontWeight:700,fontSize:12,cursor:"pointer"}},a.emoji+" "+a.name);})
    ),
    ce("div",{style:{background:acc.grad,borderRadius:14,padding:"13px 16px",marginBottom:12,color:"#fff",display:"flex",justifyContent:"space-between",alignItems:"center"}},
      ce("div",null,
        ce("div",{style:{fontSize:11,opacity:.8,marginBottom:2}},"הוצ׳ קבועות — "+acc.name),
        ce("div",{style:{fontSize:24,fontWeight:900}},fmt(Math.round(total))+"/חודש")
      ),
      ce("div",{style:{textAlign:"left",opacity:.8}},
        ce("div",{style:{fontSize:9,fontWeight:600,textTransform:"uppercase",marginBottom:2}},rows.length+" פריטים"),
        ce("div",{style:{fontSize:11,fontWeight:700}},fmt(Math.round(total*12))+"/שנה")
      )
    ),
    ce("div",{style:{display:"flex",gap:8,marginBottom:10}},
      ce("input",{value:q[0],onChange:function(e){q[1](e.target.value);},placeholder:"🔍 חפש...",style:Object.assign({},sI,{flex:1,borderColor:q[0]?BL:"#ddd"})}),
      q[0]?ce("button",{onClick:function(){q[1]("");},style:{padding:"8px 10px",borderRadius:10,border:"1px solid #ddd",background:"#fff",color:"#666",cursor:"pointer",fontFamily:F,fontSize:12}},"✕"):null,
      ce("button",{onClick:add,style:{padding:"8px 14px",borderRadius:10,border:"none",cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:13,background:acc.color,color:"#fff",flexShrink:0}},"+ הוסף")
    ),
    ce("div",{style:{fontSize:10,color:BL,marginBottom:10,padding:"5px 10px",background:"#EEF2FF",borderRadius:8,display:"inline-block"}},"✎ לחצי על שדה לעריכה"),
    filtered.length===0?ce("div",{style:{background:"#fff",borderRadius:14,padding:32,textAlign:"center",color:"#aaa"}},
      ce("div",{style:{fontSize:28,marginBottom:8}},"💳"),
      ce("div",null,q[0]?"אין תוצאות":'לחצי "+ הוסף"')
    ):null,
    ce("div",{style:{display:"flex",flexDirection:"column",gap:8}},
      filtered.map(function(e){
        var cat=CATS.find(function(c){return c.id===e.category;})||{i:"📦",l:"אחר"};
        return ce("div",{key:e.id,style:{background:"#fff",borderRadius:14,padding:"12px 14px",boxShadow:SH}},
          ce("div",{style:{display:"flex",alignItems:"center",gap:9,marginBottom:10}},
            ce("span",{style:{fontSize:20,flexShrink:0}},cat.i),
            ce("input",{value:e.name,onChange:function(ev){upd(e.id,"name",ev.target.value);},style:Object.assign({},sI,{flex:1,fontWeight:700})}),
            ce("div",{style:{textAlign:"center",minWidth:65,flexShrink:0}},
              ce("div",{style:{fontSize:15,fontWeight:900,color:acc.color}},fmt(Math.round(mAmt(e)))),
              ce("div",{style:{fontSize:9,color:"#aaa"}},"לחודש")
            )
          ),
          ce("div",{style:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}},
            ce("div",{style:{background:"#f8f9fa",borderRadius:8,padding:"8px"}},
              ce("div",{style:{fontSize:9,color:"#666",marginBottom:4}},"קטגוריה"),
              ce("select",{value:e.category,onChange:function(ev){upd(e.id,"category",ev.target.value);},style:Object.assign({},sI,{background:"#f8f9fa"})},
                CATS.map(function(c){return ce("option",{key:c.id,value:c.id},c.i+" "+c.l);}))
            ),
            ce("div",{style:{background:"#f8f9fa",borderRadius:8,padding:"8px"}},
              ce("div",{style:{fontSize:9,color:"#666",marginBottom:4}},"תדירות"),
              ce("select",{value:e.freq||"monthly",onChange:function(ev){upd(e.id,"freq",ev.target.value);},style:Object.assign({},sI,{background:"#f8f9fa"})},
                ce("option",{value:"monthly"},"חודשי"),
                ce("option",{value:"quarterly"},"רבעוני"),
                ce("option",{value:"yearly"},"שנתי")
              )
            ),
            ce("div",{style:{background:"#f8f9fa",borderRadius:8,padding:"8px"}},
              ce("div",{style:{fontSize:9,color:"#666",marginBottom:4}},"סכום ₪"),
              ce("input",{type:"number",value:e.amount||"",onChange:function(ev){upd(e.id,"amount",parseFloat(ev.target.value)||0);},placeholder:"0",style:Object.assign({},sI,{fontWeight:700,background:"#f8f9fa"})})
            ),
            ce("div",{style:{display:"flex",alignItems:"flex-end",justifyContent:"flex-end",padding:"4px 0"}},
              ce("button",{onClick:function(){delT[1](e);},style:{padding:"8px 16px",borderRadius:9,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontSize:13,fontWeight:700}},"🗑 מחק")
            )
          )
        );
      })
    ),
    filtered.length>0?ce("button",{onClick:add,style:{width:"100%",marginTop:10,padding:"11px",borderRadius:12,border:"1.5px dashed "+acc.color,background:acc.color+"08",color:acc.color,cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:13}},"+ הוסף הוצאה קבועה"):null,
    delT[0]?ce(Confirm,{name:delT[0].name,onConfirm:function(){del(delT[0].id);},onCancel:function(){delT[1](null);}}):null
  );
}

// ── INCOME ────────────────────────────────────────────────────────────────────
function IncTab(props){
  var accs=props.accs,inc=props.inc,setInc=props.setInc;
  var selA=useState(accs[0].id), delT=useState(null);
  var acc=accs.find(function(a){return a.id===selA[0];})||accs[0];
  var rows=inc[selA[0]]||[];
  var total=rows.reduce(function(s,r){return s+(parseFloat(r.amount)||0);},0);
  function add(){var ni=Object.assign({},inc);ni[selA[0]]=(ni[selA[0]]||[]).concat([{id:uid(),name:"הכנסה חדשה",amount:0,type:"salary",note:""}]);setInc(ni);}
  function del(id){var ni=Object.assign({},inc);ni[selA[0]]=(ni[selA[0]]||[]).filter(function(r){return r.id!==id;});setInc(ni);delT[1](null);}
  function upd(id,f,v){var ni=Object.assign({},inc);ni[selA[0]]=(ni[selA[0]]||[]).map(function(r){return r.id===id?Object.assign({},r,{[f]:v}):r;});setInc(ni);}
  var sI={fontFamily:F,fontSize:13,padding:"7px 10px",borderRadius:8,border:"1px solid #ddd",direction:"rtl",background:"#fff",width:"100%",boxSizing:"border-box"};
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    ce("div",{style:{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}},
      accs.map(function(a){return ce("button",{key:a.id,onClick:function(){selA[1](a.id);},style:{padding:"7px 16px",borderRadius:20,border:"1.5px solid "+(selA[0]===a.id?a.color:"#ddd"),background:selA[0]===a.id?a.grad:"#fff",color:selA[0]===a.id?"#fff":"#666",fontFamily:F,fontWeight:700,fontSize:12,cursor:"pointer"}},a.emoji+" "+a.name);})
    ),
    ce("div",{style:{background:acc.grad,borderRadius:14,padding:"14px 18px",marginBottom:14,color:"#fff"}},
      ce("div",{style:{fontSize:11,opacity:.8,marginBottom:4}},"הכנסות חודשיות — "+acc.name),
      ce("div",{style:{fontSize:28,fontWeight:900}},fmt(total))
    ),
    ce("button",{onClick:add,style:{width:"100%",padding:"11px",borderRadius:12,border:"1.5px dashed "+acc.color,background:acc.color+"08",color:acc.color,cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:13,marginBottom:12}},"+ הוסף הכנסה"),
    rows.length===0?ce("div",{style:{background:"#fff",borderRadius:14,padding:32,textAlign:"center",color:"#aaa"}},ce("div",{style:{fontSize:28,marginBottom:8}},"💵"),'לחצי "+ הוסף הכנסה"'):null,
    ce("div",{style:{display:"flex",flexDirection:"column",gap:8}},
      rows.map(function(r){
        return ce("div",{key:r.id,style:{background:"#fff",borderRadius:14,padding:"13px 15px",boxShadow:SH}},
          ce("div",{style:{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}},
            ce("input",{value:r.name,onChange:function(e){upd(r.id,"name",e.target.value);},style:Object.assign({},sI,{flex:1,fontWeight:700})}),
            ce("div",{style:{display:"flex",alignItems:"center",gap:6}},
              ce("span",{style:{fontSize:13,fontWeight:800,color:GR}},fmt(parseFloat(r.amount)||0)),
              ce("input",{type:"number",value:r.amount||"",onChange:function(e){upd(r.id,"amount",parseFloat(e.target.value)||0);},placeholder:"0",style:Object.assign({},sI,{width:90,fontWeight:800,color:GR})})
            )
          ),
          ce("div",{style:{display:"flex",justifyContent:"flex-end"}},
            ce("button",{onClick:function(){delT[1](r);},style:{padding:"6px 14px",borderRadius:8,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontSize:12,fontWeight:700}},"🗑 מחק")
          )
        );
      })
    ),
    delT[0]?ce(Confirm,{name:delT[0].name,onConfirm:function(){del(delT[0].id);},onCancel:function(){delT[1](null);}}):null
  );
}

// ── SAVINGS ───────────────────────────────────────────────────────────────────
function SavTab(props){
  var accs=props.accs,sav=props.sav||[],setSav=props.setSav;
  var delT=useState(null);
  function upd(id,f,v){setSav(sav.map(function(s){return s.id===id?Object.assign({},s,{[f]:v}):s;}));}
  function del(id){setSav(sav.filter(function(s){return s.id!==id;}));delT[1](null);}
  function add(accId){setSav(sav.concat([{id:uid(),accId:accId,name:"חיסכון חדש",current:0,target:0,monthly:0}]));}
  var tC=sav.reduce(function(s,x){return s+(parseFloat(x.current)||0);},0);
  var tT=sav.reduce(function(s,x){return s+(parseFloat(x.target)||0);},0);
  var sI={fontFamily:F,fontSize:13,padding:"6px 9px",borderRadius:8,border:"1px solid #ddd",direction:"rtl",background:"#fff",width:"100%",boxSizing:"border-box"};
  return ce("div",{style:{padding:"16px 16px 40px",fontFamily:F,direction:"rtl"}},
    ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,marginBottom:16}},
      [{l:"נצבר",v:fmt(tC),c:GR},{l:"הפקדה/חודש",v:fmt(sav.reduce(function(s,x){return s+(parseFloat(x.monthly)||0);},0)),c:BL},
       {l:"יעד כולל",v:fmt(tT),c:"#8B5CF6"},{l:"% השגה",v:tT>0?Math.round(tC/tT*100)+"%":"—",c:GO}
      ].map(function(item){
        return ce("div",{key:item.l,style:{background:"#fff",border:"1px solid "+item.c+"22",borderRadius:12,padding:"10px 8px",textAlign:"center",boxShadow:SH}},
          ce("div",{style:{fontSize:9,color:"#aaa",fontWeight:700,textTransform:"uppercase",marginBottom:3}},item.l),
          ce("div",{style:{fontSize:14,fontWeight:900,color:item.c}},item.v)
        );
      })
    ),
    accs.map(function(acc){
      var asvs=sav.filter(function(s){return s.accId===acc.id;});
      return ce("div",{key:acc.id,style:{background:"#fff",borderRadius:16,overflow:"hidden",marginBottom:12,boxShadow:SH}},
        ce("div",{style:{background:acc.grad,padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",color:"#fff"}},
          ce("span",{style:{fontSize:13,fontWeight:800}},acc.emoji+" "+acc.name),
          ce("div",{style:{display:"flex",alignItems:"center",gap:10}},
            ce("span",{style:{fontSize:11,opacity:.8,fontWeight:600}},fmt(asvs.reduce(function(s,x){return s+(parseFloat(x.current)||0);},0))+" נצבר"),
            ce("button",{onClick:function(){add(acc.id);},style:{padding:"4px 10px",borderRadius:8,border:"none",cursor:"pointer",background:"rgba(255,255,255,0.2)",color:"#fff",fontFamily:F,fontWeight:700,fontSize:11}},"+ הוסף")
          )
        ),
        asvs.length===0?ce("div",{style:{padding:16,textAlign:"center",color:"#aaa",fontSize:12}},'לחצי "+ הוסף"'):null,
        asvs.map(function(sv){
          var cur=parseFloat(sv.current)||0, tgt=parseFloat(sv.target)||0;
          var pct=tgt>0?Math.min(Math.round(cur/tgt*100),100):0;
          return ce("div",{key:sv.id,style:{padding:"12px 16px",borderBottom:"1px solid #f5f5f5"}},
            ce("div",{style:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}},
              ce("input",{value:sv.name,onChange:function(e){upd(sv.id,"name",e.target.value);},style:Object.assign({},sI,{flex:1,fontWeight:700,marginLeft:8})}),
              ce("button",{onClick:function(){delT[1](sv);},style:{padding:"5px 10px",borderRadius:7,border:"none",background:RD,color:"#fff",cursor:"pointer",fontFamily:F,fontSize:11,fontWeight:700}},"🗑 מחק")
            ),
            ce("div",{style:{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6,marginBottom:tgt>0?8:0}},
              [{l:"💰 נוכחי",f:"current",c:GR},{l:"🎯 יעד",f:"target",c:"#8B5CF6"},{l:"📅 חודשי",f:"monthly",c:BL}].map(function(it){
                return ce("div",{key:it.f},
                  ce("div",{style:{fontSize:9,color:"#aaa",marginBottom:2}},it.l),
                  ce("input",{type:"number",value:sv[it.f]||"",onChange:function(e){upd(sv.id,it.f,parseFloat(e.target.value)||0);},style:Object.assign({},sI,{fontWeight:700,color:it.c})})
                );
              })
            ),
            tgt>0?ce("div",null,
              ce("div",{style:{display:"flex",justifyContent:"space-between",fontSize:10,color:"#666",marginBottom:4}},
                ce("span",null,"התקדמות"),ce("span",{style:{fontWeight:700,color:acc.color}},pct+"%")
              ),
              ce("div",{style:{height:8,borderRadius:10,background:"#eee",overflow:"hidden"}},
                ce("div",{style:{height:"100%",borderRadius:10,background:pct>=100?GR:pct>=75?GO:acc.color,width:pct+"%",transition:"width .5s"}})
              ),
              (parseFloat(sv.monthly)||0)>0&&tgt>cur?ce("div",{style:{fontSize:10,color:"#aaa",marginTop:3}},"יעד בעוד ≈ "+Math.ceil((tgt-cur)/(parseFloat(sv.monthly)||1))+" חודשים"):null
            ):null
          );
        })
      );
    }),
    delT[0]?ce(Confirm,{name:delT[0].name,onConfirm:function(){del(delT[0].id);},onCancel:function(){delT[1](null);}}):null
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App(){
  var d=loadData(), def=buildDefaults();
  var exp=useState(d&&d.exp?d.exp:def.exp);
  var inc=useState(d&&d.inc?d.inc:def.inc);
  var mon=useState(d&&d.mon?d.mon:{});
  var fut=useState(d&&d.fut?d.fut:def.fut);
  var sav=useState(d&&d.sav?d.sav:def.sav);
  var tab=useState("summary");
  var accs=ACCS;
  useEffect(function(){saveData({exp:exp[0],inc:inc[0],mon:mon[0],fut:fut[0],sav:sav[0]});},[exp[0],inc[0],mon[0],fut[0],sav[0]]);
  var hI=accs.reduce(function(s,a){return s+(inc[0][a.id]||[]).reduce(function(s2,r){return s2+(parseFloat(r.amount)||0);},0);},0);
  var hE=accs.reduce(function(s,a){return s+(exp[0][a.id]||[]).reduce(function(s2,e){return s2+mAmt(e);},0);},0);
  var hB=hI-hE;
  var upcoming=fut[0].filter(function(f){return (f.year||NOW_Y)===NOW_Y&&(f.month===NOW_M||f.month===NOW_M+1);}).length;
  var TABS=[{id:"summary",l:"🏦 סיכום"},{id:"monthly",l:"📝 הזנה חודשית"},{id:"year",l:"📅 לוח שנתי"},{id:"income",l:"💵 הכנסות"},{id:"expenses",l:"💳 הוצאות"},{id:"savings",l:"📈 חסכונות"}];
  return ce("div",{style:{fontFamily:F,direction:"rtl",background:"#F0F4FF",minHeight:"100vh"}},
    ce("div",{style:{background:"#fff",borderBottom:"1px solid #E2E8F0",position:"sticky",top:0,zIndex:200,boxShadow:"0 1px 8px rgba(99,102,241,0.08)"}},
      ce("div",{style:{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px"}},
        ce("div",{style:{display:"flex",alignItems:"center",gap:9}},
          ce("div",{style:{width:34,height:34,borderRadius:10,background:"linear-gradient(135deg,#6366F1,#8B5CF6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17}},"💰"),
          ce("div",null,
            ce("div",{style:{fontSize:14,fontWeight:800}},"תקציב ביתי"),
            hI>0?ce("div",{style:{fontSize:10,color:hB>=0?GR:RD,fontWeight:700}},(hB>=0?"+":"")+fmt(Math.round(hB))+"/חודש"):null
          )
        ),
        ce("button",{onClick:function(){doExport(accs,exp[0],inc[0],fut[0],sav[0],mon[0]);},style:{padding:"7px 12px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:F,fontWeight:700,fontSize:11,background:GR,color:"#fff"}},"📥 אקסל")
      ),
      ce("div",{style:{display:"flex",gap:0,overflowX:"auto",padding:"0 12px"}},
        TABS.map(function(t){
          return ce("button",{key:t.id,onClick:function(){tab[1](t.id);},style:{padding:"10px 13px",border:"none",borderBottom:"2.5px solid "+(tab[0]===t.id?BL:"transparent"),cursor:"pointer",fontFamily:F,fontWeight:tab[0]===t.id?700:400,fontSize:12,color:tab[0]===t.id?BL:"#666",background:"transparent",whiteSpace:"nowrap",transition:"all .15s",position:"relative"}},t.l,
            t.id==="year"&&upcoming>0?ce("span",{style:{position:"absolute",top:5,right:0,background:GO,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:800}},upcoming):null);
        })
      )
    ),
    ce("div",{style:{maxWidth:800,margin:"0 auto"}},
      tab[0]==="summary"?ce(SummaryTab,{accs:accs,exp:exp[0],inc:inc[0],mon:mon[0],fut:fut[0]}):null,
      tab[0]==="monthly"?ce(MonthTab,{accs:accs,exp:exp[0],inc:inc[0],mon:mon[0],setMon:mon[1],fut:fut[0]}):null,
      tab[0]==="year"?ce(YearTab,{accs:accs,exp:exp[0],inc:inc[0],mon:mon[0],setMon:mon[1],fut:fut[0],setFut:fut[1]}):null,
      tab[0]==="income"?ce(IncTab,{accs:accs,inc:inc[0],setInc:inc[1]}):null,
      tab[0]==="expenses"?ce(ExpTab,{accs:accs,exp:exp[0],setExp:exp[1]}):null,
      tab[0]==="savings"?ce(SavTab,{accs:accs,sav:sav[0],setSav:sav[1]}):null
    )
  );
}
