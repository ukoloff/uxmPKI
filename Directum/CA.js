//
// Update Certificates
//

var URL='https://ekb.ru/omz/abook/users.ca/';
var Server='Directum';
var DB='Sandbox';
//var DB='Directum';

WScript.Echo('Getting CSV...');
var Ajax=new ActiveXObject("Msxml2.XMLHTTP");
Ajax.open('GET', URL+'?q=r@&sort=c&as=csv', false);
Ajax.send();
var Q=namedCSV(parseCSV(Ajax.responseText));

WScript.Echo('Connecting MS SQL...');
var SQL=goSQL();
SQL.CommandText="	\
Select	Count(*)	\
From	\
 mbUser U, mbAnalit P, mbVidAn V, MBAnValR2 T	\
Where	\
 U.NeedEncode='W' And U.UserKod=P.Dop And	\
 P.Vid=V.Vid And V.Kod='ПОЛ' And T.Analit=P.Analit	\
 And U.UserLogin=?	\
 And T.SoderT2=?	\
Select	\
 UserLogin, UserKod, UserName, P.Analit, P.Kod	\
From	\
 mbUser U, mbAnalit P, mbVidAn V	\
Where	\
 U.NeedEncode='W' And U.UserKod=P.Dop And	\
 P.Vid=V.Vid And V.Kod='ПОЛ'	\
 And U.UserLogin=?	\
";

WScript.Echo('Поиск неустановленных сертификатов...');

var Dir, sys;

for(var i in Q)
{
 var X=Q[i];
 if(X.Revoke) continue;
 SQL(0)=X.u;
 SQL(1)=X.SHA1=X.SHA1.replace(/\W/g, '');
 SQL(2)=X.u;
 var Rs=SQL.Execute();
 if(Rs.EOF) continue;
 if(Rs(0).value>0) continue;
 Rs=Rs.NextRecordset();
 if(Rs.EOF) continue;
 X.cn=X.subj.replace(/.*\/CN=/i, '').replace(/\/.*/, '');
 for(var E=new Enumerator(Rs.Fields); !E.atEnd(); E.moveNext())
   X[E.item().name]=E.item().value;
 WScript.Echo(X.u+'\t'+X.UserKod+'\t'+X.cn+'\t'+X.UserName);
 
 if(!Dir)Dir=goDirectum();
 if(!sys) sys=getSys();

 Ajax.open('GET', URL+'?n='+X.id, false);
 Ajax.send();
 var N1=sys.tmp();
 var FC=sys.fso.CreateTextFile(N1, true);
 FC.Write(Ajax.responseText);
 FC.Close();
 var N2=sys.tmp();
 sys.sh.Run('"C:/Program Files/OpenSSL/openssl.exe" x509 -in "'+
	N1+'" -outform der -out "'+N2+'"', 0, true);
 sys.fso.DeleteFile(N1);

 try{

 var POL=Dir.ReferencesFactory.ПОЛ.GetObjectByCode(X.Kod);
 var CER=POL.DetailDataSet(2);
 var cerCount=CER.RecordCount;

// Работает в Directum 4.6.1; В Directum 4.7 юзать Dir.ScriptFactory
 var Events = CER.Events;
 Events.AddCheckPoint();
 Events.Events(9/*dseBeforeInsert*/).Enabled = false;
 CER.Append();
 Events.ReleaseCheckPoint();

 CER.ISBStartObjectName='{B1B27433-D685-47F8-8500-CF9525407145}';
 CER.СтрокаТ2=X.cn;
 CER.СодержаниеТ2=X.SHA1;
 CER.Requisites('ТекстТ2').LoadFromFile(N2);
 CER.ISBCertificateInfo=X.UserName;
 CER.ISBCertificateType=cerCount? 'ЭЦП':'ЭЦП и шифрование';
 CER.ISBDefaultCert=cerCount? 'Нет':'Да';
 CER.СостояниеТ2='Действующая';
 POL.Save();

 }catch(e){
 WScript.Echo('Error: '+e.message);
 }

 sys.fso.DeleteFile(N2);
}

function goSQL()
{
 var Conn=new ActiveXObject("ADODB.Connection");
 Conn.Provider='SQLOLEDB';
 Conn.Open("Integrated Security=SSPI;Data Source="+Server);
 Conn.DefaultDatabase=DB;
 var Cmd=new ActiveXObject("ADODB.Command");
 Cmd.ActiveConnection=Conn;
 return Cmd;
}

function parseCSV(S)
{
 var L='', F=[], All=[], q=0, eol=0;
 while(S.length)
 {
  (q? /""?|$/ : /;|"|\r\n?|\n|$/).test(S);
  L+=RegExp.leftContext;
  S=RegExp.rightContext;
  if(q)
   switch(RegExp.lastMatch)
   {
    case '"': q=0; continue; 
    case '""': L+='"'; continue;
    case '': eol=1;
   }
  else
  {
   if('"'==RegExp.lastMatch) { q=1; continue; }
   eol=';'!=RegExp.lastMatch;
  }
  F.push(L); L='';
  if(!eol) continue;
  All.push(F); F=[];
 }
 return All;
}

function namedCSV(CSV)
{
 for(var i=CSV.length-1; i>=0; i--)
   if((CSV[i].length>1)||(CSV[i][0].length>0)) break;
 CSV.length=i+1;
 if(!i) return CSV;
 var F=CSV.shift(), R=[];
 while(D=CSV.shift())
 {
  var R2={};
  for(var i in F) R2[F[i]]=D[i];
  R.push(R2);
 }
 return R;
}

function goDirectum()
{
 var lp=new ActiveXObject("SBLogon.LoginPoint");
 var D=lp.GetApplication("ServerName="+Server+";DBName="+DB+";IsOSAuth=1");
 return D;
}

function rnd(N)
{
 for(var S=''; S.length<(N||12); )
 {
  var n=Math.floor(62*Math.random());
  S+=String.fromCharCode('Aa0'.charCodeAt(n/26)+n%26);
 }
 return S;
}

function getSys()
{
 var R={};
 R.fso=new ActiveXObject("Scripting.FileSystemObject");
 R.sh=new ActiveXObject("WScript.Shell")
 R.tmpPath=R.sh.ExpandEnvironmentStrings('%TEMP%/');
 R.tmp=function()
 {
  do var n=this.tmpPath+rnd(); while(this.fso.FileExists(n));
  return n;
 }
 return R;
}
