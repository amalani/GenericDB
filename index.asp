<%@LANGUAGE="JAVASCRIPT"%><!--#include file="adojavas.inc"--><%

gendb = {
    author   : 'Abhishek Malani',
    url      : 'http://www.amalani.net/',
    title    : 'Generic Database Editor',
    version  : '1',
    lastmod  : '2003/05/15',

    settings : {
        // Password
        pwd     : 'nikereebok', // Only need to modify this.
        // Relative path to database location
        path    : '../db/', // don't forget the slash at the end
        norec   : 10, // No. of records / page.

//---------P A S S W O R D    H E R E 
        maxsize : 200, // Maximum size of the memo field when viewing all records.
        sepnav  : ' &nbsp; =&gt; &nbsp;',
        script  : Request.ServerVariables("SCRIPT_NAME")
    }
};

gendb.opendb = function(file,pwd){
   conn = new ActiveXObject("ADODB.connection");
   conn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.Mappath(this.settings.path + file) );
   rs = new ActiveXObject("ADODB.Recordset");
   rs.ActiveConnection = conn;
}

gendb.closedb = function(){
   conn.close();
}

gendb.checkuser = function(){
    if( query('pwd') == this.settings.pwd ){
        Session("log") = 1;
    }

    if( Session('log') != 1){
        bar();
        Response.Write('<form action="' + this.settings.script + '" method="post">');
        Response.Write('<input type="password" name="pwd"/>\n<input type="submit" value="Login!" class="button"/>\n');
        Response.Write('</form>\n\n');
        
        call = 0;
    }
}


gendb.lstmdb = function(){
    fso = Server.CreateObject("Scripting.FileSystemObject")	
    Response.Write('<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">');
    Response.Write('<tr class="options">\n<td>Please choose a database</td>\n<td>Size</td>\n<td>Options</td></tr>\n\n');
    if(fso.FolderExists( Server.Mappath( this.settings.path ) ) ){
      fs = fso.GetFolder( Server.Mappath( this.settings.path ) );

        var e = new Enumerator(fs.Files);
        
        while(!e.atEnd()) {
            file = String( e.item() ).toLowerCase();
            var files = file;
            
            if(file.indexOf('\\')>-1){
                files	= file.substring(file.lastIndexOf('\\')+1, file.length);
            }
            
            if(files.indexOf('.mdb')>-1){
            count++;
                Response.Write('<tr class="row0">\n<td class="c">');
                Response.Write('<a href="?call=lsttbls&amp;db=' + files + '">' + files + '</a><br/>');
                objfile = fso.GetFile( Server.Mappath( this.settings.path + files) );
                Response.Write('</td>\n<td class="c">' + getsize(objfile.Size,1) + '</td>\n');
                Response.Write('<td class="c"><a href="?call=exec&amp;db=' + files + '">execute statement</a> | <a href="?call=compact&amp;db=' + files + '">compact</a> | backup</td>\n</tr>\n\n');
            }
            e.moveNext();
        }	e = null;
            if( count == 0 ){
                Response.Write('<tr class="row0">\n\t<td colspan="3" class="c">No Database in given folder, Please change the path.</td>\n</tr>');
            }
        Response.Write('</table>');
   }else{ //folder check
      Response.Write('<p>The given path could not be found, please edit <pre>' + this.settings.script + '</pre>&nbsp;and change the path.</p>');
   } //folder check
}


gendb.lsttbls = function(){
%>
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td>Tables</td>
<td><a href=""><img src="img/new.gif" alt=""/> Add a Table</a></td>
</tr>

<%
list = conn.OpenSchema( adSchemaTables );

while(!list.eof){
    if(list("table_type")=="TABLE"){
        Response.Write('<tr class="row' + count + '">\n');
        Response.Write('<td class="c"><a href="?call=lstrec&amp;db=' + db + '&amp;table='+ list("table_name") + '">' + list("table_name") + '</a></td>\n');
        Response.Write('<td class="c"><a href="?call=tblprop&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/editf.gif" alt=""/> edit</a>\n &nbsp; &nbsp; <a href="?call=clrtbl&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/empty.gif" alt=""/> clear</a>\n &nbsp; &nbsp; <a href="?call=deltbl&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/del.gif" alt=""/> del</a></td>\n');
        Response.Write('</tr>\n\n');
    count = (count==0) ? 1 : 0;
    }

    list.movenext();
}

list.close();

Response.Write('</table>\n\n');
}



//-----------------------------------------OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOPPPPPP ENDS

function replacejs(text){
   text = text.replace("'","\\'");
   text = text.replace('"','\\"');
   return text;
}

function dblink(db){
   return '[<a href="?call=lsttbls&amp;db=' + db + '">' + db + '</a>]';
}

function tbllink(db,table,page){
if(page){l = '&amp;page=' + page;}else{l = '';}
   return '[<a href="?call=lstrec&amp;db=' + db + '&amp;table=' + table + ((page) ? '&amp;page=' + page : '') + l +'">' + table + '</a>]'
}

String.prototype.makesafe = function()	{
    return this.replace(/\'/gi, "''"); //'
}

function options(opt1,opt2,p1,p2,sep1,sep2,end1,end2){
   if(p1){
        Response.Write('<div class="bar mtop15"><p>');
    }
   if(end1){
        Response.Write('[');
    }
       Response.Write('<a href="' + script + opt1 + '">' + opt2 + '</a>');
   if(sep1){
        Response.Write(' - ');
    }
   if(end2){
        Response.Write(']');
    }
   if(sep2){
        Response.Write(' &nbsp; ');
    }
   if(p2){
        Response.Write('</p></div>\n\n');
    }
}

function options2(text,p1,p2,sep1,sep2,end1,end2){
   if(p1){
        Response.Write('<div class="bar mtop15"><p>');
    }
   if(end1){
        Response.Write('[');
    }
    Response.Write(text);
   if(sep1){
        Response.Write(' - ');
    }
   if(end2){
        Response.Write(']');
    }
   if(sep2){
        Response.Write(' &nbsp; ');
    }
   if(p2){
        Response.Write('</p></div>\n\n');
    }
}


function getsize(text,size){
   text = parseInt(text);
   var sizecount = 0;
   var sizearray = new Array("Bytes", "KBytes", "MBytes", "GBytes");
   while(text>1024){
      text = text / 1024;
      sizecount++;
   }
   return Math.round(text) + ((size) ? ' ' + sizearray[sizecount] : '');
}

function gettype(txt){
switch ( parseInt(txt) ){
   case 0:
      return "adEmpty"
   case 16:
      return "adTinyInt"
   case 2:
      return "adSmallInt"
   case 3:
        return "adInteger"
   case 20:
      return "adBigInt"
   case 17:
      return "adUnsignedTinyInt"
   case 18:
      return "adUnsignedSmallInt"
   case 19:
      return "adUnsignedInt"
   case 21:
      return "adUnsignedBigInt"
   case 4:
      return "adSingle"
   case 5:
      return "adDouble"
   case 6:
      return "adCurrency"
   case 14:
      return "adDecimal"
   case 131:
      return "adNumeric"
   case 11:
      return "adBoolean"
   case 10:
      return "adError"
   case 132:
      return "adUserDefined"
   case 12:
      return "adVariant"
   case 9:
      return "adIDispatch"
   case 13:
      return "adIUnknown"
   case 72:
      return "adGUID"
   case 7:
      return "adDate"
   case 133:
      return "adDBDate"
   case 134:
      return "adDBTime"
   case 135:
      return "adDBTimeStamp"
   case 8:
      return "adBSTR"
   case 129:
      return "adChar"
   case 200:
      return "adVarChar"
   case 201:
      return "adLongVarChar"
   case 130:
      return "adWChar"
   case 202:
      return "adVarWChar"
   case 203:
      return "adLongVarWChar"
   case 128:
      return "adBinary"
   case 204:
      return "adVarBinary"
   case 205:
      return "adLongVarBinary"
   default:
      return "Undefined <em>by</em> ActiveX Data Objects"
    }
}



function getsimpletype(text){
   switch (parseInt( text )){
      case adInteger:
      case adTinyInt:
      case adSmallInt:
      case adBigInt:
      case adUnsignedTinyInt:
      case adUnsignedSmallInt:
      case adUnsignedInt:
      case adUnsignedBigInt:
      case adSingle:
      case adDouble:
      case adCurrency:
      case adNumeric:
         return 'Integer';
         break;
   
      case adDate:
      case adDBDate:
      case adDBTime:
      case adDBTimeStamp:
         return 'Date/Time';
         break;
   
      case adBoolean:
         return 'Boolean';
         break;
   
      case adLongVarChar: 
      case adLongVarWChar:
         return 'Memo';
         break;
   
      default:
         return 'Text';
         break;
   }
}

function bar(nav){
   nav = ((nav) ? nav : '');
   Response.Write('<div class="bar"><p>\n[<a href="?">Home</a>] ' + nav + '</p></div>\n\n');
}

function sTextarea(elname,elvalue){
   var elname = String(elname);
   var elvalue = String(elvalue);
   Response.Write('<textarea class="' + ((elvalue.length<200) ? 'sm' : 'bg') + '" name="' + elname + '">' + Server.HTMLEncode(elvalue) + '</textarea>');
}

function sInput(elname,elvalue){
   var elname = String(elname);
   var elvalue = String(elvalue);
   
   Response.Write('<input type="text" name="' + elname + '" value="' + Server.HTMLEncode(elvalue) + '"/>');
}


function sInputc(elname,elvalue){
   var elname = String(elname);
   var elvalue = String(elvalue);
   Response.Write('True: <input ' + ((elvalue=="true") ? 'checked="checked"' : "") + 'type="radio" class="radio" name="' + elname + '" value="True"/>');
   Response.Write('False: <input ' + ((elvalue=="false") ? 'checked="checked"' : "") + ' type="radio" class="radio" name="' + elname + '" value="False"/>');

   //   Response.Write('<input type="checkbox" name="' + elname + '" ' + ((elvalue=='true') ? 'checked="true"' : '' ) + '/>');
}


Date.prototype.getdb = function(){
   return this.getYear() + '/' + parseInt( this.getMonth()+1 ) + '/' + this.getDate() + ' ' + this.getHours() + ':' + this.getMinutes() + ':' + this.getSeconds(); 
}


function query(obj,value,int){ // (queryname, value_if_null, is_integer)
   obj = String(Request(obj));
   
   if(int){ //if querytype = integer.
      obj = parseInt( obj );
      if( isNaN(obj) ){
         if(value){ obj = value; }else{ obj = 0 };
      }
   }else{
      if( obj == "undefined" || obj == null || obj == "" ){
         if(value){ obj = value; }else{ obj = 0 };
      }
   }
   return obj;
}


//------------------------------E N D   F U N C T I O N S
%><%



//Session.Timeout = 5;

//
//
//
//          DO NOT EDIT BELOW THIS LINE.
//          
//
//


/*
----------------------------------HISTORY----------------------------------

====| DEVELOPMENTS
   =  Scratched the older version.
   =  Basic layout made
   =  lstmdb()  - db in directory
   =  lsttbl() - tables in db
   =  lstrec() - Records in table with paging
   =  viewrec() - show record
   =  clrtbl() - clear all records in table
   =  deltbl() - delete all records in table
   =  editrec() - edit record
   =  delrec() - delete record
   =  exec() - execute custom statement
   =  addtable() - insert table
   =  addrec() - add record
   =  Select box add in lstrec() so as to conserve space for big tables.
   ==\> Error in showing records from exec() removed.
    ==\> Errors in lstmdb() cuz file was not saved when system hanged.
    = tblprop() - table properties
    = 
    


---------------------------------------------------------------------------
*/




var script = Request.ServerVariables("SCRIPT_NAME");


//-----------Q U E R Y S T R I N G S
var call = query('call');
var num = query('num');
var numvalue = query('numvalue', 'asc');

var table = query('table');
var db = query('db');
if (db!=0){
   gendb.opendb( db );
}

var confirm = query('confirm');
var page = query('page', null, 1);


var autono = "id"; //changes as per table.
var orderby = autono;
var count = 0;


%><?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
   <title>Generic Database Editor</title>
   <meta http-equiv="content-type" content="application/xhtml+xml"/>
   <meta name="Author" content="Abhi Malani"/>
   <script type="text/javascript" language="javascript" src="gendb.js"></script>
   <link rel="stylesheet" type="text/css" href="css.css"/>
</head>

<body>

<div id="canvas">

<div id="header">
   <h1>generic database editor</h1>
   <h2 class="hide">by Abhi Malani - http://www.duelcom.com/malani/</h2>
</div>

<div class="bar"></div>

<%
//----------------------- C H E C K   L O G I N
gendb.checkuser();


//-------|    C O D I N G      H E R E
var nav = "";


switch ( call ){
   
   case 'logout':
    Session("log") = 0;
    Response.Redirect(script);
    break;
   
   case 'lsttbls':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + '[<a href="' + script + '?call=exec&db=' + db + '">execute custom statement</a>]';
      bar(nav);
            gendb.lsttbls();
      break;
      
   case 'addtable':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + '[Add Table]';
      bar(nav);
      addtable();
      break;
   
   case 'tblprop':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + '[Table - ' + table + ' - Properties]';
      bar(nav);
      tblprop();
      break;
      
   case 'lstrec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table);
      bar(nav);
    lstrec();
    break;
   
   case 'viewrec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table) + gendb.settings.sepnav + '[VIEWING Record # ' + numvalue + ']';
      bar(nav);
    viewrec();
    break;
   
   case 'addrec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table) + gendb.settings.sepnav + '[ADD RECORD]';
      bar(nav);
    addrec();
    break;
    
   case 'editrec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table) + gendb.settings.sepnav + '[EDITING Record # ' + numvalue + ']';;
      bar(nav);
    editrec();
    break;
    
   case 'saverec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table) + gendb.settings.sepnav + '[SAVE Record]';
      bar(nav);
    saverec();
    break;
        
   case 'delrec':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table) + gendb.settings.sepnav + ' [DELETE Record # ' + numvalue + ']';
      bar(nav);
    delrec();
    break;
   
   case 'clrtbl':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table);
      bar(nav);
    clrtbl();
    break;
   
   case 'deltbl':
      nav+= gendb.settings.sepnav + dblink( db ) + gendb.settings.sepnav + tbllink(db, table);
      bar(nav);
    deltbl();
    break;
 
   case 'compact':
      nav+= gendb.settings.sepnav + dblink( db );
      bar(nav);
        //gendb.closedb();     
    compact( db );
    break;
   
   case 'exec':
      if(db){
         nav+= gendb.settings.sepnav + '[<a href="?call=exec&amp;db=' + db + '">execute - ' + db + '</a>]'
      }
      bar(nav);
    exec();
    break;
   
   default:
      if(Session("log")==1){
            bar();
         gendb.lstmdb();
      }
    break;
}
%>



<%
//-------------L I S T    A L L   T H E   D A T A B A S E S

//-------------L I S T    A L L   T H E   T A B L E S


function lsttbls(){
%>
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td>Tables</td>
<td><a href=""><img src="img/new.gif" alt=""/> Add a Table</a></td>
</tr>

<%
list = conn.OpenSchema( adSchemaTables );

while(!list.eof){
    if(list("table_type")=="TABLE"){
        Response.Write('<tr class="row' + count + '">\n');
        Response.Write('<td class="c"><a href="?call=lstrec&amp;db=' + db + '&amp;table='+ list("table_name") + '">' + list("table_name") + '</a></td>\n');
        Response.Write('<td class="c"><a href="?call=tblprop&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/editf.gif" alt=""/> edit</a>\n &nbsp; &nbsp; <a href="?call=clrtbl&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/empty.gif" alt=""/> clear</a>\n &nbsp; &nbsp; <a href="?call=deltbl&amp;db=' + db + '&amp;table=' + list("table_name") + '"><img src="img/del.gif" alt=""/> del</a></td>\n');
        Response.Write('</tr>\n\n');
    count = (count==0) ? 1 : 0;
    }

    list.movenext();
}

list.close();

Response.Write('</table>\n\n');
}


function compact( db ){


}




//-------------A D D   A   T A B L E

function addtable(){
   if(table!=0){
      list = conn.OpenSchema( 20 );
      
      //---------check for table existense
      while(!list.eof){
         if(list("table_type")=="TABLE"){
            if( list("table_name") == table){
               count = 1;
            }
         }
      list.movenext();
      }
      list.close();   
      
      if(count!=1){
         if(num==0){ //if autonumber value passed?
            num = 'id';
         }
         
         sql = 'Create table [' + table + '] (' + num + ' int identity(1,1) PRIMARY KEY)';
         conn.execute(sql);
         
         Response.Write('<p>The table [<em class="b">' + table + '</em>] has been added succesfully. The Primary Key is [<em class="b">' + num + '</em>].</p>');
      }else{   
         Response.Write('<p>The table [<em class="b">' + table + '</em>] already exists. Please enter a different name for the table</p>');
         Response.Write('<p>' + gendb.settings.sepnav + '<a href="' + script + '?call=addtable&db=' + db + '">Add table</a>');
      }
   }else{
%>
<form method="post" action="<%=script%>?call=addtable&amp;db=<%=db%>">
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td>Table Name:</td>
<td>Autonumber:</td>
</tr>

<tr class="row0">
<td><input type="text" name="table"/></td>
<td><input type="text" name="num"/> (If this is left empty, [<em class="b">id</em>] is kept as autonumber w/. Primary Key)</td>
</tr>

<tr class="row1">
<td colspan="2"><input type="submit" class="button" value="Add!"/></td>
</tr>

</table>
</form>
<%
   }//if
}//fx (Exec)


// -----------------------------D I S P L A Y   T A B L E   P R O P E R T I E S
function tblprop(){
%>
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td>Opts: <a href="?call=addfld&amp;db=<%=db%>&amp;table=<%=table%>"><img src="img/new.gif" alt="Add Field"/></a></td>
<td>Fields</td>
<td>Allow Null</td>
<td>Default Value</td>
<td>Field-Type</td>
<td>Simple Field-Type</td>
</tr>

<%
list = conn.OpenSchema( adSchemaColumns );

while(!list.eof){
    if(table == list("table_name")) { //otherwise shows all the columns for all the tables in the db.
        Response.Write('<tr class="row' + count + '">\n');
        Response.Write('<td><a href="?call=editfld&amp;db=' + db + '&amp;table=' + table + '"><img src="img/editf.gif" alt="edit field"/></a> <a href="?call=delfld&amp;db=' + db + '&amp;table=' + table + '"><img src="img/del.gif" alt="delete field"/></a></td>\n');
        Response.Write('<td>' + list("column_name") + '</td>\n');
        Response.Write('<td>' + list("is_nullable") + '</td>\n');
        Response.Write('<td>' + list("COLUMN_DEFAULT") + '</td>\n');
        Response.Write('<td>' + gettype( list("DATA_TYPE") ) + '</td>\n');
        Response.Write('<td>' + getsimpletype( list("DATA_TYPE") ) + '</td>\n');
    
        Response.Write('</tr>\n\n');
    }	 // end table check
        list.movenext();
        count = (count == 1) ? 0 : 1;

}

list.close();
Response.Write('</table>');


}//end edit table fx.





function clrtbl(){

   if(confirm==1){
    sql = "delete * from [" + table + "]";
    Response.Write('<p>Table Cleared</p>');
    //Response.Write(sql);
    conn.execute(sql);
   }else{
        Response.Write('<p>Are you sure you want to do this? This will remove all the records from the table : [<em class="b">' + table +'</em>]</p>');
        Response.Write('<p><a class="but" href="?call=clrtbl&amp;db=' + db + '&amp;table=' + table + '&amp;confirm=1">Yes</a> &nbsp; <a class="but" href="?call=lsttbls&db=' + db + '">No</a></p>');
   }
    

} //end fx


function deltbl(){
%>
<p>
<%
   if(confirm==1){
    sql = "drop table [" + table + "]";
    Response.Write('<p>Table Deleted</p>');
    //Response.Write(sql);
    conn.execute(sql);
   }else{
        Response.Write('<p>Are you sure you want to do this? This will delete the table : [<em class="b">' + table +'</em>]</p>');
        Response.Write('<p><a class="but" href="?call=deltbl&amp;db=' + db + '&amp;table=' + table + '&amp;confirm=1">Yes</a> &nbsp; <a class="but" href="?call=lsttbls&db=' + db + '">No</a></p>');
   }
    
   Response.Write('</p>');
} //end fx







//-------------L I S T   A L L   R E C O R D S

function lstrec(sqlc){
   
   var sql = "select * from " + table;

   if(num!=0){
    sql+= " order by [" + num + "] " + numvalue;
   }

   if(sqlc){
      sqlc = String( unescape( sqlc ) );
   }
   
   sql = (sqlc) ? sqlc : sql;

//	sql = sql.makesafe();
    
   rs.open(sql, conn, 3);

   sql = escape(sql);
   rs.PageSize = gendb.settings.norec; //no. of records specified by usr
   
   if (page > rs.PageCount){
      page = rs.PageCount;
   }
   
   if (page < 1){
      page = 1;
   }
   
   options2('Records: ' + rs.recordCount + ']' + gendb.settings.sepnav + '[Pages: ' + rs.PageCount,1,null,null,1,1,1);
   
   /*
   Response.Write('<br/>total records : ' +rs.recordcount)
   Response.Write('<br/>current page<b>: ' + page + '</b>')
   Response.Write('<br/>no of records: PageSize :' + rs.PageSize)
   Response.Write('<br/>default page: AbsolutePage :' + rs.AbsolutePage)
   Response.Write('<br/>no. of pages: PageCount :' +rs.PageCount)
   Response.Write('<br/><br/>');
   */
   
   var execsql = "";
   if(sqlc){
    execsql = "?call=exec&amp;db=" + db + "&amp;memo=" + sql;
   }else{
    execsql = "?call=lstrec&amp;db=" + db + "&amp;table=" + table;
   }
   
   //Response.Write(execsql); // (check values by exec() function)

   //Links for pre, next, first, last.  {START}
   if (page>1){
      options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page='+(page-1)+'" title="Previous"','<img src="img/pre.gif" alt="Previous"/>',null,null,null,1);
   }else{
      options2('<img src="img/pre1.gif" alt="Previous"/>',null,null,null,1);
   }

   if (page<rs.PageCount){
      options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page='+(page+1)+'" title="Next"','<img src="img/next.gif" alt="Next"/>',null,null,null,1);
   }else{
      options2('<img src="img/next1.gif" alt="Next"/>',null,null,null,1);
   }
   
      options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page=1','<img src="img/empty.gif" alt="first"/>',null,null,null,1);
   
      options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page='+rs.PageCount,'<img src="img/editf.gif" alt="last"/>',null,null,null,1);

   //Links for pre, next, first, last. {END}


      //DISPLAYS THE PAGE NO NAVIGATION

   if(sqlc){
      for(i=1; i<=rs.PageCount; i++){
         if(page==i){
            options2(' ' + i + ' ',null,null,null,1);
         }else{
            options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page=' + i,'( ' + i + ' )',null,null,null,1);
         }
      }
   }else{
      Response.Write('\n\n<select name="m" onChange="jump(\'' + replacejs(execsql) + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page=' + '\',this)">\n');
      
      for(i=1; i<=rs.PageCount; i++){
         if(i==page){
            Response.Write('<option value="' + i + '" selected="selected">Page: ' + i + '</option>\n');
         }else{
            Response.Write('<option value="' + i + '">' + i + '</option>\n');
         }
      }
      Response.Write('</select>\n\n');
   }
   
   //END PAGE/SELECT NAVIGATION
   
      options(execsql + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page=' + page,'Refresh',null,1,null,null,1,1);
   
   if(!rs.eof){  //because empty 'rs' gives error
      rs.AbsolutePage = page;
   }
   
   //
   // Table Sorting removed for 'custom sql execution'
   // 
   //
   
   Response.Write('<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">');
   Response.Write('<tr class="options">\n');
   
   if(!sqlc){
      Response.Write('<td>Opts:&nbsp;<a href="?call=addrec&amp;db=' + db + '&amp;table=' + table + '"><img src="img/new.gif" alt="Insert"/></a></td>\n');
   }

   for (i=0; i<rs.fields.count; i++){
    Response.Write('<td>');
    
    if(rs(i).Properties("ISAUTOINCREMENT")==1){
        autono = String(rs(i).Name);
        orderby = autono; //change to the current autonumber field.
        num = (num==0) ? orderby : num;
    }
   
      if(!sqlc && rs(i).Type!=adLongVarChar && rs(i).Type!=adLongVarWChar){	
      Response.Write('<a href="?call=lstrec&amp;db=' + db + '&amp;table=' + table + '&amp;num=' + rs(i).Name + '&amp;numvalue=' + ((num == rs
   (i).Name) ? ((numvalue=='asc') ? 'desc' : 'asc') : 'asc') + '&amp;page=' + page + '">');
         Response.Write(rs(i).Name);
    
      if(num == rs(i).Name){
            Response.Write('&nbsp;<img src="img/' + numvalue + '.gif" alt=""/>');
      }
      Response.Write('</a>');
    
      }else{
        Response.Write(rs(i).Name);
      }

    Response.Write('</td>\n');
   }
   Response.Write('</tr>\n\n');
   
   for (j = 0; j < rs.PageSize; j++){
      if(!rs.eof){
      Response.Write('<tr class="row' + count + '">\n');
    
      if(!sqlc){
      Response.Write('<td><a href="?call=viewrec&amp;db=' + db + '&amp;table=' + table + '&amp;num=' + autono + '&amp;numvalue=' + rs(autono) + '&page=' + page + '"><img src="img/empty.gif" alt="View Record"/></a>\n');
      Response.Write(' <a href="?call=editrec&amp;db=' + db + '&amp;table=' + table + '&amp;num=' + autono + '&amp;numvalue=' + rs(autono) + '&page=' + page +'"><img src="img/editf.gif" alt="Edit"/></a>\n');
      Response.Write(' <a href="?call=delrec&amp;db=' + db + '&amp;table=' + table + '&amp;num=' + autono + '&amp;numvalue=' + rs(autono) + '&page=' + page + '"><img src="img/del.gif" alt="Delete"/></a> </td>\n');
      }	
    
    for (i=0; i<rs.fields.count; i++){
        Response.Write('<td>');
        
        var text = String(rs(i));
        
        if(text.length>300){
            text = text.substring(0, gendb.settings.maxsize) + '...[<em class="b">contd</em>]';
        }
        Response.Write(text);
        
        Response.Write('</td>');
    }
    
    Response.Write('</tr>\n\n');
      
      count = ((count==0) ? 1 : 0);
      rs.movenext();
   }
   
   }
   var pagecount = rs.PageCount;
   var absolutepage = rs.AbsolutePage;
   
   rs.close();
   Response.Write('</table>\n\n');
}

//------------------------A D D   R E C O R D
function addrec(){
   sql = "SELECT * from [" + table + "]";
   rs.open(sql, conn);
   
%>
<form method="post" action="<%=script%>?call=saverec&amp;db=<%=db%>&amp;table=<%=table%>&amp;num=<%=num%>&amp;numvalue=<%=numvalue%>&amp;confirm=1"> 
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td class="options">Fields</td>
<td>Values</td>
<td>Max Size</td>
<td>Data Type</td>
</tr>
<%
      for(i=0; i<rs.fields.count; i++){
%>
<tr class="row<%=count%>">
<td><%=rs(i).Name%></td>
<td>
<%
      switch (parseInt(rs(i).Type)){
         case adInteger:
         case adTinyInt:
         case adSmallInt:
         case adBigInt:
         case adUnsignedTinyInt:
         case adUnsignedSmallInt:
         case adUnsignedInt:
         case adUnsignedBigInt:
         case adSingle:
         case adDouble:
         case adNumeric:
            if(rs(i).Properties("ISAUTOINCREMENT")==1){
               Response.Write('Auto');
            }else{
               sInput(rs(i).Name,'');
            }
            break;
      
         case adDate:
         case adDBDate:
         case adDBTime:
         case adDBTimeStamp:
            sInput(rs(i).Name,'');
            Response.Write('[YYYY/MM/DD H:M:S] (24Hours time cycle)');
            break;
      
         case adBoolean:
            sInputc(rs(i).Name,'');
            break;
      
         case adLongVarChar: 
         case adLongVarWChar:
            sTextarea(rs(i).Name,'');
            break;
      
         default:
            sInput(rs(i).Name,'');
            break;
      }
%>
</td>
<td class="c"><%=getsize(rs(i).definedSize,1)%></td>
<td class="c"><%=getsimpletype(rs(i).Type)%> <%
      if(rs(i).Properties("ISAUTOINCREMENT")==1){
         Response.Write(' Autonumber');
      }
%></td>
</tr>
<%
      } //for
%>

<tr class="row0"><td colspan="4"><input type="submit" value="done"/></td></tr>
</table>
<%
   rs.close();
}






//------------------------E D I T   R E C O R D
function editrec(){

   sql = "SELECT * from [" + table + "] where [" + num + "] = " + numvalue;
   rs.open(sql, conn);
   
   if(!rs.eof){
%>
<form method="post" action="<%=script%>?call=saverec&amp;db=<%=db%>&amp;table=<%=table%>&amp;num=<%=num%>&amp;numvalue=<%=numvalue%>"> 
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td class="options">Fields</td>
<td>Values</td>
<td>Max Size</td>
<td>Data Type</td>
</tr>
<%
      for(i=0; i<rs.fields.count; i++){
%>
<tr class="row<%=count%>">
<td><%=rs(i).Name%></td>
<td>
<%
      switch (parseInt(rs(i).Type)){
         case adInteger:
         case adTinyInt:
         case adSmallInt:
         case adBigInt:
         case adUnsignedTinyInt:
         case adUnsignedSmallInt:
         case adUnsignedInt:
         case adUnsignedBigInt:
         case adSingle:
         case adDouble:
         case adNumeric:
            if(rs(i).Properties("ISAUTOINCREMENT")==1){
               Response.Write(rs(i).Value);
            }else{
               sInput(rs(i).Name,rs(i).Value);
            }
            break;
      
         case adDate:
         case adDBDate:
         case adDBTime:
         case adDBTimeStamp:
            dt = new Date(rs(i).Value);
            sInput(rs(i).Name,dt.getdb());
            Response.Write('[YYYY/MM/DD H:M:S] (24Hours time cycle)');
            break;
      
         case adBoolean:
            sInputc(rs(i).Name,rs(i).Value);
            break;
      
         case adLongVarChar: 
         case adLongVarWChar:
            sTextarea(rs(i).Name,rs(i).Value);
            break;
      
         default:
            sInput(rs(i).Name,rs(i).Value);
            break;
      }
%>
</td>
<td class="c"><%=getsize(rs(i).definedSize,1)%></td>
<td class="c"><%=getsimpletype(rs(i).Type)%> <%
      if(rs(i).Properties("ISAUTOINCREMENT")==1){
         Response.Write(' Autonumber');
      }
%></td>
</tr>
<%
      } //for
%>

<tr class="row0"><td colspan="4"><input type="submit" value="done"/></td></tr>
</table>
<%
   }else{ //if
      Response.Write("Record does not exist!");
   }
   rs.close();
}

//-------------------V I E W    R E C O R D

function viewrec(){
   sql = "SELECT * from [" + table + "] where [" + num + "] = " + numvalue;
   rs.open(sql, conn);
   
   if(!rs.eof){
%>
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td class="options">Fields</td>
<td>Values</td>
<td>Max Size</td>
<td>Data Type</td>
</tr>
<%
      for(i=0; i<rs.fields.count; i++){
%>
<tr class="row<%=count%>">
<td><%=rs(i).Name%></td>
<td>
<%
         if(rs(i).Attributes==234){
            Response.Write(Server.HtmlEncode(rs(i).Value));
         }else{
            Response.Write(rs(i).Value);
         }
%>
</td>
<td class="c"><%=getsize(rs(i).definedSize,1)%></td>
<td class="c"><%

         Response.Write( getsimpletype(rs(i).Type) );
      
         if(rs(i).Properties("ISAUTOINCREMENT")==1){
            Response.Write(' (Autonumber)');
         }
   
/*
for(j=0; j<rs(i).Properties.Count; j++){
   Response.Write('<br/>');
   Response.Write( rs(i).Properties(j).Name );
   Response.Write(' - ' + rs(i).Properties(j).Value );
}

Response.Write('<br/>'+rs(i).Attributes);
*/

%></td>

</tr>
<%
      } //for
      Response.Write('</table>');
   
   }else{ //if !rs.eof
      Response.Write('<p>Record does not exist!</p>');
   }

rs.close();
}//end fx




//-------------------D E L E T E    R E C O R D

function delrec(){
   sql = "SELECT * from [" + table + "] where [" + num + "] = " + numvalue;
   rs.open(sql, conn);
   
   if(confirm=="yes"){
   Response.Write('<p>');
    sql = "delete * from [" + table + "] where [" + num + "] = " + numvalue;
    conn.execute(sql);
    Response.Write('Record deleted');
    Response.Write('<br/> Back to '+ tbllink(db, table, page));
   Response.Write('</p>');
   
   }else{
        Response.Write('<p>Are you sure you want to delete the foll. record?</p>');
        Response.Write('<p><a class="but" href="?call=delrec&amp;confirm=yes&amp;db='+db+'&amp;table=' + table + '&amp;num=' + num + '&amp;numvalue=' + numvalue + '&amp;page=' + page + '">Yes</a> &nbsp; <a class="but" href="?call=lstrec&amp;db=' + db + '&amp;table=' + table + '">No</a></p>');	
   
   if(!rs.eof){
%>
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options">
<td class="options">Fields</td>
<td>Values</td>
</tr>
<%
      for(i=0; i<rs.fields.count; i++){
%>
<tr class="row<%=count%>">
<td><%=rs(i).Name%></td>
<td>
<%
         if(rs(i).Attributes==234){
            Response.Write(Server.HtmlEncode(rs(i).Value));
         }else{
            Response.Write(rs(i).Value);
         }
%>
<%=rs(i).length%>
</td>
</tr>
<%
      } //for
      Response.Write('</table>');

   }else{ //if !rs.eof
      Response.Write('<p>Record does not exist!</p>');
   }
} //confirm==yes?

rs.close();
}//end fx


//-------------------------S A V E   R E C O R D

function saverec(){
   sql = "SELECT * from [" + table + "] where [" + num + "] = " + numvalue;
   if(confirm){
      sql = "SELECT * from [" + table + "]";
   }
   
   //Response.Write(sql);
   
   rs.open(sql, conn,3,0x0002);

   if(confirm){
      rs.AddNew();
   }
%>
<%
   for(i=0; i<rs.fields.count; i++){
   
      if(rs(i).Properties("ISAUTOINCREMENT")!=1){
         switch (rs(i).Type){
            case adDate:
            dt = new Date(Request(rs(i).Name));
            rs(i).Value = dt.getdb();
            break;
            
            case adBoolean:
            var cbool = String(Request(rs(i).Name));
               switch (cbool){
                  case 'undefined':
                  case 'false':
                  case 'False':
                  rs(i).Value = 'False';
                  break;
                  
                  default:
                  rs(i).Value = 'True';
                  break;
               }
            break;
            
            case adCurrency:
            rs(i).Value = Request(rs(i).Name);
            break;
            
            default:
            rs(i).Value = Request(rs(i).Name);
            break;
        }//switch
      }else{
         num = rs(i).Name;
      }
   } //for
   rs.Update();
   rs.MoveLast();
   numvalue = parseInt(rs(num));
   if(confirm){
      Response.Write('<p>Record added succesfully! Here are the values for the same.</p>');
   }else{
      Response.Write('<p>Record updated succesfully! here are the values for the same.</p>');
   }

   rs.close();

   viewrec();
} //end function







//------------E X E C U T E   C U S T O M   S T A T E M E N T
function exec(){
   memo = query('memo');

   if(memo!=0){
   memo = unescape(memo);
    Response.Write('<p>the foll. command was executed:<br/><span class="text">' + memo + '</span></p>');
    
    if(String(memo).toLowerCase().indexOf("select") > -1){
        lstrec(memo);
    }else{
      conn.execute(memo);
    }
   }else{
%>
<form name="execcust" method="get" action="<%=script%>?">
<table border="1" bordercolor="#999999" cellpadding="0" cellspacing="0">
<tr class="options"><input type="hidden" name="call" value="exec"/>
<input type="hidden" name="db" value="<%=db%>"/>
<td>Enter SQL statements here.
</td>
<td>Examples &amp; Tablelist</td>
</tr>

<tr class="row0">
<td><textarea class="bg" name="memo"></textarea></td>
<td>
<p>
Insert Table Name: <select onChange="selectel(this);">
<option value="" selected="selected">Table Names:</option>
<%
   list = conn.OpenSchema(20);
   
   while(!list.eof){
      if(list("table_type")=="TABLE"){
         Response.Write('<option value="[' + list("table_name") + ']">' + list("table_name") + '</option>');
      }
   list.movenext();
   }
   list.close();   
   
%>
</select></p>
<p>
SELECT * from [table]<br/>
DELETE * from [table] where id = 7<br/>
CREATE Table [table]<br/>
DROP Table [table]<br/>
Create Table [table] (id int identity(1,1) PRIMARY KEY)<br/>
insert into [table] ([field]) values ('[value]')

<br/><br/>[table] = tablename
</td>
</tr>

<tr class="row1">
<td colspan="2"><input type="submit" class="button" value="Execute!"/></td>
</tr>

</table>
</form>
<%
   }//if
}//fx (Exec)
    









%>

<div class="bar mtop15">
<p>[<a href="?call=logout">Logout</a>]<%=gendb.settings.sepnav%>[<a href="?call=contact">Report Bugs</a>]<%=gendb.settings.sepnav%>[Help]</p>
</div>

<div class="bar"></div>

</div> <!-- canvas -->
</body>
</html>
<%
if(db!=0 ){
   gendb.closedb();
}
%>