<%

function replacejs(text){
   text = text.replace("'","\\'");
   text = text.replace('"','\\"');
   return text;
}

function dblink(db){
   return '[<a href="?call=listtables&amp;db=' + db + '">' + db + '</a>]';
}

function tbllink(db,table){
   return '[<a href="?call=listrec&amp;db=' + db + '&amp;table=' + table + ((page) ? '&amp;page=' + page : '') + '">' + table + '</a>]'
}

function options(opt1,opt2,p1,p2,sep1,sep2,end1,end2){
   if(p1){Response.Write('<div class="bar mtop"><p>');}
   if(end1){Response.Write('[')};
   Response.Write('<a href="' + script + opt1 + '">' + opt2 + '</a>');
   if(sep1){Response.Write(' - ');}
   if(end2){Response.Write(']')};
   if(sep2){Response.Write(' &nbsp; ');}
   if(p2){Response.Write('</p></div>\n\n');}
}

function options2(text,p1,p2,sep1,sep2,end1,end2){
   if(p1){Response.Write('<div class="bar mtop"><p>');}
   if(end1){Response.Write('[')};
   Response.Write(text);
   if(sep1){Response.Write(' - ');}
   if(end2){Response.Write(']')};
   if(sep2){Response.Write(' &nbsp; ');}
   if(p2){Response.Write('</p></div>\n\n');}
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


function gettype(text){
   switch (parseInt(text)){
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
   Response.Write('<div class="bar"><p>\n[<a href="?">Home</a>] ' + nav + '</p></div>');
}



function opendb(file){
   file = Server.Mappath(path + file)
   conn = Server.CreateObject("ADODB.connection");
   conn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file);
   rs = Server.CreateObject("ADODB.Recordset");
}

function closedb(){
   conn.close();
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
   //return parseInt(this.getMonth()+((opt) ? 0 : 1)) + '-' + this.getDate() + '-' + this.getYear() + ' ' + this.getHours() + ':' + this.getMinutes() + ':' + this.getSeconds(); 

   return this.getYear() + '/' + parseInt( this.getMonth()+1 ) + '/' + this.getDate() + ' ' + this.getHours() + ':' + this.getMinutes() + ':' + this.getSeconds(); 

}



function query(obj,value,int){

   obj = String(Request(obj));
   
   if(int){ //if querytype = integer.
      obj = parseInt(obj);
      if(isNaN(obj)){
         if(value){obj = value}else{obj = 0};
      }
   }else{
      if(obj == "undefined" || obj == null || obj == ""){
         if(value){obj = value}else{obj = 0};
      }
   }
   
   return obj;
}



%>