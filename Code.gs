const SHEET_NAME="Software";
const ADMIN_PASSWORD="CHANGE_THIS_PASSWORD";

function getSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh=ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,6).setValues([["id","name","url","category","description","position"]]);
  }
  return sh;
}

function doGet(e){
  const data=getSoftware_();
  const cb=e&&e.parameter?e.parameter.callback:"";
  if(cb){
    return ContentService.createTextOutput(cb+"("+JSON.stringify(data)+")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const data=JSON.parse(e.postData.contents||"{}");
    if(String(data.password||"")!==String(ADMIN_PASSWORD)){
      return json_({success:false,error:"Wrong password"});
    }

    let result;
    if(data.action==="add") result=addSoftware_(data.item);
    else if(data.action==="edit") result=editSoftware_(data.id,data.item);
    else if(data.action==="delete") result=deleteSoftware_(data.id);
    else if(data.action==="reorder") result=reorderSoftware_(data.ids);
    else result={success:false,error:"Unknown action"};

    return json_(result);
  }catch(err){
    return json_({success:false,error:String(err)});
  }
}

function getSoftware_(){
  const sh=getSheet_(), last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,6).getValues()
    .filter(r=>r[0])
    .map(r=>({
      id:String(r[0]),
      name:String(r[1]||""),
      url:String(r[2]||""),
      category:String(r[3]||"utility"),
      description:String(r[4]||""),
      position:Number(r[5]||0)
    }))
    .sort((a,b)=>a.position-b.position);
}

function addSoftware_(item){
  if(!item||!item.id||!item.name||!item.url)
    return {success:false,error:"Name, URL and ID are required"};

  const list=getSoftware_();
  if(list.some(x=>String(x.id)===String(item.id)))
    return {success:false,error:"Software already exists"};

  const pos=list.length?Math.max.apply(null,list.map(x=>Number(x.position)||0))+1:1;
  getSheet_().appendRow([
    String(item.id),String(item.name),String(item.url),
    String(item.category||"utility"),String(item.description||""),pos
  ]);
  return {success:true};
}

function editSoftware_(id,item){
  const sh=getSheet_(), row=findRow_(sh,id);
  if(row===-1)return {success:false,error:"Software not found"};
  const pos=sh.getRange(row,6).getValue();
  sh.getRange(row,1,1,6).setValues([[
    String(id),String(item.name||""),String(item.url||""),
    String(item.category||"utility"),String(item.description||""),pos
  ]]);
  return {success:true};
}

function deleteSoftware_(id){
  const sh=getSheet_(), row=findRow_(sh,id);
  if(row===-1)return {success:false,error:"Software not found"};
  sh.deleteRow(row);
  normalizePositions_();
  return {success:true};
}

function reorderSoftware_(ids){
  if(!Array.isArray(ids))return {success:false,error:"Invalid IDs"};
  const sh=getSheet_();
  ids.forEach((id,i)=>{
    const row=findRow_(sh,id);
    if(row!==-1)sh.getRange(row,6).setValue(i+1);
  });
  normalizePositions_();
  return {success:true};
}

function findRow_(sh,id){
  const last=sh.getLastRow();
  if(last<2)return -1;
  const ids=sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<ids.length;i++)
    if(String(ids[i][0])===String(id))return i+2;
  return -1;
}

function normalizePositions_(){
  const sh=getSheet_(), last=sh.getLastRow();
  if(last<2)return;
  const rows=sh.getRange(2,1,last-1,6).getValues();
  rows.sort((a,b)=>Number(a[5]||0)-Number(b[5]||0));
  rows.forEach((r,i)=>r[5]=i+1);
  sh.getRange(2,1,rows.length,6).setValues(rows);
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
