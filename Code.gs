const ROOT_FOLDER_ID = '';
const COMMUNITY_FOLDER_NAME = 'Community Uploads';
const MAX_FILE_SIZE_BYTES = 45 * 1024 * 1024;

const AUSTRALIA_STRUCTURE = {
  "VIC": { "VCE": ["Accounting", "Agricultural and Horticultural Studies", "Algorithmics HESS", "Ancient History", "Applied Computing Data Analytics", "Applied Computing Software Development", "Art Creative Practice", "Art Making and Exhibiting", "Australian Politics", "Global Politics", "Biology", "Business Management", "Chemistry", "Chinese Second Language", "Classical Studies", "Dance", "Drama", "Economics", "English", "English Language", "EAL", "Food Studies", "Foundation Mathematics", "General Mathematics", "Geography", "Health and Human Development", "History Revolutions", "Indonesian Second Language", "Japanese Second Language", "Legal Studies", "Literature", "Mathematical Methods", "Specialist Mathematics", "Media", "Music Contemporary Performance", "Music Inquiry", "Outdoor and Environmental Studies", "Philosophy", "Physical Education", "Physics", "Product Design and Technology", "Psychology", "Religion and Society", "Sociology", "Systems Engineering", "Theatre Studies", "Visual Communication Design", "French", "German", "Latin", "Chinese First Language", "Spanish"] },
  "QLD": { "QCE": ["Accounting", "Ancient History", "Biology", "Business", "Chemistry", "Design", "Digital Solutions", "Drama", "Earth and Environmental Science", "Economics", "English", "English as an Additional Language", "Literature", "English and Literature Extension", "Film Television and New Media", "Food and Nutrition", "Geography", "General Mathematics", "Mathematical Methods", "Specialist Mathematics", "French", "German", "Japanese", "Legal Studies", "Modern History", "Music", "Music Extension", "Physical Education", "Physics", "Psychology", "Study of Religion", "Visual Art"] },
  "WA": { "WACE": ["Accounting and Finance", "Ancient History", "Biology", "Business Management and Enterprise", "Chemistry", "Computer Science", "Dance", "Design", "Drama", "Earth and Environmental Science", "Economics", "Engineering Studies", "English", "English as an Additional Language", "Food Science and Technology", "French", "Geography", "German", "Health Studies", "Human Biology", "Japanese Second Language", "Literature", "Marine and Maritime Studies", "Materials Design and Technology", "Mathematics Applications", "Mathematics Methods", "Mathematics Specialist", "Media Production and Analysis", "Modern History", "Music", "Outdoor Education", "Philosophy and Ethics", "Physical Education Studies", "Physics", "Politics and Law", "Psychology", "Visual Arts"] },
  "SA": { "SACE": ["Accounting", "Agricultural Production", "Ancient Studies", "Biology", "Business Innovation", "Chemistry", "Child Studies", "Classical Studies", "Design and Technology", "Digital Communication Solutions", "Drama", "Economics", "English", "EAL", "English Literary Studies", "Essential English", "Food and Hospitality", "French", "General Mathematics", "Geography", "German", "Health and Wellbeing", "Japanese", "Legal Studies", "Mathematical Methods", "Media Studies", "Modern History", "Music", "Nutrition", "Outdoor Education", "Philosophy", "Physical Education", "Physics", "Psychology", "Society and Culture", "Specialist Mathematics", "Visual Arts"] },
  "TAS": { "TCE": ["Accounting", "Ancient History", "Biology", "Business Studies", "Chemistry", "Computer Science", "Economics", "English", "English as an Additional Language", "Food and Nutrition", "Geography", "Legal Studies", "General Mathematics", "Mathematics Methods", "Mathematics Specialised", "Modern History", "Music", "Philosophy", "Physical Sciences", "Physics", "Psychology", "Sociology", "Visual Art"] },
  "ACT": { "BSSS": ["Accounting", "Ancient History", "Biology", "Business", "Chemistry", "Design and Technology", "Drama", "Economics", "English", "English as an Additional Language", "Geography", "History Modern", "Legal Studies", "Literature", "Mathematical Applications", "Mathematical Methods", "Specialist Mathematics", "Media", "Music", "Philosophy", "Physical Education", "Physics", "Psychology", "Sociology"] },
  "NT": { "NTCET": ["Accounting", "Biology", "Business Innovation", "Chemistry", "Drama", "Economics", "English", "English Literary Studies", "Essential English", "Geography", "Legal Studies", "General Mathematics", "Mathematical Methods", "Specialist Mathematics", "Modern History", "Nutrition", "Physical Education", "Physics", "Psychology", "Society and Culture", "Visual Arts"] }
};

const LEAF_FOLDERS = ["Trial Papers", "Assessment Tasks", "Trial", "Assessment"];
const CACHE_TTL_DATA = 60;
const CACHE_TTL_DRIVE = 30;
const CACHE_TTL_BINNED = 120;

function hashPassword(password) {
  if (!password) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password).trim(), Utilities.Charset.UTF_8);
  let txtHash = "";
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    let byteStr = byteVal.toString(16);
    if (byteStr.length == 1) byteStr = "0" + byteStr;
    txtHash += byteStr;
  }
  return txtHash;
}

function doPost(e) {
  let lock = null;
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let response = { success: false, message: "Unknown action" };

    if (action === "get_data") {
      const cached = getCache().get("files_data");
      if (cached) {
        return jsonResponse({ success: true, data: JSON.parse(cached), cached: true });
      }
      const data = getFilesDataFast(ss, false);
      getCache().put("files_data", JSON.stringify(data), CACHE_TTL_DATA);
      response = { success: true, data: data };
    } else if (action === "list_drive") {
      const cacheKey = "drive:" + (payload.path||[]).join('/').toLowerCase();
      const cached = getCache().get(cacheKey);
      if (cached) {
        return jsonResponse(JSON.parse(cached));
      }
      const result = listDriveContentsFast(ss, payload);
      getCache().put(cacheKey, JSON.stringify(result), CACHE_TTL_DRIVE);
      response = result;
    } else if (action === "get_quota") {
      response = { success: true };
    } else {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) return jsonResponse({ success: false, message: "Server busy, try again." });
      
      if (action === "register") response = handleRegister(ss, payload);
      else if (action === "login") response = handleLoginFast(ss, payload);
      else if (action === "change_password") response = handleChangePassword(ss, payload);
      else if (action === "delete_account") response = handleDeleteAccount(ss, payload);
      else if (action === "upload") response = handleUploadCommunity(ss, payload);
      else if (action === "edit_file") response = handleEditFile(ss, payload);
      else if (action === "bin_file") response = handleBinFile(ss, payload);
      else if (action === "vote") response = handleVote(ss, payload);
      else if (action === "admin_action") response = handleAdmin(ss, payload);

      if (response.success && ["upload","edit_file","bin_file","vote"].indexOf(action) !== -1) {
        invalidateCaches(payload);
      }
    }
    return jsonResponse(response);
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString(), stack: err.stack });
  } finally {
    if (lock) try{lock.releaseLock();}catch(e){}
  }
}

function getCache(){ return CacheService.getScriptCache(); }
function invalidateCaches(payload){
  try{
    getCache().remove("files_data");
    getCache().remove("files_data_all");
    getCache().remove("binned_ids");
    const path = payload.currentPath || payload.path || [];
    for(let i=0;i<=path.length;i++){
      const key = "drive:" + path.slice(0,i).join('/').toLowerCase();
      getCache().remove(key);
    }
    getCache().remove("drive:");
  }catch(e){}
}

function jsonResponse(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function toTitleCase(str){ if (!str) return ""; return str.toString().trim().toLowerCase().replace(/(^|\s|-|\/)\S/g, t => t.toUpperCase()); }
function toUpperCaseTrim(str){ if (!str) return ""; return str.toString().trim().toUpperCase(); }

function listDriveContentsFast(ss, payload){
  const path = payload.path || [];
  try{
    let currentFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    for(let i=0;i<path.length;i++){
      const segment = String(path[i]).trim();
      if(!segment) continue;
      const found = findFolderCaseInsensitive(currentFolder, segment);
      if(!found) return { success: true, folders: [], files: [], path: path, isLeaf: false };
      currentFolder = found;
    }

    const folders=[];
    const folderIter=currentFolder.getFolders();
    while(folderIter.hasNext()){
      const f=folderIter.next();
      if(path.length===0 && f.getName()===COMMUNITY_FOLDER_NAME) continue;
      folders.push({name:f.getName(), id:f.getId()});
    }
    folders.sort((a,b)=>a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
    const isLeaf=folders.length===0;

    let binnedIds=getBinnedIdsCached(ss);

    const files=[];
    const fileIter=currentFolder.getFiles();
    while(fileIter.hasNext()){
      const file=fileIter.next();
      const name=file.getName();
      if(!name.toLowerCase().endsWith('.pdf')) continue;
      if(binnedIds.has(file.getId())) continue;
      const parsed=parseFilenameForDrive(name);
      files.push({
        id:file.getId(), driveId:file.getId(), name:name, school:parsed.school, year:parsed.year,
        hasSol:parsed.hasSol, solutions:parsed.hasSol?"Yes":"No",
        url:`https://drive.google.com/uc?export=download&id=${file.getId()}`,
        state:path[0]||"", system:path[1]||"", subject:path[2]||"", task:path[3]||"",
        path:path.join('/'), isDriveFile:true, status:"Verified"
      });
    }
    files.sort((a,b)=>a.school.toLowerCase().localeCompare(b.school.toLowerCase()));

    return { success: true, folders: folders, files: files, path: path, isLeaf: isLeaf };
  }catch(err){
    return { success: false, message: "listDrive error: " + err.toString() };
  }
}

function getBinnedIdsCached(ss){
  const cache=getCache();
  const cached=cache.get("binned_ids");
  if(cached){
    try{ return new Set(JSON.parse(cached)); }catch(e){}
  }
  const binned=new Set();
  try{
    const sheet=ss.getSheetByName("Files");
    if(sheet && sheet.getLastRow()>1){
      const lastRow=sheet.getLastRow();
      const ids=sheet.getRange(2,1,lastRow-1,1).getValues();
      const statuses=sheet.getRange(2,13,lastRow-1,1).getValues();
      for(let i=0;i<ids.length;i++){
        if(String(statuses[i][0]).toLowerCase()==="binned"){
          binned.add(String(ids[i][0]).trim());
        }
      }
    }
  }catch(e){}
  cache.put("binned_ids", JSON.stringify(Array.from(binned)), CACHE_TTL_BINNED);
  return binned;
}

function findFolderCaseInsensitive(parent, name){
  const exact=parent.getFoldersByName(name);
  if(exact.hasNext()) return exact.next();
  const all=parent.getFolders();
  const lower=name.toLowerCase();
  while(all.hasNext()){
    const f=all.next();
    if(f.getName().toLowerCase()===lower) return f;
  }
  return null;
}

function parseFilenameForDrive(filename){
  const nameNoExt=filename.replace(/\.pdf$/i, '').trim();
  const hasSol=/\bw\.?\s*sol\.?\b/i.test(nameNoExt);
  const cleanName=nameNoExt.replace(/\s*\bw\.?\s*sol\.?\b/i, '').trim();
  const yearMatch=cleanName.match(/\b(19\d{2}|20\d{2})\b/);
  const year=yearMatch?yearMatch[0]:'';
  let school=yearMatch?cleanName.substring(0,yearMatch.index).trim().replace(/[-_\s]+$/, '').trim():cleanName;
  school=school.replace(/\s{2,}/g, ' ').trim();
  return { school, year, hasSol };
}

function getFilesDataFast(ss, includeBinned){
  let sheet=ss.getSheetByName("Files");
  if(!sheet){
    sheet=ss.insertSheet("Files");
    sheet.appendRow(["ID","State","System","Subject","Task","School","Year","Solutions","DriveLink","Uploader","Upvotes","Downvotes","Status","OrigPath","VotersJSON","FileHash","OriginalName","DriveId"]);
    return [];
  }
  const lastRow=sheet.getLastRow();
  if(lastRow<=1) return [];
  const data=sheet.getRange(2,1,lastRow-1,18).getValues();
  const out=[];
  for(let i=0;i<data.length;i++){
    const r=data[i];
    const status=String(r[12]||"Unverified");
    if(!includeBinned && status==="Binned") continue;
    out.push({
      id:String(r[0]), state:String(r[1]), system:String(r[2]), subject:String(r[3]), task:String(r[4]),
      school:String(r[5]), year:String(r[6]), solutions:String(r[7]), hasSol:String(r[7]).toLowerCase()==="yes",
      url:String(r[8]), uploader:String(r[9]), up:parseInt(r[10])||0, down:parseInt(r[11])||0,
      status:status, origPath:String(r[13]), voters:String(r[14]||"{}"),
      fileHash:String(r[15]||""), originalName:String(r[16]||""), driveId:String(r[17]||""),
      isDriveFile:false, path:`${r[1]}/${r[2]}/${r[3]}/${r[4]}`
    });
  }
  return out;
}

function handleUploadCommunity(ss, data){
  const loginCheck=handleLoginFast(ss, { username: data.username, password: data.password });
  if(!loginCheck.success) return { success: false, message: "Auth failed." };
  const currentPath=data.currentPath||[];
  if(!Array.isArray(currentPath) || currentPath.length===0) return { success: false, message: "Invalid folder." };
  const origPath=currentPath.join('/').trim();
  const state=toTitleCase(currentPath[0]||"");
  let system="", subject="", task="";
  if(currentPath.length===1){
    system=""; subject=""; task="";
  } else if(currentPath.length===2){
    system=toUpperCaseTrim(currentPath[1]||"");
  } else if(currentPath.length===3){
    const lastLower=String(currentPath[2]||"").toLowerCase();
    if(lastLower.includes("trial") || lastLower.includes("assess")){
      subject=toTitleCase(currentPath[1]||"");
      task=toTitleCase(currentPath[2]||"");
      system="";
    }else{
      system=toUpperCaseTrim(currentPath[1]||"");
      subject=toTitleCase(currentPath[2]||"");
    }
  } else {
    system=toUpperCaseTrim(currentPath[1]||"");
    subject=toTitleCase(currentPath[2]||"");
    task=toTitleCase(currentPath[3]||"");
  }

  const school=toTitleCase(data.school);
  const year=String(data.year).trim();
  const solutions=data.solutions?"Yes":"No";
  const originalName=data.originalName||"";
  if(!school||!year) return { success: false, message: "School and Year required." };
  if(!data.fileBase64) return { success: false, message: "No file." };
  const approxBytes=Math.floor(data.fileBase64.length*0.75);
  if(approxBytes>MAX_FILE_SIZE_BYTES) return { success: false, message: `File too large ${(approxBytes/1024/1024).toFixed(1)}MB. Max 45MB.` };

  const sheet=ss.getSheetByName("Files");
  if(sheet && sheet.getLastRow()>1){
    const rows=sheet.getRange(2,1,sheet.getLastRow()-1,14).getValues();
    const lowerSchool=school.toLowerCase(), lowerYear=year.toLowerCase();
    const lowerOrig=origPath.toLowerCase();
    for(let i=0;i<rows.length;i++){
      const r=rows[i];
      const status=String(r[12]||"Unverified").toLowerCase();
      if(status==="binned") continue;
      const existingOrig=String(r[13]||"").toLowerCase();
      if(existingOrig===lowerOrig && String(r[5]).toLowerCase()===lowerSchool && String(r[6]).toLowerCase()===lowerYear){
        return { success: false, message: `Duplicate: ${school} ${year} already exists in ${origPath} (ID: ${r[0]}).` };
      }
    }
  }

  let currentFolder=DriveApp.getFolderById(ROOT_FOLDER_ID);
  currentFolder=getOrCreateFolder(currentFolder, COMMUNITY_FOLDER_NAME);
  for(let seg of currentPath){
    currentFolder=getOrCreateFolder(currentFolder, seg.trim());
  }

  const fileName=`${school} ${year}${solutions==="Yes" ? " w. sol" : ""}.pdf`.replace(/\//g,"-");
  const fileIter=currentFolder.getFiles();
  while(fileIter.hasNext()){
    const f=fileIter.next();
    if(f.getName().toLowerCase()===fileName.toLowerCase()){
      return { success: false, message: `Duplicate file already in ${origPath}: ${fileName}` };
    }
    const parsed=parseFilenameForDrive(f.getName());
    if(parsed.school.toLowerCase()===school.toLowerCase() && parsed.year===year && ((parsed.hasSol && solutions==="Yes") || (!parsed.hasSol && solutions==="No"))){
      return { success: false, message: `Duplicate: ${school} ${year} already exists as ${f.getName()} in ${origPath}` };
    }
  }

  const blob=Utilities.newBlob(Utilities.base64Decode(data.fileBase64), "application/pdf", fileName);
  const file=currentFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const downloadUrl=`https://drive.google.com/uc?export=download&id=${file.getId()}`;
  const fileId="F_" + new Date().getTime() + "_" + Math.floor(Math.random()*1000);
  sheet.appendRow([fileId, state, system, subject, task, school, year, solutions, downloadUrl, data.username, 0, 0, "Unverified", origPath, "{}", "", originalName, file.getId()]);
  return { success: true, message: `Uploaded to ${origPath}: ${school} ${year}` };
}

function handleLoginFast(ss, data){
  const reqUser=String(data.username||"").trim();
  const reqPass=String(data.password||"").trim();
  const reqHash=hashPassword(reqPass);
  
  if(reqUser==="ADMINUSERNAME" && reqPass==="adminpassword") return { success:true, is_admin:true, username:"ADMINUSERNAME" };
  
  const cache=getCache();
  let usersJson=cache.get("users_cache");
  let users;
  if(usersJson){
    users=JSON.parse(usersJson);
  }else{
    let sheet=ss.getSheetByName("Users");
    if(!sheet||sheet.getLastRow()<=1) return { success:false, message:"Invalid credentials." };
    users=sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
    cache.put("users_cache", JSON.stringify(users), 300);
  }
  
  const user=users.find(u=> {
    const dbUser = String(u[0]).trim().toLowerCase();
    const dbPass = String(u[1]).trim();
    return dbUser === reqUser.toLowerCase() && (dbPass === reqHash || dbPass === reqPass);
  });
  
  if(user){
    if(String(user[4]).trim().toLowerCase()==="true") return { success:false, message:"Banned." };
    return { success:true, is_admin:false, username:String(user[0]).trim() };
  }
  return { success:false, message:"Invalid credentials." };
}

function handleRegister(ss, data){
  const reqUser=String(data.username||"").trim();
  const reqPass=String(data.password||"").trim();
  if(!reqUser||!reqPass) return {success:false,message:"Username and password required."};
  let userSheet=ss.getSheetByName("Users");
  if(!userSheet){ userSheet=ss.insertSheet("Users"); userSheet.appendRow(["Username","Password","IP","Created","Banned"]); }
  const users=userSheet.getRange(2,1,Math.max(userSheet.getLastRow()-1,0),1).getValues().flat().map(r=>String(r).trim().toLowerCase());
  if(users.includes(reqUser.toLowerCase())) return {success:false,message:"Username exists."};
  
  const hashed = hashPassword(reqPass);
  userSheet.appendRow(["'"+reqUser, "'"+hashed, String(data.ip||""), new Date().toISOString(), "False"]);
  getCache().remove("users_cache");
  return {success:true,message:"Account created!"};
}

function handleChangePassword(ss, data){
  const login=handleLoginFast(ss, { username:data.username, password:data.oldPassword });
  if(!login.success) return {success:false,message:"Current password incorrect."};
  const sheet=ss.getSheetByName("Users");
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  const idx=rows.findIndex(r=>String(r).trim().toLowerCase()===String(data.username).trim().toLowerCase());
  if(idx!==-1){ 
    const hashed = hashPassword(data.newPassword);
    sheet.getRange(idx+2,2).setValue("'"+hashed); 
    getCache().remove("users_cache"); 
    return {success:true,message:"Password changed!"}; 
  }
  return {success:false,message:"User not found."};
}

function handleDeleteAccount(ss, data){
  const login=handleLoginFast(ss, { username:data.username, password:data.password });
  if(!login.success) return {success:false,message:"Incorrect password."};
  const sheet=ss.getSheetByName("Users");
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  const idx=rows.findIndex(r=>String(r).trim().toLowerCase()===String(data.username).trim().toLowerCase());
  if(idx!==-1){ sheet.deleteRow(idx+2); getCache().remove("users_cache"); return {success:true,message:"Deleted."}; }
  return {success:false,message:"Not found."};
}

function getOrCreateFolder(parent, name){
  const cleanName=String(name).trim();
  if(!cleanName) return parent;
  const exact=parent.getFoldersByName(cleanName);
  if(exact.hasNext()) return exact.next();
  const all=parent.getFolders();
  while(all.hasNext()){ const f=all.next(); if(f.getName().toLowerCase()===cleanName.toLowerCase()) return f; }
  return parent.createFolder(cleanName);
}

function handleEditFile(ss, data){
  const login=handleLoginFast(ss, { username:data.username, password:data.password });
  if(!login.success) return {success:false,message:"Auth failed."};
  const sheet=ss.getSheetByName("Files");
  const ids=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  const idx=ids.findIndex(id=>String(id).trim()===String(data.fileId).trim());
  if(idx===-1) return {success:false,message:"Not found."};
  const rowNum=idx+2;
  const uploader=sheet.getRange(rowNum,10).getValue();
  if(String(uploader).trim()!==data.username && data.username!=="ADMINUSERNAME") return {success:false,message:"Unauthorized."};
  sheet.getRange(rowNum,6).setValue(toTitleCase(data.school));
  sheet.getRange(rowNum,7).setValue(String(data.year).trim());
  sheet.getRange(rowNum,8).setValue(data.solutions?"Yes":"No");
  sheet.getRange(rowNum,11).setValue(0);
  sheet.getRange(rowNum,12).setValue(0);
  sheet.getRange(rowNum,13).setValue("Unverified");
  sheet.getRange(rowNum,15).setValue("{}");
  return {success:true,message:"Updated."};
}

function handleBinFile(ss, data){
  const login=handleLoginFast(ss, { username:data.username, password:data.password });
  if(!login.success) return {success:false,message:"Auth failed."};
  const sheet=ss.getSheetByName("Files");
  const ids=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat();
  const idx=ids.findIndex(id=>String(id).trim()===String(data.fileId).trim());
  if(idx===-1) return {success:false,message:"Not found."};
  const rowNum=idx+2;
  const uploader=sheet.getRange(rowNum,10).getValue();
  if(String(uploader).trim()!==data.username && data.username!=="ADMINUSERNAME") return {success:false,message:"Unauthorized."};
  sheet.getRange(rowNum,13).setValue("Binned");
  return {success:true,message:"Binned."};
}

function handleVote(ss, data){
  let fileSheet=ss.getSheetByName("Files");
  if(!fileSheet){
    fileSheet=ss.insertSheet("Files");
    fileSheet.appendRow(["ID","State","System","Subject","Task","School","Year","Solutions","DriveLink","Uploader","Upvotes","Downvotes","Status","OrigPath","VotersJSON","FileHash","OriginalName","DriveId"]);
  }
  let ids=[];
  if(fileSheet.getLastRow()>1){ ids=fileSheet.getRange(2,1,fileSheet.getLastRow()-1,1).getValues().flat().map(id=>String(id).trim()); }
  let idx=ids.findIndex(id=>id===String(data.fileId).trim());
  if(idx===-1 && data.isDriveFile){
    const info=data.driveInfo||{};
    fileSheet.appendRow([String(data.fileId).trim(), String(info.state||""), String(info.system||""), String(info.subject||""), String(info.task||""), String(info.school||""), String(info.year||""), String(info.solutions||"No"), String(info.url||""), String(info.uploader||"drive_owner"), 0,0, String(info.status||"Verified"), String(info.path||""), "{}", "", String(info.originalName||""), String(data.fileId).trim()]);
    ids.push(String(data.fileId).trim());
    idx=ids.length-1;
  }
  if(idx===-1) return {success:false, message:"File not found."};
  const rowNum=idx+2;
  let voters={}; try{ voters=JSON.parse(fileSheet.getRange(rowNum,15).getValue()||"{}"); }catch(e){}
  const reqUser=String(data.username).trim(); if(!reqUser) return {success:false,message:"Username required."};
  if(voters[reqUser]===data.vote) delete voters[reqUser]; else voters[reqUser]=data.vote;
  let up=0,down=0; for(let u in voters){ if(voters[u]==="up")up++; if(voters[u]==="down")down++; }
  let status=String(fileSheet.getRange(rowNum,13).getValue()).trim()||"Verified";
  if((up+down)>=50){ if(up/(up+down)>=0.90) status="Verified"; else if(down/(up+down)>=0.70) status="Binned"; }
  fileSheet.getRange(rowNum,11).setValue(up);
  fileSheet.getRange(rowNum,12).setValue(down);
  fileSheet.getRange(rowNum,13).setValue(status);
  fileSheet.getRange(rowNum,15).setValue(JSON.stringify(voters));
  return {success:true, message:"Vote updated.", up:up, down:down, status:status};
}

function handleAdmin(ss, data){
  if(data.admin_user!=="ADMINUSERNAME"||data.admin_pass!=="adminpassword") return {success:false};
  if(data.cmd==="get_bin"){ const sheet=ss.getSheetByName("Files"); const values=sheet.getRange(2,1,sheet.getLastRow()-1,18).getValues(); return {success:true,data:values.filter(r=>String(r[12]).toLowerCase()==="binned"||String(r[12]).toLowerCase()==="unverified")}; }
  if(data.cmd==="verify_file"){ const s=ss.getSheetByName("Files"); const ids=s.getRange(2,1,s.getLastRow()-1,1).getValues().flat(); const idx=ids.findIndex(id=>String(id).trim()===String(data.fileId).trim()); if(idx!==-1){ s.getRange(idx+2,13).setValue("Verified"); return {success:true}; } }
}

function createAllAustraliaFolders(){
  let root=DriveApp.getFolderById(ROOT_FOLDER_ID);
  let created=0, skipped=0;
  for(let state in AUSTRALIA_STRUCTURE){
    if(state==="NSW"){ Logger.log("Skipping NSW entirely"); continue; }
    let stateFolder=getOrCreateFolder(root,state);
    for(let system in AUSTRALIA_STRUCTURE[state]){
      let sysFolder=getOrCreateFolder(stateFolder,system);
      AUSTRALIA_STRUCTURE[state][system].forEach(subj=>{
        let subjFolder=getOrCreateFolder(sysFolder,subj);
        LEAF_FOLDERS.forEach(leaf=>{ const leafExists=subjFolder.getFoldersByName(leaf); if(!leafExists.hasNext()){ subjFolder.createFolder(leaf); created++; } else skipped++; });
      });
    }
  }
  Logger.log(`Created ${created}, skipped ${skipped}. NSW skipped. Folders are now made, this function will not run again unless you manually run it.`);
  return `Created ${created}, skipped ${skipped}. NSW skipped.`;
}

function listAllAustraliaFolders(){ let out=[]; for(let state in AUSTRALIA_STRUCTURE){ if(state==="NSW") continue; for(let system in AUSTRALIA_STRUCTURE[state]){ AUSTRALIA_STRUCTURE[state][system].forEach(subj=>{ LEAF_FOLDERS.forEach(leaf=>{ out.push(`${state}/${system}/${subj}/${leaf}`); }); }); } } Logger.log(out.join('\n')); return out; }