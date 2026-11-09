/** PURCHASE TEMPLATE — STAGE 1
 * One file. Receipt handling and fixed purchase record IDs.
 * IDs use the visible Ordered on date and current item order; no counters.
 */
const PA = {
  receivedLogHeaders:['Timestamp','User','BlockID','SourceSheet','Row','Notes - changes in order','Change','Paid details'],
  categories:['Lab Supplies','Capital Asset/Equipment','Shipping','Others'],
  dateHelp:{
    ordered:'Enter the actual order date once. Blank dates leave the Item number blank.',
    month:'Choose an accounting date. The summary uses its month and year. Leave blank until assigned.'
  },
  grantFormula:'=IF(Control!$B$5="","",Control!$B$5)',
  vendors:['Amazon','Molex','Top3DShop','Edmund Optics','Thorlabs','MacMaster','DigiKey','Mouser','Others','Shipment'],
  vendorCodeKey:'PurchaseToolsVendorCode',
  vendorOriginKey:'PurchaseToolsVendorOrigin',
  removedAmazonKey:'PurchaseToolsAmazonRemoved',
  codes:{Amazon:'AMZ',Molex:'MOX',Top3DShop:'T3D','Edmund Optics':'EDM',Thorlabs:'THR',MacMaster:'MCM',DigiKey:'DIG',Mouser:'MOU','Fiber Optic Center':'FOC',FedEx:'FDX',UPS:'UPS',USPS:'USPS',DHL:'DHL'},
  text(v){return String(v??'').replace(/\u00a0/g,' ').trim();},
  number(v){return v==null||this.text(v)===''?NaN:Number(typeof v==='number'?v:this.text(v).replace(/[$,\s]/g,''));},
  month(v,tz){
    if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,'yyyy-MM');
    const s=this.text(v),iso=s.match(/^(20\d{2}-(?:0[1-9]|1[0-2]))(?:-\d{2})?$/);
    if(iso)return iso[1];
    const m=s.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(?:\d{1,2},?\s+)?(20\d{2})$/i);
    return m?m[2]+'-'+String(['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[1].toLowerCase())+1).padStart(2,'0'):'';
  },
  project(book,required=true){
    const sh=book.getSheetByName('Control');if(!sh)throw Error('Missing Control tab. Run setup first.');
    const name=this.text(sh.getRange('B3').getValue()),code=this.text(sh.getRange('B4').getValue()).toUpperCase();
    if(required&&(!name||!/^[A-Z0-9]{2,10}$/.test(code)))throw Error(book.getName()+': fill Control!B3 (name) and B4 (unique 2–10 letter/number code).');
    return {name,code};
  },
  layout(sh,requireOrdered=true){
    const first=sh.getName()==='Shipment'?5:4;
    const headers=sh.getRange(first-1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    const find=(...names)=>{const i=headers.findIndex(h=>names.includes(this.text(h)));return i<0?0:i+1;};
    const l={first,width:headers.length,id:find('Item number'),ordered:find('Ordered on'),priority:find('Priority'),item:find('Item','Description'),
      company:find('Company'),link:find('Link'),qty:find('Quantity','Unit'),unit:find('Individual price','Unit price'),
      total:find('Final price'),category:find('Category'),month:find('Reported on'),paid:find('Paid'),
      arrived:find('Arrived'),arrival:find('Arrive_lab'),notes:find('Notes'),tracking:find('Purchase Record ID')||find('Tracking ID'),key:find('Formula Key'),position:find('Position Key')};
    if([l.id,l.item,l.qty,l.unit,l.total,l.category,l.month,l.paid,l.arrived,l.arrival,l.notes].some(c=>!c)||l.id!==1||(requireOrdered&&l.ordered!==2)||(l.priority&&l.priority!==3))
      throw Error(sh.getName()+': headers do not match the clean template. No columns are moved automatically.');
    return l;
  },
  rows(sh){
    const l=this.layout(sh),n=Math.max(0,sh.getLastRow()-l.first+1);
    return {l,rows:n?sh.getRange(l.first,1,n,l.width).getValues():[]};
  },
  vendorMetadata(sh,key){
    return sh.getDeveloperMetadata().find(entry=>entry.getKey()===key);
  },
  customVendorCode(sh){
    return this.vendorMetadata(sh,this.vendorCodeKey)?.getValue()||'';
  },
  isVendor(sh){
    return Boolean(sh&&(this.vendors.includes(sh.getName())||this.customVendorCode(sh)));
  },
  vendorNames(book){
    const aliases=new Map(),custom=[];
    for(const sh of book.getSheets()){
      const name=sh.getName();
      if(this.vendors.includes(name)||['Control','Projects'].includes(name)||!this.customVendorCode(sh))continue;
      const origin=this.vendorMetadata(sh,this.vendorOriginKey)?.getValue();
      if(!origin){custom.push(name);continue;}
      const original=book.getSheetByName(origin);
      if(aliases.has(origin)||(original&&original.getName()!==name))
        throw Error('Multiple tabs claim the built-in vendor '+origin+'.');
      aliases.set(origin,name);
    }
    const control=book.getSheetByName('Control');
    const removed=control&&this.vendorMetadata(control,this.removedAmazonKey);
    return [...this.vendors.filter(name=>name!=='Amazon'||!removed).map(name=>aliases.get(name)||name),...custom];
  },
  code(company){
    const key=this.text(company).toLowerCase().replace(/[^a-z0-9]/g,'');
    const match=Object.keys(this.codes).find(n=>n.toLowerCase().replace(/[^a-z0-9]/g,'')===key);
    const code=match?this.codes[match]:key.toUpperCase().slice(0,6);
    if(!code)throw Error('Company is required for Others/Shipment.');return code;
  },
  run(action,...args){
    const lock=LockService.getDocumentLock();
    if(!lock)throw Error('Paste this file into a spreadsheet-bound Apps Script project.');
    lock.waitLock(30000);try{return this[action](SpreadsheetApp.getActiveSpreadsheet(),...args);}finally{lock.releaseLock();}
  },
  upgrade(book){
    // Only the known Stage 1 schema is migrated. Native insertion shifts all
    // purchases, formulas, validation and the old B1 selector together.
    const pending=[];
    for(const name of this.vendors){
      const sh=book.getSheetByName(name);if(!sh)continue;
      const l=this.layout(sh,false);if(l.ordered){if(l.ordered!==2)throw Error(name+': Ordered on must be column B.');continue;}
      const expected=name==='Others'
        ?['Item number','Company','Item','Link','Quantity','Individual price','Final price','Arrive_lab','Category','Reported on','Paid','Arrived','Notes']
        :name==='Shipment'
        ?['Item number','Unit','Company','Description','Date','Unit price','Final price','Order status','Reported on','Paid','Arrived','Arrive_lab','Notes','Category']
        :['Item number','Item','Link','Quantity','Individual price','Final price','Arrive_lab','Category','Reported on','Paid','Arrived','Notes'];
      const headers=sh.getRange(l.first-1,1,1,l.width).getDisplayValues()[0].map(v=>this.text(v));
      if(JSON.stringify(headers)!==JSON.stringify(expected))throw Error(name+': unknown old layout; no automatic column insertion.');
      pending.push({sh,first:l.first});
    }
    for(const {sh,first} of pending){
      sh.insertColumnAfter(1);
      sh.getRange(first-1,3).copyTo(sh.getRange(first-1,2),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
      sh.getRange(first-1,2).setValue('Ordered on');
      sh.setColumnWidth(2,145);
    }
  },
  format(book,requested=null){
    const headers=name=>{
      const regular=['Item','Link','Quantity','Individual price','Final price','Arrive_lab','Category','Reported on','Paid','Arrived','Notes'];
      if(name==='Shipment')return ['Item number','Ordered on','Priority','Unit','Company','Description','Date','Unit price','Final price','Order status','Reported on','Paid','Arrived','Arrive_lab','Notes','Category'];
      return ['Item number','Ordered on','Priority',...(name==='Others'?['Company']:[]),...regular];
    };
    const purchaseHeaders=row=>row.filter(header=>!['Priority','Tracking ID','Purchase Record ID','Formula Key','Position Key'].includes(header)).join('|');
    const registryHeaders=['Project','Spreadsheet URL or ID','Enabled','Notes'];
    const registry=book.getSheetByName('Projects');
    if(registry&&registry.getRange(1,1,1,4).getDisplayValues()[0].join('|')!==registryHeaders.join('|'))
      throw Error('Existing Projects tab has unexpected headers. Use a fresh workbook or restore the expected headers.');
    const created=[];
    for(const name of this.vendors){
      const sh=book.getSheetByName(name);
      if(!sh)continue;
      const l=this.layout(sh);
      if(purchaseHeaders(sh.getRange(l.first-1,1,1,l.width).getDisplayValues()[0])!==purchaseHeaders(headers(name)))
        throw Error(name+': existing headers do not match the purchase workbook schema.');
    }
    for(const name of this.vendors){
      if(book.getSheetByName(name)||requested&&!requested.includes(name))continue;
      created.push(name);
      this.createVendorSheet(book.insertSheet(name),headers(name));
    }
    for(const name of requested?created:this.vendors){
      const sh=book.getSheetByName(name),l=this.layout(sh);
      const category=this.categoryRule();
      sh.getRange(l.first,l.category,sh.getMaxRows()-l.first+1,1).setDataValidation(category);
    }
    if(!book.getSheetByName('Control')&&(!requested||requested.includes('Control'))){
      const sh=book.insertSheet('Control');
      sh.getRange('A1').setValue('Purchase workbook settings').setFontWeight('bold');
      sh.getRange('A3:A4').setValues([['Project name'],['Project code']]).setFontWeight('bold');
      sh.getRange('B3:B4').setNumberFormat('@');
      sh.setColumnWidth(1,185);sh.setColumnWidth(2,475);
    }
    if(!registry&&(!requested||requested.includes('Projects'))){
      const sh=book.insertSheet('Projects');
      sh.getRange('A1:D1').setValues([registryHeaders]).setBackground('#d6dce5').setFontWeight('bold');
      sh.getRange(2,3,sh.getMaxRows()-1,1).insertCheckboxes();
      sh.setFrozenRows(1);sh.setColumnWidth(2,380);sh.setColumnWidth(4,300);
    }
  },
  createVendorSheet(sh,fields){
    const name=sh.getName(),row=name==='Shipment'?4:3;
    sh.getRange('A1').setFormula(this.grantFormula).setNote('Shared grant number from Control!B5.');
    sh.getRange(row,1,1,fields.length).setValues([fields]);
    const l=this.layout(sh),n=sh.getMaxRows()-row;
    for(const column of [l.priority,l.paid,l.arrived])sh.getRange(row+1,column,n,1).insertCheckboxes();
    return l;
  },
  vendorInput(book,request,current=null){
    const name=this.text(request?.name),code=this.text(request?.code).toUpperCase();
    if(!name||name.length>100||/^[=+@-]|[\[\]:*?\/\\]/.test(name))throw Error('Enter a valid company name for the tab.');
    if(!/^[A-Z0-9]{2,6}$/.test(code))throw Error('Use a vendor code of 2–6 letters or digits.');
    const sheets=book.getSheets(),normalize=value=>value.toLowerCase().replace(/[^a-z0-9]/g,'');
    const key=normalize(name);
    if(!key)throw Error('Include at least one letter or digit in the company name.');
    const currentName=current?.getName(),origin=current&&(this.vendorMetadata(current,this.vendorOriginKey)?.getValue()||currentName);
    if([...this.vendors,'Control','Projects',...Object.keys(this.codes)].some(existing=>existing!==origin&&normalize(existing)===key)||
        sheets.some(sh=>sh.getName()!==currentName&&normalize(sh.getName())===key))
      throw Error('A tab with that name already exists or is reserved.');
    if(Object.entries(this.codes).some(([vendor,reserved])=>vendor!==origin&&reserved===code)||
        sheets.some(sh=>sh.getName()!==currentName&&this.customVendorCode(sh)===code))
      throw Error('That vendor code is already in use. Choose another.');
    return {name,code};
  },
  addVendor(book,request){
    this.project(book);
    const {name,code}=this.vendorInput(book,request);
    const sh=book.insertSheet(name);
    try{
      const l=this.createVendorSheet(sh,['Item number','Ordered on','Priority','Item','Link','Quantity','Individual price',
        'Final price','Arrive_lab','Category','Reported on','Paid','Arrived','Notes']);
      const n=sh.getMaxRows()-l.first+1;
      sh.getRange(l.first,l.category,n,1).setDataValidation(this.categoryRule());
      for(const [column,field] of [[l.ordered,'ordered'],[l.month,'month']])
        sh.getRange(l.first,column,n,1).setNumberFormat('yyyy-mm-dd').setDataValidation(this.dateRule(field));
      sh.getRange(l.first,l.arrival,n,1).setNumberFormat('yyyy-mm-dd');
      this.restoreFinalPrice(sh,l);
      this.prepareTracking(sh,l);
      this.styleVendor(sh);
      this.statusFormat(sh);
      sh.setColumnWidth(l.priority,95);
      sh.addDeveloperMetadata(this.vendorCodeKey,code);
      this.ids(book,false,{sh,l,first:l.first,last:sh.getMaxRows()});
      return 'Created '+name+' ('+code+').';
    }catch(error){book.deleteSheet(sh);throw error;}
  },
  renameableVendor(book,tab){
    const sh=book.getSheetByName(tab);
    if(!this.isVendor(sh)||['Others','Shipment','Control','Projects'].includes(tab))
      throw Error('Open a regular vendor tab. Others and Shipment have special roles and keep their names.');
    return sh;
  },
  renameVendor(book,request){
    const sh=this.renameableVendor(book,request?.tab),oldName=sh.getName();
    if(sh.getSheetId()!==request.id)throw Error('This tab changed. Reopen Rename vendor tab.');
    const {name,code}=this.vendorInput(book,request,sh);
    const codeEntry=this.vendorMetadata(sh,this.vendorCodeKey),oldCode=codeEntry?.getValue();
    if(name===oldName&&code===(oldCode||this.code(oldName)))return 'Vendor tab is already up to date.';
    const originEntry=this.vendorMetadata(sh,this.vendorOriginKey);
    const builtIn=this.vendors.includes(oldName)&&!originEntry;
    if((originEntry?.getValue()||oldName)==='Amazon')this.repairGrantLinks(book,this.vendorNames(book));
    const rewrite=()=>{
      const l=this.layout(sh);
      this.ids(book,false,{sh,l,first:l.first,last:sh.getMaxRows()});
      this.ids(book,false,null,null,'');
    };
    try{
      if(codeEntry)codeEntry.setValue(code);
      else sh.addDeveloperMetadata(this.vendorCodeKey,code);
      if(builtIn)sh.addDeveloperMetadata(this.vendorOriginKey,oldName);
      if(name!==oldName)sh.setName(name);
      sh.getRange('A2').setValue(name.toUpperCase());
      rewrite();
    }catch(error){
      try{
        if(sh.getName()!==oldName)sh.setName(oldName);
        if(codeEntry)codeEntry.setValue(oldCode);
        else this.vendorMetadata(sh,this.vendorCodeKey)?.remove();
        if(builtIn)this.vendorMetadata(sh,this.vendorOriginKey)?.remove();
        sh.getRange('A2').setValue(oldName.toUpperCase());
        rewrite();
      }catch(repairError){throw Error(error.message+' Rename rollback needs review: '+repairError.message);}
      throw error;
    }
    return 'Vendor updated to '+name+' ('+code+').';
  },
  vendorPicker(rename=false){
    const title=rename?'Rename vendor tab':'Add vendor tab';
    const book=rename?SpreadsheetApp.getActiveSpreadsheet():null;
    const sh=book?this.renameableVendor(book,book.getActiveSheet().getName()):null;
    const target=sh?{tab:sh.getName(),id:sh.getSheetId(),code:this.customVendorCode(sh)||this.code(sh.getName())}:null;
    const html=`<!doctype html><html><head><base target="_top"><style>
      body{font:14px Arial,sans-serif;padding:16px;color:#202124}
      label{display:block;margin:0 0 14px;font-weight:bold}
      input{display:block;box-sizing:border-box;width:100%;padding:8px;margin-top:5px;font:inherit}
      button{padding:8px 14px;margin-right:6px}#status{white-space:pre-wrap}
    </style></head><body><form id="vendor">
      <label>Company name<input name="company" required maxlength="100" autofocus></label>
      <label>Vendor code<input name="code" required pattern="[A-Za-z0-9]{2,6}" maxlength="6" placeholder="2–6 letters or digits"></label>
      <p id="hint" hidden>Future orders and summaries use the new company name. Changing the code changes existing Item numbers. Previously generated files stay as saved.</p>
      <button type="submit">Create vendor tab</button>
      <button type="button" id="close">Cancel</button>
    </form><p id="status" role="status" aria-live="polite"></p>
    <script>
      const form=document.getElementById('vendor'),status=document.getElementById('status');
      const close=document.getElementById('close'),submit=form.querySelector('[type=submit]');
      const target=__TARGET__;
      if(target){
        form.elements.company.value=target.tab;form.elements.code.value=target.code;
        submit.textContent='Save vendor';document.getElementById('hint').hidden=false;
      }
      close.onclick=()=>google.script.host.close();
      form.onsubmit=event=>{
        event.preventDefault();
        if(submit.disabled)return;
        submit.disabled=close.disabled=true;status.textContent=target?'Updating vendor tab…':'Creating vendor tab…';
        const runner=google.script.run.withSuccessHandler(message=>{
          status.textContent=message;submit.hidden=true;close.textContent='Close';close.disabled=false;
        }).withFailureHandler(error=>{
          status.textContent=error.message;submit.disabled=close.disabled=false;
        });
        const fields={name:form.elements.company.value,code:form.elements.code.value};
        if(target)runner.renamePurchaseVendor({...target,...fields});
        else runner.addPurchaseVendor(fields);
      };
    </script></body></html>`;
    const payload=JSON.stringify(target).replace(/</g,'\\u003c');
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html.replace('__TARGET__',()=>payload))
      .setWidth(420).setHeight(rename?410:310),title);
  },
  removableVendor(book,tab){
    const sheet=book.getSheetByName(tab);
    if(sheet&&(tab==='Amazon'||this.vendorMetadata(sheet,this.vendorOriginKey)?.getValue()==='Amazon'))return sheet;
    if(!sheet||[...this.vendors,'Control','Projects'].includes(sheet.getName())||
        !this.customVendorCode(sheet)||this.vendorMetadata(sheet,this.vendorOriginKey))
      throw Error('Open Amazon or a custom vendor tab created with Add vendor tab.');
    return sheet;
  },
  removeVendor(book,request){
    if(request?.confirmation!=='DELETE_TAB')throw Error('Type DELETE_TAB exactly to remove the vendor tab.');
    const sheet=this.removableVendor(book,request.tab);
    if(sheet.getSheetId()!==request.id)throw Error('This tab changed. Reopen Remove vendor tab.');
    const name=sheet.getName();
    const amazon=(this.vendorMetadata(sheet,this.vendorOriginKey)?.getValue()||name)==='Amazon';
    const control=book.getSheetByName('Control');
    if(amazon)this.repairGrantLinks(book,this.vendorNames(book));
    try{
      this.ids(book,false,null,null,name);
      if(amazon)control.addDeveloperMetadata(this.removedAmazonKey,'true');
      book.deleteSheet(sheet);
    }catch(error){
      if(book.getSheetByName(name)?.getSheetId()===request.id){
        if(amazon)this.vendorMetadata(control,this.removedAmazonKey)?.remove();
        try{this.ids(book,false,null,null,'');}
        catch(repairError){throw Error(error.message+' Item number formulas could not be restored. Run Refresh / repair item IDs.');}
      }
      throw error;
    }
    return 'Removed '+name+'.';
  },
  removeVendorPicker(){
    const book=SpreadsheetApp.getActiveSpreadsheet(),sheet=this.removableVendor(book,book.getActiveSheet().getName());
    const target=JSON.stringify({tab:sheet.getName(),id:sheet.getSheetId()}).replace(/</g,'\\u003c');
    const html=`<!doctype html><html><head><base target="_top"><style>
      body{font:14px Arial,sans-serif;padding:16px;color:#202124}
      input{box-sizing:border-box;width:100%;padding:8px;margin:8px 0 14px;font:inherit}
      button{padding:8px 14px;margin-right:6px}#remove{color:#b3261e}
    </style></head><body><form id="removal">
      <p>Remove <strong id="tab"></strong>?</p>
      <p><strong>This is irreversible. Purchase Tools cannot recover the tab or its data.</strong></p>
      <p>Saved orders and reports stay in Drive. Future and regenerated summaries omit this tab's purchases. Any custom formulas that reference this tab may break.</p>
      <label>Type <strong>DELETE_TAB</strong> to confirm<input id="confirmation" required autocomplete="off"></label>
      <button type="submit" id="remove" disabled>Delete tab</button>
      <button type="button" id="close">Cancel</button>
    </form><p id="status" role="status" aria-live="polite"></p>
    <script>
      const target=__TARGET__;
      const byId=id=>document.getElementById(id);
      const input=byId('confirmation'),remove=byId('remove'),close=byId('close'),status=byId('status');
      byId('tab').textContent=target.tab;
      input.oninput=()=>{remove.disabled=input.value!=='DELETE_TAB';};
      close.onclick=()=>google.script.host.close();
      byId('removal').onsubmit=event=>{
        event.preventDefault();
        if(remove.disabled||input.value!=='DELETE_TAB')return;
        remove.disabled=close.disabled=input.disabled=true;status.textContent='Removing vendor tab…';
        google.script.run.withSuccessHandler(message=>{
          status.textContent=message;remove.hidden=true;close.textContent='Close';close.disabled=false;
        }).withFailureHandler(error=>{
          status.textContent=error.message;remove.disabled=close.disabled=input.disabled=false;
        }).removePurchaseVendor({...target,confirmation:input.value});
      };
    </script></body></html>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html.replace('__TARGET__',()=>target)).setWidth(450).setHeight(370),'Remove vendor tab');
  },
  dateRule(field){
    return SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText(this.dateHelp[field]).build();
  },
  categoryRule(){
    return SpreadsheetApp.newDataValidation().requireValueInList(this.categories,true).setAllowInvalid(true)
      .setHelpText('Choose a category, leave blank, or type a custom category.').build();
  },
  formulaLocale(book){
    // The portable workbook and report formulas use the US argument separator.
    if(book.getSpreadsheetLocale()!=='en_US'){
      book.setSpreadsheetLocale('en_US');
      SpreadsheetApp.flush();
    }
  },
  styleVendor(sh){
    const l=this.layout(sh),header=l.first-1,count=sh.getMaxRows()-header;
    const clip=SpreadsheetApp.WrapStrategy.CLIP;
    sh.getRange('A1:A2').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontWeight('bold').setFontSize(14);
    sh.getRange(2,1,1,l.width).setBackground('#eff2f7');
    sh.getRange('A2').setValue(sh.getName()==='Shipment'?'SHIPMENTS':sh.getName().toUpperCase()).setFontSize(13);
    sh.getRange(header,1,1,l.width).setBackground('#d8dde5').setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrapStrategy(clip);
    sh.setFrozenRows(header);
    sh.setRowHeights(1,2,35);
    sh.setRowHeight(header,32);
    sh.getRange(l.first,1,count,l.width).setVerticalAlignment('middle');
    for(const column of [l.id,l.ordered,l.priority,l.qty,l.unit,l.total,l.arrival,l.month,l.paid,l.arrived]){
      sh.getRange(l.first,column,count,1).setHorizontalAlignment('center');
    }
    for(const column of [l.item,l.link,l.company].filter(Boolean))sh.getRange(l.first,column,count,1).setWrapStrategy(clip);
    for(const column of [l.category,l.notes])sh.getRange(l.first,column,count,1).setWrap(true);
    for(const column of [l.unit,l.total])sh.getRange(l.first,column,count,1).setNumberFormat('$#,##0.00##');
    for(const [column,width] of [[l.id,260],[l.ordered,150],[l.item,450],[l.link,320],[l.company,160],
      [l.qty,95],[l.unit,135],[l.total,125],[l.arrival,145],[l.month,145],[l.category,180],[l.notes,310]]){
      if(column)sh.setColumnWidth(column,width);
    }
  },
  formatControl(sh,assets){
    sh.setHiddenGridlines(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,225);
    sh.setColumnWidth(2,650);
    sh.getRange('A1:B1').setBackground('#d8dde5').setFontSize(14).setFontWeight('bold');
    sh.getRange('A1').setValue('PURCHASE WORKFLOW');
    sh.getRange('A2').setValue('Use Purchase Tools to set up this workbook, create orders, and generate reports.');
    sh.getRange('A2:B26').setWrap(true).setVerticalAlignment('middle');
    sh.getRange('A3:A26').setBackground('#eff2f7').setFontWeight('bold');
    sh.getRange('A5').setValue('Grant number');
    sh.getRange('B3:B5').setBackground('#c1eac9').setNumberFormat('@').setFontSize(12).setFontWeight('bold');
    for(const [row,label,description] of [
      [6,'Monthly summary','Choose a month and year in Purchase Tools. The report includes enabled projects.'],
      [7,'Other projects','Use Sync projects in this folder to manage the Projects list.'],
      [9,'Grant','Set the grant in Control!B5 or in the setup window. All vendor tabs read it from Control.'],
      [10,'Paid / Arrived','Checking both boxes turns the row green. Arrived saves a date. Unchecking it clears the date.'],
      [11,'Received Log','Setup enables logging for individual Paid, Arrived, and Notes edits.'],
      [18,'Order selection','Open a vendor tab and enter item numbers, ranges, or ALL in the order dialog.'],
      [19,'Project code convention','Use a unique code of 2–10 letters or digits for each project.'],
      [20,'Purchase IDs','Item numbers follow row order. The fixed Purchase Record ID follows the purchase.'],
      [23,'Move purchases','Move the complete row so the order date and Purchase Record ID stay with the purchase.'],
      [26,'Setup / repair','Use Purchase Tools to repair formulas and formatting or change project settings.']
    ])sh.getRange(row,1,1,2).setValues([[label,description]]);
    this.controlLinks(sh,assets);
    sh.setRowHeight(1,30);
    sh.autoResizeRows(2,25);
  },
  controlLinks(sh,assets,check=false){
    const existing=check?sh.getRange('B12:B22').getRichTextValues():null;
    let changed=false;
    for(const [row,label,key] of [
      [12,'Order Docs','orders'],[13,'Monthly Summaries','summaries'],
      [14,'Order document template','template'],[15,'Blank workbook template','workbookTemplate'],
      [21,'Archived summaries','archive'],[22,'Shared Received Log','receivedLog']
    ]){
      if(check&&existing[row-12][0]?.getLinkUrl()===assets[key].getUrl())continue;
      sh.getRange(row,1).setValue(label);
      const cell=sh.getRange(row,2);
      if(assets){
        cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Open '+label)
          .setLinkUrl(assets[key].getUrl()).build());
        changed=true;
      }else if(!cell.getDisplayValue()){
        cell.setValue('Run Initialize Drive templates and folders.');
      }
    }
    return changed;
  },
  setupSettings(book){
    return [
      ['Control','B5','Grant number'],
      ['Control','B3','Project name'],
      ['Control','B4','Project code']
    ].map(([tab,cell,label])=>({
      tab,cell,label,current:book.getSheetByName(tab)?.getRange(cell).getDisplayValue()||''
    }));
  },
  repairGrantLinks(book,vendors,changes=[]){
    const control=book.getSheetByName('Control');
    if(!control)throw Error('Restore the Control tab before updating grant links.');
    const target=control.getRange('B5'),legacy=book.getSheetByName('Amazon')?.getRange('A1');
    if(!this.text(target.getDisplayValue())&&legacy&&legacy.getFormula()!==this.grantFormula){
      const value=this.text(legacy.getDisplayValue());
      if(value.startsWith('#')||value.startsWith('='))throw Error('Correct the legacy grant in Amazon!A1 or enter it in Control!B5.');
      if(value){
        target.setNumberFormat('@').setValue(value);
        changes.push('Moved the shared grant to Control!B5.');
      }
    }
    const label=control.getRange('A5');
    if(label.getDisplayValue()!=='Grant number'){
      label.setValue('Grant number');
      target.setBackground('#c1eac9').setNumberFormat('@').setFontSize(12).setFontWeight('bold');
      control.getRange('B9').setValue('Set the grant in Control!B5 or in the setup window. All vendor tabs read it from Control.');
      changes.push('Added the shared grant field to Control.');
    }
    for(const name of vendors){
      const grant=book.getSheetByName(name)?.getRange('A1');
      if(grant&&grant.getFormula()!==this.grantFormula){
        grant.setFormula(this.grantFormula).setNote('Shared grant number from Control!B5.');
        changes.push(name+': restored grant link.');
      }
    }
  },
  progressPicker(title,working,handler,measured=false){
    const session=measured?Utilities.getUuid():'';
    const html=`<!doctype html><html><head><base target="_top"><style>
  body{font:14px Arial,sans-serif;color:#202124;padding:16px}
  progress{width:100%;accent-color:#1a73e8}
  #status{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.4}
  #status.error{color:#b3261e}
  button{padding:8px 13px;margin-right:6px}
</style></head><body>
  <progress id="progress" aria-label="Working"></progress>
  <p id="status" role="status" aria-live="polite"></p>
  <p id="confirmation" hidden>Recreate the missing tabs as blank layouts? This cannot recover deleted purchases.</p>
  <button id="recreate" hidden>Recreate missing tabs</button><button id="close" disabled>Close</button>
  <script>
  const byId=id=>document.getElementById(id);
  const progress=byId('progress'),status=byId('status'),close=byId('close');
  const recreate=byId('recreate'),confirmation=byId('confirmation');
  const handler=${JSON.stringify(handler)},working=${JSON.stringify(working)},session=${JSON.stringify(session)};
  let attempt=0;
  close.onclick=()=>google.script.host.close();
  function run(request){
    const token=session?session+'-'+(++attempt):'';
    let finished=false,timer;
    progress.hidden=false;close.disabled=true;recreate.hidden=true;confirmation.hidden=true;
    status.className='';status.textContent=Array.isArray(request)?'Recreating missing tabs…':working;
    const finish=()=>{finished=true;if(timer)clearTimeout(timer);progress.hidden=true;close.disabled=false;};
    if(token){
      progress.max=1;progress.value=0;
      const next=()=>{if(!finished)timer=setTimeout(poll,1200);};
      const poll=()=>google.script.run.withFailureHandler(next).withSuccessHandler(update=>{
        if(finished)return;
        if(update){
          progress.max=update.total;progress.value=update.completed;
          status.textContent=update.label+' ('+update.completed+' / '+update.total+' steps complete)';
        }
        next();
      }).readPurchaseCheckProgress(token);
      next();
    }
    google.script.run.withFailureHandler(error=>{
      finish();status.className='error';status.textContent=error.message||String(error);
    }).withSuccessHandler(result=>{
      finish();status.textContent=typeof result==='string'?result:result.message;
      if(result.missing?.length){
        confirmation.hidden=false;recreate.hidden=false;
        recreate.onclick=()=>run(result.missing);
      }
    })[handler](request,token);
  }
  run(true);
  </script></body></html>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(420),title);
  },
  setupPicker(){
    const settings=this.setupSettings(SpreadsheetApp.getActiveSpreadsheet());
    const html=`<!doctype html><html><head><style>
  body{font:14px Arial,sans-serif;color:#202124;margin:0;padding:22px}
  label{display:block;font-weight:bold;margin-top:14px}
  input{box-sizing:border-box;width:100%;padding:9px;margin-top:6px;border:1px solid #dadce0;border-radius:5px;font:inherit}
  .current{color:#5f6368;margin:5px 0;overflow-wrap:anywhere}
  .actions{display:flex;justify-content:flex-end;gap:8px;margin-top:20px}
  button{font:inherit;border:1px solid #dadce0;border-radius:5px;padding:9px 14px;background:white;cursor:pointer}
  button.primary{background:#1a73e8;color:white;border-color:#1a73e8}button:disabled{opacity:.55;cursor:default}
  #status,#changes{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.4}
  #status.error{color:#b3261e}
</style></head><body>
<form id="form">
  <div id="fields"></div>
  <div class="actions">
    <button type="button" id="cancel">Cancel</button>
    <button type="button" id="skip">Skip</button>
    <button type="submit" class="primary">Continue</button>
  </div>
  <p class="current">Skip runs setup using the current values.</p>
</form>
<section id="confirm" hidden>
  <strong>Overwrite these current values?</strong>
  <p id="changes"></p>
  <div class="actions"><button type="button" id="no">No, go back</button><button type="button" class="primary" id="yes">Yes, overwrite</button></div>
</section>
<p id="status" role="status" aria-live="polite"></p>
<progress id="progress" aria-label="Initializing workbook" style="width:100%;accent-color:#1a73e8" hidden></progress>
<button type="button" id="close" hidden>Close</button>
<script>
  const settings=__SETTINGS__;
  const byId=id=>document.getElementById(id);
  const form=byId('form'),confirm=byId('confirm'),status=byId('status');
  const inputs=settings.map(setting=>{
    const label=document.createElement('label'),input=document.createElement('input'),current=document.createElement('p');
    label.textContent=setting.label+' ('+setting.tab+'!'+setting.cell+')';
    input.value=setting.current;input.required=true;
    label.appendChild(input);
    current.className='current';current.textContent='Current value: '+(setting.current||'(blank)');
    byId('fields').append(label,current);
    return input;
  });
  let proposed,busy=false;
  function run(request){
    if(busy)return;
    busy=true;form.hidden=true;confirm.hidden=true;byId('progress').hidden=false;
    status.className='';status.textContent='Setting up the workbook…';
    google.script.run.withSuccessHandler(message=>{
      byId('progress').hidden=true;status.textContent=message;byId('close').hidden=false;
    }).withFailureHandler(error=>{
      byId('progress').hidden=true;busy=false;form.hidden=false;status.className='error';status.textContent=error.message;
    }).setupPurchaseWorkbook(request);
  }
  form.addEventListener('submit',event=>{
    event.preventDefault();
    proposed=inputs.map(input=>input.value.trim());
    proposed[2]=proposed[2].toUpperCase();
    if(proposed.some(value=>!value||value.startsWith('='))||!(/^[A-Z0-9]{2,10}$/).test(proposed[2])){
      status.className='error';status.textContent='Enter all three values. Project code must use 2–10 letters or numbers.';return;
    }
    const changes=settings.flatMap((setting,i)=>setting.current&&setting.current!==proposed[i]
      ?[setting.label+': '+setting.current+' → '+proposed[i]]:[]);
    if(changes.length){
      byId('changes').textContent=changes.join(String.fromCharCode(10));
      status.textContent='';form.hidden=true;confirm.hidden=false;
    }else save(false);
  });
  function save(overwrite){run({values:proposed,previous:settings.map(s=>s.current),overwrite});}
  byId('yes').onclick=()=>save(true);
  byId('no').onclick=()=>{confirm.hidden=true;form.hidden=false;};
  byId('skip').onclick=()=>run({skip:true});
  byId('cancel').onclick=byId('close').onclick=()=>google.script.host.close();
</script></body></html>`;
    const safeSettings=JSON.stringify(settings).replace(/</g,'\\u003c');
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html.replace('__SETTINGS__',()=>safeSettings)).setWidth(510).setHeight(530),
      'Initialize workbook'
    );
  },
  removeEmptyStarter(book,starter){
    if(!starter||book.getSheets().length<=1||starter.getLastRow()!==0)return;
    const notes=starter.getRange(1,1,starter.getMaxRows(),starter.getMaxColumns()).getNotes();
    if(!notes.some(row=>row.some(Boolean)))book.deleteSheet(starter);
  },
  setup(book,request,templateOnly=false){
    const before=book.getSheets();
    if(!templateOnly&&before.some(sh=>sh.getLastRow()>0||this.vendors.includes(sh.getName())))
      throw Error('Initialize workbook is for blank workbooks. Use Check / repair workbook for an existing project.');
    let settings=[],values=[];
    if(request&&request.skip!==true){
      settings=this.setupSettings(book);
      if(!Array.isArray(request.values)||request.values.length!==settings.length||
         !request.values.every(value=>typeof value==='string')){
        throw Error('Enter the grant number, project name, and project code.');
      }
      values=request.values.map(value=>this.text(value));
      values[2]=values[2].toUpperCase();
      if(values.some(value=>!value||value.startsWith('='))||!/^[A-Z0-9]{2,10}$/.test(values[2])){
        throw Error('Enter all three values. Project code must use 2–10 letters or numbers.');
      }
      if(!Array.isArray(request.previous)||settings.some((setting,i)=>setting.current!==request.previous[i])){
        throw Error('Setup values changed while this window was open. Close it and reopen setup.');
      }
      if(request.overwrite!==true&&settings.some((setting,i)=>setting.current&&setting.current!==values[i])){
        throw Error('Confirm before overwriting existing setup values.');
      }
    }
    this.formulaLocale(book);
    this.upgrade(book); // Insert Ordered on in older Stage 1 copies before applying the new template.
    this.format(book);
    const warnings=[];
    const checkboxRule=SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    const orderedRule=this.dateRule('ordered'),reportedRule=this.dateRule('month');
    for(const name of this.vendors){
      const sh=book.getSheetByName(name);
      const previous=this.layout(sh);
      if(!previous.priority){
        sh.insertColumnAfter(2);
        sh.getRange(previous.first-1,4).copyTo(sh.getRange(previous.first-1,3),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
        sh.getRange(previous.first-1,3).setValue('Priority');
        sh.getRange(previous.first,3,sh.getMaxRows()-previous.first+1,1).setValue(false);
      }
      const l=this.layout(sh);
      // The order dialog replaces the old selector cells.
      for(const address of ['C1','D1','E1']){
        sh.getRange(address).clearContent().clearNote();
      }
      sh.setColumnWidth(l.priority,95);
      const purchaseRows=sh.getMaxRows()-l.first+1;
      for(const column of [l.priority,l.paid,l.arrived]){
        sh.getRange(l.first,column,purchaseRows,1).setDataValidation(checkboxRule);
      }
      for(const column of [l.arrival,l.month]){
        sh.getRange(l.first,column,purchaseRows,1).setNumberFormat('yyyy-mm-dd');
      }
      sh.getRange(l.first,l.ordered,purchaseRows,1).setNumberFormat('yyyy-mm-dd').setDataValidation(orderedRule);
      sh.getRange(l.first,l.month,purchaseRows,1).setDataValidation(reportedRule);
      sh.getRange(l.first-1,l.ordered).setNote('Enter the actual order date once. Move it with the complete row. Blank dates are not filled automatically.');
      sh.getRange(l.first-1,l.arrival).setNote('Checking Arrived saves a fixed arrival date. Unchecking clears it; rechecking stamps a new date. Stored date values do not change daily.');
      this.restoreFinalPrice(sh,l);
      this.prepareTracking(sh,l);
      this.styleVendor(sh);
      this.statusFormat(sh);
    }
    settings.forEach((setting,i)=>{
      if(setting.current!==values[i]){
        book.getSheetByName(setting.tab).getRange(setting.cell).setNumberFormat('@').setValue(values[i]);
      }
    });
    this.repairGrantLinks(book,this.vendors);
    const repaired=this.ids(book);
    try{
      if(!templateOnly)this.enableReceivedLog(book);
    }catch(error){
      warnings.push('Could not enable Received Log: '+error.message);
    }
    let assets;
    if(!templateOnly){
      try{assets=this.driveAssets(book,'references');}
      catch(error){warnings.push('Could not refresh Control links: '+error.message);}
    }
    this.formatControl(book.getSheetByName('Control'),assets);
    if(!templateOnly&&before.length===1&&/^(?:Sheet\s*1|Hoja\s*1)$/i.test(before[0].getName())){
      this.removeEmptyStarter(book,before[0]);
    }
    const message=[
      'Template formatting, formulas and tracking checked.',
      'ID cells updated or cleared: '+repaired.toLocaleString('en-US'),
      '',
      'HOW IDs WORK',
      '• New items receive a number automatically.',
      '• Numbers follow the current item order.',
      '• Enter Ordered on once; its date stays with the row.',
      '',
      'BEFORE GENERATING ORDERS OR SUMMARIES',
      'Check the project name/code in Control!B3/B4',
      'and the grant in Control!B5.',
      '',
      'Paid + Arrived turns the row green. Arrived stamps a fixed date; unchecking clears it.',
      'Use the Purchase Tools prompt to select items for an order.',
      ...(warnings.length?['','NOTICES',...warnings.map(w=>'• '+w)]:[])
    ].join('\n');
    if(request)return 'Setup complete\n\n'+message;
    const ui=SpreadsheetApp.getUi();
    ui.alert('Setup complete',message,ui.ButtonSet.OK);
  },
  finalPriceFormula(l){
    const letter=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
    const [item,qty,unit]=[l.item,l.qty,l.unit].map(letter);
    const span=column=>column+l.first+':'+column;
    const [items,quantities,prices]=[item,qty,unit].map(span);
    return '=ARRAYFORMULA(IF(('+items+'="")+('+quantities+'="")+('+prices+'=""),"",IFERROR('+quantities+'*'+prices+',"")))';
  },
  checkWorkbook(book,progress=()=>{}){
    const changes=[],notices=[],missing=[],layouts=new Map();
    const vendors=this.vendorNames(book);
    const managedTabs=[...vendors,'Control','Projects'];
    const total=vendors.length+3;
    let completed=0;
    progress(completed,total,'Checking workbook structure');
    if(managedTabs.every(name=>!book.getSheetByName(name)))
      throw Error('No purchase tabs found. Use Initialize workbook on a blank spreadsheet.');
    if(book.getSpreadsheetLocale()!=='en_US'){
      this.formulaLocale(book);
      changes.push('Restored the US formula locale.');
    }
    for(const name of managedTabs){
      if(!book.getSheetByName(name))missing.push(name);
    }
    if(!missing.includes('Control'))this.repairGrantLinks(book,vendors,changes);
    for(const setting of this.setupSettings(book)){
      if(!missing.includes(setting.tab)&&!this.text(setting.current))
        notices.push(setting.label+' is blank in '+setting.tab+'!'+setting.cell+'.');
    }
    completed++;
    for(const name of vendors){
      progress(completed++,total,'Checking '+name);
      const sh=book.getSheetByName(name);
      if(!sh)continue;
      let l;
      try{l=this.layout(sh);}
      catch(error){notices.push(error.message);continue;}
      if(!l.tracking||sh.getRange(l.first-1,l.tracking).getDisplayValue()!=='Purchase Record ID'){
        const action=l.tracking?'renamed':'added';
        this.prepareTracking(sh,l,false);
        l=this.layout(sh);
        changes.push(name+': '+action+' Purchase Record ID column.');
      }
      const last=sh.getLastRow(),activeRows=new Set();
      const values=last>=l.first?sh.getRange(l.first,2,last-l.first+1,l.width-1).getValues():[];
      let missingRecordId=false;
      values.forEach((row,i)=>{
        const item=this.text(row[l.item-2]),id=this.text(row[l.tracking-2]);
        if(item||id)activeRows.add(l.first+i);
        if(item&&!id)missingRecordId=true;
      });
      layouts.set(name,{sh,l,last,activeRows});
      if(missingRecordId){
        this.fillTracking(sh,l,l.first,last,true);
        changes.push(name+': assigned missing Purchase Record IDs.');
      }
      this.checkFinalPrice(sh,l,last,values,changes,notices);
      this.checkVendorStyle(sh,l,activeRows,changes);
    }
    progress(completed++,total,'Checking Item number formulas');
    if(missing.includes('Control'))notices.push('Item number repair requires the Control tab.');
    else{
      const repaired=this.ids(book,false,null,layouts);
      if(repaired)changes.push('Repaired '+repaired+' Item number formula(s).');
      if(layouts.size!==vendors.length)notices.push('Company-based Item number checks need all vendor tabs.');
    }
    progress(completed,total,'Checking Drive links and Received Log');
    let assets;
    try{assets=this.driveAssets(book,'references');}
    catch(error){notices.push('Drive assets: '+error.message);}
    const control=book.getSheetByName('Control');
    if(control){
      if(control.getRange('A1').getDisplayValue()!=='PURCHASE WORKFLOW'||control.getRange('A5').getDisplayValue()!=='Grant number'){
        this.formatControl(control,assets);
        changes.push('Restored Control layout.');
      }else if(assets&&this.controlLinks(control,assets,true))changes.push('Refreshed Control links.');
    }
    if(layouts.size){
      try{if(this.enableReceivedLog(book))changes.push('Enabled Received Log.');}
      catch(error){notices.push('Received Log: '+error.message);}
    }
    progress(total,total,'Workbook check complete');
    return {missing,message:[changes.length?changes.join('\n'):'No repairs made.',
      ...(missing.length?['Missing tabs: '+missing.join(', ')+'.']:[]),...notices].join('\n\n')};
  },
  checkFinalPrice(sh,l,last,rows,changes,notices){
    const anchor=sh.getRange(l.first,l.total),expected=this.finalPriceFormula(l);
    if(anchor.getFormula()===expected&&!anchor.getDisplayValue().startsWith('#'))return;
    const n=Math.max(1,last-l.first+1),range=sh.getRange(l.first,l.total,n,1);
    const formulas=range.getFormulas(),values=range.getValues(),typed=[];
    const col=anchor.getA1Notation().replace(/\d+$/,'');
    values.forEach(([value],i)=>{
      if(formulas[i][0]||value==='')return;
      const row=rows[i]||[],qty=this.number(row[l.qty-2]),unit=this.number(row[l.unit-2]);
      const expectedTotal=Number.isFinite(qty*unit)?(qty*unit).toFixed(2):'missing inputs';
      typed.push(col+(l.first+i)+': '+value+' (expected '+expectedTotal+')');
    });
    if(typed.length){
      notices.push(sh.getName()+': review Final price cells '+typed.slice(0,12).join(', ')+
        (typed.length>12?' and '+(typed.length-12)+' more':'')+'. Values may be typed totals or results of a displaced array.');
      return;
    }
    if(JSON.stringify(range.getFormulas())!==JSON.stringify(formulas)||JSON.stringify(range.getValues())!==JSON.stringify(values)){
      notices.push(sh.getName()+': Final price changed during the check. Run it again.');
      return;
    }
    range.clearContent();
    anchor.setFormula(expected);
    changes.push(sh.getName()+': restored Final price calculation.');
  },
  checkVendorStyle(sh,l,activeRows,changes){
    const header=sh.getRange(l.first-1,1,1,l.width);
    if(header.getBackground()!=='#d8dde5'||header.getFontWeight()!=='bold'){
      header.setBackground('#d8dde5').setFontWeight('bold').setHorizontalAlignment('left')
        .setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      changes.push(sh.getName()+': restored header styling.');
    }
    if(sh.getFrozenRows()!==l.first-1){
      sh.setFrozenRows(l.first-1);
      changes.push(sh.getName()+': restored frozen headers.');
    }
    const criteria=SpreadsheetApp.DataValidationCriteria,n=sh.getMaxRows()-l.first+1;
    const columns=[
      ...[l.priority,l.paid,l.arrived].filter(Boolean).map(c=>[c,criteria.CHECKBOX,()=>SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()]),
      [l.ordered,criteria.DATE_IS_VALID_DATE,()=>this.dateRule('ordered')],
      [l.month,criteria.DATE_IS_VALID_DATE,()=>this.dateRule('month')],
      [l.category,criteria.VALUE_IN_LIST,()=>this.categoryRule()]
    ];
    let last=l.first;
    for(const row of activeRows)last=row;
    const samples=[...new Set([l.first,last])].map(row=>sh.getRange(row,1,1,l.width).getDataValidations()[0]);
    for(const [column,type,make] of columns){
      if(samples.some(row=>{
        const rule=row[column-1];
        return !rule||rule.getCriteriaType()!==type||rule.getAllowInvalid()!==(type===criteria.VALUE_IN_LIST)||
          type===criteria.VALUE_IN_LIST&&JSON.stringify(rule.getCriteriaValues()[0])!==JSON.stringify(this.categories);
      })){
        sh.getRange(l.first,column,n,1).setDataValidation(make());
        changes.push(sh.getName()+': restored validation in column '+column+'.');
      }
    }
    if(l.tracking&&samples.some(row=>JSON.stringify(row[l.tracking-1]?.getCriteriaValues()[0])===JSON.stringify(this.categories))){
      sh.getRange(l.first,l.tracking,n,1).clearDataValidations();
      changes.push(sh.getName()+': removed a misplaced category dropdown from Purchase Record ID.');
    }
    const expected=this.statusFormula(l);
    const managed=sh.getConditionalFormatRules().filter(rule=>rule.getBooleanCondition()?.getCriteriaValues().some(value=>String(value).includes('PA_PAID_ARRIVED_V1')));
    const bounds=managed.length===1?managed[0].getRanges():[];
    if(managed.length!==1||managed[0].getBooleanCondition().getCriteriaValues()[0]!==expected||
      bounds.length!==1||bounds[0].getRow()!==l.first||bounds[0].getColumn()!==1||
      bounds[0].getNumRows()!==n||bounds[0].getNumColumns()!==l.width){
      this.statusFormat(sh,l);
      changes.push(sh.getName()+': restored Paid and Arrived color rule.');
    }
  },
  recreateMissing(book,requested,progress=()=>{}){
    const managed=new Set([...this.vendorNames(book),'Control','Projects']);
    const missing=requested.filter(name=>managed.has(name)&&!book.getSheetByName(name));
    if(!missing.length)return 'The requested tabs are already present.';
    progress(0,this.vendors.length+3,'Recreating missing tabs');
    this.format(book,missing);
    for(const name of missing.filter(name=>this.vendors.includes(name))){
      const sh=book.getSheetByName(name);
      this.prepareTracking(sh,this.layout(sh),false);
      this.styleVendor(sh);
    }
    const control=book.getSheetByName('Control');
    if(missing.includes('Control'))this.formatControl(control);
    return 'Recreated as blank: '+missing.join(', ')+'.\n\n'+this.checkWorkbook(book,progress).message;
  },
  restoreFinalPrice(sh,l){
    const formula=this.finalPriceFormula(l);
    const anchor=sh.getRange(l.first,l.total);
    if(anchor.getFormula()===formula&&!anchor.getDisplayValue().startsWith('#'))return;
    // Final price is a calculated column. Clear the whole spill area, including its old anchor.
    sh.getRange(l.first,l.total,sh.getMaxRows()-l.first+1,1).clearContent();
    anchor.setFormula(formula);
    sh.getRange(l.first-1,l.total).setNote('Calculated for the vendor tab from Quantity and Individual price. Leave this column free of manual entries.');
  },
  prepareTracking(sh,l,initialize=true){
    if(initialize)for(const column of [l.key,l.position].filter(Boolean).sort((a,b)=>b-a))sh.deleteColumn(column);
    const current=this.layout(sh),col=current.tracking||sh.getLastColumn()+1,row=l.first-1;
    if(!current.tracking){
      if(col>sh.getMaxColumns())sh.insertColumnsAfter(sh.getMaxColumns(),1);
      sh.getRange(row,current.notes).copyTo(sh.getRange(row,col),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    }
    sh.getRange(row,col).setValue('Purchase Record ID')
      .setNote('Internal purchase reference. Assigned when an item is entered. Move the full row. The ID clears only when all purchase data is cleared.');
    sh.setColumnWidth(col,170);
    sh.getRange(current.first,col,sh.getMaxRows()-current.first+1,1).setNumberFormat('@');
    if(initialize)this.fillTracking(sh,this.layout(sh),current.first,sh.getLastRow());
  },
  fillTracking(sh,l,first,last,preserve=false){
    if(!l.tracking||last<first)return;
    const rows=sh.getRange(first,2,last-first+1,l.width-1).getValues();
    const totalFormulas=sh.getRange(first,l.total,last-first+1,1).getFormulas();
    const spilled=/ARRAYFORMULA\s*\(/i.test(sh.getRange(l.first,l.total).getFormula());
    let used;
    const ignored=new Set([l.tracking,l.key,l.position,l.paid,l.arrived,l.priority]),changes=[];
    rows.forEach((values,i)=>{
      const row=first+i,current=this.text(values[l.tracking-2]);
      const hasData=values.some((value,j)=>!ignored.has(j+2)&&
        !(j+2===l.total&&(spilled||totalFormulas[i][0]))&&this.text(value)!=='');
      if(!hasData&&current){if(!preserve)changes.push({row,value:''});return;}
      if(current||!this.text(values[l.item-2]))return;
      if(!used){
        used=new Set(sh.getRange(l.first,l.tracking,Math.max(1,sh.getLastRow()-l.first+1),1).getValues()
          .map(([id])=>this.text(id)).filter(Boolean));
      }
      let id;
      for(let attempt=0;attempt<8;attempt++){
        id='T'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
        if(!used.has(id))break;
      }
      if(used.has(id))throw Error('Could not allocate a unique Purchase Record ID. Run setup again.');
      used.add(id);
      changes.push({row,value:id});
    });
    for(let i=0;i<changes.length;){
      let end=i+1;
      while(end<changes.length&&changes[end].row===changes[end-1].row+1)end++;
      sh.getRange(changes[i].row,l.tracking,end-i,1).setValues(changes.slice(i,end).map(change=>[change.value]));
      i=end;
    }
  },
  trackingEdit(e){
    if(!e?.range)return;
    const sh=e.range.getSheet();
    if(!this.isVendor(sh))return;
    const l=this.layout(sh),first=Math.max(l.first,e.range.getRow()),last=e.range.getLastRow();
    if(last<first)return;
    const lock=LockService.getDocumentLock();
    if(!lock?.tryLock(5000))return;
    try{
      const touches=columns=>columns.some(c=>c>=e.range.getColumn()&&c<=e.range.getLastColumn());
      if(touches([l.item,l.ordered,l.company])&&
        sh.getRange(first,l.id,last-first+1,1).getFormulas().some(([formula])=>!formula)){
        this.ids(sh.getParent(),false,{sh,l,first,last});
      }
      this.fillTracking(sh,l,first,last);
      SpreadsheetApp.flush();
    }finally{lock.releaseLock();}
  },
  statusFormula(l){
    const column=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
    const at=c=>'INDEX($'+column(c)+':$'+column(c)+',ROW())';
    return '=AND('+at(l.paid)+'=TRUE,'+at(l.arrived)+'=TRUE,LEN(TRIM(TO_TEXT('+at(l.item)+')))>0,N("PA_PAID_ARRIVED_V1")=0)';
  },
  statusFormat(sh,l=this.layout(sh)){
    const marker='PA_PAID_ARRIVED_V1';
    const others=sh.getConditionalFormatRules().filter(rule=>{
      const condition=rule.getBooleanCondition();
      return !condition||!condition.getCriteriaValues().some(value=>String(value).includes(marker));
    });
    const rule=SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(this.statusFormula(l)).setBackground('#e2efda')
      .setRanges([sh.getRange(l.first,1,sh.getMaxRows()-l.first+1,l.width)]).build();
    sh.setConditionalFormatRules([rule,...others]);
  },
  receivedLogSheet(book){
    const {receivedLog}=this.driveAssets(book,'log');
    const sh=SpreadsheetApp.openById(receivedLog.getId()).getSheetByName('Log');
    const header=this.receivedLogHeaders.join('|');
    if(!sh||sh.getRange(1,1,1,8).getValues()[0].join('|')!==header)throw Error('Received Log is missing its expected Log headers.');
    return sh;
  },
  enableReceivedLog(book){
    const ready=this.vendorNames(book).some(name=>{
      const sh=book.getSheetByName(name);
      if(!sh)return false;
      try{return Boolean(this.layout(sh).tracking);}catch(error){return false;}
    });
    if(!ready)throw Error('Initialize or repair the workbook before enabling Received Log.');
    const settings=PropertiesService.getScriptProperties(),key='RECEIVED_LOG_TRIGGER_ID',saved=settings.getProperty(key);
    const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='trackPurchaseLog');
    const current=triggers.find(t=>t.getEventType()===ScriptApp.EventType.ON_EDIT&&t.getTriggerSourceId()===book.getId()&&t.getUniqueId()===saved);
    if(current)return;
    const log=this.receivedLogSheet(book);
    if(!log.getRange(1,1,1,8).canEdit())throw Error('The account enabling logging needs edit access to Received Log.');
    const enabled=ScriptApp.newTrigger('trackPurchaseLog').forSpreadsheet(book).onEdit().create();
    settings.setProperty(key,enabled.getUniqueId());
    for(const trigger of triggers){
      if(trigger.getUniqueId()!==enabled.getUniqueId())ScriptApp.deleteTrigger(trigger);
    }
    book.toast('Paid, Arrived and Notes edits are logged automatically.','Received Log',8);
    return true;
  },
  recordReceivedEdit(e){
    if(!e?.range||e.range.getNumRows()!==1||e.range.getNumColumns()!==1)return;
    const sh=e.range.getSheet(),name=sh.getName();
    if(!this.isVendor(sh))return;
    const l=this.layout(sh),row=e.range.getRow(),col=e.range.getColumn();
    if(row<l.first||!l.tracking||![l.paid,l.arrived,l.notes].includes(col))return;
    const purchase=sh.getRange(row,1,1,l.width);
    const values=purchase.getValues()[0];
    const id=this.text(values[l.tracking-1]);
    if(!id)return;
    const itemNumber=this.text(values[l.id-1]),block=itemNumber?itemNumber+' ['+id+']':'['+id+']';
    const checked=String(e.value).toUpperCase();
    if(col!==l.notes&&(!['TRUE','FALSE'].includes(checked)||String(e.oldValue).toUpperCase()===checked))return;
    const action=col===l.notes?'Note updated':
      col===l.paid?(checked==='TRUE'?'Marked paid':'Unchecked payment'):
        checked==='TRUE'?'Marked arrived':'Unchecked arrival';
    const literal=value=>/^[=+\-@]/.test(value)?"'"+value:value;
    const display=purchase.getDisplayValues()[0];
    const link=l.link?sh.getRange(row,l.link).getRichTextValue()?.getLinkUrl()||display[l.link-1]:'';
    const details=[display[l.ordered-1],display[l.item-1],link,display[l.qty-1],display[l.unit-1],display[l.total-1]]
      .map(value=>{
        const field=this.text(value).replace(/[\r\n]+/g,' ');
        return /[",]/.test(field)?'"'+field.replace(/"/g,'""')+'"':field;
      }).join(', ');
    this.receivedLogSheet(e.source).appendRow([new Date(),e.user?.getEmail()||'',literal(block),
      name,row,literal(this.text(values[l.notes-1])),action,literal(details)]);
  },
  receiptEdit(e){
    if(!e||!e.range||!e.source)return;
    const sh=e.range.getSheet();if(!this.isVendor(sh))return;
    const l=this.layout(sh),first=Math.max(l.first,e.range.getRow()),last=e.range.getLastRow();
    if(first>last||l.arrived<e.range.getColumn()||l.arrived>e.range.getLastColumn())return;
    const lock=LockService.getDocumentLock();if(!lock)return;
    if(!lock.tryLock(1000)){
      e.source.toast('Arrival date not updated: another action is running. Repeat the Arrived checkbox change after it finishes.','Purchase Tools',10);return;
    }
    try{
      const rows=sh.getRange(first,1,last-first+1,l.width).getValues();
      const formulas=sh.getRange(first,l.arrival,last-first+1,1).getFormulas(),changed=[],now=new Date();
      rows.forEach((row,i)=>{
        const value=row[l.arrival-1],formula=formulas[i][0];
        if(row[l.arrived-1]===false&&(this.text(value)!==''||formula))changed.push({row:first+i,value:''});
        else if(row[l.arrived-1]===true&&this.text(row[l.item-1])){
          if(this.text(value)==='')changed.push({row:first+i,value:now});
          // Freeze an existing date formula to its valid result as well, so
          // no TODAY()/NOW() formula remains after an explicit receipt edit.
          else if(formula)changed.push({row:first+i,value:value instanceof Date&&!isNaN(value.getTime())?value:now});
        }
      });
      for(let i=0;i<changed.length;){
        let j=i+1;while(j<changed.length&&changed[j].row===changed[j-1].row+1)j++;
        sh.getRange(changed[i].row,l.arrival,j-i,1).setValues(changed.slice(i,j).map(change=>[change.value])).setNumberFormat('yyyy-mm-dd');i=j;
      }
    }catch(error){e.source.toast('Arrival date not updated: '+error.message,'Purchase Tools',10);throw error;}
    finally{lock.releaseLock();}
  },
  ids(book,refresh=false,edited=null,checked=null,excludedVendor=null){
    // One row formula per available row. No fixed suffix or stored counters.
    const names=edited&&!edited.l.company?[edited.sh.getName()]:this.vendorNames(book).filter(name=>name!==excludedVendor);
    const layouts=checked||new Map(names.map(name=>{
      if(edited?.sh.getName()===name)return [name,edited];
      const sh=book.getSheetByName(name);if(!sh)throw Error('Missing '+name+' tab. Run setup first.');
      return [name,{sh,l:this.layout(sh)}];
    }));
    const letter=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
    const quote=name=>"'"+name.replace(/'/g,"''")+"'";
    const vendorCodes=new Map(),used=new Set();
    for(const [name,{sh,l}] of layouts){
      if(l.company)continue;
      const code=this.customVendorCode(sh)||(this.vendors.includes(name)?this.code(name):'');
      if(!/^[A-Z0-9]{2,6}$/.test(code)||used.has(code))throw Error('Invalid or duplicate vendor code on '+name+'.');
      vendorCodes.set(name,code);used.add(code);
    }
    const companyCode=ref=>{
      const key='LOWER(REGEXREPLACE(LOWER(TRIM(TO_TEXT('+ref+'))),"[^a-z0-9]",""))';
      const choices=Object.entries({...this.codes,...Object.fromEntries(vendorCodes)})
        .filter(([name])=>/[a-z0-9]/i.test(name))
        .map(([name,code])=>'"'+name.toLowerCase().replace(/[^a-z0-9]/g,'')+'","'+code+'"').join(',');
      return 'IF('+key+'="","",SWITCH('+key+','+choices+',UPPER(LEFT('+key+',6))))';
    };
    const present=ref=>'N(LEN(TRIM(TO_TEXT('+ref+')))>0)';
    let repaired=0;
    for(const [name,{sh,l,activeRows,last:checkedLast}] of layouts){
      if(excludedVendor!==null&&!l.company)continue;
      if(edited&&name!==edited.sh.getName()||checked&&l.company&&layouts.size!==names.length)continue;
      const item=letter(l.item),currentItem='INDEX('+item+':'+item+',ROW())',ordered='INDEX('+letter(l.ordered)+':'+letter(l.ordered)+',ROW())';
      const code=l.company?companyCode('INDEX('+letter(l.company)+':'+letter(l.company)+',ROW())'):'"'+vendorCodes.get(name)+'"';
      let sequence='SUMPRODUCT(ARRAYFORMULA('+present('INDEX('+item+':'+item+','+l.first+'):INDEX('+item+':'+item+',ROW())')+'))';
      if(l.company){
        const count=(tab,running)=>{
          const {l:x}=layouts.get(tab),prefix=tab===name?'':quote(tab)+'!';
          const span=c=>'INDEX('+prefix+c+':'+c+','+x.first+'):INDEX('+prefix+c+':'+c+','+(running?'ROW()':'ROWS('+prefix+c+':'+c+')')+')';
          return 'SUMPRODUCT(ARRAYFORMULA(N(('+companyCode(span(letter(x.company)))+')=vcode)*'+present(span(letter(x.item)))+'))';
        };
        const prior=[...layouts].filter(([tab,{l:x}])=>!x.company).map(([tab,{l:x}])=>{
          const col=quote(tab)+'!'+letter(x.item)+':'+letter(x.item);
          return 'IF(vcode="'+vendorCodes.get(tab)+'",SUMPRODUCT(ARRAYFORMULA('+present('INDEX('+col+','+x.first+'):INDEX('+col+',ROWS('+col+'))')+')),0)';
        });
        if(name==='Shipment')prior.push(count('Others',false));
        sequence=[...prior,count(name,true)].join('+');
      }
      const formula='=IF(OR(LEN(TRIM(TO_TEXT('+currentItem+')))=0,NOT(ISNUMBER('+ordered+')),Control!$B$4=""),"",LET(vcode,'+code+
        ',seq,'+sequence+',IF(vcode="","",vcode&"-"&TEXT('+ordered+',"yyyymmdd")&"-"&UPPER(TRIM(Control!$B$4))&"-"&TEXT(seq,"0000"))))';
      const first=edited?.first??l.first,last=edited?.last??checkedLast??sh.getMaxRows();
      if(last<first)continue;
      const n=last-first+1,range=sh.getRange(first,1,n,1),existing=range.getFormulas();
      const needs=i=>(!activeRows||activeRows.has(first+i))&&(refresh||existing[i][0]!==formula);
      const snapshot=checked&&existing.some((_,i)=>needs(i))?range.getValues():null;
      for(let i=0;i<n;){
        if(!needs(i)){i++;continue;}
        let end=i+1;
        while(end<n&&needs(end))end++;
        const target=sh.getRange(first+i,1,end-i,1);
        if(snapshot&&(target.getFormulas().some(([value],j)=>value!==existing[i+j][0])||
          target.getValues().some(([value],j)=>!existing[i+j][0]&&value!==snapshot[i+j][0]))){
          throw Error(name+': Item number formulas changed during repair. Run the check again.');
        }
        target.setFormulas(Array.from({length:end-i},()=>[formula]));
        repaired+=end-i;
        i=end;
      }
      if(!edited&&!checked)sh.getRange(l.first-1,1).setNote('Live formula ID: uses the visible Ordered on date and current item order. Enter the date once and move complete rows. Numbers change with order; no stored counters.');
    }
    SpreadsheetApp.flush();
    return repaired;
  },
  refreshIds(book){
    this.project(book); // Explain missing project settings before rewriting tags.
    const changed=this.ids(book,true),missing=[];
    // Reading after flush checks the displayed result of the rewritten formulas.
    for(const name of this.vendorNames(book)){
      const sh=book.getSheetByName(name),{l,rows}=this.rows(sh);
      rows.forEach((r,i)=>{
        if(this.text(r[l.item-1])&&!/^[A-Z0-9]+-\d{8}-[A-Z0-9]{2,10}-\d{4}$/.test(this.text(r[0])))
          missing.push(name+' row '+(l.first+i)+(!r[l.ordered-1]?' — enter Ordered on':l.company&&!this.text(r[l.company-1])?' — enter Company':''));
      });
    }
    return [
      missing.length?'Item IDs need attention.':'Item IDs refreshed.',
      'ID cells refreshed or repaired: '+changed.toLocaleString('en-US'),
      'Numbers rebuilt from the current item order.',
      '',
      ...(missing.length?[
        'TAGS STILL MISSING OR INVALID',
        ...missing.slice(0,8).map(where=>'• '+where),
        ...(missing.length>8?['…and '+(missing.length-8)+' more.']:[]),
        '',
        'Check the listed rows and project settings.',
        'If Sheets is still calculating, try this command again.'
      ]:['All populated items currently show valid tags.']),
      '',
      'ID dates come from Ordered on; they do not change daily.'
    ].join('\n');
  },
  records(book,project=this.project(book)){
    const tz=book.getSpreadsheetTimeZone(),out=[];
    for(const tab of this.vendorNames(book)){
      const sh=book.getSheetByName(tab);if(!sh)continue;const {l,rows}=this.rows(sh);
      const grant=this.text(sh.getRange('A1').getDisplayValue());
      const grantMissing=!grant||/^#(?:REF!|N\/A|VALUE!|ERROR!|NAME\?|NUM!|DIV\/0!|NULL!)$/.test(grant);
      rows.forEach((r,i)=>{
        const item=this.text(r[l.item-1]),row=l.first+i;
        // Ignore unused template rows, including their unchecked checkboxes.
        // A partial purchase must remain visible to validation even without Item.
        const inputs=[l.id,l.ordered,l.item,l.company,l.link,l.qty,l.unit,l.total,l.category,l.month].filter(Boolean);
        if(!inputs.some(c=>this.text(r[c-1])!==''))return;
        const id=this.text(r[0]),qty=this.number(r[l.qty-1]),unit=this.number(r[l.unit-1]),total=this.number(r[l.total-1]);
        const company=l.company?this.text(r[l.company-1]):tab,month=this.month(r[l.month-1],tz),issues=[];
        const cell=(col,label)=>{let a='';for(let n=col;n;n=Math.floor((n-1)/26))a=String.fromCharCode(65+(n-1)%26)+a;return a+row+' ('+label+')';};
        const required=[[l.id,'Item number'],[l.ordered,'Ordered on'],[l.item,tab==='Shipment'?'Description':'Item'],
          [l.company,'Company'],[l.link,'Link'],[l.qty,tab==='Shipment'?'Unit':'Quantity'],
          [l.unit,tab==='Shipment'?'Unit price':'Individual price'],[l.total,'Final price'],[l.category,'Category']];
        for(const [col,label] of required)if(col&&this.text(r[col-1])==='')issues.push(cell(col,label)+': required');
        const orderValue=r[l.ordered-1],ordered=orderValue instanceof Date&&!isNaN(orderValue.getTime())?Utilities.formatDate(orderValue,tz,'yyyy-MM-dd'):'';
        if(this.text(orderValue)&&!ordered)issues.push(cell(l.ordered,'Ordered on')+': enter a valid date');
        if(id&&!/^[A-Z0-9]+-\d{8}-[A-Z0-9]{2,10}-\d{4}$/.test(id))issues.push(cell(l.id,'Item number')+': invalid ID; refresh item IDs');
        if(id&&id.split('-')[2]!==project.code)issues.push(cell(l.id,'Item number')+': project tag differs from Control!B4');
        if(grantMissing)issues.push('A1 (Grant number): enter a grant number or fix the linked grant cell');
        if(this.text(r[l.qty-1])&&(!Number.isFinite(qty)||qty<=0))issues.push(cell(l.qty,'Quantity')+': must be positive');
        if(this.text(r[l.unit-1])&&(!Number.isFinite(unit)||unit<0))issues.push(cell(l.unit,'Unit price')+': invalid price');
        if(this.text(r[l.total-1])&&(!Number.isFinite(total)||total<0))issues.push(cell(l.total,'Final price')+': invalid price');
        if([qty,unit,total].every(Number.isFinite)&&Math.abs(total-qty*unit)>0.02)issues.push(cell(l.total,'Final price')+': differs from quantity × unit price');
        const arrival=r[l.arrival-1] instanceof Date?Utilities.formatDate(r[l.arrival-1],tz,'yyyy-MM-dd'):this.text(r[l.arrival-1]);
        out.push({id,recordId:l.tracking?this.text(r[l.tracking-1]):'',ordered,item,company,qty,unit,total,grant,month,reported:!!this.text(r[l.month-1]),
          priority:l.priority>0&&r[l.priority-1]===true,
          category:this.text(r[l.category-1]),link:l.link?this.text(r[l.link-1]):'',paid:r[l.paid-1]===true,
          arrived:r[l.arrived-1]===true,arrival,notes:this.text(r[l.notes-1]),tab,row:l.first+i,
          source:book.getId(),project:project.name,code:project.code,issues});
      });
    }return out;
  },
  warn(title,errors,inline=false){
    if(!errors.length)return false;
    if(inline)throw Error(title+'\n\n'+errors.slice(0,20).join('\n')+'\n\nNo output was generated or replaced.');
    const ui=SpreadsheetApp.getUi();
    ui.alert(title,[
      'Complete or correct the cells below, then generate again.',
      '',
      ...errors.slice(0,20),
      ...(errors.length>20?['','…and '+(errors.length-20)+' more affected rows.']:[]),
      '',
      'No output was generated or replaced.'
    ].join('\n'),ui.ButtonSet.OK);
    return true;
  },
  select(records,input){
    input=this.text(input).toUpperCase();if(!records.length)throw Error('No purchases on this vendor tab.');
    if(input==='ALL')return records.slice();if(!input)throw Error('Enter current IDs, ranges, or ALL in the order prompt.');
    const result=new Map();
    for(const token of input.split(',').map(t=>t.trim())){
      const range=token.match(/^(\d{1,4})-(\d{1,4})$/),terms=[];
      if(range){const a=+range[1],b=+range[2];if(a>b||b-a>999)throw Error('Invalid/excessive selection range: '+token);for(let n=a;n<=b;n++)terms.push(String(n));}
      else terms.push(token);
      for(const term of terms){
        const hits=records.filter(r=>/^\d{1,4}$/.test(term)?Number(r.id.split('-').pop())===Number(term):r.id===term);
        if(hits.length!==1)throw Error(hits.length?'Ambiguous suffix '+term+'; use full ID.':'ID not found: '+term);
        result.set(hits[0].id,hits[0]);
      }
    }return [...result.values()].sort((a,b)=>a.row-b.row);
  },
  orders(book,selection,tab=book.getActiveSheet().getName(),inline=false){
    if(!this.isVendor(book.getSheetByName(tab)))throw Error('Open a vendor tab first.');
    const input=this.text(selection);
    if(!input)throw Error('Enter current IDs, ranges, or ALL in the order prompt.');
    this.ids(book);
    const candidates=this.records(book).filter(r=>r.tab===tab);let items;
    const describe=r=>tab+' row '+r.row+':\n'+r.issues.map(issue=>'• '+issue).join('\n');
    try{items=this.select(candidates,input);}
    catch(error){
      // A missing date/item can leave A blank, making suffix selection impossible.
      // Explain those unlabeled partial rows instead of only saying ID not found.
      if(candidates.some(r=>!r.id&&r.issues.length)&&this.warn('Order not generated — missing cells',
        [error.message,...candidates.filter(r=>!r.id&&r.issues.length).map(describe)],inline))return [];
      throw error;
    }
    if(this.warn('Order not generated — missing or invalid cells',items.filter(r=>r.issues.length).map(describe),inline))return [];
    const {template,orders:folder}=this.driveAssets(book,'orders');
    const body=DocumentApp.openById(template.getId()).getBody();
    const a=body.findText('\\{\\{ITEM_START\\}\\}'),b=body.findText('\\{\\{ITEM_END\\}\\}');
    if(!a||!b)throw Error('Order template needs {{ITEM_START}} and {{ITEM_END}} markers.');
    const start=body.getChildIndex(a.getElement().getParent()),end=body.getChildIndex(b.getElement().getParent());
    if(end<=start+1)throw Error('Order template has no repeating item paragraphs.');
    const patterns=[];for(let n=start+1;n<end;n++){
      if(body.getChild(n).getType()!==DocumentApp.ElementType.PARAGRAPH)throw Error('Repeating template content must consist of paragraphs.');
      patterns.push(body.getChild(n).copy().asParagraph());
    }
    const groups=new Map();for(const r of items){const key=r.company+'|'+r.grant;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);}
    const urls=[];
    for(const group of groups.values()){
      group.sort((a,b)=>Number(b.priority)-Number(a.priority)||a.row-b.row);
      const first=group[0],total=group.reduce((n,r)=>n+r.total,0),categories=[...new Set(group.map(r=>r.category))];
      const stamp=Utilities.formatDate(new Date(),book.getSpreadsheetTimeZone(),'yyyyMMdd_HHmmss');
      const file=template.makeCopy(first.code+'_'+first.company.replace(/[^A-Za-z0-9_-]/g,'_')+'_'+stamp+'_Grant_'+first.grant+'_'+group.length+'_Items',folder);
      const doc=DocumentApp.openById(file.getId()),target=doc.getBody();
      const head={Category:categories.length===1?categories[0]:'Mixed',GrantNo:first.grant,BlockID:group.length===1?first.id:'Selected: '+input,Company:first.company,TotalItems:String(group.length),TotalCost:total.toFixed(2)};
      for(const [k,v]of Object.entries(head))target.replaceText('\\{\\{'+k+'\\}\\}',String(v));
      // Re-resolve boundaries after replacing header text.
      const from=target.getChildIndex(target.findText('\\{\\{ITEM_START\\}\\}').getElement().getParent());
      const to=target.getChildIndex(target.findText('\\{\\{ITEM_END\\}\\}').getElement().getParent());
      for(let n=to;n>=from;n--)target.removeChild(target.getChild(n));
      let pos=from;group.forEach((r,i)=>{
        const values={Item_Number:String(i+1),ItemID:r.id,OrderedOn:r.ordered,ItemName:r.item,ItemLink:r.link,UnitCost:r.unit.toFixed(tab==='DigiKey'?4:2),Qty:String(r.qty),Subtotal:r.total.toFixed(2),TotalItems:String(group.length),TotalCost:total.toFixed(2)};
        for(const pattern of patterns){const text=target.insertParagraph(pos++,pattern.copy()).editAsText();
          for(const [k,v]of Object.entries(values))text.replaceText('\\{\\{'+k+'\\}\\}',v);
          const re=/https?:\/\/[^\s]+/g;let match;while((match=re.exec(text.getText())))text.setLinkUrl(match.index,match.index+match[0].length-1,match[0]);
          if(r.priority&&text.getText())text.setBackgroundColor('#f4cccc');
        }target.insertParagraph(pos++,'');
      });doc.saveAndClose();urls.push(file.getUrl());
    }
    if(!inline)SpreadsheetApp.getUi().alert('Created '+urls.length+' order document(s):\n'+urls.join('\n'));
    return urls;
  },
  orderPicker(){
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),tab=sheet.getName();
    if(!this.isVendor(sheet))throw Error('Open a vendor tab first.');
    const html=`<!doctype html><html><head><style>
  body{font:14px Arial,sans-serif;color:#202124;margin:0;padding:22px}
  label{display:block;font-weight:bold;margin-bottom:7px}
  input{box-sizing:border-box;width:100%;padding:10px;border:1px solid #dadce0;border-radius:5px;font:inherit}
  p{line-height:1.5;color:#5f6368}
  .actions{display:flex;justify-content:flex-end;gap:10px;margin-top:22px}
  button{font:inherit;border:1px solid #dadce0;border-radius:5px;padding:10px 16px;background:white;cursor:pointer}
  button.primary{background:#1a73e8;color:white;border-color:#1a73e8}button:disabled{opacity:.55;cursor:default}
  #status{white-space:pre-wrap;line-height:1.4;margin-top:18px}#status.error{color:#b3261e}
  #orders a{display:block;color:#1a73e8;font-weight:bold;margin-top:10px}
</style></head><body>
<form id="form">
  <label for="selection">Item IDs or ranges</label>
  <input id="selection" type="text" required autofocus aria-describedby="examples">
  <p id="examples">Enter <strong>0001-0003,0007</strong> for a range and an individual item, <strong>0001</strong> or a full ID for one item, or <strong>ALL</strong> for every purchase on this tab.</p>
  <div class="actions"><button type="button" id="cancel">Cancel</button><button type="submit" class="primary" id="generate">Generate</button></div>
</form>
<div id="status" role="status" aria-live="polite"></div>
<div id="orders"></div>
<script>
  const selection=document.getElementById('selection'),generate=document.getElementById('generate');
  const cancel=document.getElementById('cancel'),status=document.getElementById('status'),orders=document.getElementById('orders');
  let busy=false;
  cancel.addEventListener('click',()=>{if(!busy)google.script.host.close();});
  document.getElementById('form').addEventListener('submit',event=>{
    event.preventDefault();
    if(busy||generate.hidden)return;
    const input=selection.value.trim();
    if(!input){status.className='error';status.textContent='Enter item IDs, ranges, or ALL.';return;}
    busy=true;
    selection.disabled=generate.disabled=cancel.disabled=true;
    status.className='';status.textContent='Generating purchase orders…';
    google.script.run.withSuccessHandler(urls=>{
      busy=false;cancel.disabled=false;
      status.textContent=urls.length===1?'Purchase order ready.':urls.length+' purchase orders ready.';
      urls.forEach((url,i)=>{
        const link=document.createElement('a');
        link.href=url;link.target='_blank';link.rel='noopener';
        link.textContent=urls.length===1?'Open purchase order':'Open purchase order '+(i+1);
        orders.appendChild(link);
      });
      generate.hidden=true;cancel.textContent='Close';
    }).withFailureHandler(error=>{
      busy=false;selection.disabled=generate.disabled=cancel.disabled=false;
      status.className='error';status.textContent=error.message;
    }).generatePurchaseOrder(input,__TAB__);
  });
</script></body></html>`;
    const output=HtmlService.createHtmlOutput(html.replace('__TAB__',JSON.stringify(tab))).setWidth(440).setHeight(440);
    SpreadsheetApp.getUi().showModalDialog(output,'Generate purchase order for '+tab);
  },
  monthPicker(){
    const book=SpreadsheetApp.getActiveSpreadsheet();
    const current=Utilities.formatDate(new Date(),book.getSpreadsheetTimeZone(),'yyyy-MM');
    const [year,month]=current.split('-').map(Number);
    const names=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const months=names.map((name,i)=>'<option value="'+String(i+1).padStart(2,'0')+'"'+(i+1===month?' selected':'')+'>'+name+'</option>').join('');
    const years=Array.from({length:100},(_,i)=>2000+i).map(y=>'<option value="'+y+'"'+(y===year?' selected':'')+'>'+y+'</option>').join('');
    const html=`<!doctype html>
<html><head><base target="_top"><style>
  body{font:14px Arial,sans-serif;color:#202124;margin:0;padding:22px}
  p{line-height:1.5;margin:0 0 18px}.fields{display:flex;gap:16px}
  label{display:block;font-weight:bold;margin-bottom:7px}.field{flex:1}.field:first-child{flex:1.5}
  select{box-sizing:border-box;width:100%;padding:10px;border:1px solid #dadce0;border-radius:5px;background:white;font:inherit}
  .note{color:#5f6368;font-size:12px;margin-top:18px}
  .actions{display:flex;justify-content:flex-end;gap:10px;margin-top:22px}
  button{font:inherit;border:1px solid #dadce0;border-radius:5px;padding:10px 16px;background:white;cursor:pointer}
  button.primary{background:#1a73e8;color:white;border-color:#1a73e8}button:disabled{opacity:.55;cursor:default}
  #status{white-space:pre-wrap;line-height:1.4;margin-top:18px}#status.error{color:#b3261e}
  #report{color:#1a73e8;font-weight:bold}a[hidden]{display:none}
</style></head><body>
<p>Select the month and year to generate. Purchases are included using their <strong>Reported on</strong> value.</p>
<form id="form">
<div class="fields">
  <div class="field"><label for="month">Month</label><select id="month">__MONTHS__</select></div>
  <div class="field"><label for="year">Year</label><select id="year">__YEARS__</select></div>
</div>
<p class="note">Each year has one workbook. If this month already has a tab, its previous version will be archived before it is replaced.</p>
<div class="actions"><button type="button" id="cancel">Cancel</button><button type="submit" class="primary" id="generate">Generate</button></div>
</form>
<div id="status" role="status" aria-live="polite"></div>
<p><a id="report" hidden target="_blank" rel="noopener">Open monthly summary</a></p>
<script>
  const month=document.getElementById('month'),year=document.getElementById('year');
  const generate=document.getElementById('generate'),cancel=document.getElementById('cancel');
  const status=document.getElementById('status'),report=document.getElementById('report');
  let busy=false,pendingConfirmation=null;
  function controls(disabled){month.disabled=disabled;year.disabled=disabled;generate.disabled=disabled;cancel.disabled=disabled;}
  cancel.addEventListener('click',function(){if(!busy)google.script.host.close();});
  document.getElementById('form').addEventListener('submit',function(event){
    event.preventDefault();if(busy)return;
    busy=true;controls(true);status.className='';report.hidden=true;
    const value=year.value+'-'+month.value;
    const label=month.options[month.selectedIndex].text+' '+year.value;
    const confirmation=pendingConfirmation;
    pendingConfirmation=null;generate.textContent='Generate';
    status.textContent='Generating '+label+'…';
    google.script.run.withSuccessHandler(function(result){
      busy=false;
      if(result&&result.confirmation){
        pendingConfirmation=result.confirmation;
        status.textContent=result.message+'\\n\\nContinue to archive and replace this month, or cancel to keep it.';
        generate.textContent='Archive and replace';generate.disabled=false;cancel.disabled=false;
        return;
      }
      if(!result){controls(false);status.className='error';status.textContent='No summary was generated. Close this window, correct any reported cells, then try again.';return;}
      status.textContent=['Summary ready: '+label+'.',...result.notices].join('\\n\\n');
      report.href=result.url;report.hidden=false;generate.hidden=true;cancel.disabled=false;cancel.textContent='Close';
    }).withFailureHandler(function(error){
      busy=false;controls(false);status.className='error';
      status.textContent=error&&error.message?error.message:'Could not generate the summary. Please try again.';
    }).generateMonthlySummary(value,confirmation);
  });
</script></body></html>`;
    const output=HtmlService.createHtmlOutput(html.replace('__MONTHS__',months).replace('__YEARS__',years)).setWidth(440).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(output,'Generate monthly summary');
  },
  projectFolder(book){
    return this.singleParent(DriveApp.getFileById(book.getId()),'This workbook');
  },
  singleParent(item,label){
    const parents=item.getParents();
    if(!parents.hasNext())throw Error(label+' has no accessible parent folder.');
    const folder=parents.next();
    if(parents.hasNext())throw Error(label+' has more than one parent folder.');
    return folder;
  },
  writeOrderTemplate(body){
    body.clear();
    body.setPageWidth(595.28).setPageHeight(841.89)
      .setMarginTop(72).setMarginBottom(72).setMarginLeft(72).setMarginRight(72);
    const style=(paragraph,size,font)=>paragraph.setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0)
      .editAsText().setFontFamily(font).setFontSize(size).setBold(false);
    const first=body.getParagraphs()[0]||body.appendParagraph('');
    first.setText('Company: {{Company}}');
    style(first,15,'Times New Roman').setBold(true);
    for(const line of ['Grant Number: {{GrantNo}}','{{Category}}','A total of   {{TotalItems}}   items to order']){
      const text=style(body.appendParagraph(line),15,'Times New Roman').setBold(true);
      const start=line.indexOf('{{TotalItems}}');
      if(start>=0)text.setUnderline(start-2,start+'{{TotalItems}}'.length+1,true);
    }
    body.appendParagraph('').setSpacingBefore(12).setSpacingAfter(0);
    body.appendParagraph('{{ITEM_START}}');
    for(const [label,value,spacingBefore] of [
      ['{{Item_Number}}. Item: ','{{ItemName}}',12],['Link:','{{ItemLink}}',0],
      ['Unit Cost: ','${{UnitCost}}',0],['Quantity: ','{{Qty}}',0],['Subtotal:',' ${{Subtotal}}',0]
    ]){
      const paragraph=body.appendParagraph(label+value);
      style(paragraph,10,'Arial').setBold(0,label.length-1,true);
      paragraph.setSpacingBefore(spacingBefore);
    }
    body.appendParagraph('{{ITEM_END}}');
    body.appendParagraph('').setSpacingBefore(12).setSpacingAfter(12);
    const total=body.appendParagraph('Total: ${{TotalCost}}');
    style(total,15,'Arial').setBold(true);
    total.setSpacingBefore(12).setSpacingAfter(12);
  },
  driveAssets(book,mode){
    const create=mode==='initialize';
    const all=create||mode==='references';
    const unique=(iterator,name)=>{
      const item=iterator.hasNext()?iterator.next():null;
      if(iterator.hasNext())throw Error('More than one Drive item named '+name+'. Give each item a unique name.');
      if(!item&&!create)throw Error('Missing '+name+'. Run Purchase Tools → Initialize Drive templates and folders.');
      return item;
    };
    const folder=(parent,name)=>unique(parent.getFoldersByName(name),name)||parent.createFolder(name);
    const parent=this.projectFolder(book),projectsName='General_Purchase_Orders';
    const alreadyInProjects=parent.getName()===projectsName;
    if(!alreadyInProjects&&!create)throw Error('Run Purchase Tools → Initialize Drive templates and folders first.');
    const root=alreadyInProjects?this.singleParent(parent,projectsName):parent;
    const projects=alreadyInProjects?parent:folder(root,projectsName);
    const assets={root};
    if(all||mode==='summary'){
      assets.summaries=folder(root,'Monthly_summaries');
      assets.archive=folder(assets.summaries,'Archive_Summary_changes');
    }
    if(all||mode==='orders'){
      assets.orders=folder(root,'Orders_forms');
      const templates=folder(root,'Templates'),name='Template_form';
      assets.template=unique(templates.getFilesByName(name),name);
      const existingTemplate=assets.template;
      if(!existingTemplate){
        const doc=DocumentApp.create(name);
        this.writeOrderTemplate(doc.getBody());
        doc.saveAndClose();
        assets.template=DriveApp.getFileById(doc.getId()).moveTo(templates);
      }
      if(assets.template.getMimeType()!==MimeType.GOOGLE_DOCS)throw Error(name+' must be a native Google Doc.');
      if(create&&existingTemplate){
        const doc=DocumentApp.openById(assets.template.getId()),body=doc.getBody();
        const generic=[
          'Purchase order','Vendor: {{Company}}','Grant: {{GrantNo}}   Category: {{Category}}','{{BlockID}}',
          '{{ITEM_START}}','{{Item_Number}}. {{ItemName}}','{{ItemID}}   Ordered on: {{OrderedOn}}','{{ItemLink}}',
          'Quantity: {{Qty}}   Unit cost: $ {{UnitCost}}   Subtotal: $ {{Subtotal}}',
          '{{ITEM_END}}','Items: {{TotalItems}}   Total: $ {{TotalCost}}'
        ].join('\n');
        const compact=[
          'Company: {{Company}}','Grant Number: {{GrantNo}}','{{Category}}',
          'A total of   {{TotalItems}}   items to order','{{ITEM_START}}',
          '{{Item_Number}}. Item: {{ItemName}}','Link: {{ItemLink}}','Unit Cost: ${{UnitCost}}',
          'Quantity: {{Qty}}','Subtotal: ${{Subtotal}}','{{ITEM_END}}','Total: ${{TotalCost}}'
        ].join('\n');
        if([generic,compact].includes(body.getText().trim())){
          this.writeOrderTemplate(body);
          doc.saveAndClose();
        }
      }
      if(all){
        const sheetName='Template_for_Purchase_Orders';
        let sheetFile=unique(templates.getFilesByName(sheetName),sheetName);
        if(!sheetFile){
          const templateBook=SpreadsheetApp.create(sheetName),blank=templateBook.getSheets()[0];
          this.setup(templateBook,{skip:true},true);
          templateBook.deleteSheet(blank);
          sheetFile=DriveApp.getFileById(templateBook.getId()).moveTo(templates);
          assets.templateBook=templateBook;
        }else if(create&&sheetFile.getMimeType()===MimeType.GOOGLE_SHEETS){
          assets.templateBook=SpreadsheetApp.openById(sheetFile.getId());
          this.setup(assets.templateBook,{skip:true},true);
        }
        if(sheetFile.getMimeType()!==MimeType.GOOGLE_SHEETS)throw Error(sheetName+' must be a native Google Sheet.');
        assets.workbookTemplate=sheetFile;
      }
    }
    if(all||mode==='log'){
      const templates=folder(root,'Templates'),name='Received_Log';
      assets.receivedLog=unique(templates.getFilesByName(name),name);
      if(!assets.receivedLog){
        const book=SpreadsheetApp.create(name),sh=book.getSheets()[0].setName('Log');
        sh.getRange('A1:H1').setValues([this.receivedLogHeaders]).setBackground('#d6dce5').setFontWeight('bold');
        sh.setFrozenRows(1);
        [170,215,350,130,65,310,185,360].forEach((width,i)=>sh.setColumnWidth(i+1,width));
        sh.getRange(2,6,sh.getMaxRows()-1,1).setWrap(true);
        sh.getRange(2,8,sh.getMaxRows()-1,1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        SpreadsheetApp.flush();
        assets.receivedLog=DriveApp.getFileById(book.getId()).moveTo(templates);
      }
      if(assets.receivedLog.getMimeType()!==MimeType.GOOGLE_SHEETS)throw Error(name+' must be a native Google Sheet.');
    }
    if(!alreadyInProjects)DriveApp.getFileById(book.getId()).moveTo(projects);
    return assets;
  },
  initializeDrive(book){
    const assets=this.driveAssets(book,'initialize');
    this.formatControl(assets.templateBook.getSheetByName('Control'),assets);
    const control=book.getSheetByName('Control');
    if(control)this.formatControl(control,assets);
    return 'Drive templates and folders are ready in '+assets.root.getName()+'.';
  },
  projectSourceId(value,required=true){
    const raw=this.text(value),id=raw.match(/\/spreadsheets\/d\/([\w-]+)/)?.[1]||raw;
    if(!/^[\w-]{20,}$/.test(id)){
      if(required)throw Error('Invalid spreadsheet URL/ID in Projects.');
      return '';
    }
    return id;
  },
  scanProjects(book){
    const folder=this.projectFolder(book),files=folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const projects=[],unreadable=[];let count=0;
    while(files.hasNext()){
      const file=files.next(),id=file.getId(),fileName=file.getName();
      count++;
      try{
        const source=id===book.getId()?book:SpreadsheetApp.openById(id);
        if(!source.getSheetByName('Control'))continue;
        const {name,code}=this.project(source,false);
        if(!name||!code)continue;
        const issue=!/^[A-Z0-9]{2,10}$/.test(code)?'Invalid project code':
          !source.getSheetByName('Projects')?'Run setup in this workbook first':'';
        projects.push({id,fileName,name,code,issue,current:id===book.getId()});
      }catch(error){unreadable.push(fileName);}
    }
    return {folder:folder.getName(),folderId:folder.getId(),count,unreadable,
      projects:projects.sort((a,b)=>a.fileName.localeCompare(b.fileName))};
  },
  syncProjectRegistries(book,request){
    if(!Array.isArray(request.projects)||request.projects.length<2)throw Error('Select at least two workbooks.');
    const scan=this.scanProjects(book);
    if(scan.folderId!==request.folderId)throw Error('The workbook changed folders. Reopen the dialog and scan again.');
    const selected=new Map(),codes=new Map();
    for(const entry of request.projects){
      const p=scan.projects.find(candidate=>candidate.id===entry.id);
      if(!p||p.issue||p.name!==entry.name||p.code!==entry.code)throw Error('A selected workbook changed or is unavailable. Reopen the dialog and scan again.');
      if(selected.has(p.id))throw Error('A workbook was selected twice.');
      if(codes.has(p.code))throw Error('Duplicate project code '+p.code+' in '+codes.get(p.code)+' and '+p.fileName+'. Select one or correct the code.');
      selected.set(p.id,p);codes.set(p.code,p.fileName);
    }
    const externalCodes=new Map();
    const plans=[];
    for(const p of selected.values()){
      try{
        const source=p.current?book:SpreadsheetApp.openById(p.id),sh=source.getSheetByName('Projects');
        const header=sh.getRange(1,1,1,3).getDisplayValues()[0];
        if(header.join('|')!=='Project|Spreadsheet URL or ID|Enabled')throw Error('Unexpected Projects headers. Check columns A through C.');
        const last=sh.getLastRow(),registry=last>1?sh.getRange(2,1,last-1,4):null;
        const rows=registry?registry.getValues():[],formulas=registry?registry.getFormulas():[],notes=registry?registry.getNotes():[];
        const known=new Map(),empty=[],usedCodes=new Map([...selected.values()].map(other=>[other.code,other.id]));
        rows.forEach((row,i)=>{
          if(row[0]===''&&row[1]===''&&(row[2]===''||row[2]===false)&&row[3]===''&&
            !formulas[i].some(Boolean)&&!notes[i].some(Boolean))empty.push(i+2);
          if(!this.text(row[1]))return;
          const id=this.projectSourceId(row[1],row[2]===true);
          if(selected.has(id)){
            if(known.has(id))throw Error('Projects contains repeated entries for '+selected.get(id).fileName+'. Keep one entry and retry.');
            known.set(id,i+2);return;
          }
          if(row[2]!==true)return;
          if(!externalCodes.has(id))externalCodes.set(id,this.project(SpreadsheetApp.openById(id)).code);
          const code=externalCodes.get(id);
          if(usedCodes.has(code)&&usedCodes.get(code)!==id)throw Error('Enabled project in row '+(i+2)+' conflicts with project code '+code+'.');
          usedCodes.set(code,id);
        });
        const entries=[];let next=last+1;
        for(const other of selected.values()){
          const existing=known.get(other.id);
          const row=existing&&!(empty[0]<existing)?existing:(empty.shift()??next++);
          const move=existing&&existing!==row?existing:null;
          const values=[other.code,other.id,true];
          const update=move||!existing||values.some((value,i)=>rows[existing-2][i]!==value)||
            sh.getRange(existing,1).getFontWeight()!=='bold';
          if((update&&row<=sh.getMaxRows()&&!sh.getRange(row,1,1,move?4:3).canEdit())||
            (move&&!sh.getRange(move,1,1,4).canEdit()))throw Error('No edit access to the required Projects cells.');
          entries.push({row,move,values,update});
        }
        const extra=next-1-sh.getMaxRows();
        if(extra>0&&!sh.getRange(sh.getMaxRows(),1,1,3).canEdit())throw Error('No edit access to extend Projects.');
        plans.push({p,sh,extra,entries});
      }catch(error){throw Error(p.fileName+': '+error.message+' No project entries were changed.');}
    }
    return plans.map(({p,sh,extra,entries})=>{
      try{
        if(extra>0)sh.insertRowsAfter(sh.getMaxRows(),extra);
        const checkbox=SpreadsheetApp.newDataValidation().requireCheckbox().build();
        for(const {row,move,values,update} of entries){
          if(!update)continue;
          if(move)sh.getRange(move,1,1,4).moveTo(sh.getRange(row,1,1,4));
          sh.getRange(row,3).setDataValidation(checkbox);
          sh.getRange(row,1,1,2).setNumberFormat('@');
          sh.getRange(row,1).setFontWeight('bold');
          sh.getRange(row,1,1,3).setValues([values]);
        }
        SpreadsheetApp.flush();
        for(const {row,values} of entries){
          if(sh.getRange(row,1,1,3).getValues()[0].some((value,i)=>value!==values[i]))throw Error('Could not verify Projects row '+row+'.');
        }
        const locations=' · Projects rows '+entries.map(e=>e.row).sort((a,b)=>a-b).join(', ');
        return {fileName:p.fileName,status:(entries.some(e=>e.update)?'Synced':'Already up to date')+locations};
      }catch(error){return {fileName:p.fileName,status:'Could not finish: '+error.message};}
    });
  },
  projectSyncPicker(){
    const html=`<!doctype html><html><head><base target="_top"><style>
  body{font:14px Arial,sans-serif;color:#202124;padding:16px}
  #list{max-height:280px;overflow:auto} .project{padding:9px 0;border-bottom:1px solid #dadce0;overflow-wrap:anywhere}
  label{display:block;margin-top:5px} #status{white-space:pre-wrap;overflow-wrap:anywhere}
  button{padding:8px 13px;margin-right:6px} .actions{margin-top:16px}
</style></head><body>
  <p id="count">Scanning this workbook’s Drive folder…</p>
  <progress id="progress" aria-label="Scanning or syncing workbooks" style="width:100%;accent-color:#1a73e8"></progress>
  <div id="list"></div>
  <p id="status" role="status"></p>
  <div class="actions"><button id="sync" disabled>Sync selected projects</button><button id="close">Close</button></div>
  <script>
  const byId=id=>document.getElementById(id),sync=byId('sync'),status=byId('status'),close=byId('close');
  const options=[];
  const progress=byId('progress');
  const failure=error=>{progress.hidden=true;status.textContent=error.message||String(error);close.disabled=false;};
  close.onclick=()=>google.script.host.close();
  google.script.run.withFailureHandler(failure).withSuccessHandler(scan=>{
    progress.hidden=true;
    byId('count').textContent=scan.folder+' · '+scan.count+' spreadsheets found · '+scan.projects.length+' with project settings';
    const duplicates=new Set(scan.projects.filter(p=>scan.projects.some(other=>other.id!==p.id&&other.code===p.code)).map(p=>p.code));
    for(const p of scan.projects){
      const row=document.createElement('div'),title=document.createElement('strong'),label=document.createElement('label'),box=document.createElement('input');
      row.className='project';
      title.textContent=p.fileName+(p.current?' (this workbook)':'')+' — '+p.name+' ('+p.code+')';
      box.type='checkbox';box.disabled=Boolean(p.issue);box.checked=p.current&&!p.issue;
      label.append(box,p.issue||(' Include in sync'+(duplicates.has(p.code)?' · duplicate code in folder':'')));
      row.append(title,label);byId('list').append(row);options.push({p,box});
    }
    const instructions='Select at least two workbooks. Each selected workbook will show the full selection in Projects. Existing entries outside the selection are preserved.';
    status.textContent=instructions+(scan.unreadable.length?'\\nCould not inspect: '+scan.unreadable.join(', '):'');
    const update=()=>{sync.disabled=options.filter(o=>o.box.checked).length<2;};
    options.forEach(o=>o.box.onchange=update);
    sync.onclick=()=>{
      const projects=options.filter(o=>o.box.checked).map(o=>o.p);
      if(new Set(projects.map(p=>p.code)).size!==projects.length){
        status.textContent='Select only one workbook per project code, or correct the duplicate codes and reopen this dialog.';return;
      }
      sync.disabled=true;close.disabled=true;progress.hidden=false;status.textContent='Syncing selected workbooks…';
      options.forEach(o=>o.box.disabled=true);
      const unlock=()=>{progress.hidden=true;close.disabled=false;options.forEach(o=>o.box.disabled=Boolean(o.p.issue));update();};
      google.script.run.withFailureHandler(error=>{unlock();failure(error);}).withSuccessHandler(results=>{
        unlock();status.textContent=results.map(r=>r.fileName+' — '+r.status).join('\\n');
      }).syncPurchaseProjects({folderId:scan.folderId,projects});
    };
  }).syncPurchaseProjects({scan:true});
  </script></body></html>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(580).setHeight(520),'Sync projects in this folder');
  },
  summary(book,month,showMessage=true,confirmation=null){
    // The month picker supplies the reporting period explicitly.
    if(typeof month!=='string'||!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month))throw Error('Select a month and year in Purchase Tools → Generate monthly summary.');
    const currentProject=this.project(book);
    this.ids(book);
    const sources=new Map([[book.getId(),book]]),registry=book.getSheetByName('Projects');
    if(registry&&registry.getLastRow()>1)for(const r of registry.getRange(2,1,registry.getLastRow()-1,3).getValues()){
      if(r[2]!==true)continue;
      const label=this.text(r[0])||'(unnamed project)';
      try{
        const id=this.projectSourceId(r[1]);
        if(!sources.has(id))sources.set(id,SpreadsheetApp.openById(id));
      }catch(error){throw Error('Cannot read enabled project '+label+'. Check its spreadsheet ID and your access. The current summary was kept.');}
    }
    const all=[],errors=[],projects=[],codes=new Map(),seen=new Set(),notices=[];
    let paidPending=0,pendingAmount=0,pendingWithoutAmount=0;
    for(const [id,source] of sources){
      const p=source===book?currentProject:this.project(source);
      if(codes.has(p.code))throw Error('Duplicate project code '+p.code);
      codes.set(p.code,id);
      projects.push(p.name+' ('+p.code+')');
      for(const r of this.records(source,p)){
        if(!r.month){
          if(r.reported)errors.push(r.project+'/'+r.tab+' row '+r.row+': invalid Reported on date');
          else if(r.paid){
            paidPending++;
            if(Number.isFinite(r.total)&&r.total>=0)pendingAmount+=r.total;
            else pendingWithoutAmount++;
          }
          continue;
        }
        if(r.month!==month)continue;
        if(r.issues.length)errors.push(r.project+'/'+r.tab+' row '+r.row+':\n'+r.issues.map(issue=>'• '+issue).join('\n'));
        const key=r.source+'|'+r.id;if(seen.has(key))errors.push('Duplicate current ID: '+r.id);seen.add(key);all.push(r);
      }
    }
    if(this.warn('Summary not generated — missing or invalid cells',errors,!showMessage))return '';
    if(!all.length)throw Error('No purchases assigned to '+month+'. Existing reports unchanged.');
    const grid=this.summaryGrid(all,month,projects);
    if(paidPending){
      const amount=new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(pendingAmount);
      notices.push(paidPending+' paid '+(paidPending===1?'purchase has':'purchases have')+' no Reported on date. '+
        'Known total: '+amount+'.'+(pendingWithoutAmount?' Missing or invalid final price: '+pendingWithoutAmount+'.':''));
    }
    const url=this.publish(book,month,codes,grid,confirmation,notices);
    if(typeof url!=='string')return url;
    if(showMessage)SpreadsheetApp.getUi().alert(['Monthly summary: '+all.length+' purchases.',...notices,url].join('\n\n'));
    return {url,notices};
  },
  summaryGrid(records,month,projects){
    const data=[
      ['MONTHLY PURCHASE SUMMARY — '+month],
      ['Projects: '+projects.join(', ')],
      ['Green = Paid (confirmed)','','','Orange = Payment unconfirmed (verify with finance)'],
      []
    ];
    const details=[],finance=[];const add=r=>{data.push(r||[]);return data.length;};
    const byGrant=new Map();
    for(const record of records){
      if(!byGrant.has(record.grant))byGrant.set(record.grant,[]);
      byGrant.get(record.grant).push(record);
    }
    const grants=[...byGrant.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    for(const grant of grants){
      const items=byGrant.get(grant).sort((a,b)=>a.company.localeCompare(b.company)||a.project.localeCompare(b.project)||a.row-b.row);
      add(['GRANT '+grant]);add(['Item number','Ordered on','Company','Item','Link','Quantity','Unit price','Final price','Order status','Category','Paid','Arrive_lab','Project','Vendor tab','Source spreadsheet','Notes']);
      const first=data.length+1;
      items.forEach(r=>add([r.recordId?r.id+' ['+r.recordId+']':r.id,r.ordered,r.company,r.item,r.link,r.qty,r.unit,r.total,r.arrived?'Received':'Not received',r.category,r.paid?'Paid':'Unconfirmed',r.arrival,r.project,r.tab,r.source,r.notes]));
      const last=data.length;add(['','TOTAL ('+grant+')','','','','','','=SUM(H'+first+':H'+last+')']);add();details.push({grant,first,last,items});
    }
    add(['FINANCIAL RECONCILIATION']);const cross=data.length+1;
    for(const d of details){
      add(['','Summary ('+d.grant+')','Our total','Department total','Difference']);
      const first=data.length+1,vendors=[...new Set(d.items.map(r=>r.company))];
      for(const vendor of vendors){const row=data.length+1;add(['',vendor,'=SUMIF(C'+d.first+':C'+d.last+',B'+row+',H'+d.first+':H'+d.last+')','','=IF(D'+row+'="","",ROUND(C'+row+'-D'+row+',2))']);}
      const last=data.length,row=data.length+1;
      add(['','TOTAL','=SUM(C'+first+':C'+last+')','=IF(COUNT(D'+first+':D'+last+')='+vendors.length+',SUM(D'+first+':D'+last+'),"")','=IF(D'+row+'="","",ROUND(C'+row+'-D'+row+',2))']);add();finance.push({first,last});
    }
    const put=(r,c,v)=>{while(data.length<r)add();data[r-1][c-1]=v;};
    ['Grant','Lab Supplies','Capital Asset/Equipment','Shipping','Others / Custom','Total','Paid','Unconfirmed'].forEach((v,i)=>put(cross,7+i,v));
    details.forEach((d,i)=>{const row=cross+i+1;put(row,7,d.grant);
      ['Lab Supplies','Capital Asset/Equipment','Shipping'].forEach((cat,j)=>put(row,8+j,'=SUMIF(J'+d.first+':J'+d.last+',"'+cat+'",H'+d.first+':H'+d.last+')'));
      put(row,12,'=SUM(H'+d.first+':H'+d.last+')');put(row,11,'=L'+row+'-SUM(H'+row+':J'+row+')');
      put(row,13,'=SUMIF(K'+d.first+':K'+d.last+',"Paid",H'+d.first+':H'+d.last+')');put(row,14,'=L'+row+'-M'+row);
    });
    const crossTotal=cross+details.length+1;
    put(crossTotal,7,'ALL GRANTS');
    for(const [index,column] of [...'HIJKLMN'].entries()){
      put(crossTotal,8+index,'=SUM('+column+(cross+1)+':'+column+(crossTotal-1)+')');
    }
    const grantTotals=column=>finance.map(f=>column+(f.last+1)).join(',');
    const connieTotals=grantTotals('D');
    add();
    const grandRow=data.length+1;
    add([
      '',
      'GRAND TOTAL',
      '=SUM('+grantTotals('C')+')',
      '=IF(COUNT('+connieTotals+')='+finance.length+',SUM('+connieTotals+'),"")',
      '=IF(ISNUMBER(D'+grandRow+'),ROUND(SUM('+grantTotals('E')+'),2),"")'
    ]);
    return {data:data.map(r=>Array.from({length:16},(_,i)=>r[i]??'')),details,finance,cross,crossTotal};
  },
  writeReport(sh,grid){
    const rows=grid.data.length;if(sh.getMaxRows()<rows)sh.insertRowsAfter(sh.getMaxRows(),rows-sh.getMaxRows());
    if(sh.getMaxColumns()<16)sh.insertColumnsAfter(sh.getMaxColumns(),16-sh.getMaxColumns());
    const range=sh.getRange(1,1,rows,16);
    range.setValues(grid.data).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('TOP');
    sh.setFrozenRows(4);sh.setHiddenGridlines(true);
    sh.getRange('A1:P1').merge().setBackground('#1f4e79').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14);
    sh.getRange(2,1,1,16).setFontColor('#555555');
    sh.getRange('A3:C3').merge().setBackground('#6aa84f').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange('D3:H3').merge().setBackground('#ff9900').setFontColor('#000000').setFontWeight('bold');
    for(const d of grid.details){
      sh.getRange(d.first-2,1,1,16).setBackground('#b7b7b7').setFontWeight('bold');
      sh.getRange(d.first-1,1,1,16).setBackground('#d9d9d9').setFontWeight('bold');
      sh.getRange(d.last+1,1,1,16).setBackground('#eeeeee').setFontWeight('bold');
      sh.getRange(d.first,7,d.last-d.first+1,2).setNumberFormat('$#,##0.00');
      sh.getRange(d.first,2,d.last-d.first+1,1).setNumberFormat('yyyy-mm-dd');
      sh.getRange(d.first,1,d.last-d.first+1,1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      sh.getRange(d.first,4,d.last-d.first+1,2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      sh.getRange(d.first,15,d.last-d.first+1,1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    sh.getRange(grid.cross-1,1,1,16).setBackground('#b7b7b7').setFontWeight('bold');
    const palette=['#d9d2e9','#c9daf8','#fce5cd','#d9ead3','#ffe599'];
    grid.finance.forEach((f,i)=>{
      const color=palette[i%palette.length];
      for(const row of [f.first-1,f.last+1])sh.getRange(row,2,1,4).setBackground(color).setFontWeight('bold');
      sh.getRange(f.first-1,4).setFontColor('#c00000');
      sh.getRange(f.first,3,f.last-f.first+2,3).setNumberFormat('$#,##0.00');
      sh.getRange(f.first,3,f.last-f.first+1,1).setFontWeight('bold');
      sh.getRange(f.first,4,f.last-f.first+1,1).setBackground('#fff2f2').setFontColor('#c00000').setNote('Enter independently confirmed finance values.');
    });
    sh.getRange(grid.cross,7,1,8).setBackground('#c9daf8').setFontWeight('bold').setWrap(true);
    if(grid.crossTotal)sh.getRange(grid.crossTotal,7,1,8).setBackground('#d9ead3').setFontWeight('bold');
    sh.getRange(rows,2,1,4).setBackground('#d9e2f3').setFontWeight('bold');
    sh.getRange(rows,3,1,3).setNumberFormat('$#,##0.00');
    sh.getRange(grid.cross+1,8,grid.details.length+(grid.crossTotal?1:0),7).setNumberFormat('$#,##0.00');
    const widths=[365,145,160,410,260,95,120,130,135,210,120,145,150,150,200,480];widths.forEach((w,i)=>sh.setColumnWidth(i+1,w));
    sh.getRange(1,10,rows,1).setWrap(true);sh.getRange(1,16,rows,1).setWrap(true);
    const tintRanges=[sh.getRange(1,1,rows,10),sh.getRange(1,12,rows,5)];
    const statusRange=sh.getRange(1,11,rows,1),rules=[];
    for(const [status,tint,badge,text] of [['Paid','#e2efda','#6aa84f','#ffffff'],['Unconfirmed','#fff2cc','#ff9900','#000000']]){
      const condition='=AND($K1="'+status+'",ISNUMBER($H1),$A1<>"")';
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(condition).setBackground(tint).setRanges(tintRanges).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(condition).setBackground(badge).setFontColor(text).setRanges([statusRange]).build());
    }
    sh.setConditionalFormatRules(rules);
    SpreadsheetApp.flush();const values=range.getValues(),formulas=range.getFormulas(),tz=sh.getParent().getSpreadsheetTimeZone();
    // Sheets rewrites a one-cell range such as H7:H7 as H7 (also in
    // SUMIF/COUNT). Compare that harmless canonical form, keeping quoted
    // text unchanged. Other references, operators and amounts remain exact.
    const formulaKey=formula=>formula.split(/("(?:[^"]|"")*")/g).map((part,i)=>i%2?part:
      part.replace(/(^|[^A-Z0-9_])(\$?[A-Z]{1,3}\$?[1-9]\d*):\2(?![A-Z0-9_])/g,'$1$2')).join('');
    for(let r=0;r<rows;r++)for(let c=0;c<16;c++){
      const expected=grid.data[r][c];let wanted=expected,actual=values[r][c];
      if(typeof expected==='string'&&expected.startsWith('=')){wanted=formulaKey(expected);actual=formulaKey(formulas[r][c]);}
      else if(actual instanceof Date&&/^\d{4}-\d{2}-\d{2}$/.test(expected))actual=Utilities.formatDate(actual,tz,'yyyy-MM-dd');
      if(String(actual??'')!==String(wanted))throw Error('Report readback differs at row '+(r+1)+', column '+(c+1)+
        '. Expected '+JSON.stringify(expected)+'; read '+JSON.stringify(formulas[r][c]||values[r][c])+'. No existing month was replaced.');
      if(typeof values[r][c]==='string'&&/^#(?:REF!|DIV\/0!|VALUE!|N\/A|ERROR!|NAME\?|NUM!)/.test(values[r][c]))throw Error('Formula error in staged report. Old current report retained.');
    }
  },
  departmentEntries(sheet){
    const range=sheet.getRange(1,2,sheet.getLastRow(),4);
    const values=range.getValues(),formulas=range.getFormulas(),entries=new Map();
    let grant=null;
    values.forEach(([label,our,department,difference],i)=>{
      const header=String(label).match(/^Summary \((.+)\)$/);
      if(header&&['Our total','Our maths'].includes(our)&&['Department total','Connie'].includes(department)&&difference==='Difference'){
        grant=header[1];return;
      }
      if(!label||(label==='TOTAL'&&/^=SUM\(/i.test(formulas[i][1])))grant=null;
      if(grant===null||department==='')return;
      const key=JSON.stringify([grant,label]);
      if(entries.has(key))throw Error('Multiple department amounts for '+grant+' / '+label+'. Review the current month before regenerating.');
      entries.set(key,department);
    });
    return entries;
  },
  publish(book,month,codes,grid,confirmation,notices=[]){
    const {summaries:folder,archive}=this.driveAssets(book,'summary');
    const year=month.slice(0,4),name='Monthly_Summary_'+year;
    const owner='PurchaseTemplate annual summary; year='+year+'; folder='+folder.getId();
    const findAnnual=()=>{
      const found=folder.getFilesByName(name),file=found.hasNext()?found.next():null;
      if(found.hasNext())throw Error('Multiple annual workbooks named '+name+'. Resolve the duplicate before generating.');
      return file;
    };
    let file=findAnnual();
    let report,starter;if(file){
      if(file.getDescription()!==owner)throw Error('A file named '+name+' already exists and is not registered as this folder’s annual summary.');
      report=SpreadsheetApp.openById(file.getId());
      starter=report.getSheetByName('_annual_starter');
    }else{
      report=SpreadsheetApp.create('_new_'+name+'_'+Utilities.getUuid());
      starter=report.getSheets()[0].setName('_annual_starter');
      this.formulaLocale(report);
      file=DriveApp.getFileById(report.getId());file.setDescription(owner);file.moveTo(folder);
      try{
        if(findAnnual())throw Error('The annual workbook was just created by another run. Retry generation.');
        file.setName(name);
        if(findAnnual()?.getId()!==file.getId())throw Error('The annual workbook changed. Retry generation.');
      }catch(error){file.setTrashed(true);throw error;}
    }
    const projectSet=JSON.stringify([...codes].sort(([a],[b])=>a.localeCompare(b)));
    const old=report.getSheetByName(month),stageName='_build_'+month,previousName='_previous_'+month;
    if(report.getSheetByName(stageName)||report.getSheetByName(previousName))
      throw Error('Another run or interrupted replacement needs review for '+month+' in '+report.getUrl());
    const oldSet=old?old.getRange('A2').getNote():null;
    if(old&&oldSet!==projectSet){
      let prior;
      try{prior=new Map(JSON.parse(oldSet||'[]'));}catch(error){prior=new Map();}
      const removed=[...prior.keys()].filter(code=>!codes.has(code));
      const added=[...codes.keys()].filter(code=>!prior.has(code));
      const changed=[...codes.keys()].filter(code=>prior.has(code)&&prior.get(code)!==codes.get(code));
      const token=JSON.stringify([report.getId(),month,old.getSheetId(),oldSet,projectSet]);
      if(confirmation!==token)return {confirmation:token,message:[
        'Project coverage changed for '+month+'.',
        !prior.size?'The current tab has no recorded project list.':'',
        removed.length?'Removed: '+removed.join(', '):'',
        added.length?'Added: '+added.join(', '):'',
        changed.length?'Different workbook for: '+changed.join(', '):''
      ].filter(Boolean).join('\n')};
    }
    // A fixed staging name gives runs from different source workbooks one slot per month.
    const stage=report.insertSheet(stageName);
    let remaining;
    try{
      if((report.getSheetByName(month)?.getSheetId()??null)!==(old?.getSheetId()??null)||
          (old&&old.getRange('A2').getNote()!==oldSet)||report.getSheetByName(previousName))
        throw Error('The current month changed while generating. Retry to review the latest version.');
      const entries=old&&grid.finance?this.departmentEntries(old):null;
      remaining=entries?new Map(entries):null;
      if(remaining)grid.finance.forEach(({first,last},i)=>{
        const grant=grid.details[i].grant;
        for(let row=first;row<=last;row++){
          const cells=grid.data[row-1],key=JSON.stringify([grant,cells[1]]);
          if(!remaining.has(key))continue;
          const amount=remaining.get(key);
          if(typeof amount!=='number'||!Number.isFinite(amount))
            throw Error('Enter a numeric department total for '+grant+' / '+cells[1]+' in the current month before regenerating.');
          cells[3]=amount;
          remaining.delete(key);
        }
      });
      this.writeReport(stage,grid);
      stage.getRange('A2').setNote(projectSet);
      if(old)this.archiveMonth(report,old,month,archive);
      if(entries&&JSON.stringify([...entries])!==JSON.stringify([...this.departmentEntries(old)]))
        throw Error('Department totals changed during generation. The current month was kept. Retry to include the latest entries.');
      if(old)old.setName(previousName);
      stage.setName(month);
      SpreadsheetApp.flush();
    }catch(error){
      if(stage.getName()===month)stage.setName(stageName);
      if(old&&old.getName()===previousName)old.setName(month);
      report.deleteSheet(stage);
      throw error;
    }
    if(old)report.deleteSheet(old);
    if(starter&&starter.getLastRow()===0)report.deleteSheet(starter);
    if(remaining){
      for(const key of remaining.keys()){
        const [grant,vendor]=JSON.parse(key);
        notices.push('Department total kept in the archive only: '+grant+' / '+vendor+'.');
      }
    }
    return report.getUrl()+'#gid='+stage.getSheetId();
  },
  archiveMonth(report,old,month,archive){
    const name='Monthly_Summary_'+month,re=new RegExp('^'+name+'(?:_[^/]+)?_recreated_v(\\d+)$');
    let version=0;
    const files=archive.getFiles();
    while(files.hasNext()){
      const match=files.next().getName().match(re);
      if(match)version=Math.max(version,Number(match[1])+1);
    }
    const saved=SpreadsheetApp.create('_archive_'+name+'_'+Utilities.getUuid());
    const file=DriveApp.getFileById(saved.getId());
    try{
      saved.setSpreadsheetLocale(report.getSpreadsheetLocale());
      saved.setSpreadsheetTimeZone(report.getSpreadsheetTimeZone());
      const starter=saved.getSheets()[0],copy=old.copyTo(saved).setName(month);
      saved.deleteSheet(starter);
      SpreadsheetApp.flush();
      const snapshot=sheet=>{
        const range=sheet.getDataRange();
        return JSON.stringify([range.getValues(),range.getFormulas(),range.getNotes()]);
      };
      if(snapshot(copy)!==snapshot(old))throw Error('Archive verification failed. The current month was kept.');
      file.moveTo(archive);
      file.setName(name+'_recreated_v'+version);
    }catch(error){file.setTrashed(true);throw error;}
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Purchase Tools')
    .addItem(
      'Generate order from current vendor tab',
      'generatePurchaseOrder'
    )
    .addItem(
      'Generate monthly summary',
      'generateMonthlySummary'
    )
    .addSeparator()
    .addItem(
      'Initialize workbook',
      'setupPurchaseWorkbook'
    )
    .addItem('Check / repair workbook','checkPurchaseWorkbook')
    .addItem(
      'Refresh / repair item IDs',
      'refreshPurchaseIds'
    )
    .addItem('Sync projects in this folder','syncPurchaseProjects')
    .addSeparator()
    .addItem('Add vendor tab','addPurchaseVendor')
    .addItem('Remove vendor tab','removePurchaseVendor')
    .addItem('Rename vendor tab','renamePurchaseVendor')
    .addSeparator()
    .addSeparator()
    .addItem('Initialize Drive templates and folders','initializePurchaseDrive')
    .addToUi();
}

function addPurchaseVendor(request) {
  if(request===undefined)return PA.vendorPicker();
  return PA.run('addVendor',request);
}

function removePurchaseVendor(request) {
  if(request===undefined)return PA.removeVendorPicker();
  return PA.run('removeVendor',request);
}

function renamePurchaseVendor(request) {
  if(request===undefined)return PA.vendorPicker(true);
  return PA.run('renameVendor',request);
}

function setupPurchaseWorkbook(request) {
  if(request===undefined)return PA.setupPicker();
  return PA.run('setup',request);
}

function checkPurchaseWorkbook(request,token) {
  if(request===undefined)return PA.progressPicker('Check / repair workbook','Checking workbook and applying repairs…','checkPurchaseWorkbook',true);
  const progress=(completed,total,label)=>{
    if(!/^[\w-]{1,80}$/.test(token||''))return;
    try{CacheService.getUserCache().put('purchase-check-'+token,JSON.stringify({completed,total,label}),600);}
    catch(error){console.warn('Could not publish workbook progress: '+error.message);}
  };
  if(Array.isArray(request))return PA.run('recreateMissing',request,progress);
  return PA.run('checkWorkbook',progress);
}

function readPurchaseCheckProgress(token) {
  if(!/^[\w-]{1,80}$/.test(token||''))return null;
  const value=CacheService.getUserCache().get('purchase-check-'+token);
  return value?JSON.parse(value):null;
}

function generatePurchaseOrder(selection,tab) {
  if(selection===undefined)return PA.orderPicker();
  return PA.run('orders',selection,tab,true);
}

function generateMonthlySummary(month,confirmation) {
  if(month===undefined)return PA.monthPicker();
  return PA.run('summary',month,false,confirmation);
}

function refreshPurchaseIds(request) {
  if(request===undefined)return PA.progressPicker('Refresh / repair item IDs','Refreshing item IDs across all vendor tabs…','refreshPurchaseIds');
  return PA.run('refreshIds');
}

function syncPurchaseProjects(request) {
  if(request===undefined)return PA.projectSyncPicker();
  if(request.scan===true)return PA.scanProjects(SpreadsheetApp.getActiveSpreadsheet());
  return PA.run('syncProjectRegistries',request);
}

function initializePurchaseDrive(request) {
  if(request===undefined)return PA.progressPicker('Initialize Drive templates and folders','Preparing Drive templates and folders…','initializePurchaseDrive');
  return PA.run('initializeDrive');
}

function trackPurchaseLog(e) {
  const enabled=PropertiesService.getScriptProperties().getProperty('RECEIVED_LOG_TRIGGER_ID');
  if(enabled&&String(e?.triggerUid)===enabled)PA.recordReceivedEdit(e);
}

// Simple bound-sheet trigger; no manual trigger installation is needed.
function onEdit(e) {
  PA.receiptEdit(e);
  PA.trackingEdit(e);
}
