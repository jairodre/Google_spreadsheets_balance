const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

const script = readFileSync(join(__dirname, '..', 'PurchaseTemplate.gs'), 'utf8');
const iterator = items => {
  let index = 0;
  return { hasNext: () => index < items.length, next: () => items[index++] };
};

function app(services) {
  return vm.runInNewContext(script + '\nPA', services);
}

test('annual reports keep other months and archive the replaced project set', () => {
  const files = [], reports = new Map();
  let nextId = 0, failWrite = false;
  const folder = {
    getId: () => 'summary-folder',
    getFilesByName: name => iterator(files.filter(f => f.folder === folder && f.name === name))
  };
  const archive = { getFiles: () => iterator(files.filter(f => f.folder === archive)) };
  const create = name => {
    const id = ++nextId, tabs = [];
    const file = {
      name, getId: () => id, getName() { return this.name; },
      getDescription() { return this.description; }, setDescription(value) { this.description = value; },
      moveTo(value) { this.folder = value; }, setName(value) { this.name = value; },
      setTrashed() { this.folder = null; }
    };
    const report = {
      getId: () => id, getUrl: () => 'report/' + id, getSheets: () => tabs,
      getSpreadsheetLocale: () => 'en_US', getSpreadsheetTimeZone: () => 'America/Denver',
      setSpreadsheetLocale() {}, setSpreadsheetTimeZone() {},
      getSheetByName: title => tabs.find(tab => tab.getName() === title) || null,
      insertSheet(title) {
        if (this.getSheetByName(title)) throw Error('Duplicate tab');
        const sheetId = ++nextId;
        const tab = {
          note: '', value: '', getName: () => title, getSheetId: () => sheetId,
          setName(value) { title = value; return this; }, getLastRow() { return this.value ? 1 : 0; },
          getRange() { return { getNote: () => this.note, setNote: value => { this.note = value; } }; },
          getDataRange() { return { getValues: () => [[this.value]], getFormulas: () => [['']], getNotes: () => [[this.note]] }; },
          copyTo(target) {
            const copy = target.insertSheet('Copy of ' + title);
            copy.value = this.value; copy.note = this.note;
            return copy;
          }
        };
        tabs.push(tab);
        return tab;
      },
      deleteSheet(tab) { tabs.splice(tabs.indexOf(tab), 1); }
    };
    report.insertSheet('Hoja 1');
    files.push(file); reports.set(id, report);
    return report;
  };
  const purchase = app({
    SpreadsheetApp: { create, openById: id => reports.get(id), flush() {} },
    DriveApp: { getFileById: id => files.find(f => f.getId() === id) },
    Utilities: { getUuid: () => 'mock-' + nextId }
  });
  purchase.driveAssets = () => ({ summaries: folder, archive });
  purchase.writeReport = (tab, value) => {
    if (failWrite) throw Error('Staging failed');
    tab.value = value;
  };
  const both = new Map([['FD3', 'fd3'], ['CAL', 'cal']]), cal = new Map([['CAL', 'cal']]);
  const publish = (month, projects, value, confirmation) => purchase.publish({}, month, projects, value, confirmation);
  assert.match(publish('2026-11', both, 'combined'), /#gid=/);
  publish('2026-12', both, 'December');
  publish('2027-01', cal, 'January');
  const annual = reports.get(folder.getFilesByName('Monthly_Summary_2026').next().getId());
  const before = publish('2026-11', cal, 'CAL only');
  assert.match(before.message, /Removed: FD3/);
  assert.equal(annual.getSheetByName('2026-11').value, 'combined');
  assert.equal(archive.getFiles().hasNext(), false);
  failWrite = true;
  assert.throws(() => publish('2026-11', cal, 'CAL only', before.confirmation), /Staging failed/);
  assert.equal(annual.getSheetByName('2026-11').value, 'combined');
  failWrite = false;
  publish('2026-11', cal, 'CAL only', before.confirmation);
  publish('2026-11', cal, 'CAL only');
  assert.deepEqual(annual.getSheets().map(tab => tab.getName()).sort(), ['2026-11', '2026-12']);
  assert.equal(annual.getSheetByName('2026-12').value, 'December');
  const versions = files.filter(f => f.folder === archive);
  assert.deepEqual(versions.map(f => f.name), ['Monthly_Summary_2026-11_recreated_v0', 'Monthly_Summary_2026-11_recreated_v1']);
  assert.deepEqual(versions.map(f => reports.get(f.getId()).getSheets().map(tab => tab.value)), [['combined'], ['CAL only']]);
  const entries = new Map([
    [JSON.stringify(['Grant-1', 'Thorlabs']), 0],
    [JSON.stringify(['Grant-1', 'Mouser']), 42]
  ]);
  purchase.departmentEntries = () => entries;
  const grid = purchase.summaryGrid([{grant:'Grant-1',company:'Thorlabs',project:'CAL',total:50}], '2026-11', ['CAL']);
  const notices = [];
  purchase.publish({}, '2026-11', cal, grid, null, notices);
  const first = grid.finance[0].first - 1;
  assert.deepEqual(Array.from(grid.data[first - 1].slice(2,5)), ['Our total','Department total','Difference']);
  assert.equal(grid.data[first][3], 0);
  assert.match(grid.data[first][4], /=IF/);
  assert.match(notices[0], /Grant-1 \/ Mouser/);
});

test('department reconciliation reads legacy entries and skips calculated totals', () => {
  const values = [
    ['Summary (Grant-1)', 'Our maths', 'Connie', 'Difference'],
    ['Thorlabs', 50, 0, 50],
    ['TOTAL', 50, 0, 50],
    ['GRAND TOTAL', 50, 0, 50]
  ];
  const formulas = values.map(() => ['', '', '', '']);
  formulas[2][1] = '=SUM(C2)';
  const entries = app({}).departmentEntries({
    getLastRow: () => values.length,
    getRange: () => ({getValues: () => values, getFormulas: () => formulas})
  });
  assert.equal(entries.size, 1);
  assert.equal(entries.get(JSON.stringify(['Grant-1','Thorlabs'])), 0);
});

function sheet(name) {
  const rows = new Map();
  let writes=0;
  const range = (row, column, height = 1, width = 1) => {
    if (typeof row === 'string') {
      const [, firstCol, firstRow, lastCol, lastRow] = row.match(/^([A-Z])(\d+)(?::([A-Z])(\d+))?$/);
      return range(Number(firstRow), firstCol.charCodeAt(0) - 64,
        lastRow ? Number(lastRow) - Number(firstRow) + 1 : 1,
        lastCol ? lastCol.charCodeAt(0) - firstCol.charCodeAt(0) + 1 : 1);
    }
    const result = {
      getDisplayValues: () => Array.from({ length: height }, (_, i) =>
        Array.from({ length: width }, (_, j) => rows.get(row + i)?.[column + j - 1] ?? '')),
      setValues(values) {
        writes++;
        values.forEach((valuesRow, i) => {
          if (!rows.has(row + i)) rows.set(row + i, []);
          valuesRow.forEach((value, j) => { rows.get(row + i)[column + j - 1] = value; });
        });
        return this;
      },
      setValue(value) { return this.setValues([[value]]); },
      setFormula(value) { return this.setValue(value); },
      setFormulas(values) { return this.setValues(values); },
      setRichTextValue(value) { return this.setValue(value); },
      getDisplayValue() { return this.getDisplayValues()[0][0]; },
      getFormulas() { return this.getDisplayValues().map(values=>values.map(value=>typeof value==='string'&&value.startsWith('=')?value:'')); },
      getFormula() { return this.getFormulas()[0][0]; },
      getA1Notation() { return String.fromCharCode(64+column)+row; },
      clearContent() { return this.setValues(Array.from({length:height},()=>Array(width).fill(''))); }
    };
    for (const method of ['setBackground', 'setFontWeight', 'setFontSize', 'setHorizontalAlignment', 'setVerticalAlignment',
      'setNumberFormat', 'setNote', 'insertCheckboxes', 'setDataValidation', 'setWrap', 'setWrapStrategy']) {
      result[method] = () => result;
    }
    result.getValues = result.getDisplayValues;
    return result;
  };
  return {
    getName: () => name, getRange: range, getMaxRows: () => 30,
    getLastRow: () => Math.max(0,...rows.keys()),
    getWriteCount: () => writes,
    setName(value) { name = value; return this; },
    getLastColumn: () => Math.max(1, ...Array.from(rows.values(), row => row.length)),
    setFrozenRows() {}, setColumnWidth() {}, setRowHeights() {}, setRowHeight() {}, setHiddenGridlines() {}, autoResizeRows() {}
  };
}

test('the portable schema creates all tabs and preserves existing project settings on repeat', () => {
  const validation = {
    requireValueInList: () => validation,
    setAllowInvalid: () => validation,
    setHelpText: () => validation,
    build: () => ({})
  };
  const purchase = app({ SpreadsheetApp: { newDataValidation: () => validation } });
  const sheets = new Map();
  const book = {
    getSheetByName: name => sheets.get(name),
    insertSheet: name => {
      assert(!sheets.has(name));
      const tab = sheet(name);
      sheets.set(name, tab);
      return tab;
    }
  };
  purchase.format(book);
  assert.equal(sheets.size, 12);
  for (const vendor of purchase.vendors) {
    const layout = purchase.layout(sheets.get(vendor));
    assert.equal(layout.ordered, 2);
    assert.equal(layout.priority, 3);
  }
  assert.deepEqual(sheets.get('Projects').getRange('A1:D1').getDisplayValues()[0],
    ['Project', 'Spreadsheet URL or ID', 'Enabled', 'Notes']);
  const control = sheets.get('Control');
  control.getRange('B3:B4').setValues([['Example project'], ['DEMO']]);
  purchase.format(book);
  assert.deepEqual(control.getRange('B3:B4').getDisplayValues(), [['Example project'], ['DEMO']]);
});

test('Starter tab cleanup preserves notes, content, and the last tab', () => {
  const purchase=app({});
  let rowCount=0,notes=[['']],tabs;
  const starter={getName:()=> 'Hoja 1',getLastRow:()=>rowCount,getMaxRows:()=>1,getMaxColumns:()=>1,
    getRange:()=>({getNotes:()=>notes})};
  const report={};
  const book={getSheets:()=>tabs,deleteSheet:tab=>{tabs=tabs.filter(item=>item!==tab);}};
  tabs=[starter];
  purchase.removeEmptyStarter(book,starter);
  assert.deepEqual(tabs,[starter]);
  tabs.push(report);
  notes=[['Keep this']];
  purchase.removeEmptyStarter(book,starter);
  assert.deepEqual(tabs,[starter,report]);
  notes=[['']];rowCount=1;
  purchase.removeEmptyStarter(book,starter);
  assert.deepEqual(tabs,[starter,report]);
  rowCount=0;
  purchase.removeEmptyStarter(book,starter);
  assert.deepEqual(tabs,[report]);
});

test('Progress dialog keeps warnings, confirmation, and server errors visible', () => {
  let html;
  const output={setWidth(){return this;},setHeight(){return this;}};
  const purchase=app({
    Utilities:{getUuid:()=> 'test-progress'},
    HtmlService:{createHtmlOutput(value){html=value;return output;}},
    SpreadsheetApp:{getUi:()=>({showModalDialog(){}})}
  });
  purchase.progressPicker('Check / repair workbook','Checking…','checkPurchaseWorkbook',true);
  const elements=Object.fromEntries(['progress','status','close','recreate','confirmation'].map(id=>[id,{}]));
  const requests=[];let poll;
  const run={
    withFailureHandler(callback){this.failure=callback;return this;},
    withSuccessHandler(callback){this.success=callback;return this;},
    checkPurchaseWorkbook(request,token){requests.push([request,token]);this.finish=this.success;this.fail=this.failure;},
    readPurchaseCheckProgress(token){
      assert.equal(token,'test-progress-1');
      this.success({completed:4,total:13,label:'Checking Mouser'});
    }
  };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1],{
    document:{getElementById:id=>elements[id]},google:{script:{run,host:{close(){}}}},
    setTimeout(callback){poll=callback;return 1;},clearTimeout(){}
  });
  assert.equal(elements.progress.hidden,false);
  assert.deepEqual(requests[0],[true,'test-progress-1']);
  assert.equal(elements.progress.value,0);
  poll();
  assert.equal(elements.progress.value,4);
  assert.equal(elements.progress.max,13);
  assert.match(elements.status.textContent,/Checking Mouser/);
  run.finish({message:'Review a typed Final price.',missing:['Mouser']});
  assert.equal(elements.progress.hidden,true);
  assert.equal(elements.status.textContent,'Review a typed Final price.');
  assert.equal(elements.recreate.hidden,false);
  elements.recreate.onclick();
  assert.deepEqual(requests[1],[['Mouser'],'test-progress-2']);
  assert.equal(elements.progress.hidden,false);
  run.fail({message:'Drive permission denied.'});
  assert.equal(elements.progress.hidden,true);
  assert.equal(elements.status.textContent,'Drive permission denied.');
  assert.equal(elements.close.disabled,false);
});

test('Repair adds Reported on date validation and clears a misplaced category dropdown', () => {
  const criteria={CHECKBOX:'checkbox',DATE_IS_VALID_DATE:'date',VALUE_IN_LIST:'list'};
  const built=[];
  const purchase=app({SpreadsheetApp:{DataValidationCriteria:criteria,newDataValidation:()=>({
    requireDate(){this.type=criteria.DATE_IS_VALID_DATE;return this;},
    setAllowInvalid(value){this.allowInvalid=value;return this;},
    setHelpText(value){this.help=value;return this;},
    build(){built.push(this);return this;}
  })}});
  const l={first:4,width:15,priority:3,item:4,ordered:2,category:10,month:11,paid:12,arrived:13,tracking:15};
  const rule=(type,allow=false,values=[])=>({getCriteriaType:()=>type,getAllowInvalid:()=>allow,getCriteriaValues:()=>values});
  const sample=[];
  sample[1]=rule(criteria.DATE_IS_VALID_DATE);
  for(const column of [3,12,13])sample[column-1]=rule(criteria.CHECKBOX);
  sample[9]=sample[14]=rule(criteria.VALUE_IN_LIST,true,[purchase.categories]);
  const applied=[],cleared=[];
  const bounds={getRow:()=>4,getColumn:()=>1,getNumRows:()=>27,getNumColumns:()=>15};
  const vendor={
    getName:()=> 'Shipment',getMaxRows:()=>30,getFrozenRows:()=>3,
    getConditionalFormatRules:()=>[{getBooleanCondition:()=>({getCriteriaValues:()=>[purchase.statusFormula(l)]}),getRanges:()=>[bounds]}],
    getRange:(row,column)=>({getBackground:()=> '#d8dde5',getFontWeight:()=> 'bold',
      getDataValidations:()=>[sample],setDataValidation:value=>applied.push([column,value]),
      clearDataValidations:()=>cleared.push(column)})
  };
  purchase.checkVendorStyle(vendor,l,new Set(),[]);
  assert.equal(applied.length,1);
  assert.equal(applied[0][0],l.month);
  assert.equal(built[0].type,criteria.DATE_IS_VALID_DATE);
  assert.equal(built[0].allowInvalid,false);
  assert.deepEqual(cleared,[l.tracking]);
});

test('Workbook check skips healthy rows, preserves typed totals and reports missing tabs', () => {
  const purchase=app({SpreadsheetApp:{flush(){}}});
  purchase.vendors=['Amazon','Mouser'];
  purchase.checkVendorStyle=()=>{};
  purchase.enableReceivedLog=()=>{};
  purchase.controlLinks=()=>false;
  purchase.driveAssets=()=>({});
  const vendor=sheet('Amazon'),control=sheet('Control'),projects=sheet('Projects');
  const tabs=new Map([['Amazon',vendor],['Control',control],['Projects',projects]]);
  const book={getSheetByName:name=>tabs.get(name),getSheets:()=>[...tabs.values()],getSpreadsheetLocale:()=>'en_US'};
  vendor.getRange(3,1,1,15).setValues([['Item number','Ordered on','Priority','Item','Link','Quantity','Individual price',
    'Final price','Arrive_lab','Category','Reported on','Paid','Arrived','Notes','Purchase Record ID']]);
  control.getRange('A1').setValue('PURCHASE WORKFLOW');
  vendor.getRange('D4').setValue('Part');
  vendor.getRange('F4:G4').setValues([[2,5]]);
  vendor.getRange('O4').setValue('DEMO-RECORD');
  vendor.getRange('H4').setValue(12);
  const steps=[];
  const result=purchase.checkWorkbook(book,(completed,total,label)=>steps.push([completed,total,label]));
  assert.deepEqual(steps.map(([completed])=>completed),[0,1,2,3,4,5]);
  assert.equal(steps.at(-1)[1],5);
  assert.deepEqual(Array.from(result.missing),['Mouser']);
  assert.match(result.message,/H4: 12 \(expected 10\.00\)/);
  assert.equal(vendor.getRange('H4').getValues()[0][0],12);
  assert.equal(vendor.getRange('O4').getValues()[0][0],'DEMO-RECORD');
  assert(vendor.getRange('A4').getFormula().startsWith('=IF('));
  assert.equal(vendor.getRange('A5').getFormula(),'');
  vendor.getRange('H4').clearContent();
  purchase.checkWorkbook(book);
  assert(vendor.getRange('H4').getFormula().startsWith('=ARRAYFORMULA('));
  const before=vendor.getWriteCount();
  assert.match(purchase.checkWorkbook(book).message,/No repairs made/);
  assert.equal(vendor.getWriteCount(),before);
  assert.throws(()=>purchase.setup(book,{skip:true}),/for blank workbooks/);
});

test('US formula locale and Stage 1 vendor styling apply without changing purchase inputs', () => {
  let locale = 'es_ES', changes = 0;
  const purchase = app({ SpreadsheetApp: { WrapStrategy: { CLIP: 'clip' }, flush() {} } });
  const book = {
    getSpreadsheetLocale: () => locale,
    setSpreadsheetLocale(value) { locale = value; changes++; }
  };
  purchase.formulaLocale(book);
  purchase.formulaLocale(book);
  assert.equal(locale, 'en_US');
  assert.equal(changes, 1);
  const vendor = sheet('Thorlabs');
  vendor.getRange(3, 1, 1, 15).setValues([['Item number', 'Ordered on', 'Priority', 'Item', 'Link',
    'Quantity', 'Individual price', 'Final price', 'Arrive_lab', 'Category', 'Reported on',
    'Paid', 'Arrived', 'Notes', 'Purchase Record ID']]);
  vendor.getRange('A1').setValue('00123');
  vendor.getRange('D4').setValue('Sample part');
  vendor.getRange('L4:M4').setValues([[true, true]]);
  purchase.styleVendor(vendor);
  assert.deepEqual(vendor.getRange('A1:A2').getValues(), [['00123'], ['THORLABS']]);
  assert.equal(vendor.getRange('D4').getValues()[0][0], 'Sample part');
  assert.deepEqual(vendor.getRange('L4:M4').getValues(), [[true, true]]);
});

test('Control shows portable asset links and preserves project settings', () => {
  const purchase = app({ SpreadsheetApp: { newRichTextValue: () => ({
    setText(text) { this.text = text; return this; },
    setLinkUrl(url) { this.url = url; return this; },
    build() { return {text:this.text,url:this.url}; }
  }) } });
  const control = sheet('Control');
  control.getRange('B3:B4').setValues([['Example project'], ['DEMO']]);
  const keys=['orders','summaries','template','workbookTemplate','archive','receivedLog'];
  const assets=Object.fromEntries(keys.map(key=>[key,{getUrl:()=> 'https://example.invalid/'+key}]));
  purchase.formatControl(control,assets);
  assert.deepEqual(control.getRange('B3:B4').getValues(), [['Example project'], ['DEMO']]);
  [12,13,14,15,21,22].forEach((row,i)=>assert.equal(control.getRange(row,2).getValues()[0][0].url,assets[keys[i]].getUrl()));
});

function folder(name, parent) {
  const folders = [], files = [];
  return {
    getName: () => name,
    getParents: () => iterator(parent ? [parent] : []),
    getFoldersByName: value => iterator(folders.filter(item => item.getName() === value)),
    getFilesByName: value => iterator(files.filter(item => item.name === value)),
    createFolder(value) { const child = folder(value, this); folders.push(child); return child; },
    addFile: file => files.push(file)
  };
}

test('Drive initialization creates reusable assets beside the workbook', () => {
  const root = folder('Example purchase system'), paragraphs = [];
  let workbookParent = root, moves = 0;
  const workingFile = {
    getParents: () => iterator([workbookParent]),
    moveTo(destination) { workbookParent = destination; moves++; return this; }
  };
  let docsCreated = 0, sheetsCreated = 0;
  const template = {
    name: 'Template_form', getId: () => 'demo-doc', getMimeType: () => 'google-docs',
    moveTo(destination) { destination.addFile(this); return this; }
  };
  const workbookTemplate = { ...template, name: 'Template_for_Purchase_Orders', getId: () => 'demo-template', getMimeType: () => 'google-sheets' };
  const receivedLog = { ...workbookTemplate, name: 'Received_Log', getId: () => 'demo-log' };
  const blank = {};
  const templateBook = { getId: () => 'demo-template', getSheets: () => [blank], deleteSheet: tab => assert.equal(tab, blank) };
  const logTab = sheet('Blank');
  const logBook = { getId: () => receivedLog.getId(), getSheets: () => [logTab], getSheetByName: name => name === logTab.getName() ? logTab : null };
  const textStyle = {};
  for(const method of ['setFontFamily','setFontSize','setBold','setUnderline'])textStyle[method]=()=>textStyle;
  const body={
    clear(){paragraphs.length=0;return this;},
    getParagraphs:()=>[],
    getText:()=>paragraphs.join('\n'),
    appendParagraph(value){
      const index=paragraphs.push(value)-1;
      return {setText(text){paragraphs[index]=text;},setLineSpacing(){return this;},
        setSpacingBefore(){return this;},setSpacingAfter(){return this;},editAsText:()=>textStyle};
    }
  };
  for(const method of ['setPageWidth','setPageHeight','setMarginTop','setMarginBottom','setMarginLeft','setMarginRight'])body[method]=()=>body;
  const doc={getId:()=>template.getId(),getBody:()=>body,saveAndClose(){}};
  const purchase = app({
    MimeType: { GOOGLE_DOCS: 'google-docs', GOOGLE_SHEETS: 'google-sheets' },
    SpreadsheetApp: {
      WrapStrategy: { CLIP: 'clip' }, flush() {},
      create(name) {
        sheetsCreated++;
        if (name === receivedLog.name) return logBook;
        assert.equal(name, workbookTemplate.name);
        return templateBook;
      },
      openById(id) {
        if (id === workbookTemplate.getId()) return templateBook;
        assert.equal(id, receivedLog.getId());
        return logBook;
      }
    },
    DriveApp: {
      getFileById(id) {
        if (id === 'demo-doc') return template;
        if (id === 'demo-template') return workbookTemplate;
        if (id === receivedLog.getId()) return receivedLog;
        assert.equal(id, 'demo-workbook');
        return workingFile;
      }
    },
    DocumentApp: {
      create() {
        docsCreated++;
        return doc;
      },
      openById(id){assert.equal(id,template.getId());return doc;}
    }
  });
  purchase.setup = (book, options, templateOnly) => {
    assert.equal(book, templateBook);
    assert.equal(options.skip, true);
    assert.equal(templateOnly, true);
  };
  const book = { getId: () => 'demo-workbook' };
  assert.throws(() => purchase.driveAssets(book, 'summary'), /Run Purchase Tools/);
  const first = purchase.driveAssets(book, 'initialize');
  const projects = root.getFoldersByName('General_Purchase_Orders').next();
  assert.equal(workbookParent, projects);
  const second = purchase.driveAssets(book, 'initialize');
  assert.equal(moves, 1);
  assert.equal(first.root, second.root);
  assert.equal(first.template, second.template);
  assert.equal(docsCreated, 1);
  assert.equal(sheetsCreated, 2);
  assert.equal(first.receivedLog, second.receivedLog);
  assert.equal(purchase.receivedLogSheet(book), logTab);
  assert.equal(logTab.getRange('H1').getDisplayValues()[0][0], 'Paid details');
  assert.equal(root.getFoldersByName('Templates').next().getFilesByName(workbookTemplate.name).next(), workbookTemplate);
  assert.equal(first.summaries.getFoldersByName('Archive_Summary_changes').next(), first.archive);
  assert.equal(purchase.driveAssets(book, 'summary').archive, first.archive);
  assert.equal(purchase.driveAssets(book, 'orders').orders, first.orders);
  assert.equal(purchase.driveAssets(book, 'references').workbookTemplate, workbookTemplate);
  assert(paragraphs.indexOf('{{ITEM_END}}') > paragraphs.indexOf('{{ITEM_START}}') + 1);
  assert.equal(paragraphs[0],'Company: {{Company}}');
  assert.deepEqual(paragraphs.slice(paragraphs.indexOf('{{ITEM_START}}')+1,paragraphs.indexOf('{{ITEM_END}}')),
    ['{{Item_Number}}. Item: {{ItemName}}','Link:{{ItemLink}}','Unit Cost: ${{UnitCost}}','Quantity: {{Qty}}','Subtotal: ${{Subtotal}}']);
  assert.equal(paragraphs[paragraphs.indexOf('{{ITEM_START}}')-1],'');
  assert.equal(paragraphs[paragraphs.indexOf('{{ITEM_END}}')+1],'');
  assert.equal(paragraphs.at(-1),'Total: ${{TotalCost}}');
  body.clear();
  [
    'Purchase order','Vendor: {{Company}}','Grant: {{GrantNo}}   Category: {{Category}}','{{BlockID}}',
    '{{ITEM_START}}','{{Item_Number}}. {{ItemName}}','{{ItemID}}   Ordered on: {{OrderedOn}}','{{ItemLink}}',
    'Quantity: {{Qty}}   Unit cost: $ {{UnitCost}}   Subtotal: $ {{Subtotal}}',
    '{{ITEM_END}}','Items: {{TotalItems}}   Total: $ {{TotalCost}}'
  ].forEach(line=>body.appendParagraph(line));
  purchase.driveAssets(book,'initialize');
  assert.equal(paragraphs[0],'Company: {{Company}}');
  const compact=paragraphs.filter(Boolean).map(line=>line==='Link:{{ItemLink}}'?'Link: {{ItemLink}}':line);
  paragraphs.splice(0,paragraphs.length,...compact);
  purchase.driveAssets(book,'initialize');
  assert(paragraphs.includes('Link:{{ItemLink}}'));
  assert.equal(paragraphs[paragraphs.indexOf('{{ITEM_START}}')-1],'');
  body.appendParagraph('A custom order note');
  purchase.driveAssets(book,'initialize');
  assert.equal(paragraphs.at(-1),'A custom order note');
  root.createFolder('Monthly_summaries');
  assert.throws(() => purchase.driveAssets(book, 'summary'), /More than one Drive item/);
});
