const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');
const id = '123e4567-e89b-42d3-a456-426614174000';
const payload = {schemaVersion:2, submissionId:id, age:30, gender:'ذكر', answersV2:['نعم','نعم','لم يُطرح','لم يُطرح','لم يُطرح'], result:'إنفلونزا محتملة'};

function fixture(submissions, time) {
  let writes = 0;
  let released = 0;
  const sheet = {
    getLastRow:()=>submissions + 1,
    getRange(row,column) {
      return {
        getValue:()=>row===1?'Submission ID':null,
        setValue(){}, setValues(){writes++;}, setNumberFormat(){},
        createTextFinder:()=>({matchEntireCell:()=>({findNext:()=>null})})
      };
    },
    hideColumns(){}
  };
  const context = vm.createContext({
    Date:class extends Date {static now(){return time;}},
    SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet})},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){released++;}})},
    ContentService:{MimeType:{JSON:'json',JAVASCRIPT:'js'},createTextOutput:(text)=>({text,setMimeType(){return this;}})},
    console
  });
  vm.runInContext(source,context);
  return {post:()=>JSON.parse(context.doPost({postData:{contents:JSON.stringify(payload)}}).text),
    status:()=>context.doGet({parameter:{callback:'quizStatus_test'}}).text,
    stats:()=>({writes,released})};
}

const before = Date.parse('2027-11-18T21:59:59Z');
const after = Date.parse('2027-11-18T22:00:00Z');
test('the 2500th result saves, but the 2501st is rejected',()=>{
  const available=fixture(2499,before);
  assert.equal(available.post().ok,true);
  assert.equal(available.stats().writes,1);
  const full=fixture(2500,before);
  assert.equal(full.post().locked,true);
  assert.equal(full.stats().writes,0);
  assert.equal(full.stats().released,1);
  assert.match(full.status(),/"locked":true/);
});
test('the quiz closes at midnight on 19 November 2027 in Egypt',()=>{
  assert.equal(fixture(0,before).post().ok,true);
  const expired=fixture(0,after);
  assert.equal(expired.post().locked,true);
  assert.equal(expired.stats().writes,0);
  assert.match(expired.status(),/"locked":true/);
});
