const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');
const browserSource = fs.readFileSync(path.join(__dirname, 'script.js'), 'utf8');
const id = '123e4567-e89b-42d3-a456-426614174000';
const payload = {schemaVersion:2, submissionId:id, age:30, gender:'ذكر', answersV2:['نعم','نعم','لم يُطرح','لم يُطرح','لم يُطرح'], result:'إنفلونزا محتملة'};

function fixture(submissions, time, { lastRow = submissions + 1, maxRows = 3000 } = {}) {
  const timestamps = Array.from({length: lastRow - 1}, (_, index) => index < submissions ? new Date(time) : '');
  const ids = new Set();
  const events = [];
  const writes = [];
  const sheet = {
    getLastRow:()=>lastRow,
    getMaxRows:()=>maxRows,
    insertRowsAfter(row,count) {
      assert.equal(row,maxRows);
      maxRows += count;
    },
    getRange(row,column,height=1,width=1) {
      assert(row + height - 1 <= maxRows, 'range fits the sheet grid');
      return {
        getValues:()=>timestamps.map(timestamp=>[timestamp]),
        setValues([values]) {
          assert.equal(column,1);
          assert.equal(width,10);
          writes.push(values);
          timestamps.push(values[0]);
          ids.add(values[9]);
          lastRow = row;
          events.push('write');
        },
        setNumberFormat(){},
        createTextFinder(query) {
          assert.equal(column,10);
          return {matchEntireCell:()=>({findNext:()=>ids.has(query) ? {} : null})};
        }
      };
    }
  };
  const context = vm.createContext({
    Date:class extends Date {static now(){return time;}},
    SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet}),flush:()=>events.push('flush')},
    LockService:{getScriptLock:()=>({waitLock:()=>events.push('acquire'),releaseLock:()=>events.push('release')})},
    ContentService:{MimeType:{JSON:'json',JAVASCRIPT:'js'},createTextOutput:(text)=>({text,setMimeType(){return this;}})},
    console
  });
  vm.runInContext(source,context);
  return {
    post:(data=payload)=>JSON.parse(context.doPost({postData:{contents:JSON.stringify(data)}}).text),
    status:(submissionId)=>JSON.parse(context.doGet({parameter:{submissionId}}).text),
    jsonp:()=>context.doGet({parameter:{callback:'quizStatus_test'}}).text,
    writes, events
  };
}

const before = Date.parse('2027-11-18T21:59:59Z');
const after = Date.parse('2027-11-18T22:00:00Z');
test('the 2500th result saves, but the 2501st is rejected',()=>{
  const available=fixture(2499,before);
  assert.equal(available.post().ok,true);
  assert.deepEqual(available.events,['acquire','write','flush','release']);
  assert.equal(available.post({...payload,submissionId:'223e4567-e89b-42d3-a456-426614174000'}).locked,true);
  assert.equal(available.post().duplicate,true);
  assert.equal(available.writes.length,1);
  assert.deepEqual(available.status(id),{ok:true,locked:true,accepted:true});
  assert.match(available.jsonp(),/^quizStatus_test\(\{.*"locked":true.*\}\);$/);
});
test('the quiz closes at midnight on 19 November 2027 in Egypt',()=>{
  assert.equal(fixture(0,before).post().ok,true);
  const expired=fixture(0,after);
  assert.equal(expired.post().locked,true);
  assert.equal(expired.writes.length,0);
  assert.equal(expired.status().locked,true);
});

test('submissions expand the existing 1,000-row grid as needed',()=>{
  const current=fixture(999,before,{maxRows:1000});
  assert.equal(current.post().ok,true);
  assert.equal(current.writes.length,1);
});

test('blank timestamp rows do not count toward the submission limit',()=>{
  const current=fixture(2499,before,{lastRow:2501});
  assert.equal(current.status().locked,false);
  assert.equal(current.post().ok,true);
  assert.equal(current.status().locked,true);
});

test('the server stores all six browser paths in the Sheet1 column order',()=>{
  const screen={addEventListener(){}};
  const browser=vm.createContext({document:{getElementById:()=>screen}});
  vm.runInContext(`${browserSource}\nglobalThis.quiz = {questions,state,submissionPayload};`,browser);
  const {questions,state,submissionPayload}=browser.quiz;
  const paths=[
    {fever:'yes',beyondNose:'yes'},
    {fever:'yes',beyondNose:'no'},
    {fever:'no',discharge:'thick'},
    {fever:'no',discharge:'watery',allergySigns:'no'},
    {fever:'no',discharge:'watery',allergySigns:'yes',air:'yes'},
    {fever:'no',discharge:'watery',allergySigns:'yes',air:'no'}
  ];

  state.participant={age:'30',gender:'ذكر'};
  state.submissionId=id;
  for (const answers of paths) {
    state.answers=Object.assign(Object.fromEntries(questions.map(question=>[question.key,null])),answers);
    const browserPayload=submissionPayload();
    const current=fixture(0,before);
    assert.equal(current.post(browserPayload).ok,true);
    assert.deepEqual(Array.from(current.writes[0]).slice(1),[
      30,'ذكر',...browserPayload.answersV2,browserPayload.result,id
    ]);
  }
});

test('invalid routes and mismatched results are rejected',()=>{
  const current=fixture(0,before);
  assert.equal(current.post({...payload,answersV2:['نعم','نعم','خفيفة وشفافة ومائية','لا','لم يُطرح']}).ok,false);
  assert.equal(current.post({...payload,result:'نزلة برد محتملة'}).ok,false);
  assert.equal(current.writes.length,0);
});
