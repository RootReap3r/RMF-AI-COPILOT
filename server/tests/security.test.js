import {test} from 'node:test';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {once} from 'node:events';
import net from 'node:net';
const token='test-only-token-'.repeat(4);
async function port(){const server=net.createServer();server.listen(0,'127.0.0.1');await once(server,'listening');const p=server.address().port;await new Promise(r=>server.close(r));return p;}
test('missing authentication configuration stops startup',async()=>{
 const child=spawn(process.execPath,['index.js'],{env:{...process.env,AUTH_TOKEN:''},stdio:'pipe'});
 const [code]=await once(child,'exit');assert.notEqual(code,0);
});
test('API authentication and bounded document parsing',async()=>{
 const p=await port(), base=`http://127.0.0.1:${p}`;
 const child=spawn(process.execPath,['index.js'],{env:{...process.env,AUTH_TOKEN:token,PORT:String(p)},stdio:'pipe'});
 let logs='';child.stderr.on('data',b=>logs+=b);
 try{
  let ready=false;
  for(let n=0;n<100;n++){try{if((await fetch(base+'/api/health')).ok){ready=true;break;}}catch{}await new Promise(r=>setTimeout(r,50));}
  assert.ok(ready,logs);
  assert.equal((await fetch(base+'/api/storage')).status,401);
  assert.equal((await fetch(base+'/api/storage',{headers:{'X-Api-Key':'wrong'}})).status,401);
  const headers={'X-Api-Key':token};
  assert.equal((await fetch(base+'/api/storage',{headers})).status,200);
  const form=new FormData();form.append('files',new Blob(['safe text']),'test.txt');
  const parsed=await fetch(base+'/api/parse-document',{method:'POST',headers,body:form});
  assert.equal(parsed.status,200);assert.match((await parsed.json()).text,/safe text/);
  const large=new FormData();large.append('files',new Blob([new Uint8Array(5*1024*1024+1)]),'large.txt');
  assert.equal((await fetch(base+'/api/parse-document',{method:'POST',headers,body:large})).status,400);
  const many=new FormData();for(let n=0;n<5;n++)many.append('files',new Blob(['a']),n+'.txt');
  assert.equal((await fetch(base+'/api/parse-document',{method:'POST',headers,body:many})).status,400);
 }finally{child.kill();await once(child,'exit');}
});
