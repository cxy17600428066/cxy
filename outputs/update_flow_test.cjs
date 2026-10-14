const fs=require('fs');let s=fs.readFileSync('outputs/check_mes_v2.cjs','utf8');
s=s.replace("for(const r of ['region','center','yan']){await p.selectOption('#role',r);await ev(()=>batchDialog([1]));", "for(const r of ['region','center','yan']){await p.selectOption('#role',r);await p.evaluate(r=>batchDialog(r==='region'?[1,2]:[1,2,3]),r);");
fs.writeFileSync('outputs/check_mes_v2.cjs',s);
let h=fs.readFileSync('outputs/check_mes_hierarchy.cjs','utf8').replace("['dan','渠道']","['dan','平台']");fs.writeFileSync('outputs/check_mes_hierarchy.cjs',h);
