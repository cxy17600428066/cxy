import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
const ROOT='C:/Users/admin/Documents/New project';
const BUILD=ROOT+'/.japan-travel-build';
const OUT=ROOT+'/日本5日游宣传';
const RUNTIME='C:/Program Files/WindowsApps/OpenAI.CodexPrimaryRuntime.v26-1119-907-0_26.1119.907.0_x64__3k8sg7r9htsxt/dependencies';
const {Presentation,PresentationFile}=await import(pathToFileURL(RUNTIME+'/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs'));
await fs.mkdir(OUT,{recursive:true});
await fs.mkdir(BUILD+'/renders',{recursive:true});
const p=Presentation.create({slideSize:{width:1280,height:720}});
const C={paper:'#F5F2EA',navy:'#173B53',blue:'#426E89',red:'#B75242',ink:'#233F4E',muted:'#687B82',white:'#FFFFFF'};
const FONT='Microsoft YaHei';
function text(s,str,x,y,w,h,size=26,color=C.ink,bold=false){
 const box=s.shapes.add({geometry:'textbox',name:str.split('\n')[0].slice(0,24),position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 box.text=str; box.text.style={typeface:FONT,fontSize:size,color,bold,insets:0,verticalAlignment:'top',autoFit:'none',wrap:'square'};
 return box;
}
async function photo(s,name,x,y,w,h){
 const bytes=await fs.readFile(BUILD+'/assets/'+name+'.jpg');
 return s.images.add({blob:bytes,contentType:'image/jpeg',alt:({cover:'河口湖与冬日富士山倒影',lake:'河口湖畔的富士山',pagoda:'新仓山浅间公园的忠灵塔与富士山',tokyo:'东京塔夜景'})[name],fit:'cover',position:{left:x,top:y,width:w,height:h}});
}
function slide(){const s=p.slides.add();s.background.fill=C.paper;return s;}
function notes(s,str){s.speakerNotes.text=str;}
function folio(s,n,color=C.muted){text(s,String(n).padStart(2,'0'),1190,667,40,25,18,color);}

// 01. A minimal editorial cover with editable type.
{
 const s=slide();
 await photo(s,'cover',486,0,794,720);
 text(s,'JAPAN  /  5 DAYS',62,66,390,32,21,C.blue);
 text(s,'富士山\n与东京',60,177,410,205,78,C.navy,true);
 text(s,'杭州出发  5天4晚',65,430,370,44,29,C.navy);
 text(s,'河口湖住一晚\n沿湖看山，走进东京',65,507,360,82,25,C.muted);
 text(s,'TOKYO + KAWAGUCHIKO',65,653,380,25,17,C.blue);
 notes(s,'行程来自本次对话确认的东京与富士山5日4晚方案。旅行日期尚未确定，照片展示代表性景色，不承诺出行当日天气或季节景观。图片来源：富士河口湖町 https://www.town.fujikawaguchiko.lg.jp/iju/environment/ 图片 https://www.town.fujikawaguchiko.lg.jp/iju/images/environment/winter3.jpg');
}
// 02. Five-day itinerary, one continuous typographic composition.
{
 const s=slide();
 text(s,'5日行程',65,57,760,70,52,C.navy,true);
 text(s,'杭州出发，东京往返，富士山安排两天',68,139,820,37,26,C.muted);
 const rows=[
 ['01','抵达东京','入住新宿。早到可逛涩谷，晚到休息'],
 ['02','河口湖与富士山','大石公园、湖畔散步，夜宿河口湖'],
 ['03','经典观景点','新仓山浅间公园，午后返回东京'],
 ['04','东京地标漫步','浅草寺、隅田川、银座与东京塔'],
 ['05','东京返回杭州','根据航班安排返程，预留机场时间']
 ];
 rows.forEach((r,i)=>{const y=220+i*89;text(s,r[0],68,y,65,51,35,C.red);text(s,r[1],168,y-1,660,42,29,C.navy,true);text(s,r[2],168,y+40,700,37,23,C.muted);});
 text(s,'05',945,222,270,170,132,C.blue);
 text(s,'DAYS',950,393,270,42,28,C.blue);
 text(s,'东京 3晚\n河口湖 1晚',950,476,255,96,27,C.navy);
 folio(s,2);
 notes(s,'每日安排为本次对话中的建议路线，日期未定。第1天是否游览涩谷取决于航班抵达时间，第5天以返程航班为准。东京观景参考 https://www.gotokyo.org/jp/spot/506/index.html 。富士山区域参考 https://www.japan.travel/en/spot/1329/ 。');
}
// 03. The scenic lake day.
{
 const s=slide();
 await photo(s,'lake',0,0,770,720);
 text(s,'DAY 02',828,62,360,32,21,C.red);
 text(s,'河口湖\n与富士山',826,127,392,147,52,C.navy,true);
 text(s,'上午抵达',830,320,350,39,27,C.navy,true);
 text(s,'新宿出发，约2小时车程\n抵达后先寄存行李',830,363,355,76,24,C.muted);
 text(s,'下午沿湖看山',830,466,360,40,27,C.navy,true);
 text(s,'大石公园与河口湖北岸\n留时间散步、喝咖啡',830,511,355,76,24,C.muted);
 text(s,'晚上入住温泉酒店',830,619,360,38,25,C.navy);
 folio(s,3);
 notes(s,'交通时长为大致行程时间，巴士可能堵车，班次以出行日为准。富士回游列车 https://www.fujikyu-railway.jp/fujikaiyuu/ 。新宿河口湖巴士 https://bus.fujikyu.co.jp/lp/en-shinjuku-kawaguchiko/ 。景点 https://www.japan.travel/en/spot/1329/ 。大石公园 https://www.fujisan-kyokai.jp/areaguide/2495 。照片来源 JNTO https://res.cloudinary.com/jnto/image/upload/w_720%2Ch_480%2Cc_fill%2Cf_auto%2Cfl_lossy%2Cq_60/v1/media/filer_public/89/cc/89cce4ff-8576-48b2-8867-e120757a8c7c/pixta_62963427_l_sxllah 。照片为河口湖代表性景观，不指代特定酒店或大石公园机位。');
}
// 04. The iconic mountain and pagoda composition.
{
 const s=slide();
 await photo(s,'pagoda',0,0,1280,473);
 text(s,'DAY 03',65,512,190,32,21,C.red);
 text(s,'新仓山浅间公园',63,558,710,66,48,C.navy,true);
 text(s,'五重塔与富士山的经典同框',66,634,710,38,25,C.muted);
 text(s,'清晨湖边看山，随后前往公园\n午后回东京，晚上逛新宿',847,520,375,89,23,C.navy);
 text(s,'观景台约398级台阶\n按体力调整，轻装出行',847,620,375,69,21,C.muted);
 notes(s,'观景台约398级台阶，信息来源富士吉田官方旅游指南 https://en.fujiyoshida.net/spot/index.php?p=12 。清晨湖景后前往公园为建议顺序，需以交通与天气调整。图片来源 https://www.relocate.world/destinations/japan/starting-a-business-in-Japan 图片 https://storage.googleapis.com/relocate-production-store/japanbusiness3_07455a58dc/japanbusiness3_07455a58dc.jpg 。秋色为照片拍摄时景观，不代表未知出行日期的景色。');
}
// 05. Tokyo's night panorama with a concise route.
{
 const s=slide(); s.background.fill='#10222D';
 await photo(s,'tokyo',495,0,785,720);
 text(s,'DAY 04',65,62,350,32,21,'#DFC7A4');
 text(s,'东京\n经典地标',63,133,415,155,58,C.white,true);
 text(s,'上午  浅草与隅田川',66,339,405,38,26,C.white,true);
 text(s,'浅草寺、雷门、仲见世街\n河畔散步，远眺晴空塔',66,388,405,79,23,'#C3D2D9');
 text(s,'下午至夜晚  银座与东京塔',66,519,405,40,25,C.white,true);
 text(s,'银座街头漫步\n增上寺、芝公园与东京塔夜景',66,567,405,82,23,'#C3D2D9');
 notes(s,'行程与前一轮旅行方案一致。隅田公园可欣赏晴空塔，来源 https://www.gotokyo.org/jp/spot/506/index.html 。芝公园周边地理参考 https://www.tokyo-park.or.jp/park/siba/assets/files/shiba_cn.pdf 。照片来源 MOTENAS JAPAN 页面 https://en.motenas-japan.jp/tokyo-nightlife/ ，图像标注 Yu Kato / Unsplash，图片 https://en.motenas-japan.jp/wp-content/uploads/2024/09/yu-kato-e8vEEav7yPU-unsplash-scaled.jpg 。照片为东京城市夜景示意，拍摄位置不等于当日地面游览机位。');
}
// 06. Practical ending, with a clearly stated planning budget.
{
 const s=slide();
 text(s,'住宿与旅行预算',65,57,1080,74,52,C.navy,true);
 text(s,'¥7,000–10,000',64,185,750,112,78,C.blue,true);
 text(s,'每人预算估算',68,307,580,43,29,C.navy);
 text(s,'按两人同行、合住一间房计算\n含往返机票、4晚住宿、交通、餐饮和常规门票',68,375,740,92,24,C.muted);
 text(s,'东京住新宿 3晚',881,198,335,43,28,C.navy,true);
 text(s,'前后尽量订同一家酒店\n提前询问大行李寄存',882,249,335,85,23,C.muted);
 text(s,'河口湖住温泉酒店 1晚',881,379,335,43,25,C.navy,true);
 text(s,'优先考虑车站接送\n想看山须确认房型朝向',882,431,335,88,23,C.muted);
 text(s,'观景受天气影响，河口湖住一晚能增加尝试机会',68,553,1135,42,27,C.navy,true);
 text(s,'日期尚未确定。预算不含购物与签证费用，并非实时机酒报价。\n旺季、房型与航班会影响总价。照片为代表性景色，能见度与季节景观以实际为准。',68,613,1125,70,21,C.muted);
 folio(s,6);
 notes(s,'预算为前一轮方案的规划估算，不是特定日期供应商报价。两人合住，每人7000至10000元人民币，含机票、4晚住宿、交通、餐饮、常规门票。不含购物与签证费用。富士山观景受天气影响，冬季晴天和清晨通常更适宜，来源 https://faq.japan-travel.jnto.go.jp/en/fuji-guide/views-of-fuji/ 。杭州东京航线以出行日期查询为准，航空公司班表 https://www.ana.co.jp/ja/cn/plan-book/promotions/multiple-routes-flight-resumption/ 。中国大陆普通护照需提前办妥签证，浙江常住居民可参考 https://www.shanghai.cn.emb-japan.go.jp/itpr_zh/00_000903.html 。');
}
for(let i=0;i<p.slides.items.length;i++){
 const s=p.slides.items[i];
 await fs.writeFile(BUILD+'/renders/slide-'+(i+1)+'.png',new Uint8Array(await (await p.export({slide:s,format:'png',scale:1})).arrayBuffer()));
 await fs.writeFile(BUILD+'/renders/slide-'+(i+1)+'.layout.json',await (await s.export({format:'layout'})).text());
 console.log('Rendered slide',i+1);
}
await (await PresentationFile.exportPptx(p)).save(BUILD+'/candidate.pptx');
console.log('Draft exported',BUILD+'/candidate.pptx');
