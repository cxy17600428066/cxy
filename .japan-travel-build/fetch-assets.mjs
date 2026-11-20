import fs from 'node:fs/promises';
import path from 'node:path';
const out = 'C:/Users/admin/Documents/New project/.japan-travel-build/assets';
await fs.mkdir(out, {recursive:true});
const assets = {
  cover: 'https://www.town.fujikawaguchiko.lg.jp/iju/images/environment/winter3.jpg',
  lake: 'https://res.cloudinary.com/jnto/image/upload/w_720%2Ch_480%2Cc_fill%2Cf_auto%2Cfl_lossy%2Cq_60/v1/media/filer_public/89/cc/89cce4ff-8576-48b2-8867-e120757a8c7c/pixta_62963427_l_sxllah',
  pagoda: 'https://storage.googleapis.com/relocate-production-store/japanbusiness3_07455a58dc/japanbusiness3_07455a58dc.jpg',
  tokyo: 'https://www.gotokyo.org/en/see-and-do/images/img_2.jpg'
};
for (const [name,url] of Object.entries(assets)) {
  const res = await fetch(url);
  if(!res.ok) throw new Error(`${name}: ${res.status}`);
  const bytes = Buffer.from(await res.arrayBuffer());
  await fs.writeFile(path.join(out, name+'.jpg'),bytes);
  console.log(name,bytes.length,res.headers.get('content-type'));
}
