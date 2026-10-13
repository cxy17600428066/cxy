from pathlib import Path
p=Path('outputs/mes_v2.js')
s=p.read_text(encoding='utf-8')
s=s.replace("p.type==='online'&&p.owner==='林青'", "p.type==='online'&&(p.owner==='林青'||!p.owner&&['天猫旗舰店','抖音官方店'].includes(p.channel))")
s=s.replace("s.enabled&&s.channels.includes(type)&&s.start<=c.end&&s.end>=c.start", "s.enabled&&s.channels.includes(type)&&s.start<=c.end&&s.end>=c.start||p&&s.code===p.sku")
s=s.replace("'<button onclick=\"edit('+id+')\">返回填报</button>'", "'<button onclick=\"closeModal()\">返回填报</button>'")
s=s.replace('onchange="adjust(Number(this.value))"', 'onchange="let chosen=Number(this.value);closeAll();adjust(chosen)"')
s=s.replace("function analysis(){let a=visible()", "function analysis(){if(type==='offline'&&period==='month'&&dimension==='客户')return monthlyCustomers();let a=visible()")
insert='''
function monthlyCustomers(){let a=visible().filter(p=>['已确认','已接收'].includes(p.status)),os=relevantOrders().filter(o=>a.some(p=>p.id===o.plan));return title('达成分析与计划归档','月度省区计划没有客户目标，不强行拆分客户计划量。')+'<div class="alert">客户维度仅展示已关联月计划的实际销售；月计划的计划量与达成率保留在省区 / 商品维度。订单回冲规则待确认，以下为模拟净履约量。</div><div class="card"><div class="cardhead"><h2>月计划按客户下钻</h2><select aria-label="分析维度" onchange="dimension=this.value;render()">'+['中心','大区','省区','客户','商品'].map(x=>'<option '+(dimension===x?'selected':'')+'>'+x+'</option>').join('')+'</select></div><div class="tablewrap"><table><thead><tr><th>客户</th><th>客户目标量</th><th>实际完成预览</th><th>达成率</th><th>下钻</th></tr></thead><tbody>'+group(os,o=>o.customer).map(([k,x])=>'<tr><td>'+esc(k)+'</td><td>— · 月计划未拆客户目标</td><td>'+fmt(sum(x)-sum(x,'returns'))+' 箱</td><td>—</td><td><button class="link" onclick="customerOrders('+JSON.stringify(x.map(o=>o.id)).replaceAll('"','&quot;')+')">订单明细 →</button></td></tr>').join('')+'</tbody></table></div></div>'}
function customerOrders(ids){modal('客户销售订单下钻',orderTable(relevantOrders().filter(o=>ids.includes(o.id))))}
'''
s=s.replace('// Keep entered form values',insert+'\n// Keep entered form values')
p.write_text(s,encoding='utf-8')
'''
'''
