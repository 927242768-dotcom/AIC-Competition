/* ZhiGuYunXi AIC deck. Native/editable PowerPoint, generated with PptxGenJS.
 * npm install pptxgenjs
 * node build_deck.js
 * Put the verified project PNG files in ./assets before rebuilding.
 */
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
let H=null;
try { H=require(process.env.PPTX_HELPERS || '/home/oai/skills/slides/pptxgenjs_helpers'); } catch(e) {}
const pptx=new pptxgen();
pptx.layout='LAYOUT_WIDE';
pptx.author='智骨云析项目组'; pptx.subject='AIC AI+学科交叉：骨科CT空间认知教学';
pptx.title='智骨云析｜CT智能分割与空间认知教学实验平台';
pptx.company='智骨云析项目组'; pptx.lang='zh-CN';
pptx.theme={headFontFace:'Microsoft YaHei',bodyFontFace:'Microsoft YaHei',lang:'zh-CN'};
const W=40/3, HH=7.5, M=.6;
const C={ink:'153044',muted:'5B7080',teal:'087F83',aqua:'50B6B2',pale:'E8F3F3',bg:'F6F8FA',line:'D7E2E8',white:'FFFFFF',navy:'102A3A',amber:'AD6E2A',sand:'FBF3E7',soft:'EEF2F6'};
const font='Microsoft YaHei'; const numfont='Aptos Display';
const ST=pptx.ShapeType;
const root=__dirname, assets=path.join(root,'assets');
const dims=JSON.parse(fs.readFileSync(path.join(root,'asset_dimensions.json'),'utf8'));
const sourceBase='D:/国创项目/';
const model='experiments/20260917_185941_ctspine1k_v7_segformer3d_hr_hard_negative_precision/';
const upstream='https://openaccess.thecvf.com/content/CVPR2024W/DEF-AI-MIA/html/Perera_SegFormer3D_An_Efficient_Transformer_for_3D_Medical_Image_Segmentation_CVPRW_2024_paper.html';
const dataPaper='https://arxiv.org/abs/2105.14711';
const eduPaper='https://doi.org/10.1038/s41598-023-35046-2';
const rubric='https://www.aicomp.cn/tracks/tracks-2/3806.html';
const allTitles=[];
function rect(s,x,y,w,h,fill=C.white,line=null,r=0){s.addShape(r?ST.roundRect:ST.rect,{x,y,w,h,rectRadius:r,fill:{color:fill},line:{color:line||fill,width:line?.75:0},radius:r});}
function line(s,x1,y1,x2,y2,color=C.line,width=1,dash){s.addShape(ST.line,{x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color,width,...(dash?{dashType:'dash'}:{})}});}
function tx(s,t,x,y,w,h,size=18,color=C.ink,bold=false,opts={}){
 const o={x,y,w,h,fontFace:font,fontSize:size,color,bold,margin:0,breakLine:false,valign:'mid',paraSpaceAfterPt:0,lineSpacingMultiple:1.12,fit:'shrink',...opts};
 s.addText(t,o);
}
function tag(s,t,x,y,w=1.4,color=C.teal,fill=C.pale){rect(s,x,y,w,.3,fill);tx(s,t,x+.09,y+.02,w-.18,.26,10.5,color,true);}
function arrow(s,x,y,w=.35,color=C.teal){s.addShape(ST.chevron,{x,y,w,h:.18,fill:{color},line:{color,width:0}});}
function img(s,name,x,y,w,h,mode='contain'){
 const file=path.join(assets,name);let o;
 if(H){o=mode==='crop'?H.imageSizingCrop(file,x,y,w,h):H.imageSizingContain(file,x,y,w,h);}
 else {const [iw,ih]=dims[name],q=Math.min(w/iw,h/ih),ww=iw*q,hh=ih*q;o={x:x+(w-ww)/2,y:y+(h-hh)/2,w:ww,h:hh};}
 s.addImage({path:file,...o,altText:'真实项目素材：'+name});
}
function notes(s,text,sources=[]){s.addNotes(`${text}\n\n[Sources]\n${sources.map(v=>v.startsWith('http')?v:sourceBase+v).join('\n')}\n[/Sources]`);}
function base(section,title,sub='',dark=false){
 const s=pptx.addSlide();s.background={color:dark?C.navy:C.bg};s._tag=title;allTitles.push(title);
 tx(s,section,M,.32,8,.26,10.5,dark?C.aqua:C.teal,true,{charSpacing:1.2});
 tx(s,'ZHIGU YUNXI',10.9,.32,1.83,.26,10,dark?'B8CFD8':C.muted,false,{align:'right',charSpacing:1.5});
 tx(s,title,M,.87,12.05,.59,29,dark?C.white:C.ink,true);
 if(sub)tx(s,sub,M,1.48,12.08,.44,13.5,dark?'B8CFD8':C.muted);
 line(s,M,6.98,12.73,6.98,dark?'315061':C.line,.65);
 tx(s,'智骨云析  /  AI + 医学影像教学',M,7.08,8,.18,9,dark?'96B6C5':C.muted);
 tx(s,String(pptx._slides.length).padStart(2,'0'),12.13,7.06,.6,.24,10,dark?'96B6C5':C.muted,false,{align:'right'});
 return s;
}
function foot(s,text,dark=false){tx(s,text,M,6.56,12.12,.26,10.5,dark?'A7C1CE':C.muted);}
function metric(s,value,label,x,y,w=2.7,col=C.teal){tx(s,value,x,y,w,.76,43,col,true,{fontFace:numfont});tx(s,label,x,y+.85,w,.44,14,C.muted);}
function item(s,n,title,body,x,y,w=3.6){tx(s,n,x,y,.5,.38,15,C.teal,true,{fontFace:numfont});tx(s,title,x+.62,y,w-.62,.4,19,C.ink,true);tx(s,body,x+.62,y+.53,w-.62,.88,15,C.muted);}
function table(s,headers,rows,x,y,widths,rowh=.5){
 const total=widths.reduce((a,b)=>a+b,0);rect(s,x,y,total,rowh,C.navy);
 let xx=x;headers.forEach((v,i)=>{tx(s,v,xx+.14,y+.04,widths[i]-.28,rowh-.08,12,C.white,true,{align:i?'center':'left'});xx+=widths[i];});
 rows.forEach((r,k)=>{let yy=y+(k+1)*rowh;rect(s,x,yy,total,rowh,k%2?C.white:'EDF3F5');let xx=x;r.forEach((v,i)=>{tx(s,String(v),xx+.14,yy+.04,widths[i]-.28,rowh-.08,13,i===0?C.ink:C.muted,i===0,{align:i?'center':'left'});xx+=widths[i];});});
}
// 01. Cover: transparent patient-specific v7 mesh, editorial rather than generic technology wallpaper.
{
 const s=pptx.addSlide();s.background={color:C.navy};allTitles.push('智骨云析');
 rect(s,8.3,0,5.033,7.5,'143547');
 line(s,8.35,.55,8.35,6.9,'2B5162',.8);
 tx(s,'AIC 2026  /  AI + 学科交叉',.68,.58,7,.38,12,C.aqua,true,{charSpacing:1.5});
 tx(s,'智骨云析',.68,1.55,7.5,1.1,54,C.white,true);
 tx(s,'CT 智能分割与\n三维空间认知教学实验平台',.72,2.9,7.1,1.2,26,C.white,true);
 tx(s,'SegFormer3D 多尺度编码  ×  高分辨率解码  ×  物理空间交互',.73,4.36,6.95,.65,15,'BBD1DB');
 line(s,.73,5.3,7.35,5.3,'345362',1);
 tx(s,'队长  贾泽铭     队员  林荣盛、彭彦博',.73,5.63,7,.4,15,'D3E2E8');
 tx(s,'全球校园人工智能算法精英大赛 · 智青春·算未来',.73,6.72,7.2,.28,10.5,'9CB9C7');
 img(s,'mesh_hero.png',9.19,.66,2.1,6.02);
 tag(s,'v7  PREDICTION',8.88,.18,2.0,C.aqua,'183F50');
 tx(s,'来自真实 CT 的三维分割表面',8.7,6.93,4.0,.27,10,'A9C7D3',false,{align:'center'});
 notes(s,'开场：我们聚焦医学生从二维断层影像建立三维空间关系的学习过程。平台把已有三维分割结果转化为可逐步调用、可与专家标注核验的影像参照。右侧为 v7 在 validation197 病例 0663 上保存的 prediction 重建，不是通用示意图。模型任务为椎骨前景二分类。',[model+'full_volume_validation197/summary.json','docs/13_v7_final_model_freeze.md','../AIC比赛/01_作品说明文档/智骨云析_作品说明书.docx','../AIC比赛/99_历史备份_非提交/重制工作区_20260923/assets/mesh_summary.json']);
}
// 02. Needs, linked to actual learning actions rather than invented survey statistics.
{
 const s=base('01  需求分析','把“看见切片”转化为“理解空间”','学科场景：医学影像与骨科解剖入门 · 教师演示、实验实训、学生自主练习');
 const ys=[2.17,3.52,4.87];
 [['01','断层定位','一张切片中的局部形态，难以直接对应完整三维结构。','三平面联动 + 连续切片追踪'],['02','提示时机','没有参照容易卡住；直接展示标注又跳过了自主观察。','先判断，再逐级调用分割参考'],['03','错误复核','只知道答案对错，仍缺少重新定位和比较的操作路径。','回到原始 CT、专家 GT 与三维空间']].forEach((r,i)=>{
  tx(s,r[0],.65,ys[i],.53,.49,24,C.teal,true,{fontFace:numfont});
  tx(s,r[1],1.35,ys[i],2.0,.44,20,C.ink,true);
  tx(s,r[2],3.28,ys[i]-.04,5.22,.56,16,C.muted);
  tx(s,r[3],3.28,ys[i]+.66,5.22,.36,15,C.teal,true);
  if(i<2)line(s,.66,ys[i]+1.15,8.7,ys[i]+1.15);
 });
 rect(s,9.16,2.09,3.57,4.16,C.navy);img(s,'raw_sagittal.png',9.38,2.3,3.13,3.46);
 tx(s,'真实矢状位 CT',9.35,5.87,3.2,.23,11,'C6DCE5',false,{align:'center'});
 foot(s,'需求来自教学任务分析；空间能力与解剖学习的关联参考 Koh 等（2023），不作为本平台教学效果证明。');
 notes(s,'说明三个具体教学动作：在单层判断、沿相邻层面追踪、在三维空间核验。已有研究讨论空间能力和三维模型呈现对解剖学习的影响，但该研究不是本项目或 CT 教学的验证。本项目尚无师生需求问卷或教学试点统计，不展示虚构百分比。',[eduPaper,'web/frontend/teaching.js','web/frontend/practice.js']);
}
// 03. Solution and outcome, readable flow above the product image.
{
 const s=base('02  解决方案','同一病例，贯通观察、提示、核验与复盘','AI 的作用是提供结构参照；学生的判断仍回到真实影像中完成。');
 const labels=[['自主观察','Raw CT'],['分级提示','规则提示 → Prediction'],['空间核验','连续切片 + MPR'],['三维理解','物理空间表面'],['练习复盘','解析 → 对应实验']];
 labels.forEach((r,i)=>{let x=.65+i*2.46;tag(s,String(i+1).padStart(2,'0'),x,2.05,.48);tx(s,r[0],x,2.54,2.1,.44,20,C.ink,true);tx(s,r[1],x,3.03,2.12,.35,12.7,C.muted);if(i<4)arrow(s,x+2.15,2.65,.2);});
 rect(s,.65,3.67,7.66,2.57,C.white,C.line);img(s,'home_ui.png',.75,3.73,7.46,2.45,'crop');
 tx(s,'一个教学实验环境',8.8,3.9,3.8,.42,22,C.teal,true);
 tx(s,'病例与影像有真实来源\n分割与标注保持独立切换\n二维、三维共享病例上下文',8.8,4.53,3.75,1.52,18,C.ink);
 foot(s,'当前定位：医学影像教学与科研辅助原型；不是自动疾病诊断，也不是通用大模型问答系统。');
 notes(s,'产品流程分成五个动作。前两级提示来自教学规则，第三级才展示当前病例已保存的模型 prediction。MPR 与 3D 通过 evaluation ID 和 case ID 关联。基础练习目前为 6 道规则判分题，错题提供章节/实验入口，尚未形成完整个体学习画像。',['web/frontend/teaching.js','web/frontend/practice.js','web/frontend/index.html']);
}
// 04. Three actual technical contributions, explicitly distinct from the upstream encoder.
{
 const s=base('03  技术创新','三项改进，连接模型输出与教学操作','以现有 SegFormer3D 编码器为基础，自研工作集中于解码、工程优化与教学交互。');
 const cc=[['01','高分辨率三维解码','C1–C4 逐级融合\n原始 CT 浅层分支\n恢复至输入空间尺寸','解决分割细节表达问题','segformer3d_hr_decoder.py'],['02','全体积误检控制','FP 加权 Tversky + CE\n完整体积滑窗复核\n固定连通域后处理','处理 patch 与全体积的差异','joint_loss.py / evaluate.py'],['03','可核验的影像实验','按需显示模型分割\nGT 定位反馈与三平面联动\nPrediction → 物理空间 3D','把 AI 结果转为学生操作','teaching.js / mesh.py']];
 cc.forEach((r,i)=>{let x=.65+i*4.12;rect(s,x,2.09,3.87,4.16,C.white,C.line);rect(s,x,2.09,3.87,.065,i===1?C.ink:C.teal);tx(s,r[0],x+.23,2.37,1,.57,32,C.aqua,true,{fontFace:numfont});tx(s,r[1],x+.23,3.12,3.39,.47,20,C.ink,true);tx(s,r[2],x+.23,3.86,3.4,1.29,16,C.muted);tx(s,r[3],x+.23,5.45,3.4,.31,13.5,C.teal,true);tx(s,r[4],x+.23,5.91,3.4,.2,9.7,C.muted);});
 foot(s,'开源编码器、标准算法与本项目适配分别注明来源；不将常用技术组合描述为全领域首次。');
 notes(s,'创新陈述限定为本项目的结构改进、错误驱动工程优化和教学场景适配。SegFormer3D 编码器不是自主原创。Tversky、Marching Cubes、GroupNorm 等为已有方法。v6 与 v7 是组合 pipeline 对照，不是单模块控制变量实验。残余假阳性精修 v7.1 完成但未被采纳，不列作性能成功点。',[upstream,'src/modeling/segformer3d_hr_decoder.py','src/modeling/joint_loss.py','web/frontend/teaching.js','src/reconstruction/mesh.py']);
}
// 05. Native editable architecture diagram, exact 64^3 patch feature scales.
{
 const s=base('03  技术创新 / 网络结构','多尺度语义与浅层空间细节逐级融合','SegFormer3D Encoder + High-Resolution 3D Decoder · 以 64³ 输入为例');
 rect(s,.67,2.13,1.25,1.1,C.navy);tx(s,'CT',.85,2.3,.88,.39,22,C.white,true,{align:'center'});tx(s,'1 × 64³',.76,2.82,1.05,.23,12,'B9D2DF',false,{align:'center'});
 const xs=[2.45,4.48,6.51,8.54],sc=['16³','8³','4³','2³'],ch=['32','64','160','256'];
 tx(s,'上游多尺度编码器',2.44,2.04,8.1,.27,12,C.muted);
 xs.forEach((x,i)=>{rect(s,x,2.53,1.55,1.02,'E1EBF0');tx(s,'C'+(i+1),x+.1,2.66,1.35,.36,19,C.ink,true,{align:'center'});tx(s,sc[i]+' · '+ch[i]+' ch',x+.08,3.12,1.39,.22,11.3,C.muted,false,{align:'center'});if(i<3)arrow(s,x+1.7,2.91,.19,C.muted);});
 arrow(s,2.08,2.87,.2,C.muted);
 const dx=[10.7,8.54,6.51,4.48,2.45,.67];
 const dl=[['投影 C4','64 ch'],['融合 C3','64 ch'],['融合 C2','48 ch'],['融合 C1','32 ch'],['融合 1/2','24 ch'],['融合 1×','16 ch']];
 line(s,10.11,3.04,11.47,3.04,C.teal,1.4);line(s,11.47,3.04,11.47,4.22,C.teal,1.4);
 dx.forEach((x,i)=>{rect(s,x,4.22,1.5,.91,i===0?C.teal:C.pale,i===0?null:C.teal);tx(s,dl[i][0],x+.05,4.34,1.4,.27,13,i===0?C.white:C.teal,true,{align:'center'});tx(s,dl[i][1],x+.06,4.73,1.38,.23,11.5,i===0?'CDEDEC':C.muted,false,{align:'center'});if(i<5){s.addShape(ST.chevron,{x:x-.39,y:4.58,w:.18,h:.17,rotate:180,fill:{color:C.teal},line:{color:C.teal,width:0}});}});
 [[8.54,2],[6.51,3],[4.48,4]].forEach(([x,i])=>{});
 [[6.51,8.54],[4.48,6.51],[2.45,4.48]].forEach(([x,z])=>{line(s,x+.77,3.55,x+.77,3.83,C.aqua,1.3);line(s,x+.77,3.83,z+.75,3.83,C.aqua,1.3);line(s,z+.75,3.83,z+.75,4.22,C.aqua,1.3);});
 line(s,1.28,3.23,1.28,3.87,C.teal,1.1);line(s,1.28,3.87,3.19,3.87,C.teal,1.1);line(s,3.19,3.87,3.19,4.22,C.teal,1.1);line(s,1.28,3.87,1.28,4.22,C.teal,1.1);
 tag(s,'CT 浅层分支',.66,5.6,1.7);tx(s,'1/2 与 1× 分辨率特征参与解码',2.58,5.6,4.88,.35,15,C.ink);
 tag(s,'输出对齐',8.39,5.6,1.18);tx(s,'2 × 64³ logits',9.8,5.61,2.55,.3,15,C.ink,true);
 foot(s,'融合块：三线性上采样 → 特征拼接 → Conv3D / GroupNorm / GELU；输出与输入保持同空间尺寸。');
 notes(s,'模型读取编码器 C1-C4 feature maps，而不是简单上采样原 all-MLP decoder 的低分辨率 logits。逐级与 skip feature 拼接，再融合原始 CT 的 half/full-resolution shallow path。图中 C1-C4 尺度及 decoder 通道数来自配置和源码。该结构的独立边界增益尚缺完整同协议消融，因此描述机制与实现，不单独归因数值提升。',['src/modeling/segformer3d_hr_decoder.py','configs/orthopedic_ct_full_large_scale_v7_segformer3d_hr_hard_negative_precision.yaml',upstream]);
}
// 06. Training rationale and transparent attribution.
{
 const s=base('03  技术创新 / 优化策略','将假阳性控制写入优化目标与推理流程','问题由完整体积暴露：v6 召回较高，但 Precision 仅 0.3434，预测前景膨胀至 3.71 倍。');
 rect(s,.65,2.12,5.5,3.87,C.navy);
 tx(s,'FP 加权的区域优化',.91,2.4,4.95,.47,23,C.white,true);
 tx(s,'Tversky = TP / (TP + 0.7FP + 0.3FN)',.91,3.22,4.95,.65,19,'E4F2F5',false,{fontFace:'Aptos'});
 tx(s,'L = 0.8 × (1 − Tversky) + 1.2 × CE',.91,4.18,4.95,.54,19,'E4F2F5',false,{fontFace:'Aptos'});
 tx(s,'对误检施加更强惩罚，同时保留区域重叠与分类约束。',.91,5.04,4.91,.6,16,'BBD3DE');
 item(s,'01','保留背景判别','64³ patch 训练；前景概率 0.75，记录实际前景 / 背景采样统计。',6.69,2.16,5.85);
 item(s,'02','完整体积滑窗','96³ ROI、50% overlap，检查局部学习与完整 CT 的差异。',6.69,3.53,5.85);
 item(s,'03','固定结构后处理','剔除小于 8,192 体素的连通域，主协议确定后冻结。',6.69,4.89,5.85);
 foot(s,'v6 → v7 同时涉及训练配置和后处理变化，属于系统级改进；不能视为某个模块的独立消融收益。');
 notes(s,'损失定义以 joint_loss.py 为准，alpha 对应 FP，beta 对应 FN。v7 hard_sampling 配置保留 high_loss 接口，但历史 guidance 仅有 7 例，不能声称 v7 对 train610 完成全量困难样本引导；v7.1 才实际完成 610/610 residual-FP guidance。主讲只展示确实落实的优化目标、采样统计与固定后处理。',['src/modeling/joint_loss.py','configs/orthopedic_ct_full_large_scale_v7_segformer3d_hr_hard_negative_precision.yaml',model+'sampling_stats.csv',model+'full_volume_validation197/summary.json']);
}
// 07. Data rigor without conflating the historical pilot test.
{
 const s=base('04  数据与工程实施','807 例开发数据，先建立可追溯的影像基础','CTSpine1K · 按病例划分开发集 · 当前任务为椎骨前景二分类分割');
 metric(s,'610','训练病例',.75,2.07,2.2);metric(s,'197','验证病例',3.21,2.07,2.2);metric(s,'1.5 mm','各向同性缓存间距',6.05,2.07,3.0);metric(s,'807 / 807','全量缓存与审计完成',9.34,2.07,3.4);
 const blocks=[['稳定读取','完整性 / CRC\n异常读取回退'],['强度处理','HU：−1000 ～ 2000\n固定 min-max 缓存'],['空间一致','shape / spacing\naffine / 标签对齐'],['标签约束','原标签 1–25\n训练统一映射为前景']];
 blocks.forEach((r,i)=>{let x=.65+i*3.1;rect(s,x,4.04,2.87,1.79,C.white,C.line);tx(s,r[0],x+.2,4.26,2.47,.36,18,C.ink,true);tx(s,r[1],x+.2,4.84,2.47,.67,14,C.muted);});
 foot(s,'本轮 train610 / validation197 未将 test198 用于训练或选参；历史 10 例 pilot 与当前主线分开归档。');
 notes(s,'当前开发缓存为 0.6.0-compact-u16-fixed-hu-minmax-1p5mm-nibsafe，不能误用早期 1 mm、case-wise z-score 作为 v7 协议。807 例全量 audit missing/QC bad/geometry bad/suspicious HU 均记录为 0。原始椎体类别合并为 binary foreground，因此不能将 Prediction 解释为椎弓根、椎板或单椎体编号自动分类。历史 pilot 曾访问 liver_169，本页不声称所有 test 数据在项目历史上从未被访问。',['docs/14_ctspine1k_v6_fullscale_stage_20260916.md','configs/orthopedic_ct_full_large_scale_v7_segformer3d_hr_hard_negative_precision.yaml','PROJECT_STATUS.md',dataPaper]);
}
// 08. Runtime architecture and feasibility.
{
 const s=base('05  方案可行性 / 系统架构','训练与课堂解耦：离线生成，交互读取','课堂不等待重新训练或完整体积推理，而是使用可回溯到模型版本的病例资产。');
 const cols=[.67,4.82,8.98];
 [['离线算法层','CT 缓存 / 固定 split\nSegFormer3D-HR 训练\n全体积评估与版本冻结'],['病例资产层','Prediction + 逐例指标\n原始 CT + 专家 GT\n物理几何 + Mesh / SDF'],['教学交互层','FastAPI / MPR 图像接口\n学生实验 / 教师演示\nWebGL2 三维观察与测量']].forEach((r,i)=>{rect(s,cols[i],2.22,3.68,2.55,i===1?C.navy:C.white,i===1?null:C.line);tag(s,'0'+(i+1),cols[i]+.22,2.46,.47,i===1?C.aqua:C.teal,i===1?'224859':C.pale);tx(s,r[0],cols[i]+.22,3.04,3.24,.45,21,i===1?C.white:C.ink,true);tx(s,r[1],cols[i]+.22,3.69,3.25,.83,15,i===1?'C4D8E1':C.muted);if(i<2)arrow(s,cols[i]+3.84,3.49,.18);});
 rect(s,.67,5.18,12.0,.92,C.pale);tx(s,'共享 evaluation_id + case_id',.92,5.39,4.29,.46,20,C.teal,true,{fontFace:'Aptos'});tx(s,'数值指标、二维叠加与三维表面保持同源关联',5.5,5.38,6.82,.48,17,C.ink);
 foot(s,'演示读取已有真实产物；当前 /infer 接口并不等于已开放任意新病例的实时模型推理。');
 notes(s,'这是已经运行的工程架构，不把本地可运行等同于无需病例数据的独立便携软件。Web API 健康检查存在，saved evaluation 资产可加载。病例迁移依赖缓存、predictions、路径配置和许可范围。v7 validation 平均 inference_seconds 约 92.69 s，仅代表该本机 CPU 评估记录，非页面响应时间。',['web/backend/app.py','web/frontend/teaching.js','web/frontend/research_3d.js','web/desktop_shell.py',model+'full_volume_validation197/summary.json']);
}
// 09. Workload and documented milestones.
{
 const s=base('06  项目实施 / 研发工作量','用数据、实验和软件产物记录研发过程','从真实 CT 接入，到算法迭代，再到可操作的教学原型。');
 const yy=2.6,xs=[.89,3.31,5.73,8.15,10.57];line(s,.98,yy,11.94,yy,C.line,2);
 const r=[['08.15–08.16','真实数据接入','NIfTI / 几何 / QC\n10 例工程链验证'],['08.26–08.29','机制与消融','归一化 / 采样 / 损失\n精修与重建验证'],['09.16','扩展开发规模','807 例缓存审计\n610 / 197 固定划分'],['09.17–09.22','v7 主模型冻结','197 例全体积评估\nv7.1 负结果归档'],['09.22–09.23','教学原型集成','MPR / 定位反馈\n三维与课堂模式']];
 r.forEach((a,i)=>{s.addShape(ST.ellipse,{x:xs[i],y:yy-.07,w:.14,h:.14,fill:{color:C.teal},line:{color:C.teal}});tx(s,a[0],xs[i]-.18,2.12,2.04,.29,12,C.teal,true);tx(s,a[1],xs[i]-.18,3.03,2.08,.38,18,C.ink,true);tx(s,a[2],xs[i]-.18,3.64,2.14,.9,14,C.muted);});
 line(s,.66,4.92,12.67,4.92);
 metric(s,'807','例开发缓存 / 几何审计',.74,5.11,3.52);
 metric(s,'610','例残余 FP 引导生成与审计',4.94,5.11,3.6);
 metric(s,'3 × 197','v6 / v7 / v7.1 全体积评估',9.12,5.11,3.62);
 foot(s,'3 × 197 指同一验证队列的三个版本评估，不是 591 名独立患者；v7.1 产物保留为负向实验。');
 notes(s,'工作量用实际产物体现，避免把文件数或代码行数当作创新指标。日期来自台账和 experiment 路径。当前检查复用了冻结模型结果，不重新运行 validation197 或 test。历史消融多数属于 10-case pilot，不与扩样后的 v7 主结果混合。',['PROJECT_STATUS.md','docs/14_ctspine1k_v6_fullscale_stage_20260916.md','docs/14_training_line_final_handoff.md']);
}
// 10. Collaboration, no invented personal allocation.
{
 const s=base('06  项目实施 / 协作与质量','以模块交付和交叉复核组织工程协作','参赛成员：队长 贾泽铭 · 队员 林荣盛、彭彦博');
 const rows=[['数据与质控','病例清单、几何审计、标签规则','进入训练前核对来源与空间一致性'],['模型与评估','配置、checkpoint、逐病例指标','训练选择器与完整体积结果分列'],['三维与 Web','MPR、Mesh / SDF、教学交互','在同一病例链路核验真实预测'],['材料与发布','说明书、演示、PPT、Git 记录','数字、截图、版本信息交叉复核']];
 table(s,['协作模块','交付物','复核要求'],rows,.67,2.17,[2.08,4.32,5.61],.66);
 tag(s,'质量记录',.7,5.85,1.2);tx(s,'已有自动化测试与 CI；本次补充关键模块回归：7 passed',2.17,5.82,10.3,.46,17,C.teal,true);
 foot(s,'协作过程保留配置、结果与版本记录；模块交付以可运行产物和可复核证据为准。');
 notes(s,'成员姓名来自本地作品说明书。资料未明确三个人分别负责哪个模块，不能臆造个人贡献或专业背景。图表描述工程的模块协作与复核要求。7 passed 是本次已执行的 test_reconstruction_mesh / test_segformer3d_hr_decoder / test_web_health 关键回归记录，不是全量当前测试数。历史 138 passed 仅属于当时阶段，未写成最新全量结果。正式报名材料如要求逐人责任，应由团队补充确认。',['../AIC比赛/01_作品说明文档/智骨云析_作品说明书.docx','TASKS.md','CONTRIBUTING.md','../AIC比赛/99_历史备份_非提交/重制工作区_20260923/qa_tests.txt']);
}
// 11. Main result. No patch metric passed off as a full-volume result.
{
 const s=base('07  应用效果 / 算法验证','完整 CT 体积验证：Dice 0.7933','v7 · 固定 validation197 · 197 / 197 病例完成 · 逐病例指标取均值',true);
 tx(s,'0.7933',.72,2.21,6.0,1.53,86,C.white,true,{fontFace:numfont});
 tx(s,'Mean Dice',.83,3.91,5.7,.48,23,C.aqua,false,{fontFace:'Aptos'});
 line(s,6.83,2.32,6.83,5.92,'315363',1);
 [['0.7862','Precision'],['0.8245','Recall'],['0.7347','IoU']].forEach((r,i)=>{let y=2.25+i*1.21;tx(s,r[0],7.4,y,3.18,.65,36,C.white,true,{fontFace:numfont});tx(s,r[1],10.59,y+.15,2.0,.31,15,'B8D0DD',false,{fontFace:'Aptos'});});
 tx(s,'37.95 mm',.84,5.12,3.0,.54,29,C.aqua,true,{fontFace:numfont});tx(s,'HD95',.84,5.77,2.65,.3,13,'B8D0DD',false,{fontFace:'Aptos'});
 tx(s,'8.81 mm',3.93,5.12,2.64,.54,29,C.aqua,true,{fontFace:numfont});tx(s,'ASSD',3.93,5.77,2.65,.3,13,'B8D0DD',false,{fontFace:'Aptos'});
 foot(s,'主协议：96³ ROI · overlap 0.5 · 连通域阈值 8,192 体素；属于开发期验证结果，不是独立测试或临床结论。',true);
 notes(s,'强调结果的评价集合和性质。0.9402 是 patch selector，不能作为本页主性能。v7 full-volume Dice mean 为 0.7933430437，std 为 0.3033803606，后续一页专门展示困难病例分布。HD95、ASSD 的均值仍较高，当前不足以支撑临床测量或自动诊断。',[model+'full_volume_validation197/summary.json','docs/13_v7_final_model_freeze.md']);
}
// 12. Native shape bar chart and protocol footnote.
{
 const s=base('07  应用效果 / 版本对照','从前景过预测，转向更均衡的系统输出','在同一 validation197 上比较已保存的 v6 与 v7 完整体积结果。');
 tag(s,'v6',.77,2.05,.61,C.muted,'E4EBEF');tag(s,'v7',1.6,2.05,.61,C.teal,C.pale);
 const vals=[[.4845,.7933],[.3434,.7862],[.8686,.8245]],labs=['Dice','Precision','Recall'];
 const x0=.94,ybase=5.89,maxh=3.0;
 [0,.5,1].forEach(v=>{let y=ybase-v*maxh;line(s,.85,y,8.71,y,'DCE5E9',.7);tx(s,v.toFixed(1),.45,y-.11,.32,.21,9.5,C.muted);});
 vals.forEach((v,i)=>{let x=x0+i*2.65;v.forEach((z,j)=>{rect(s,x+j*.84,ybase-z*maxh,.65,z*maxh,j?C.teal:'A8BBC7');tx(s,z.toFixed(4),x+j*.84-.12,ybase-z*maxh-.38,.91,.26,12.5,j?C.teal:C.muted,true,{align:'center',fontFace:'Aptos'});});tx(s,labs[i],x-.07,6.02,1.85,.31,14,C.ink,true,{align:'center',fontFace:'Aptos'});});
 rect(s,9.16,2.11,3.55,4.14,C.white,C.line);
 tx(s,'误检与表面偏差',9.4,2.42,3.06,.42,20,C.ink,true);
 [['前景体积比','3.71× → 1.68×'],['HD95','166.10 → 37.95 mm'],['ASSD','38.33 → 8.81 mm']].forEach((r,i)=>{let y=3.15+i*.88;tx(s,r[0],9.4,y,3.03,.25,12,C.muted);tx(s,r[1],9.4,y+.33,3.03,.38,17,C.teal,true,{fontFace:'Aptos'});});
 foot(s,'v6 无连通域过滤，v7 使用 8,192 体素阈值；图中改善包含训练与后处理贡献，不是网络结构的单独增益。');
 notes(s,'v6 mean Dice 0.4845271987、Precision 0.3433617778、Recall 0.8686251163。v7 分别为 0.7933430437、0.7862013885、0.8244633247。Dice 绝对差约 0.3088，但没有同协议单因素消融，不能把它归为 HR decoder 或 hard mining 的独立提升。两个版本均有 HR decoder；改善并不是从无 HR 到有 HR。',['experiments/20260916_163356_ctspine1k_v6_segformer3d_hr_full610_val197/full_volume_validation197/summary.json',model+'full_volume_validation197/summary.json']);
}
// 13. Distribution and limitations are integral, not tiny fine print.
{
 const s=base('07  应用效果 / 失败分析','平均值之外，保留困难病例的真实分布','v7 validation197 · Dice 分箱来自 metrics_per_case.csv，不重新推理或重新选参。');
 const n=[32,5,3,22,135],l=['< 0.50','0.50–0.70','0.70–0.80','0.80–0.90','≥ 0.90'];
 const baseY=5.7;
 n.forEach((v,i)=>{let x=.92+i*1.46,hh=v/135*2.96;rect(s,x,baseY-hh,.92,hh,i===0?'B48750':i===4?C.teal:'A7C1CA');tx(s,String(v),x-.09,baseY-hh-.42,1.1,.31,17,i===0?C.amber:C.ink,true,{align:'center',fontFace:numfont});tx(s,l[i],x-.18,5.92,1.3,.31,10.7,C.muted,false,{align:'center',fontFace:'Aptos'});});
 line(s,.71,5.71,8.26,5.71,C.line,.9);
 rect(s,8.83,2.12,3.88,4.12,C.white,C.line);
 tx(s,'32 / 197',9.09,2.48,3.35,.73,37,C.amber,true,{fontFace:numfont});
 tx(s,'病例 Dice < 0.50',9.09,3.32,3.35,.36,17,C.ink,true);
 tx(s,'高分病例与严重失败并存。\n\n低质量结果需要明确提示、专家参考与回到原始 CT 的核验路径。',9.09,4.05,3.32,1.58,16,C.muted);
 foot(s,'二分类模型不能自动识别椎骨亚结构；尚无外部 / 多中心验证，当前不报告教学显著性或临床可用性。');
 notes(s,'本次对既有 CSV 按边界 [0,0.5,0.7,0.8,0.9,1.000001] 统计，病例数为 32、5、3、22、135，总计 197。中位数约 0.9410，但不能用中位数掩盖 32 例低于 0.5 的失效。展示全体分布比只展示最佳病例更能解释当前工程边界。',[model+'full_volume_validation197/metrics_per_case.csv','docs/14_training_line_final_handoff.md']);
}
// 14. Exact v7 case, independent sources Raw/Prediction/GT.
{
 const s=base('07  应用效果 / 真实影像','同一病例，对照原始 CT、v7 预测与专家标注','示例病例：CTSpine1K COLONOG · 0663 · 矢状位 · 相同空间位置');
 ['原始 CT','v7 Prediction','专家 GT'].forEach((label,i)=>{let x=.66+i*4.12;rect(s,x,2.18,3.86,3.89,C.navy);img(s,['raw_sagittal.png','prediction_sagittal.png','gt_sagittal.png'][i],x+.13,2.31,3.6,3.43);tx(s,label,x+.1,5.77,3.65,.23,12,'D6E6EC',true,{align:'center'});});
 foot(s,'该示例 Dice 0.9640，属于表现较好的单例，用于说明结果关联；不代表全部病例平均性能。');
 notes(s,'示例全量 case ID 为 ctspine1k-colonog-1.3.6.1.4.1.9328.50.4.0663。Raw/Prediction/GT 图片从 v7 evaluation API 重新提取，避免使用旧素材中的 v7.1 图却配 v7 指标。该病例 Dice 0.9639699031、Precision 0.9702431462、Recall 0.9577772599、HD95 1.5 mm，均为单例结果。GT 来源为数据集标注，与模型输出独立展示。',[model+'full_volume_validation197/metrics_per_case.csv','../AIC比赛/99_历史备份_非提交/重制工作区_20260923/assets/evidence.json']);
}
// 15. Geometry integration and one directly verified simplification example.
{
 const s=base('03  技术创新 / 三维工程','从体素分割到毫米空间中的可操作表面','同一 v7 示例的真实 Prediction → Marching Cubes → 特征加权网格简化');
 rect(s,.67,2.13,3.74,4.21,'E5EFF1');img(s,'mesh_hero.png',1.65,2.24,1.26,3.91);
 tx(s,'物理坐标映射',4.88,2.2,6.95,.4,21,C.ink,true);
 tx(s,'p = origin + direction × diag(spacing) × index',4.88,2.82,7.18,.53,19,C.teal,true,{fontFace:'Aptos'});
 tx(s,'法向变化加权的顶点聚类，在简化时保留局部几何特征。',4.88,3.58,7.39,.52,17,C.muted);
 metric(s,'81,432','原始顶点',4.88,4.31,2.12);metric(s,'62,889','简化顶点',7.44,4.31,2.22);metric(s,'22.77%','顶点减少',10.0,4.31,2.53);
 tx(s,'1.5 mm 聚类 · 工程 vertex-HD95 0.607 mm',4.88,5.97,7.65,.32,14,C.teal,true,{fontFace:'Aptos'});
 foot(s,'网格简化误差相对原始预测表面计算，不是预测相对 GT 的分割误差；SDF 另有连通域保护机制。');
 notes(s,'本次从已保存 v7 prediction 构建网格，未重新运行模型。原始 81432 vertices/162844 faces，简化 62889/125946，vertex reduction 0.2277114648，vertex-nearest ASSD 0.17930 mm，HD95 0.60654 mm。该指标是工程近似，不能冒充临床表面距离或解剖精度。三维图真实绑定 v7 病例0663。SDF 和连通域保护属于已有代码能力，本页数据为 mask mesh，不把 SDF 描述为这张图的实际来源。',['src/reconstruction/mesh.py','src/reconstruction/sdf_surface.py','src/reconstruction/measurement.py','../AIC比赛/99_历史备份_非提交/重制工作区_20260923/assets/mesh_summary.json']);
}
// 16. Product, large real workspace crop.
{
 const s=base('07  应用效果 / 教学交互','先定位，再调用模型参照与专家核验','已实现学生实验 / 教师演示、三平面联动、分级提示与 GT 定位反馈。');
 rect(s,.65,2.12,7.69,4.13,C.white,C.line);img(s,'teaching_ui.png',.73,2.2,7.53,3.97,'crop');
 item(s,'01','在原始 CT 中定位','学生先点击判断，系统依据专家 GT 反馈是否落在椎骨前景。',8.79,2.2,3.93);
 item(s,'02','按需查看参考','两级规则提示后，可显示保存的模型 Prediction。',8.79,3.58,3.93);
 item(s,'03','回到空间中复核','连续切片与三平面确认位置，再切换到同病例三维观察。',8.79,4.96,3.93);
 foot(s,'定位命中是体素级 GT 反馈，不等于完整解剖能力评分；当前没有将模型不确定性误当作学生掌握度。');
 notes(s,'teaching.js 的 locateStart 先切回 raw；locateAi 第1级方向提示、第2级空间提示、第3级 Prediction，前两级是规则文本而非生成式大模型。teaching-hit 根据归一化 XYZ 查询真实 GT；MPR 使用相同病例与坐标更新。teacher/student 为两种交互模式，不是已实现完整权限系统。旧截图页面中的内部 case 标识用于可追溯，截图是真实界面，不制作虚假 UI。',['web/frontend/teaching.js','web/backend/app.py','web/frontend/teaching.html']);
}
// 17. Evaluation evidence versus prospective educational study.
{
 const s=base('07  应用效果 / 证据边界','算法与工程已验证，教学收益需要真实试点','把“能运行”和“能提升学习效果”分开评价。');
 rect(s,.67,2.13,5.78,3.89,C.white,C.line);tag(s,'已有证据',.92,2.38,1.15);
 tx(s,'模型有效性',.92,3.04,5.25,.37,20,C.ink,true);tx(s,'197 例完整体积分割，逐病例指标可追溯。',.92,3.57,5.17,.5,16,C.muted);
 tx(s,'工程可操作性',.92,4.39,5.25,.37,20,C.ink,true);tx(s,'真实 Prediction / GT / MPR / 3D 贯通，\n定位反馈与课堂模式可操作。',.92,4.92,5.18,.7,16,C.muted);
 rect(s,6.84,2.13,5.83,3.89,C.pale);tag(s,'下一步试点',7.09,2.38,1.48,C.teal,'D6EAEA');
 tx(s,'对照设计',7.09,3.04,5.31,.37,20,C.ink,true);tx(s,'相同病例与课时：静态课件 / 常规阅片，\n对比平台任务训练。',7.09,3.57,5.31,.72,16,C.muted);
 tx(s,'主要观察',7.09,4.56,5.31,.37,20,C.ink,true);tx(s,'无 AI 提示的新病例定位准确率、完成时间，\n以及延迟测验中的空间理解与迁移。',7.09,5.08,5.31,.72,16,C.muted);
 foot(s,'尚未取得教学试点、师生反馈或课堂效率统计；不将算法 Dice 直接解释为教学增益。');
 notes(s,'这里回应学科应用效果，但如实区分已完成和计划。教学试点需要教师审核题目、匹配内容和时长，并记录前测、后测、延迟测和无提示迁移。定位准确性需要按任务定义，不仅仅在很大的前景区域任意点中就算学会。不得使用未开展的样本数、提升比例或统计显著性。',[rubric,'web/frontend/teaching.js','web/frontend/practice.js',model+'full_volume_validation197/summary.json']);
}
// 18. Resource and deployment feasibility, no fake budget.
{
 const s=base('05  方案可行性 / 部署与复用','以病例资产复用，降低课堂运行门槛','先面向本地教学实验，再逐步扩展院校与课程使用。');
 const xx=[.67,4.81,8.95];
 [['现有资源','普通 Windows 工作站\nCPU 已完成实际评估\n本地 Web 与桌面入口','不要求课堂重新训练'],['运行条件','病例 CT / GT 缓存\n已保存模型 Prediction\n匹配的路径与版本配置','软件与数据一起迁移'],['主要投入','计算与病例存储\n教师审核与任务设计\n部署维护与数据授权','尚无正式经费预算']].forEach((r,i)=>{rect(s,xx[i],2.2,3.74,3.37,C.white,C.line);tx(s,r[0],xx[i]+.23,2.5,3.26,.47,23,C.ink,true);tx(s,r[1],xx[i]+.23,3.34,3.28,1.19,16,C.muted);tag(s,r[2],xx[i]+.23,4.98,3.19);});
 line(s,.67,5.99,12.68,5.99);tx(s,'复用路径',.72,6.13,1.68,.34,15,C.teal,true);tx(s,'教师审核病例包  →  课程任务配置  →  小规模试用  →  同类课程推广',2.39,6.13,10.28,.34,15,C.ink);
 foot(s,'上游代码与病例数据保留来源和使用许可；真实影像、权重及大体积缓存不默认进入公开 Git。');
 notes(s,'现有 CPU full-volume 评估可运行不是推理高效的结论。没有核验 GPU 设备或价格，因此不填写 GPU 加速倍数、云端费用与虚构预算。软件单 exe 不包含全部研究数据，迁移需要相关病例资产、路径配置与授权核验。公共数据与 SegFormer3D 上游许可必须分别遵守。',['PROJECT_STATUS.md','web/desktop_shell.py','SECURITY.md','third_party/README.md','docs/09_public_repository_manifest.md']);
}
// 19. Concrete roadmap and acceptance criteria.
{
 const s=base('08  总结与展望','从可运行原型，走向可验证的教学工具','下一阶段围绕教学证据与泛化验证，不继续在同一 validation197 上无止境调参。');
 const r=[['01','近期：完善病例实验','教师审核病例与任务\n记录提示使用、错误类型、操作路径','验收：一套可复用的课程病例包'],['02','中期：开展教学试点','前后测与无提示新病例迁移\n比较准确性、耗时及延迟保持','验收：真实数据与效应量报告'],['03','后续：验证算法泛化','同协议跨架构基线\n外部数据与更细标签任务','验收：独立研究协议与外部结果']];
 r.forEach((a,i)=>{let x=.67+i*4.14;tx(s,a[0],x,2.19,1.14,.65,38,C.aqua,true,{fontFace:numfont});line(s,x,3.08,x+3.76,3.08,C.teal,2);tx(s,a[1],x,3.35,3.73,.47,20,C.ink,true);tx(s,a[2],x,4.13,3.7,1.1,16,C.muted);tx(s,a[3],x,5.69,3.72,.65,14,C.teal,true);});
 foot(s,'当前边界：椎骨前景二分类 · 困难病例仍有失效 · 无外部临床验证 · 无已完成教学效果试验。');
 notes(s,'v7 是唯一冻结工程主模型，v7.1 已归档，后续研究按新的协议开展。可以考虑 nnU-Net 等跨架构基线，但本项目未报告实测对比成绩。教学方面先补齐可评价任务和学习日志，再证明对学生无提示迁移能力的帮助，而不是只增加页面功能。',['docs/13_v7_final_model_freeze.md','docs/14_training_line_final_handoff.md',rubric]);
}
// 20. Closing statement.
{
 const s=base('智骨云析','让分割结果成为可核验的学习参照','',true);
 tx(s,'从二维切片出发，\n在真实三维空间中完成理解。',.7,2.1,8.43,1.45,34,C.white,true);
 const labs=[['算法','高分辨率分割'],['几何','物理坐标一致'],['教学','观察、核验、复盘']];
 labs.forEach((a,i)=>{let x=.74+i*2.76;tx(s,a[0],x,4.56,2.49,.44,23,C.aqua,true);tx(s,a[1],x,5.18,2.54,.39,15,'B8D0DD');});
 img(s,'mesh_hero.png',10.12,1.74,1.46,4.72);
 tx(s,'谢谢',.75,6.3,7,.37,19,'D8E9F0',true);
 notes(s,'结束时回到学科价值：不是用模型替学生完成判断，而是通过真实 Prediction、专家标注和物理空间工具，使学习过程能够反复观察和核验。当前成绩是开发期算法与工程证据，真实教学成效是下一步验证重点。',['docs/13_v7_final_model_freeze.md','web/frontend/teaching.js']);
}
// 21. Backup: negative experiment, not promoted to deployed capability.
{
 const s=base('补充材料 A / 版本选择','v7.1 已完成验证，但不替代冻结的 v7','Residual false-positive guidance：610 / 610 训练病例；validation / test 不参与引导生成。');
 table(s,['指标','v7 主协议','v7.1 主协议'],[['Dice','0.7933','0.7027'],['Precision','0.7862','0.6122'],['Recall','0.8245','0.8797'],['HD95（mm）','37.95','132.53'],['ASSD（mm）','8.81','24.19'],['预测 / GT 前景比','1.6849×','2.8933×']],.67,2.08,[4.3,3.87,3.87],.5);
 tx(s,'决策：保留 v7',.77,5.93,3.31,.43,23,C.teal,true);tx(s,'召回提高，但重叠、误检和表面距离综合退化；新版本并不自动成为主模型。',4.1,5.92,8.5,.5,15,C.muted);
 foot(s,'v7.1 于 epoch 5 early stop，best epoch 2；全部 197 例全体积验证完成后作为负结果归档。');
 notes(s,'v7.1 的 full-volume 结果较 v7 主协议退化；固定32768阈值的 secondary comparison 也未超过 v7，因此不重新开启阈值搜索。完成残余FP采样引导是实际工作量，但不等于部署性能收益。主模型、主结果、展示素材统一为 v7。',['docs/13_v7_final_model_freeze.md','docs/14_training_line_final_handoff.md']);
}
// 22. Backup: reproducibility, task and sources.
{
 const s=base('补充材料 B / 复现信息','锁定任务、配置与结果口径','模型配置、数据集合和结果产物一一对应，避免使用不同阶段的指标或截图。');
 table(s,['项目','冻结主线'],[['任务','椎骨前景 binary semantic；2 类'],['数据','CTSpine1K train610 / validation197；1.5 mm 缓存'],['训练','64³ patch；batch 1；AdamW；GroupNorm 解码'],['推理','96³ ROI；overlap 0.5；8,192 体素连通域过滤'],['结果','full_volume_validation197 / summary.json'],['权重','20260917_185941... / checkpoint / best.pt']],.67,2.08,[2.5,9.54],.52);
 tx(s,'0.9402 ≠ 0.7933',.77,5.94,4.2,.42,22,C.amber,true,{fontFace:'Aptos'});tx(s,'前者是局部 patch 选择器；后者才是完整 CT 体积的验证均值。',5.0,5.94,7.59,.48,16,C.muted);
 foot(s,'v7 calibration 未开启，主验证未保存 uncertainty 产物；旧 pilot 的不确定性结果不能直接移植为 v7 结论。');
 notes(s,'v7 canonical path: experiments/20260917_185941_ctspine1k_v7_segformer3d_hr_hard_negative_precision。配置 inference.batchnorm_mode=running 为通用执行模式参数；HR decoder 使用 GroupNorm，不能把BN开关写成网络全部使用BN。当前 v7 uncertainty_available=false，源CSV校准字段为空，故主讲不展示虚构 v7 uncertainty热图或校准结果。',['configs/orthopedic_ct_full_large_scale_v7_segformer3d_hr_hard_negative_precision.yaml',model+'summary.json',model+'full_volume_validation197/summary.json',model+'full_volume_validation197/metrics_per_case.csv']);
}
// 23. Source registry, legible and concise; full URLs in notes.
{
 const s=base('补充材料 C / 来源与引用','研究依据、开源来源与项目证据','每页备注附来源；演示图像、模型指标和工程网格分别记录其对应产物。');
 const refs=[['01','SegFormer3D','Perera, Navard & Yilmaz · CVPR Workshops 2024\n采用开源层次化编码器；本项目扩展高分辨率解码与空间分支。'],['02','CTSpine1K','Deng 等 · arXiv:2105.14711\n公开脊柱 CT 数据集；本项目使用固定 610 / 197 开发划分。'],['03','空间能力与解剖教学','Koh, Tan & Mogali · Scientific Reports 2023\n用于教学设计背景，不作为本平台教学有效性的证据。'],['04','项目内可追溯材料','v7 冻结文档 · 完整体积 CSV / JSON · 源码与配置\n本次提取 v7 病例素材；旧版和负向实验保留独立归档。']];
 refs.forEach((r,i)=>{let y=2.11+i*1.02;tx(s,r[0],.7,y,.6,.37,19,C.teal,true,{fontFace:numfont});tx(s,r[1],1.49,y,3.11,.4,18,C.ink,true);tx(s,r[2],4.79,y-.04,7.79,.67,13.8,C.muted);if(i<3)line(s,.69,y+.83,12.66,y+.83);});
 foot(s,'视觉参考：极简学术演示与医疗研究模板；本稿为原创排版，不套用第三方院校标识或付费模板素材。');
 notes(s,'外部文献和模板只用于来源说明与视觉参考，不复制他人结果或套用开源backbone为原创。视觉参考为 pmichaillat/latex-presentation 的留白与图表优先原则，SlidesCarnival Minimal Medical 的克制医疗配色；未使用其模板文件或图库。官方评分规则用于组织内容，评分映射仅保留在内部交付说明，不放入主讲页。',[upstream,dataPaper,eduPaper,rubric,'https://github.com/OSUPCVLab/SegFormer3D','https://github.com/pmichaillat/latex-presentation','https://www.slidescarnival.com/template/cerimon-free-presentation-template/1100','docs/13_v7_final_model_freeze.md']);
}
// Check all bounds. Containment (text inside cards, images in frames) is intentional.
if(H){
 for(const s of pptx._slides){
   H.warnIfSlideElementsOutOfBounds(s,pptx);
   const objects=s._slideObjects;
   // Text-only overlap pass avoids intentionally superimposed card fills and diagrams.
   s._slideObjects=objects.filter(o=>o.shape===undefined && o.text!==undefined);
   try{H.warnIfSlideHasOverlaps(s,pptx);}finally{s._slideObjects=objects;}
 }
}
fs.mkdirSync(path.join(root,'output'),{recursive:true});
fs.writeFileSync(path.join(root,'slide_titles.json'),JSON.stringify(allTitles,null,2));
pptx.writeFile({fileName:path.join(root,'output','ZhiGuYunXi_AIC_Redesign.pptx')});
