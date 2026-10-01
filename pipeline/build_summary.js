const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
        ShadingType, VerticalAlign } = require('docx');

const D = JSON.parse(fs.readFileSync(__dirname + '/summary_data.json'));
const OUT = process.argv[2] || ('Budget Utilization Summary ' + D.asof + '.docx');
const dollars = n => '$' + (n < 0 ? '(' + Math.abs(Math.round(n)).toLocaleString() + ')'
                                  : Math.round(n).toLocaleString());

const NAVY="1F3864", HDRTXT="FFFFFF", ALT="F2F5FA", TOTAL="D9E1F2";
const bd = { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
const COLS = [3060, 2100, 2100, 2100];

const cell = (text,{bold=false,fill=null,align=AlignmentType.LEFT,color="000000"}={},w) =>
  new TableCell({ borders, width:{size:w,type:WidthType.DXA},
    margins:{top:60,bottom:60,left:110,right:110},
    shading: fill?{fill,type:ShadingType.CLEAR}:undefined, verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({alignment:align,children:[new TextRun({text,bold,color,size:18,font:"Arial"})]})] });

function pmoTable(rows, totOv, totBud, totWO, budHdr) {
  const head = new TableRow({ tableHeader:true, children:[
    cell("PMO",{bold:true,color:HDRTXT,fill:NAVY},COLS[0]),
    cell("Overrun $",{bold:true,color:HDRTXT,fill:NAVY,align:AlignmentType.RIGHT},COLS[1]),
    cell(budHdr,{bold:true,color:HDRTXT,fill:NAVY,align:AlignmentType.RIGHT},COLS[2]),
    cell("Weighted Overrun %",{bold:true,color:HDRTXT,fill:NAVY,align:AlignmentType.RIGHT},COLS[3]) ]});
  const body = rows.map((r,i)=>{ const f=i%2?ALT:null; return new TableRow({children:[
    cell(r[0],{fill:f},COLS[0]),
    cell(dollars(r[1]),{fill:f,align:AlignmentType.RIGHT},COLS[1]),
    cell(dollars(r[2]),{fill:f,align:AlignmentType.RIGHT},COLS[2]),
    cell(r[3].toFixed(2)+"%",{fill:f,align:AlignmentType.RIGHT},COLS[3]) ]}); });
  const tot = new TableRow({children:[
    cell("Company-wide",{bold:true,fill:TOTAL},COLS[0]),
    cell(dollars(totOv),{bold:true,fill:TOTAL,align:AlignmentType.RIGHT},COLS[1]),
    cell(dollars(totBud),{bold:true,fill:TOTAL,align:AlignmentType.RIGHT},COLS[2]),
    cell(totWO.toFixed(2)+"%",{bold:true,fill:TOTAL,align:AlignmentType.RIGHT},COLS[3]) ]});
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:COLS, rows:[head,...body,tot] });
}

const T = (text,o={}) => new TextRun({text,font:"Arial",size:22,...o});
const P = (runs,opts={}) => new Paragraph({...opts,children:runs});
const bullet = (runs,level=0) => new Paragraph({numbering:{reference:"b",level},children:runs});
const SMALL = t => new TextRun({text:t,font:"Arial",size:18,color:"595959"});
const SECT = t => new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:240,after:120},
  children:[new TextRun({text:t,bold:true,font:"Arial",size:26,color:NAVY})]});

const children = [
  P([T("Project Budget Utilization Summary",{bold:true,size:30})]),
  P([SMALL("As of " + D.asof)], {spacing:{after:200}}),
  P([T("Hi All,")], {spacing:{after:120}}),
  P([T("Please see the attached reports. The overrun-by-PMO tables are included below each section. For reference, the weighted overrun % calculations are as follows:")], {spacing:{after:120}}),
  bullet([T("Original Budget Summary: ",{bold:true}), T("Weighted Overrun % = ((Total Costs − Original Budget) / Sum of All Original Budgets) × 100")]),
  bullet([T("Revised ETC Budget Summary: ",{bold:true}), T("Weighted Overrun % = ((Total Costs − Total Revised Budget, i.e., Original Budget + Budget Adjustments/ETC) / Sum of All Revised Budgets) × 100")]),

  SECT("Original Cost Budget Summary"),
  bullet([T(`We have ${D.orig_n.toLocaleString()} active projects with an original budget, with ${D.orig_overrun_pct}% overrun (target 5% or less).`)]),
  bullet([T(`Weighted Overrun % = ${D.orig_weighted.toFixed(2)}%${D.orig_weighted_delta}`)]),
  P([SMALL("Overrun by PMO (original budget) — projects where costs exceed the original budget:")], {spacing:{before:120,after:80}}),
  pmoTable(D.orig_tbl, D.orig_T, D.orig_D, D.orig_WO, "Original Budget"),

  SECT("Revised (ETC) Budget Summary"),
  bullet([T(`We have ${D.total.toLocaleString()} active projects, with ${D.rev_overrun_pct}% overrun (target 5% or less).`)]),
  bullet([T(`Weighted Overrun % = ${D.rev_weighted.toFixed(2)}%${D.rev_weighted_delta}`)]),
  bullet([T(`${D.gt70} projects with over 70% of budget spent (additional filters summarized in the report)${D.gt70_delta}.`)]),
  bullet([T(`${D.gt90_50k} projects with over 90% of budget spent (additional filters summarized in the report) and contract values greater than $50K${D.gt90_delta}.`)]),
  bullet([T(`${D.ff100} active fixed-fee projects with 100% revenue recognition${D.ff100_delta}.`)]),
  bullet([T("Of the late-stage projects:")]),
  bullet([T(`${D.late.p100}% of projects with 100%+ revenue recognition have ETCs or are on-budget (target 100%).`)],1),
  bullet([T(`${D.late.p90}% of projects with 90% revenue recognition have ETCs or are on-budget (target 100%).`)],1),
  bullet([T(`${D.late.p70}% of projects with 70% revenue recognition have ETCs or are on-budget (target 100%).`)],1),
  P([SMALL("Overrun by PMO (revised/ETC budget) — projects where costs exceed the total revised budget:")], {spacing:{before:120,after:80}}),
  pmoTable(D.rev_tbl, D.rev_T, D.rev_D, D.rev_WO, "Revised Budget"),
];

const doc = new Document({
  styles:{ default:{document:{run:{font:"Arial",size:22}}},
    paragraphStyles:[{id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:26,bold:true,font:"Arial"},paragraph:{spacing:{before:240,after:120},outlineLevel:1}}]},
  numbering:{config:[{reference:"b",levels:[
    {level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:460,hanging:260}}}},
    {level:1,format:LevelFormat.BULLET,text:"◦",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:980,hanging:260}}}}]}]},
  sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},children}],
});
Packer.toBuffer(doc).then(b=>{ fs.writeFileSync(OUT,b); console.log("doc written:",OUT); });
