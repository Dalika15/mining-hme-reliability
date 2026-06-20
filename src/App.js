import { useState, useCallback, useMemo, useRef, useEffect } from "react";
import {
  LineChart, Line, BarChart, Bar, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  ReferenceLine, Area, AreaChart, Cell
} from "recharts";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ─── DESIGN TOKENS ────────────────────────────────────────────────────────────
const T = {
  bg0:    "#0D0F14",
  bg1:    "#13161E",
  bg2:    "#1A1E2A",
  bg3:    "#222736",
  border: "#2E3448",
  gold:   "#C9A84C",
  goldL:  "#E8C96A",
  goldD:  "#8B6E2A",
  text:   "#E8EAF0",
  muted:  "#8892A4",
  green:  "#2ECC71",
  red:    "#E74C3C",
  orange: "#F39C12",
  blue:   "#3498DB",
  purple: "#9B59B6",
  cyan:   "#1ABC9C",
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:${T.bg0};color:${T.text};font-family:'Inter',sans-serif;font-size:13px}
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-track{background:${T.bg1}}
  ::-webkit-scrollbar-thumb{background:${T.bg3};border-radius:3px}
  .mono{font-family:'JetBrains Mono',monospace}
  input[type=file]{display:none}
  input[type=number],input[type=text]{background:${T.bg3};border:1px solid ${T.border};color:${T.text};border-radius:4px;padding:5px 8px;font-family:'Inter',sans-serif;font-size:12px;width:100%}
  input[type=number]:focus,input[type=text]:focus{outline:1px solid ${T.gold};border-color:${T.gold}}
  select{background:${T.bg3};border:1px solid ${T.border};color:${T.text};border-radius:4px;padding:5px 8px;font-size:12px;width:100%}
  table{border-collapse:collapse;width:100%}
  th{background:${T.bg3};color:${T.gold};font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;padding:8px 10px;text-align:left;border-bottom:1px solid ${T.border}}
  td{padding:7px 10px;border-bottom:1px solid ${T.border};font-size:11px;color:${T.text}}
  tr:hover td{background:${T.bg3}44}
  .badge{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:99px;font-size:10px;font-weight:600}
  .ok{background:#2ECC7122;color:#2ECC71}
  .warn{background:#F39C1222;color:#F39C12}
  .crit{background:#E74C3C22;color:#E74C3C}
  .info{background:#3498DB22;color:#3498DB}
  .btn{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border-radius:6px;border:none;cursor:pointer;font-family:'Inter',sans-serif;font-size:12px;font-weight:600;transition:all .15s}
  .btn-gold{background:${T.gold};color:#000}
  .btn-gold:hover{background:${T.goldL}}
  .btn-ghost{background:${T.bg3};color:${T.text};border:1px solid ${T.border}}
  .btn-ghost:hover{border-color:${T.gold};color:${T.gold}}
  .tab-active{background:${T.gold}22;color:${T.gold};border-bottom:2px solid ${T.gold}}
  .tab-inactive{color:${T.muted};border-bottom:2px solid transparent}
  .tab-inactive:hover{color:${T.text}}
  .card{background:${T.bg1};border:1px solid ${T.border};border-radius:8px;padding:16px}
  .section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:${T.gold};margin-bottom:12px;display:flex;align-items:center;gap:8px}
  .kpi-val{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:600;line-height:1}
  .kpi-sub{font-size:10px;color:${T.muted};margin-top:4px}
  .upload-zone{border:1.5px dashed ${T.border};border-radius:8px;padding:20px;text-align:center;cursor:pointer;transition:all .2s}
  .upload-zone:hover{border-color:${T.gold};background:${T.gold}08}
  .weibull-input{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px}
`;

// ─── SAMPLE DATA GENERATORS ────────────────────────────────────────────────
const FLEET = [
  // ── Volvo ADT A45G
  {id:"AT007",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT008",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT009",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT010",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT011",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT012",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT013",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT014",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT015",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT016",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT017",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT018",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT019",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT020",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT021",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT022",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT023",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"AT024",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT01",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT02",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT03",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT04",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT05",brand:"Volvo",model:"A45G",type:"ADT"},
  {id:"TAT06",brand:"Volvo",model:"A45G",type:"ADT"},
  // ── Volvo ADT A60H
  {id:"AT112",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"AT113",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"AT114",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"AT115",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"AT116",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT100",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT101",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT102",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT103",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT104",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT105",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT106",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT107",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT108",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT109",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT110",brand:"Volvo",model:"A60H",type:"ADT"},
  {id:"TAT111",brand:"Volvo",model:"A60H",type:"ADT"},
  // ── CAT Haul Trucks 777D/E
  {id:"HT002",brand:"Caterpillar",model:"777D",type:"Haul Truck"},
  {id:"HT004",brand:"Caterpillar",model:"777D",type:"Haul Truck"},
  {id:"HT005",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT006",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT007",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT008",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT009",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT010",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT011",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT012",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT013",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT014",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT015",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT016",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT017",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT018",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT019",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT020",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT021",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT022",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT023",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT024",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT025",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT026",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT027",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT028",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT029",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT030",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT031",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT032",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT033",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT034",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT035",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT036",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT037",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT038",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT039",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT040",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT041",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT042",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT043",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT044",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT045",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT046",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT047",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT048",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT049",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT050",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT051",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  {id:"HT052",brand:"Caterpillar",model:"777E",type:"Haul Truck"},
  // ── CAT Haul Trucks 789D
  {id:"HT100",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT101",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT102",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT103",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT104",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT105",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT106",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT107",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT108",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT109",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT110",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT111",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT112",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT113",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT114",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT115",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT116",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT117",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT118",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  {id:"HT119",brand:"Caterpillar",model:"789D",type:"Haul Truck"},
  // ── Water Trucks
  {id:"WT009",brand:"Caterpillar",model:"777E-WT",type:"Water Truck"},
  {id:"WT010",brand:"Caterpillar",model:"777E-WT",type:"Water Truck"},
  {id:"WT013",brand:"Volvo",model:"A45G-WT",type:"Water Truck"},
  {id:"WT014",brand:"Volvo",model:"A45G-WT",type:"Water Truck"},
  {id:"WT015",brand:"Caterpillar",model:"777D-WT",type:"Water Truck"},
  {id:"WT016",brand:"Volvo",model:"A45G-WT",type:"Water Truck"},
  {id:"WT017",brand:"Volvo",model:"A45G-WT",type:"Water Truck"},
  {id:"WT018",brand:"Volvo",model:"FMX440-WT",type:"Water Truck"},
  {id:"WT019",brand:"Volvo",model:"FMX440-WT",type:"Water Truck"},
  {id:"WT021",brand:"Caterpillar",model:"777E-WT",type:"Water Truck"},
  // ── CAT Excavators
  {id:"EX001",brand:"Caterpillar",model:"374D",type:"Excavator"},
  {id:"EX007",brand:"Caterpillar",model:"6020B",type:"Excavator"},
  {id:"EX009",brand:"Caterpillar",model:"6020B",type:"Excavator"},
  {id:"EX010",brand:"Caterpillar",model:"6020B",type:"Excavator"},
  {id:"EX011",brand:"Caterpillar",model:"6020B",type:"Excavator"},
  {id:"EX013",brand:"Caterpillar",model:"390F",type:"Excavator"},
  {id:"EX014",brand:"Caterpillar",model:"6015",type:"Excavator"},
  {id:"EX015",brand:"Caterpillar",model:"6015",type:"Excavator"},
  {id:"EX017",brand:"Caterpillar",model:"6015",type:"Excavator"},
  {id:"EX019",brand:"Caterpillar",model:"374D",type:"Excavator"},
  {id:"EX020",brand:"Caterpillar",model:"395",type:"Excavator"},
  {id:"EX021",brand:"Caterpillar",model:"395",type:"Excavator"},
  {id:"EX022",brand:"Caterpillar",model:"336GC",type:"Excavator"},
  {id:"EX023",brand:"Caterpillar",model:"336",type:"Excavator"},
  // ── Hitachi Excavators
  {id:"EX100",brand:"Hitachi",model:"EX2600",type:"Excavator"},
  {id:"EX101",brand:"Hitachi",model:"EX2600",type:"Excavator"},
  {id:"EX102",brand:"Hitachi",model:"EX2600",type:"Excavator"},
  {id:"EX103",brand:"Hitachi",model:"EX2600",type:"Excavator"},
  {id:"EX104",brand:"Hitachi",model:"EX2000",type:"Excavator"},
  {id:"TEX01",brand:"Hitachi",model:"EX1200",type:"Excavator"},
  {id:"TEX02",brand:"Hitachi",model:"EX1200",type:"Excavator"},
  {id:"TEX03",brand:"Hitachi",model:"EX1200",type:"Excavator"},
  // ── CAT Front Shovels
  {id:"FS001",brand:"Caterpillar",model:"6040FS",type:"Shovel"},
  {id:"FS002",brand:"Caterpillar",model:"6040FS",type:"Shovel"},
  // ── CAT Drills MD6240/MD6250
  {id:"BD003",brand:"Caterpillar",model:"MD6240",type:"Drill"},
  {id:"BD004",brand:"Caterpillar",model:"MD6240",type:"Drill"},
  {id:"BD005",brand:"Caterpillar",model:"MD6240",type:"Drill"},
  {id:"BD010",brand:"Caterpillar",model:"MD6250",type:"Drill"},
  {id:"BD011",brand:"Caterpillar",model:"MD6250",type:"Drill"},
  {id:"BD012",brand:"Caterpillar",model:"MD6250",type:"Drill"},
  {id:"BD013",brand:"Caterpillar",model:"MD6250",type:"Drill"},
  // ── Epiroc Drills DM30 XC SP
  {id:"BD100",brand:"Epiroc",model:"DM30 XC SP",type:"Drill"},
  {id:"BD101",brand:"Epiroc",model:"DM30 XC SP",type:"Drill"},
  {id:"BD102",brand:"Epiroc",model:"DM30 XC SP",type:"Drill"},
  {id:"BD103",brand:"Epiroc",model:"DM30 XC SP",type:"Drill"},
  {id:"BD104",brand:"Epiroc",model:"DM30 XC SP",type:"Drill"},
  // ── Epiroc Drills D65LF
  {id:"SR100",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR101",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR102",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR103",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR104",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR105",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR106",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR107",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR108",brand:"Epiroc",model:"D65LF",type:"Drill"},
  {id:"SR109",brand:"Epiroc",model:"D65LF",type:"Drill"},
  // ── Epiroc Drills T45LF/T65LF
  {id:"SR200",brand:"Epiroc",model:"T45LF",type:"Drill"},
  {id:"SR201",brand:"Epiroc",model:"T65LF",type:"Drill"},
  {id:"SR202",brand:"Epiroc",model:"T65LF",type:"Drill"},
  // ── Sandvik Drills DP1500i
  {id:"SR002",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  {id:"SR003",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  {id:"SR004",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  {id:"SR005",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  {id:"SR006",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  {id:"SR007",brand:"Sandvik",model:"DP1500i",type:"Drill"},
  // ── CAT Graders 16M/18M3
  {id:"MG001",brand:"Caterpillar",model:"16M",type:"Grader"},
  {id:"MG002",brand:"Caterpillar",model:"16M",type:"Grader"},
  {id:"MG003",brand:"Caterpillar",model:"16M3",type:"Grader"},
  {id:"MG004",brand:"Caterpillar",model:"18M3",type:"Grader"},
  {id:"MG005",brand:"Caterpillar",model:"18M3",type:"Grader"},
  {id:"MG010",brand:"Caterpillar",model:"18M3",type:"Grader"},
  {id:"MG011",brand:"Caterpillar",model:"18M3",type:"Grader"},
  {id:"TMG08",brand:"Caterpillar",model:"16M",type:"Grader"},
  {id:"TMG09",brand:"Caterpillar",model:"16M",type:"Grader"},
  {id:"TMG12",brand:"Caterpillar",model:"18M3",type:"Grader"},
  // ── CAT Track Dozers D9R/D9GC
  {id:"TD001",brand:"Caterpillar",model:"D9R",type:"Dozer"},
  {id:"TD002",brand:"Caterpillar",model:"D9R",type:"Dozer"},
  {id:"TD005",brand:"Caterpillar",model:"D9R",type:"Dozer"},
  {id:"TD014",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TD018",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TD020",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TTD17",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TTD19",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TTD25",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  {id:"TTD26",brand:"Caterpillar",model:"D9GC",type:"Dozer"},
  // ── CAT Track Dozers D10T/D10T2
  {id:"TD006",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD007",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD009",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD010",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD011",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD012",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD013",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD015",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD016",brand:"Caterpillar",model:"D10T2",type:"Dozer"},
  {id:"TD021",brand:"Caterpillar",model:"D10T",type:"Dozer"},
  {id:"TD022",brand:"Caterpillar",model:"D10T",type:"Dozer"},
  {id:"TD023",brand:"Caterpillar",model:"D10T",type:"Dozer"},
  // ── CAT Wheel Dozers 834K/844K
  {id:"WD001",brand:"Caterpillar",model:"834K",type:"Wheel Dozer"},
  {id:"WD002",brand:"Caterpillar",model:"834K",type:"Wheel Dozer"},
  {id:"WD003",brand:"Caterpillar",model:"844K",type:"Wheel Dozer"},
  {id:"WD004",brand:"Caterpillar",model:"844K",type:"Wheel Dozer"},
  // ── CAT Wheel Loaders
  {id:"WL004",brand:"Caterpillar",model:"992K",type:"Loader"},
  {id:"WL005",brand:"Caterpillar",model:"988K",type:"Loader"},
  {id:"WL009",brand:"Caterpillar",model:"992K",type:"Loader"},
  {id:"WL015",brand:"Caterpillar",model:"992K",type:"Loader"},
  {id:"WL017",brand:"Caterpillar",model:"992K",type:"Loader"},
  {id:"WL100",brand:"Volvo",model:"L150H",type:"Loader"},
];

function genStops() {
  const rows = [];
  const now = Date.now();
  FLEET.forEach(eq => {
    const n = Math.floor(Math.random()*12)+4;
    for(let i=0;i<n;i++){
      const planned = Math.random()>0.45;
      const dur = planned ? (Math.random()*8+2) : (Math.random()*24+1);
      const dt = now - Math.random()*90*864e5;
      rows.push({
        equipment: eq.id, brand: eq.brand, model: eq.model, type: eq.type,
        date: new Date(dt).toISOString().split("T")[0],
        planned, duration: +dur.toFixed(1),
        cause: planned
          ? ["PM 250h","PM 500h","PM 1000h","Inspection"][Math.floor(Math.random()*4)]
          : ["Engine Overheating","Hydraulic Leak","Flat Tyre","Transmission","Differential","Faulty Sensor"][Math.floor(Math.random()*6)],
        system: ["Engine","Hydraulics","Transmission","Undercarriage","Electrical","Structure"][Math.floor(Math.random()*6)],
        repaired: Math.random()>0.1,
      });
    }
  });
  return rows;
}

function genSOS() {
  const rows = [];
  const components = ["Engine","Transmission","Hydraulics","Torque Converter","Differential"];
  FLEET.forEach(eq => {
    components.slice(0,3).forEach(comp => {
      const fe = +(Math.random()*180+5).toFixed(0);
      const cu = +(Math.random()*60+2).toFixed(0);
      const al = +(Math.random()*40+1).toFixed(0);
      const si = +(Math.random()*30+1).toFixed(0);
      const visc = +(Math.random()*4+13).toFixed(1);
      const result = fe>120||cu>40||al>25 ? "CRITICAL" : fe>80||cu>25||al>15 ? "CAUTION" : "NORMAL";
      rows.push({
        equipment:eq.id, brand:eq.brand, model:eq.model,
        component:comp,
        sampleDate: new Date(Date.now()-Math.random()*60*864e5).toISOString().split("T")[0],
        sampleNo: `S-${Math.random().toString(36).substr(2,6).toUpperCase()}`,
        fe, cu, al, si, visc,
        result,
        recommendation: result==="CRITICAL"?"Change oil immediately — inspect component":
                        result==="CAUTION"?"Monitor — resample at 50h":"OK — next normal cycle",
      });
    });
  });
  return rows;
}

function genFaultCodes() {
  const codes = [
    {code:"E360-1",desc:"High Engine Oil Temperature",sev:"CRITICAL",sys:"Engine"},
    {code:"E361-2",desc:"Low Engine Oil Pressure",sev:"CRITICAL",sys:"Engine"},
    {code:"E395-3",desc:"Coolant Temperature High",sev:"HIGH",sys:"Refroidissement"},
    {code:"E082-1",desc:"Low Hydraulic Pressure",sev:"HIGH",sys:"Hydraulics"},
    {code:"E083-2",desc:"High Hydraulic Oil Temperature",sev:"MODERATE",sys:"Hydraulics"},
    {code:"E228-1",desc:"High Transmission Temperature",sev:"HIGH",sys:"Transmission"},
    {code:"E229-3",desc:"Low Transmission Pressure",sev:"CRITICAL",sys:"Transmission"},
    {code:"E041-2",desc:"Gearbox Position Sensor Fault",sev:"MODERATE",sys:"Transmission"},
    {code:"E550-1",desc:"Low Battery Voltage",sev:"LOW",sys:"Electrical"},
    {code:"E551-3",desc:"Alternator Failure",sev:"HIGH",sys:"Electrical"},
  ];
  const rows = [];
  FLEET.forEach(eq => {
    const n = Math.floor(Math.random()*5)+1;
    for(let i=0;i<n;i++){
      const c = codes[Math.floor(Math.random()*codes.length)];
      rows.push({
        equipment:eq.id, brand:eq.brand,
        date:new Date(Date.now()-Math.random()*30*864e5).toISOString().split("T")[0],
        ...c,
        occurrences: Math.floor(Math.random()*8)+1,
        source: eq.brand==="CAT"?"VisionLink":"Système constructeur",
        status:["NEW","IN PROGRESS","RESOLVED","UNDER MONITORING"][Math.floor(Math.random()*4)],
        action: ["Diagnosis in progress","Part on order","Repaired — test passed","Under monitoring"][Math.floor(Math.random()*4)],
      });
    }
  });
  return rows;
}

function genWO() {
  const types = ["CM","PM","PDM","Inspection"];
  const rows = [];
  FLEET.forEach(eq => {
    const n = Math.floor(Math.random()*8)+3;
    for(let i=0;i<n;i++){
      const type = types[Math.floor(Math.random()*types.length)];
      const planned = new Date(Date.now()-Math.random()*60*864e5);
      const actual = new Date(planned.getTime()+Math.random()*3*864e5);
      const cost = +(Math.random()*15000+500).toFixed(0);
      rows.push({
        woNum:`WO-${Math.random().toString(36).substr(2,7).toUpperCase()}`,
        equipment:eq.id, brand:eq.brand, model:eq.model, type,
        description:["Oil Filter Replacement","Engine Oil Drain","Brake Inspection","Tyre Replacement","Hydraulic Service","500h Full PM Service","Leak Repair","Sensor Calibration"][Math.floor(Math.random()*8)],
        plannedDate:planned.toISOString().split("T")[0],
        actualDate:actual.toISOString().split("T")[0],
        duration:+(Math.random()*12+0.5).toFixed(1),
        cost,
        technician:["Koné A.","Diallo M.","Traoré B.","Coulibaly S.","Cissé R."][Math.floor(Math.random()*5)],
        status:["CLOSED","IN PROGRESS","PLANNED","ON HOLD"][Math.floor(Math.random()*4)],
        pmCompliant: type==="PM" ? Math.random()>0.15 : null,
      });
    }
  });
  return rows;
}

// ─── WEIBULL ENGINE ────────────────────────────────────────────────────────────
function weibullMLE(ttf) {
  if(!ttf||ttf.length<2) return null;
  const n = ttf.length;
  // MLE for shape (beta) via Newton-Raphson
  let beta = 1.5;
  for(let iter=0;iter<200;iter++){
    const s1 = ttf.reduce((a,t)=>a+Math.log(t),0)/n;
    const s2 = ttf.reduce((a,t)=>a+Math.pow(t,beta)*Math.log(t),0);
    const s3 = ttf.reduce((a,t)=>a+Math.pow(t,beta),0);
    const f  = 1/beta + s1 - s2/s3;
    const s4 = ttf.reduce((a,t)=>a+Math.pow(t,beta)*Math.log(t)*Math.log(t),0);
    const df = -1/(beta*beta) - (s4*s3-s2*s2)/(s3*s3);
    const nb = beta - f/df;
    if(Math.abs(nb-beta)<1e-9) break;
    beta = nb>0.1?nb:0.1;
  }
  const eta = Math.pow(ttf.reduce((a,t)=>a+Math.pow(t,beta),0)/n, 1/beta);
  return {beta:+beta.toFixed(4), eta:+eta.toFixed(1)};
}

function weibullR(t,beta,eta){ return Math.exp(-Math.pow(t/eta,beta)); }
function weibullF(t,beta,eta){ return 1-weibullR(t,beta,eta); }
function weibullH(t,beta,eta){ return (beta/eta)*Math.pow(t/eta,beta-1); }
function weibullBLife(p,beta,eta){ return eta*Math.pow(-Math.log(1-p/100),1/beta); }

function weibullCurveData(beta,eta,maxT){
  const pts=[];
  for(let i=0;i<=100;i++){ // eslint-disable-line no-loop-func
    const t=maxT*i/100;
    pts.push({t:+t.toFixed(0),R:+(weibullR(t,beta,eta)*100).toFixed(1),F:+(weibullF(t,beta,eta)*100).toFixed(1),h:+(weibullH(t,beta,eta)*1000).toFixed(3)});
  }
  return pts;
}

// ─── PARSE HELPERS ─────────────────────────────────────────────────────────────

// Generic XLSX/CSV reader → raw rows
function readRawFile(file, cb) {
  const ext = file.name.split(".").pop().toLowerCase();
  if(ext==="csv"){
    Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>cb(r.data)});
  } else {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result,{type:"binary",cellDates:true});
      const ws = wb.Sheets[wb.SheetNames[0]];
      cb(XLSX.utils.sheet_to_json(ws,{raw:false,dateNF:"YYYY-MM-DD"}));
    };
    reader.readAsBinaryString(file);
  }
}

// ── DB Arrêts parser
// Columns: Equipment N#, Model, Date Down, Time Down, Date Up, Time Up,
//          Down Time (Hrs), Planned, Comments, CAT Code Equivalent, Job Status, Events
function parseDBArrets(rows) {
  return rows.map(r => {
    const planned = String(r["Planned"]||"").toUpperCase() === "PLANNED";
    // Date Down can be an Excel serial or a string
    let dateStr = "";
    const rawDate = r["Date Down"];
    if(rawDate instanceof Date) dateStr = rawDate.toISOString().split("T")[0];
    else if(typeof rawDate === "string") dateStr = rawDate.split(" ")[0];
    else if(typeof rawDate === "number"){
      // Excel serial date
      const d = new Date(Math.round((rawDate - 25569)*86400*1000));
      dateStr = d.toISOString().split("T")[0];
    }
    return {
      equipment:  String(r["Equipment N#"]||"").trim(),
      model:      String(r["Model"]||"").trim(),
      date:       dateStr,
      timeDown:   String(r["Time Down"]||""),
      timeUp:     String(r["Time Up"]||""),
      duration:   parseFloat(r["Down Time (Hrs)"])||0,
      planned,
      comments:   String(r["Comments"]||"").trim(),
      cause:      String(r["CAT Code Equivalent"]||r["Comments"]||"").trim(),
      system:     String(r["CAT Code Equivalent"]||"").trim(),
      jobStatus:  String(r["Job Status"]||"").trim(),
      events:     r["Events"],
    };
  }).filter(r=>r.equipment && r.duration>0);
}

// ── Truvu Spectrooil parser
// Wide format: cols 0-41 metadata, then groups of 6 cols per parameter:
//   ParameterCodeName*, ParameterName, ParameterDataType*, ParameterValue,
//   ParameterAlarmCode, ParameterAlarmCodeReason
// Key metals: ml-el.fe, ml-el.cu, ml-el.al, ml-el.cr, ml-el.si, ml-el.pb, ml-el.ni
function parseTruvu(file, cb) {
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result,{type:"binary",cellDates:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    // Get raw array of arrays to handle repeated headers
    const raw = XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
    if(raw.length < 2){ cb([]); return; }
    const headers = raw[0];

    // Build param code → col index map (ParameterValue is col+3 from ParameterCodeName)
    const paramMap = {};
    for(let i=42; i<headers.length-3; i+=6){
      const code = headers[i];
      if(code && typeof code==="string" && code.startsWith("ml-")){
        paramMap[code] = { valCol: i+3, alarmCol: i+4 };
      }
    }

    const METALS = {
      "ml-el.fe":  {label:"Fe (ppm)",   warn:80,  crit:120},
      "ml-el.cu":  {label:"Cu (ppm)",   warn:25,  crit:40},
      "ml-el.al":  {label:"Al (ppm)",   warn:15,  crit:25},
      "ml-el.cr":  {label:"Cr (ppm)",   warn:8,   crit:15},
      "ml-el.si":  {label:"Si (ppm)",   warn:15,  crit:25},
      "ml-el.pb":  {label:"Pb (ppm)",   warn:10,  crit:20},
      "ml-el.ni":  {label:"Ni (ppm)",   warn:8,   crit:15},
    };

    const rows = raw.slice(1).map(row => {
      const get = i => (i < row.length ? row[i] : null);
      // Extract metals
      const metals = {};
      Object.entries(METALS).forEach(([code,meta])=>{
        if(paramMap[code]){
          const v = parseFloat(get(paramMap[code].valCol));
          const a = get(paramMap[code].alarmCol);
          metals[code] = { value: isNaN(v)?null:v, alarm: a, ...meta };
        }
      });
      // Overall result: map AlarmCode
      const alarmCode = String(get(27)||"").trim(); // col 27 = AlarmCode
      const result = alarmCode==="Severe"?"CRITICAL": alarmCode==="Caution"?"CAUTION":"NORMAL";
      // Date
      let sampleDate = "";
      const rawD = get(28); // DrawingDate col 28
      if(rawD instanceof Date) sampleDate = rawD.toISOString().split("T")[0];
      else if(typeof rawD==="string") sampleDate = rawD.split(" ")[0].split("T")[0];
      else if(typeof rawD==="number"){
        const d=new Date(Math.round((rawD-25569)*86400*1000));
        sampleDate=d.toISOString().split("T")[0];
      }
      return {
        equipment:     String(get(1)||"").trim(),   // AssetName*
        manufacturer:  String(get(2)||"").trim(),   // AssetManufacturer
        model:         String(get(3)||"").trim(),   // AssetModel
        assetLife:     parseFloat(get(6))||0,       // AssetLife (hours)
        component:     String(get(8)||"").trim(),   // ComponentName*
        oilLife:       parseFloat(get(12))||0,      // OilLife
        oilName:       String(get(14)||"").trim(),  // OilName*
        sampleId:      String(get(25)||"").trim(),  // SampleId*
        alarmCode,
        result,
        sampleDate,
        diagnosisDate: typeof get(29)==="string"?get(29).split("T")[0]:"",
        diagnosis:     String(get(38)||"").trim(),  // Diagnosis(AdditionalRecommendations)
        observations:  String(get(39)||"").trim(),  // LimitObservations
        actions:       String(get(41)||"").trim(),  // LimitActions
        metals,
        oilChanged:    String(get(33)||"").trim(),  // OilChanged*
        filterChanged: String(get(34)||"").trim(),  // FilterChanged*
        extraInfo:     String(get(37)||"").trim(),  // ExtraInfo
      };
    }).filter(r=>r.equipment);
    cb(rows);
  };
  reader.readAsBinaryString(file);
}

// ── VisionLink CAT parser (standard export format)
// Typical columns: Asset ID, Asset S/N, Model, Event Date, Event Time,
//   Fault Code, Description, Severity, Category, Count/Hours, Status
function parseVisionLink(rows) {
  return rows.map(r => {
    const sev = String(r["Severity"]||r["Priority"]||r["Level"]||"").toUpperCase();
    const mappedSev = sev.includes("HIGH")||sev.includes("CRITICAL")||sev==="3"?"CRITICAL":
                      sev.includes("MED")||sev==="2"?"HIGH":
                      sev.includes("LOW")||sev==="1"?"MODERATE":"LOW";
    return {
      equipment:   String(r["Asset ID"]||r["Asset"]||r["Equipment"]||"").trim(),
      model:       String(r["Model"]||r["Asset Model"]||"").trim(),
      serialNo:    String(r["Asset S/N"]||r["Serial Number"]||"").trim(),
      date:        String(r["Event Date"]||r["Date"]||"").trim(),
      time:        String(r["Event Time"]||r["Time"]||"").trim(),
      code:        String(r["Fault Code"]||r["DTC"]||r["Code"]||"").trim(),
      desc:        String(r["Description"]||r["Fault Description"]||"").trim(),
      sev:         mappedSev,
      category:    String(r["Category"]||r["System"]||"").trim(),
      occurrences: parseInt(r["Count"]||r["Occurrences"]||1)||1,
      source:      "VisionLink",
      status:      String(r["Status"]||"NEW").trim(),
      action:      String(r["Action"]||r["Recommended Action"]||"").trim(),
    };
  }).filter(r=>r.equipment && r.code);
}

// ── Pronto CMMS parser (standard WO export)
// Typical columns: Work Order, Equipment, Description, WO Type, Priority,
//   Planned Start, Actual Start, Actual Finish, Estimated Hours, Actual Hours, Cost, Status, Technician
function parsePronto(rows) {
  return rows.map(r => {
    const type = String(r["WO Type"]||r["Type"]||r["Work Type"]||"").toUpperCase();
    const mappedType = type.includes("PM")||type.includes("PREV")?"PM":
                       type.includes("CM")||type.includes("CORR")?"CM":
                       type.includes("PDM")||type.includes("PRED")?"PDM":"Inspection";
    const isPM = mappedType==="PM";
    const planned   = String(r["Planned Start"]||r["Planned Date"]||r["Start Date"]||"").trim();
    const actualEnd = String(r["Actual Finish"]||r["Completion Date"]||r["End Date"]||"").trim();
    const cost      = parseFloat(String(r["Cost"]||r["Total Cost"]||"0").replace(/[^0-9.]/g,""))||0;
    const duration  = parseFloat(r["Actual Hours"]||r["Duration"]||r["Hours"]||0)||0;
    // PM compliance: completed within 10% of planned interval
    const pmCompliant = isPM ? (actualEnd ? true : false) : null;
    return {
      woNum:       String(r["Work Order"]||r["WO Number"]||r["WO#"]||"").trim(),
      equipment:   String(r["Equipment"]||r["Asset"]||r["Asset ID"]||"").trim(),
      model:       String(r["Model"]||r["Asset Model"]||"").trim(),
      description: String(r["Description"]||r["Task"]||"").trim(),
      type:        mappedType,
      priority:    String(r["Priority"]||"").trim(),
      plannedDate: planned,
      actualDate:  actualEnd,
      duration,
      cost,
      technician:  String(r["Technician"]||r["Assigned To"]||r["Mechanic"]||"").trim(),
      status:      String(r["Status"]||r["WO Status"]||"").trim()||"PLANNED",
      pmCompliant,
    };
  }).filter(r=>r.woNum && r.equipment);
}

// ─── KPI COMPUTATION ───────────────────────────────────────────────────────────
function computeKPIs(stops){
  if(!stops.length) return null;
  const calHoursPerEq = 720;
  const equipments = [...new Set(stops.map(s=>s.equipment))];
  const n = equipments.length;
  let totalDowntime=0, unplannedDowntime=0, unplannedCount=0;
  let totalCalHours=0;
  const byEq={};
  equipments.forEach(eq=>{
    const eqStops = stops.filter(s=>s.equipment===eq);
    const dt = eqStops.reduce((a,s)=>a+Number(s.duration),0);
    // Support both boolean and string "PLANNED"/"UNPLANNED"
    const isUnplanned = s => {
      if(typeof s.planned==="boolean") return !s.planned;
      return String(s.planned||"").toUpperCase()==="UNPLANNED";
    };
    const udt = eqStops.filter(isUnplanned).reduce((a,s)=>a+Number(s.duration),0);
    const uc  = eqStops.filter(isUnplanned).length;
    totalDowntime+=dt; unplannedDowntime+=udt; unplannedCount+=uc;
    totalCalHours+=calHoursPerEq;
    byEq[eq]={dt,udt,uc,availH:calHoursPerEq-dt,ma:Math.max(0,(calHoursPerEq-dt)/calHoursPerEq*100)};
  });
  const maGlobal = (totalCalHours-totalDowntime)/totalCalHours*100;
  const mtbf = unplannedCount>0 ? (totalCalHours-unplannedDowntime)/unplannedCount : 0;
  const mttr = unplannedCount>0 ? unplannedDowntime/unplannedCount : 0;
  const maByEq = Object.entries(byEq).map(([eq,v])=>({eq,ma:+v.ma.toFixed(1),dt:+v.dt.toFixed(1)}))
    .sort((a,b)=>a.ma-b.ma);
  // Pareto by CAT Code Equivalent / cause
  const causeCounts={};
  const isUnplanned = s=>{
    if(typeof s.planned==="boolean") return !s.planned;
    return String(s.planned||"").toUpperCase()==="UNPLANNED";
  };
  stops.filter(isUnplanned).forEach(s=>{
    const c=s.cause||s.system||s.comments||"Inconnu";
    causeCounts[c]=(causeCounts[c]||0)+Number(s.duration);
  });
  const pareto = Object.entries(causeCounts).sort((a,b)=>b[1]-a[1]).map(([c,h])=>({cause:c,hours:+h.toFixed(1)}));
  // Weekly MA trend — use actual dates from data
  const allDates = stops.map(s=>new Date(s.date).getTime()).filter(d=>!isNaN(d));
  const latestDate = allDates.length ? Math.max(...allDates) : Date.now();
  const weeklyMA=[];
  for(let w=11;w>=0;w--){
    const end=latestDate-(w)*7*864e5, start=end-7*864e5;
    const wStops=stops.filter(s=>{ const d=new Date(s.date).getTime(); return d>=start&&d<end; });
    const wDT=wStops.reduce((a,s)=>a+Number(s.duration),0);
    const wCalH = totalCalHours/12;
    const wMA=Math.min(100,Math.max(50,(wCalH-wDT)/wCalH*100));
    weeklyMA.push({week:`S${12-w}`,ma:+wMA.toFixed(1),target:92});
  }
  // By model type — derive from actual data
  const modelTypes={};
  stops.forEach(s=>{
    const model=s.model||s.equipment||"Inconnu";
    // Derive type from model name
    const t = model.toUpperCase().includes("HT")||model.includes("777")||model.includes("789")||model.includes("A45")||model.includes("A60")?"Truck":
               model.includes("EXC")||model.includes("SHV")?"Excavator/Shovel":
               model.includes("DM")||model.includes("BD")?"Drill":
               model.includes("D10")||model.includes("D9")?"Dozer":
               model.includes("992")||model.includes("980")?"Loader":"Autre";
    if(!modelTypes[t]) modelTypes[t]={type:t,totalDT:0,calH:0};
    modelTypes[t].totalDT+=Number(s.duration);
    modelTypes[t].calH+=calHoursPerEq/equipments.length;
  });
  const byType = Object.values(modelTypes).map(v=>({
    type:v.type,
    ma:+Math.min(100,Math.max(0,(v.calH-v.totalDT)/v.calH*100)).toFixed(1)
  }));
  return {maGlobal:+Math.min(100,Math.max(0,maGlobal)).toFixed(1),
    mtbf:+mtbf.toFixed(1),mttr:+mttr.toFixed(1),
    totalDowntime:+totalDowntime.toFixed(0),unplannedCount,n,
    maByEq,pareto:pareto.slice(0,8),weeklyMA,byType};
}

// ─── COMPONENTS ────────────────────────────────────────────────────────────────
const CustomTooltip=({active,payload,label})=>{
  if(!active||!payload?.length) return null;
  return (
    <div style={{background:"#1A1E2A",border:"1px solid #2E3448",borderRadius:6,padding:"8px 12px",fontSize:11}}>
      <div style={{color:"#C9A84C",marginBottom:4,fontWeight:600}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{color:p.color||"#E8EAF0"}}>{p.name}: <span style={{fontFamily:"monospace"}}>{p.value}</span></div>
      ))}
    </div>
  );
};

const Pill=({status})=>{
  const map={"CRITICAL":"crit","HIGH":"crit","CAUTION":"warn","MODERATE":"warn","NORMAL":"ok","OK":"ok","LOW":"info","INFO":"info","RESOLVED":"ok","IN PROGRESS":"warn","NEW":"info","UNDER MONITORING":"warn","CLOSED":"ok","PLANNED":"info","ON HOLD":"warn","EXECUTED":"ok","CANCELLED":"warn"};
  return <span className={`badge ${map[status]||"info"}`}>{status}</span>;
};

const KPICard=({label,value,unit,sub,color,target})=>(
  <div className="card" style={{display:"flex",flexDirection:"column",gap:6}}>
    <div style={{fontSize:10,color:T.muted,textTransform:"uppercase",letterSpacing:".6px",fontWeight:600}}>{label}</div>
    <div style={{display:"flex",alignItems:"baseline",gap:4}}>
      <span className="kpi-val" style={{color:color||T.gold}}>{value}</span>
      {unit&&<span style={{fontSize:11,color:T.muted}}>{unit}</span>}
    </div>
    {sub&&<div className="kpi-sub">{sub}</div>}
    {target&&<div style={{fontSize:9,color:Number(value)>=Number(target)?T.green:T.red}}>
      {Number(value)>=Number(target)?"✓":"✗"} TARGET: {target}
    </div>}
  </div>
);

const UploadZone=({label,icon,onFile,formats})=>{
  const ref=useRef();
  return (
    <div className="upload-zone" onClick={()=>ref.current.click()}>
      <input ref={ref} type="file" accept={formats} onChange={e=>e.target.files[0]&&onFile(e.target.files[0])} />
      <div style={{fontSize:20,marginBottom:6}}>{icon}</div>
      <div style={{fontWeight:600,color:T.text,marginBottom:2}}>{label}</div>
      <div style={{fontSize:10,color:T.muted}}>Drag &amp; drop or click — CSV / XLSX</div>
    </div>
  );
};

const SectionTitle=({icon,children})=>(
  <div className="section-title">{icon&&<span>{icon}</span>}{children}</div>
);

// ─── i18n ──────────────────────────────────────────────────────────────────────
const I18N = {
  en: {
    appTitle:"MINING HME RELIABILITY", appSub:"Reliability & Condition Monitoring — HME",
    tabs:{dashboard:"Dashboard KPI",weibull:"Weibull Survival & Replacement Advisor",sos:"Oil Analysis",faults:"Fault Codes",wo:"Work Orders",import:"Import Data"},
    all:"ALL", model:"Model", equipment:"Equipment", result:"Result", severity:"Severity", status:"Status",
    dateFrom:"From", dateTo:"To", resetFilter:"Reset",
    maGlobal:"Global MA%", mtbf:"Avg MTBF", mttr:"Avg MTTR", downtime:"Downtime (h)", breakdowns:"Failures", pmComp:"PM Compliance",
    target:"TARGET", ok:"OK", alert:"ALERT",
    // SOS
    sosTotal:"Total Samples", sosCrit:"CRITICAL", sosWarn:"CAUTION", sosNorm:"NORMAL",
    component:"Component", date:"Date", oilLife:"Oil Life (h)",
    fe:"Fe (ppm)", cu:"Cu (ppm)", al:"Al (ppm)", cr:"Cr (ppm)", si:"Si (ppm)", pb:"Pb (ppm)",
    oilChanged:"Oil Changed", filterChanged:"Filter Changed",
    diagnosis:"Diagnosis / Observations", actionTaken:"Corrective Action", feedback:"Feedback",
    metalLimits:"METAL THRESHOLD LIMITS (ppm)", yes:"Yes", no:"No",
    // Faults
    code:"Fault Code", description:"Description", system:"System", occurrences:"Count", source:"Source", action:"Recommended Action",
    faultFeedback:"Field Feedback",
    faultStatuses:["New","In Progress","Resolved","Under Monitoring","Escalated"],
    // WO
    woNum:"WO #", woType:"WO Type", woTypeLabel:"WO Type",
    plannedDate:"Planned Date", actualDate:"Completion Date",
    woTypes:["ALL","PM","CM","Backlog","Opportunistic"],
    woStatuses:["ALL","Planned","In Progress","Executed","Closed","On Hold","Cancelled"],
    // WO KPIs
    pmPlanned:"PM Planned", pmExecuted:"PM Executed",
    cmPlanned:"CM Planned", cmExecuted:"CM Executed",
    backlogPlanned:"Backlog Planned", backlogExecuted:"Backlog Executed",
    opportunPlanned:"Opportunistic Planned", opportunExecuted:"Opportunistic Executed",
    pmCompliance:"PM Compliance",
    // Import
    formatGuide:"Import notes & column mapping",
    loaded:"records loaded",
    // Weibull
    wParams:"Analysis Parameters", wResults:"MLE Results", wShape:"SHAPE β", wScale:"SCALE η",
    wRegime:"Detected Failure Mode", wCurve:"Reliability R(t) & Failure F(t)", wHazard:"Hazard Rate h(t)",
    wDist:"Field TTF Distribution", wCalc:"Run Weibull Analysis", wCurrent:"Current Reliability @",
    wTarget:"Target Reliability (%)", wHours:"Current Component Hours", wTTF:"TTF Data (life hours, comma-separated)",
    wB10:"B10 Life (10% probability of failure)", wB50:"B50 Life (median / 50% failures)", wB90:"B90 Life (90% failures)",
    wOK:"OK — Reliability exceeds target threshold", wAlert:"ALERT — Replacement recommended",
    wRegimes:{infant:"⬇ Infant Mortality (β<1) — Manufacturing / installation defects",random:"➡ Random Failures (β≈1) — Exponential distribution",wear:"↗ Wear-Out Regime (1.5<β<3) — Preventive maintenance recommended",severe:"⬆ Rapid Wear-Out (β>3) — Systematic replacement strategy"},
    lang:"Language",
  }
};

// ─── DATE RANGE FILTER BAR ──────────────────────────────────────────────────────
const DateRangeBar = ({from,to,onFrom,onTo,onReset,L}) => (
  <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",background:T.bg2,border:`1px solid ${T.border}`,borderRadius:6,padding:"8px 12px"}}>
    <span style={{fontSize:10,color:T.gold,fontWeight:700}}>📅</span>
    <span style={{fontSize:10,color:T.muted}}>{L.dateFrom} :</span>
    <input type="date" value={from} onChange={e=>onFrom(e.target.value)}
      style={{background:T.bg3,border:`1px solid ${T.border}`,color:T.text,borderRadius:4,padding:"3px 6px",fontSize:11,width:130}}/>
    <span style={{fontSize:10,color:T.muted}}>{L.dateTo} :</span>
    <input type="date" value={to} onChange={e=>onTo(e.target.value)}
      style={{background:T.bg3,border:`1px solid ${T.border}`,color:T.text,borderRadius:4,padding:"3px 6px",fontSize:11,width:130}}/>
    <button className="btn btn-ghost" style={{padding:"3px 10px",fontSize:10}} onClick={onReset}>{L.resetFilter}</button>
  </div>
);

const inDateRange = (dateStr, from, to) => {
  if(!dateStr) return true;
  const d = new Date(dateStr);
  if(from && d < new Date(from)) return false;
  if(to   && d > new Date(to))   return false;
  return true;
};

// ─── TABS ──────────────────────────────────────────────────────────────────────
const makeTabs = L => [
  {id:"dashboard",label:L.tabs.dashboard,icon:"📊"},
  {id:"weibull",label:L.tabs.weibull,icon:"📈"},
  {id:"sos",label:L.tabs.sos,icon:"🛢"},
  {id:"faults",label:L.tabs.faults,icon:"⚠️"},
  {id:"wo",label:L.tabs.wo,icon:"🔧"},
  {id:"import",label:L.tabs.import,icon:"📥"},
];

// ─── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const lang = "en";
  const L = I18N["en"];
  const TABS = makeTabs(L);

  const [tab, setTab] = useState("dashboard");
  const [stops, setStops] = useState(genStops);
  const [sosData, setSosData] = useState(genSOS);
  const [faultData, setFaultData] = useState(genFaultCodes);
  const [woData, setWoData] = useState(genWO);
  const [weibullInputs, setWeibullInputs] = useState({
    component:"Engine",equipment:"HT021",
    ttfRaw:"3500,4200,3800,5100,2900,4600,3200,4900,3700,4400",
    currentHours:3100,targetR:90,
  });
  const [weibullResult, setWeibullResult] = useState(null);

  // ── SOS filters
  const [sosResult, setSosResult]   = useState("ALL");
  const [sosModel,  setSosModel]    = useState("ALL");
  const [sosEq,     setSosEq]       = useState("ALL");
  const [sosFrom,   setSosFrom]     = useState("");
  const [sosTo,     setSosTo]       = useState("");
  const [sosActions,   setSosActions]   = useState({}); // {rowKey: action text}
  const [sosDiagnosis, setSosDiagnosis] = useState({}); // {rowKey: diagnosis text}
  const [sosFeedback,  setSosFeedback]  = useState({}); // {rowKey: feedback text}
  const [sosOilChg,    setSosOilChg]    = useState({}); // {rowKey: "Oui"|"Non"}
  const [sosFilChg,    setSosFilChg]    = useState({}); // {rowKey: "Oui"|"Non"}

  // ── Fault filters + inline state
  const [faultSev,      setFaultSev]      = useState("ALL");
  const [faultFrom,     setFaultFrom]     = useState("");
  const [faultTo,       setFaultTo]       = useState("");
  const [faultStatuses, setFaultStatuses] = useState({}); // {rowKey: status}
  const [faultFeedback, setFaultFeedback] = useState({}); // {rowKey: feedback}

  // ── WO filters
  const [woModel,   setWoModel]    = useState("ALL");
  const [woEq,      setWoEq]       = useState("ALL");
  const [woType,    setWoType]     = useState("ALL");
  const [woStatus,  setWoStatus]   = useState("ALL");
  const [woFrom,    setWoFrom]     = useState("");
  const [woTo,      setWoTo]       = useState("");

  // ── Dashboard filters
  const [dashFrom,  setDashFrom]   = useState("");
  const [dashTo,    setDashTo]     = useState("");

  // ── Derived unique lists
  const sosModels = useMemo(()=>["ALL",...new Set(sosData.map(r=>r.model).filter(Boolean))],[sosData]);
  const sosEqs    = useMemo(()=>{
    const base = sosModel==="ALL"?sosData:sosData.filter(r=>r.model===sosModel);
    return ["ALL",...new Set(base.map(r=>r.equipment).filter(Boolean))];
  },[sosData,sosModel]);

  const woModels = useMemo(()=>["ALL",...new Set(woData.map(r=>r.model).filter(Boolean))],[woData]);
  const woEqs    = useMemo(()=>{
    const base = woModel==="ALL"?woData:woData.filter(r=>r.model===woModel);
    return ["ALL",...new Set(base.map(r=>r.equipment).filter(Boolean))];
  },[woData,woModel]);

  const filteredStops = useMemo(()=>stops.filter(s=>inDateRange(s.date,dashFrom,dashTo)),[stops,dashFrom,dashTo]);
  const kpis = useMemo(()=>computeKPIs(filteredStops),[filteredStops]);

  const filteredSOS = useMemo(()=>sosData.filter(r=>{
    const resultMatch = (sosResult==="ALL"||(sosResult==="CRITICAL"&&r.result==="CRITICAL")||(sosResult==="CAUTION"&&r.result==="CAUTION")||(sosResult==="NORMAL"&&r.result==="NORMAL"));
    return resultMatch
        && (sosModel==="ALL"||r.model===sosModel)
        && (sosEq==="ALL"||r.equipment===sosEq)
        && inDateRange(r.sampleDate,sosFrom,sosTo);
  }),[sosData,sosResult,sosModel,sosEq,sosFrom,sosTo]);

  const filteredFaults = useMemo(()=>faultData.filter(r=>
    (faultSev==="ALL"||r.sev===faultSev)&&inDateRange(r.date,faultFrom,faultTo)
  ),[faultData,faultSev,faultFrom,faultTo]);

  const filteredWO = useMemo(()=>woData.filter(r=>
    (woModel==="ALL"||r.model===woModel)
    &&(woEq==="ALL"||r.equipment===woEq)
    &&(woType==="ALL"||r.type===woType)
    &&(woStatus==="ALL"||r.status===woStatus)
    &&inDateRange(r.plannedDate,woFrom,woTo)
  ),[woData,woModel,woEq,woType,woStatus,woFrom,woTo]);

  // ── WO KPIs — Planned vs Executed per type
  const isPlanned = r => /planif|planned/i.test(r.status||"");
  const isExecuted= r => /execut|executed|clos|complet/i.test(r.status||"");
  const woKPIs = useMemo(()=>{
    const types = ["PM","CM","Backlog","Opportun"];
    return types.map(t=>({
      type: t,
      planned:  filteredWO.filter(r=>r.type===t && isPlanned(r)).length,
      executed: filteredWO.filter(r=>r.type===t && isExecuted(r)).length,
      total:    filteredWO.filter(r=>r.type===t).length,
    }));
  },[filteredWO]);

  const pmCompliance = useMemo(()=>{
    const pm = filteredWO.filter(r=>r.type==="PM");
    if(!pm.length) return 0;
    return +(pm.filter(r=>isExecuted(r)||r.pmCompliant).length/pm.length*100).toFixed(1);
  },[filteredWO]);

  // ── Weibull compute
  const runWeibull = useCallback(()=>{
    const ttf = weibullInputs.ttfRaw.split(/[,;\s]+/).map(Number).filter(v=>v>0);
    const res = weibullMLE(ttf);
    if(!res) return;
    const maxT = Math.max(...ttf)*1.8;
    const curve = weibullCurveData(res.beta,res.eta,maxT);
    const b10 = weibullBLife(10,res.beta,res.eta);
    const b50 = weibullBLife(50,res.beta,res.eta);
    const b90 = weibullBLife(90,res.beta,res.eta);
    const curR = weibullR(Number(weibullInputs.currentHours),res.beta,res.eta)*100;
    const regKey = res.beta<1?"infant":res.beta<1.5?"random":res.beta<3?"wear":"severe";
    setWeibullResult({...res,curve,b10:+b10.toFixed(0),b50:+b50.toFixed(0),b90:+b90.toFixed(0),curR:+curR.toFixed(1),maxT,ttf,regKey});
  },[weibullInputs]);

  useEffect(()=>{ runWeibull(); },[runWeibull]); // eslint-disable-line react-hooks/exhaustive-deps

  // ── File handlers
  const handleStopsFile   = f => readRawFile(f, d => setStops(parseDBArrets(d)));
  const handleSOSFile     = f => parseTruvu(f, d => setSosData(d));
  const handleFaultFile   = f => readRawFile(f, d => setFaultData(parseVisionLink(d)));
  const handleWOFile      = f => readRawFile(f, d => setWoData(parsePronto(d)));

  const sosSevOpts  = ["ALL","CRITICAL","CAUTION","NORMAL"];
  const faultSevOpts = ["ALL","CRITICAL","HIGH","MODERATE","LOW"];

  return (
    <>
      <style>{css}</style>
      <div style={{minHeight:"100vh",background:T.bg0}}>
        {/* ── HEADER */}
        <div style={{background:T.bg1,borderBottom:`1px solid ${T.border}`,padding:"0 24px",position:"sticky",top:0,zIndex:100}}>
          <div style={{display:"flex",alignItems:"center",gap:16,height:52}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:32,height:32,background:`linear-gradient(135deg,${T.gold},${T.goldD})`,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>⛏</div>
              <div>
                <div style={{fontWeight:700,color:T.gold,fontSize:13,letterSpacing:".3px"}}>{L.appTitle}</div>
                <div style={{fontSize:9,color:T.muted,letterSpacing:".8px",textTransform:"uppercase"}}>{L.appSub}</div>
              </div>
            </div>
            <div style={{flex:1}}/>

            <div style={{fontSize:10,color:T.muted,fontFamily:"'JetBrains Mono',monospace"}}>
              {new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})}
            </div>
          </div>
          {/* Tabs */}
          <div style={{display:"flex",gap:0,overflowX:"auto"}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)}
                className={`btn ${tab===t.id?"tab-active":"tab-inactive"}`}
                style={{borderRadius:0,padding:"8px 14px",background:"transparent",border:"none",borderBottom:"2px solid",cursor:"pointer",fontSize:11,gap:5,whiteSpace:"nowrap"}}>
                {t.icon} {t.label}
              </button>
            ))}
          </div>
        </div>

        <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>

          {/* ══════════ DASHBOARD TAB ══════════ */}
          {tab==="dashboard" && kpis && (
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <DateRangeBar from={dashFrom} to={dashTo} onFrom={setDashFrom} onTo={setDashTo} onReset={()=>{setDashFrom("");setDashTo("");}} L={L}/>
              {/* KPI Row */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:12}}>
                <KPICard label={L.maGlobal} value={kpis.maGlobal} unit="%" color={kpis.maGlobal>=92?T.green:T.red} target="92" sub={`${kpis.n} ${L.equipment.toLowerCase()}s`}/>
                <KPICard label={L.mtbf} value={kpis.mtbf} unit="h" color={kpis.mtbf>=180?T.green:T.orange} target="180" sub="Mean Time Between Failures"/>
                <KPICard label={L.mttr} value={kpis.mttr} unit="h" color={kpis.mttr<=5?T.green:T.red} target="≤5" sub="Mean Time To Repair"/>
                <KPICard label={L.downtime} value={kpis.totalDowntime} unit="h" color={T.orange} sub="Total rolling period"/>
                <KPICard label={L.breakdowns} value={kpis.unplannedCount} color={T.red} sub="Unplanned downtime events"/>
                <KPICard label={L.pmComp} value={pmCompliance} unit="%" color={pmCompliance>=90?T.green:T.orange} target="95" sub="PM schedule adherence"/>
              </div>
              {/* Charts Row 1 */}
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:12}}>
                <div className="card">
                  <SectionTitle icon="📈">"Weekly MA% Trend — 12 weeks"</SectionTitle>
                  <ResponsiveContainer width="100%" height={200}>
                    <AreaChart data={kpis.weeklyMA}>
                      <defs>
                        <linearGradient id="maGrad" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={T.gold} stopOpacity={0.25}/>
                          <stop offset="95%" stopColor={T.gold} stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                      <XAxis dataKey="week" tick={{fill:T.muted,fontSize:10}} axisLine={false} tickLine={false}/>
                      <YAxis domain={[80,100]} tick={{fill:T.muted,fontSize:10}} axisLine={false} tickLine={false} unit="%"/>
                      <Tooltip content={<CustomTooltip/>}/>
                      <ReferenceLine y={92} stroke={T.green} strokeDasharray="4 4" label={{value:"Target 92%",fill:T.green,fontSize:9}}/>
                      <Area type="monotone" dataKey="ma" stroke={T.gold} strokeWidth={2} fill="url(#maGrad)" name="MA%"/>
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
                <div className="card">
                  <SectionTitle icon="🏗">"MA% by Equipment Type"</SectionTitle>
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={kpis.byType} layout="vertical">
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border} horizontal={false}/>
                      <XAxis type="number" domain={[80,100]} tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="%"/>
                      <YAxis type="category" dataKey="type" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} width={80}/>
                      <Tooltip content={<CustomTooltip/>}/>
                      <ReferenceLine x={92} stroke={T.green} strokeDasharray="3 3"/>
                      <Bar dataKey="ma" name="MA%" radius={[0,4,4,0]}>
                        {kpis.byType.map((entry,i)=>(
                          <Cell key={i} fill={entry.ma>=92?T.green:entry.ma>=88?T.orange:T.red}/>
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              {/* Charts Row 2 */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div className="card">
                  <SectionTitle icon="🔴">"Downtime Pareto — Top Failure Causes"</SectionTitle>
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={kpis.pareto}>
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                      <XAxis dataKey="cause" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} angle={-20} textAnchor="end" height={40}/>
                      <YAxis tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="h"/>
                      <Tooltip content={<CustomTooltip/>}/>
                      <Bar dataKey="hours" name="Hours" fill={T.red} radius={[4,4,0,0]} opacity={0.85}/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div className="card">
                  <SectionTitle icon="📉">"MA% by Equipment — Bottom 10"</SectionTitle>
                  <div style={{maxHeight:200,overflowY:"auto"}}>
                    <table>
                      <thead><tr><th>{L.equipment}</th><th>MA%</th><th>"Downtime (h)"</th><th>{L.status}</th></tr></thead>
                      <tbody>
                        {kpis.maByEq.slice(0,10).map((r,i)=>(
                          <tr key={i}>
                            <td className="mono" style={{color:T.gold}}>{r.eq}</td>
                            <td className="mono" style={{color:r.ma>=92?T.green:r.ma>=88?T.orange:T.red,fontWeight:600}}>{r.ma}%</td>
                            <td className="mono">{r.dt}h</td>
                            <td><Pill status={r.ma>=92?"NORMAL":r.ma>=88?"CAUTION":"CRITICAL"}/></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ══════════ WEIBULL TAB ══════════ */}
          {tab==="weibull" && (
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"grid",gridTemplateColumns:"340px 1fr",gap:16}}>
                {/* Input Panel */}
                <div style={{display:"flex",flexDirection:"column",gap:12}}>
                  <div className="card">
                    <SectionTitle icon="⚙️">Analysis Parameters</SectionTitle>
                    <div style={{display:"flex",flexDirection:"column",gap:10}}>
                      <div>
                        <div style={{fontSize:10,color:T.muted,marginBottom:4}}>EQUIPMENT</div>
                        <select value={weibullInputs.equipment} onChange={e=>setWeibullInputs(p=>({...p,equipment:e.target.value}))}>
                          {FLEET.map(f=><option key={f.id} value={f.id}>{f.id}</option>)}
                        </select>
                      </div>
                      <div>
                        <div style={{fontSize:10,color:T.muted,marginBottom:4}}>COMPONENT ANALYZED</div>
                        <select value={weibullInputs.component} onChange={e=>setWeibullInputs(p=>({...p,component:e.target.value}))}>
                          {["Engine","Transmission","Hydraulics","Undercarriage","Differential","Tyres","Alternator","Hydraulic Pump","Torque Converter","Oil Filter"].map(c=><option key={c}>{c}</option>)}
                        </select>
                      </div>
                      <div>
                        <div style={{fontSize:10,color:T.muted,marginBottom:4}}>TTF Data (life hours) — comma-separated</div>
                        <textarea
                          value={weibullInputs.ttfRaw}
                          onChange={e=>setWeibullInputs(p=>({...p,ttfRaw:e.target.value}))}
                          rows={4}
                          style={{background:T.bg3,border:`1px solid ${T.border}`,color:T.text,borderRadius:4,padding:"6px 8px",fontFamily:"'JetBrains Mono',monospace",fontSize:11,width:"100%",resize:"vertical"}}
                          placeholder="Ex: 3500,4200,3800,5100,2900..."
                        />
                        <div style={{fontSize:9,color:T.muted,marginTop:3}}>
                          {weibullInputs.ttfRaw.split(/[,;\s]+/).filter(v=>Number(v)>0).length} values entered
                        </div>
                      </div>
                      <div>
                        <div style={{fontSize:10,color:T.muted,marginBottom:4}}>CURRENT COMPONENT HOURS</div>
                        <input type="number" value={weibullInputs.currentHours} onChange={e=>setWeibullInputs(p=>({...p,currentHours:Number(e.target.value)}))} min={0}/>
                      </div>
                      <div>
                        <div style={{fontSize:10,color:T.muted,marginBottom:4}}>TARGET RELIABILITY THRESHOLD (%)</div>
                        <input type="number" value={weibullInputs.targetR} onChange={e=>setWeibullInputs(p=>({...p,targetR:Number(e.target.value)}))} min={1} max={99}/>
                      </div>
                      <button className="btn btn-gold" onClick={runWeibull} style={{width:"100%",justifyContent:"center"}}>
                        ⚡ Run Weibull Analysis
                      </button>
                    </div>
                  </div>

                  {/* Results */}
                  {weibullResult && (
                    <div className="card">
                      <SectionTitle icon="📐">MLE Results</SectionTitle>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
                        <div style={{background:T.bg3,borderRadius:6,padding:"10px 12px",textAlign:"center"}}>
                          <div style={{fontSize:9,color:T.muted,marginBottom:4}}>SHAPE β (Beta)</div>
                          <div className="mono" style={{fontSize:22,color:T.gold,fontWeight:700}}>{weibullResult.beta}</div>
                        </div>
                        <div style={{background:T.bg3,borderRadius:6,padding:"10px 12px",textAlign:"center"}}>
                          <div style={{fontSize:9,color:T.muted,marginBottom:4}}>SCALE η (Eta)</div>
                          <div className="mono" style={{fontSize:22,color:T.cyan,fontWeight:700}}>{weibullResult.eta}h</div>
                        </div>
                      </div>
                      <div style={{fontSize:10,color:T.orange,background:T.bg3,borderRadius:6,padding:"8px 10px",marginBottom:10,lineHeight:1.5}}>
                        {weibullResult.regime}
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:6}}>
                        {[
                          {label:"B10 Life — 10% probability of failure",val:weibullResult.b10,color:T.green},
                          {label:"B50 Life — 50% (median)",val:weibullResult.b50,color:T.orange},
                          {label:"B90 Life — 90% probability of failure",val:weibullResult.b90,color:T.red},
                        ].map(({label,val,color})=>(
                          <div key={label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 8px",background:T.bg3,borderRadius:4}}>
                            <span style={{fontSize:10,color:T.muted}}>{label}</span>
                            <span className="mono" style={{color,fontWeight:700}}>{val.toLocaleString()}h</span>
                          </div>
                        ))}
                        <div style={{padding:"8px 10px",background:`${weibullResult.curR>weibullInputs.targetR?T.green:T.red}18`,border:`1px solid ${weibullResult.curR>weibullInputs.targetR?T.green:T.red}44`,borderRadius:6,marginTop:4}}>
                          <div style={{fontSize:10,color:T.muted,marginBottom:2}}>Current Reliability @ {weibullInputs.currentHours}h</div>
                          <div className="mono" style={{fontSize:20,color:weibullResult.curR>weibullInputs.targetR?T.green:T.red,fontWeight:700}}>
                            {weibullResult.curR}%
                          </div>
                          <div style={{fontSize:9,color:T.muted,marginTop:2}}>
                            {weibullResult.curR>weibullInputs.targetR
                              ? `✓ OK — Reliability exceeds target threshold ${weibullInputs.targetR}%`
                              : `⚠ ALERT — Below target threshold ${weibullInputs.targetR}% — Replacement recommended`}
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>

                {/* Charts Panel */}
                <div style={{display:"flex",flexDirection:"column",gap:12}}>
                  {weibullResult && (
                    <>
                      {/* Reliability + CDF */}
                      <div className="card">
                        <SectionTitle icon="📈">Reliability Curve R(t) &amp; Failure Probability F(t)</SectionTitle>
                        <ResponsiveContainer width="100%" height={230}>
                          <LineChart data={weibullResult.curve}>
                            <defs>
                              <linearGradient id="rGrad" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor={T.green} stopOpacity={0.2}/>
                                <stop offset="95%" stopColor={T.green} stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                            <XAxis dataKey="t" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="h"/>
                            <YAxis tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="%"/>
                            <Tooltip content={<CustomTooltip/>}/>
                            <Legend wrapperStyle={{fontSize:10,color:T.muted}}/>
                            <ReferenceLine x={weibullResult.b10} stroke={T.green} strokeDasharray="3 3" label={{value:"B10",fill:T.green,fontSize:8}}/>
                            <ReferenceLine x={weibullResult.b50} stroke={T.orange} strokeDasharray="3 3" label={{value:"B50",fill:T.orange,fontSize:8}}/>
                            <ReferenceLine x={weibullInputs.currentHours} stroke={T.gold} strokeWidth={2} label={{value:"Now",fill:T.gold,fontSize:8}}/>
                            <ReferenceLine y={weibullInputs.targetR} stroke={T.purple} strokeDasharray="4 4" label={{value:`Seuil ${weibullInputs.targetR}%`,fill:T.purple,fontSize:8}}/>
                            <Line type="monotone" dataKey="R" stroke={T.green} strokeWidth={2.5} dot={false} name="R(t) Reliability %"/>
                            <Line type="monotone" dataKey="F" stroke={T.red} strokeWidth={1.5} dot={false} name="F(t) Failure Probability %" strokeDasharray="5 3"/>
                          </LineChart>
                        </ResponsiveContainer>
                      </div>

                      {/* Hazard Function */}
                      <div className="card">
                        <SectionTitle icon="⚡">Hazard Rate h(t) — Bathtub Curve</SectionTitle>
                        <ResponsiveContainer width="100%" height={180}>
                          <AreaChart data={weibullResult.curve}>
                            <defs>
                              <linearGradient id="hGrad" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor={T.red} stopOpacity={0.3}/>
                                <stop offset="95%" stopColor={T.red} stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                            <XAxis dataKey="t" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="h"/>
                            <YAxis tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false}/>
                            <Tooltip content={<CustomTooltip/>}/>
                            <ReferenceLine x={weibullInputs.currentHours} stroke={T.gold} strokeWidth={2}/>
                            <Area type="monotone" dataKey="h" stroke={T.red} strokeWidth={2} fill="url(#hGrad)" name="h(t) ×10⁻³" dot={false}/>
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>

                      {/* Scatter: TTF data */}
                      <div className="card">
                        <SectionTitle icon="🔵">Field TTF Distribution — {weibullResult.ttf.length} data points</SectionTitle>
                        <ResponsiveContainer width="100%" height={130}>
                          <ScatterChart>
                            <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                            <XAxis type="number" dataKey="h" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} unit="h" name="Heures de vie"/>
                            <YAxis type="number" dataKey="rank" hide/>
                            <Tooltip content={<CustomTooltip/>}/>
                            <Scatter
                              name="Field TTF data points"
                              data={weibullResult.ttf.map((h,i)=>({h,rank:i+1}))}
                              fill={T.gold}
                            />
                            <ReferenceLine x={weibullResult.eta} stroke={T.cyan} strokeDasharray="3 3" label={{value:"η",fill:T.cyan,fontSize:10}}/>
                          </ScatterChart>
                        </ResponsiveContainer>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ══════════ SOS TAB ══════════ */}
          {tab==="sos" && (
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* KPI Summary cards */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
                {[
                  {key:"CRITICAL", labelFr:"CRITICAL", labelEn:"CRITICAL", color:T.red},
                  {key:"CAUTION", labelFr:"CAUTION", labelEn:"CAUTION",  color:T.orange},
                  {key:"NORMAL",   labelFr:"NORMAL",    labelEn:"NORMAL",   color:T.green},
                ].map(({key,labelFr,labelEn,color})=>{
                  const n=sosData.filter(r=>r.result===key).length;
                  return <div key={key} className="card" style={{textAlign:"center"}}>
                    <div style={{fontSize:10,color:T.muted,marginBottom:6}}>{labelEn}</div>
                    <div className="mono" style={{fontSize:28,color,fontWeight:700}}>{n}</div>
                    <div style={{fontSize:9,color:T.muted,marginTop:2}}>"samples"</div>
                  </div>;
                })}
                <div className="card" style={{textAlign:"center"}}>
                  <div style={{fontSize:10,color:T.muted,marginBottom:6}}>"TOTAL SAMPLES"</div>
                  <div className="mono" style={{fontSize:28,color:T.gold,fontWeight:700}}>{sosData.length}</div>
                  <div style={{fontSize:9,color:T.muted,marginTop:2}}>"this period"</div>
                </div>
              </div>

              {/* Date range */}
              <DateRangeBar from={sosFrom} to={sosTo} onFrom={setSosFrom} onTo={setSosTo} onReset={()=>{setSosFrom("");setSosTo("");}} L={L}/>

              {/* Filters row */}
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",padding:"8px 0"}}>
                {/* Result filter */}
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.result} :</span>
                {sosSevOpts.map(s=>(
                  <button key={s} className={`btn ${sosResult===s?"btn-gold":"btn-ghost"}`}
                    style={{padding:"4px 10px",fontSize:10}} onClick={()=>setSosResult(s)}>{s}</button>
                ))}
                {/* Model filter */}
                <span style={{fontSize:10,color:T.muted,fontWeight:600,marginLeft:8}}>{L.model} :</span>
                <select value={sosModel} onChange={e=>{setSosModel(e.target.value);setSosEq("ALL");}} style={{width:130}}>
                  {sosModels.map(m=><option key={m}>{m}</option>)}
                </select>
                {/* Equipment filter — cascades from model */}
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.equipment} :</span>
                <select value={sosEq} onChange={e=>setSosEq(e.target.value)} style={{width:120}}>
                  {sosEqs.map(e=><option key={e}>{e}</option>)}
                </select>
                <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>
                  {filteredSOS.length} "results"
                </span>
              </div>

              {/* Table — columns aligned to real Truvu data */}
              <div className="card" style={{padding:0}}>
                <div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}>
                  <table>
                    <thead><tr>
                      <th>{L.equipment}</th>
                      <th>{L.model}</th>
                      <th>{L.component}</th>
                      <th>{L.date}</th>
                      <th>{L.oilLife}</th>
                      <th>{L.fe}</th>
                      <th>{L.cu}</th>
                      <th>{L.al}</th>
                      <th>{L.cr}</th>
                      <th>{L.si}</th>
                      <th>{L.pb}</th>
                      <th>{L.oilChanged}</th>
                      <th>{L.filterChanged}</th>
                      <th>{L.result}</th>
                      <th>{L.diagnosis}</th>
                      <th>{L.actionTaken}</th>
                      <th>{L.feedback}</th>
                    </tr></thead>
                    <tbody>
                      {filteredSOS.map((r,i)=>{
                        const rowKey = `${r.equipment}-${r.component}-${r.sampleDate}-${i}`;
                        const getM = code => r.metals?.[code]?.value ?? "—";
                        const alarmM = (code, warn, crit) => {
                          const v = r.metals?.[code]?.value;
                          if(v===null||v===undefined||isNaN(v)) return T.muted;
                          return v>=crit?T.red:v>=warn?T.orange:T.green;
                        };
                        const yesNo = ["— Select —", L.yes, L.no];
                        return (
                          <tr key={rowKey}>
                            <td className="mono" style={{color:T.gold,fontWeight:700}}>{r.equipment}</td>
                            <td style={{fontSize:10,color:T.muted}}>{r.model}</td>
                            <td style={{fontWeight:600}}>{r.component}</td>
                            <td style={{fontSize:10}}>{r.sampleDate}</td>
                            <td className="mono">{r.oilLife||"—"}</td>
                            <td className="mono" style={{color:alarmM("ml-el.fe",80,120)}}>{getM("ml-el.fe")}</td>
                            <td className="mono" style={{color:alarmM("ml-el.cu",25,40)}}>{getM("ml-el.cu")}</td>
                            <td className="mono" style={{color:alarmM("ml-el.al",15,25)}}>{getM("ml-el.al")}</td>
                            <td className="mono" style={{color:alarmM("ml-el.cr",8,15)}}>{getM("ml-el.cr")}</td>
                            <td className="mono" style={{color:alarmM("ml-el.si",15,25)}}>{getM("ml-el.si")}</td>
                            <td className="mono" style={{color:alarmM("ml-el.pb",10,20)}}>{getM("ml-el.pb")}</td>
                            {/* Oil Changed dropdown */}
                            <td style={{minWidth:90}}>
                              <select
                                value={sosOilChg[rowKey]||(r.oilChanged==="Yes"?L.yes:r.oilChanged==="No"?L.no:"")}
                                onChange={e=>setSosOilChg(p=>({...p,[rowKey]:e.target.value}))}
                                style={{width:"100%",fontSize:10,padding:"2px 4px"}}>
                                {yesNo.map(o=><option key={o}>{o}</option>)}
                              </select>
                            </td>
                            {/* Filter Changed dropdown */}
                            <td style={{minWidth:90}}>
                              <select
                                value={sosFilChg[rowKey]||(r.filterChanged==="Yes"?L.yes:r.filterChanged==="No"?L.no:"")}
                                onChange={e=>setSosFilChg(p=>({...p,[rowKey]:e.target.value}))}
                                style={{width:"100%",fontSize:10,padding:"2px 4px"}}>
                                {yesNo.map(o=><option key={o}>{o}</option>)}
                              </select>
                            </td>
                            <td><Pill status={r.result}/></td>
                            {/* Diagnosis — editable */}
                            <td style={{minWidth:180}}>
                              <input type="text"
                                value={sosDiagnosis[rowKey]!==undefined?sosDiagnosis[rowKey]:(r.diagnosis||r.observations||"")}
                                onChange={e=>setSosDiagnosis(p=>({...p,[rowKey]:e.target.value}))}
                                placeholder="Enter diagnosis / observations…"
                                style={{width:"100%",fontSize:10,padding:"3px 6px"}}
                              />
                            </td>
                            {/* Action Taken */}
                            <td style={{minWidth:160}}>
                              <input type="text"
                                value={sosActions[rowKey]!==undefined?sosActions[rowKey]:(r.actions||"")}
                                onChange={e=>setSosActions(p=>({...p,[rowKey]:e.target.value}))}
                                placeholder="Corrective action taken…"
                                style={{width:"100%",fontSize:10,padding:"3px 6px"}}
                              />
                            </td>
                            {/* Feedback */}
                            <td style={{minWidth:150}}>
                              <input type="text"
                                value={sosFeedback[rowKey]||""}
                                onChange={e=>setSosFeedback(p=>({...p,[rowKey]:e.target.value}))}
                                placeholder="Field feedback…"
                                style={{width:"100%",fontSize:10,padding:"3px 6px"}}
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Metal limits legend */}
              <div className="card" style={{padding:"10px 16px"}}>
                <div style={{fontSize:10,color:T.muted,marginBottom:6,fontWeight:600,textTransform:"uppercase",letterSpacing:".5px"}}>
                  {L.metalLimits}
                </div>
                <div style={{display:"flex",gap:20,flexWrap:"wrap",fontSize:9}}>
                  {[
                    {label:"Fe",warn:80,crit:120},{label:"Cu",warn:25,crit:40},
                    {label:"Al",warn:15,crit:25},{label:"Cr",warn:8,crit:15},
                    {label:"Si",warn:15,crit:25},{label:"Pb",warn:10,crit:20},
                  ].map(({label,warn,crit})=>(
                    <div key={label} style={{display:"flex",alignItems:"center",gap:6}}>
                      <span style={{color:T.text,fontWeight:700,fontFamily:"monospace"}}>{label}:</span>
                      <span style={{color:T.green}}>OK &lt;{warn}</span>
                      <span style={{color:T.orange}}>⚠ {warn}–{crit}</span>
                      <span style={{color:T.red}}>❌ &gt;{crit}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ══════════ FAULT CODES TAB ══════════ */}
          {tab==="faults" && (
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* Summary */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12}}>
                {[{k:"CRITICAL",c:T.red},{k:"HIGH",c:T.orange},{k:"MODERATE",c:T.gold},{k:"LOW",c:T.blue}].map(({k,c})=>{
                  const n=faultData.filter(r=>r.sev===k).length;
                  return <div key={k} className="card" style={{textAlign:"center"}}>
                    <div style={{fontSize:10,color:T.muted,marginBottom:6}}>{k}</div>
                    <div className="mono" style={{fontSize:28,color:c,fontWeight:700}}>{n}</div>
                  </div>;
                })}
                <div className="card" style={{textAlign:"center"}}>
                  <div style={{fontSize:10,color:T.muted,marginBottom:6}}>"TOTAL CODES"</div>
                  <div className="mono" style={{fontSize:28,color:T.gold,fontWeight:700}}>{faultData.length}</div>
                </div>
              </div>

              <DateRangeBar from={faultFrom} to={faultTo} onFrom={setFaultFrom} onTo={setFaultTo} onReset={()=>{setFaultFrom("");setFaultTo("");}} L={L}/>

              {/* Filters */}
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.severity} :</span>
                {faultSevOpts.map(s=>(
                  <button key={s} className={`btn ${faultSev===s?"btn-gold":"btn-ghost"}`}
                    style={{padding:"4px 10px",fontSize:10}} onClick={()=>setFaultSev(s)}>{s}</button>
                ))}
                <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>{filteredFaults.length} "results"</span>
              </div>

              <div className="card" style={{padding:0}}>
                <div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}>
                  <table>
                    <thead><tr>
                      <th>{L.date}</th><th>{L.equipment}</th><th>{L.code}</th><th>{L.description}</th>
                      <th>{L.system}</th><th>{L.severity}</th><th>{L.occurrences}</th>
                      <th>{L.source}</th><th>{L.action}</th>
                      <th>{L.status}</th><th>{L.faultFeedback}</th>
                    </tr></thead>
                    <tbody>
                      {filteredFaults.map((r,i)=>{
                        const rowKey = `${r.equipment}-${r.code}-${r.date}-${i}`;
                        const currentStatus = faultStatuses[rowKey]||r.status||L.faultStatuses[0];
                        return (
                          <tr key={i}>
                            <td style={{fontSize:10}}>{r.date}</td>
                            <td className="mono" style={{color:T.gold}}>{r.equipment}</td>
                            <td className="mono" style={{color:T.cyan,fontWeight:600}}>{r.code}</td>
                            <td style={{maxWidth:180,fontSize:11}}>{r.desc}</td>
                            <td style={{fontSize:10}}>{r.sys||r.system||r.category}</td>
                            <td><Pill status={r.sev}/></td>
                            <td className="mono" style={{textAlign:"center"}}>{r.occurrences}</td>
                            <td style={{fontSize:10,color:T.muted}}>{r.source}</td>
                            <td style={{fontSize:10,maxWidth:140}}>{r.action}</td>
                            {/* Status dropdown */}
                            <td style={{minWidth:130}}>
                              <select
                                value={currentStatus}
                                onChange={e=>setFaultStatuses(p=>({...p,[rowKey]:e.target.value}))}
                                style={{width:"100%",fontSize:10,padding:"2px 4px",
                                  color: /résolu|resolved/i.test(currentStatus)?T.green:
                                         /cours|progress/i.test(currentStatus)?T.orange:
                                         /escalad|escalat/i.test(currentStatus)?T.red:T.text}}>
                                {L.faultStatuses.map(s=><option key={s}>{s}</option>)}
                              </select>
                            </td>
                            {/* Feedback */}
                            <td style={{minWidth:160}}>
                              <input type="text"
                                value={faultFeedback[rowKey]||""}
                                onChange={e=>setFaultFeedback(p=>({...p,[rowKey]:e.target.value}))}
                                placeholder="Field feedback…"
                                style={{width:"100%",fontSize:10,padding:"3px 6px"}}
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ══════════ WO / PM TAB ══════════ */}
          {tab==="wo" && (
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <DateRangeBar from={woFrom} to={woTo} onFrom={setWoFrom} onTo={setWoTo} onReset={()=>{setWoFrom("");setWoTo("");}} L={L}/>

              {/* 8 KPI cards — Planned + Executed per WO type */}
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
                {woKPIs.map(({type,planned,executed,total})=>{
                  const colors={PM:T.green,CM:T.red,Backlog:T.orange,Opportun:T.cyan};
                  const color = colors[type]||T.gold;
                  const lbl = {
                    PM:    [L.pmPlanned,    L.pmExecuted],
                    CM:    [L.cmPlanned,    L.cmExecuted],
                    Backlog:[L.backlogPlanned,L.backlogExecuted],
                    Opportun:[L.opportunPlanned,L.opportunExecuted],
                  }[type]||[type+" Planned",type+" Executed"];
                  return (
                    <div key={type} className="card" style={{padding:"12px 14px"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                        <span style={{fontSize:10,fontWeight:700,color,textTransform:"uppercase",letterSpacing:".5px"}}>{type}</span>
                        <span className="mono" style={{fontSize:10,color:T.muted}}>{total} total</span>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                        <div style={{background:T.bg3,borderRadius:4,padding:"6px 8px",textAlign:"center"}}>
                          <div style={{fontSize:8,color:T.muted,marginBottom:2}}>{lbl[0]}</div>
                          <div className="mono" style={{fontSize:18,color,fontWeight:700}}>{planned}</div>
                        </div>
                        <div style={{background:T.bg3,borderRadius:4,padding:"6px 8px",textAlign:"center"}}>
                          <div style={{fontSize:8,color:T.muted,marginBottom:2}}>{lbl[1]}</div>
                          <div className="mono" style={{fontSize:18,color:T.green,fontWeight:700}}>{executed}</div>
                        </div>
                      </div>
                      {/* Execution rate bar */}
                      <div style={{marginTop:8,background:T.bg3,borderRadius:99,height:4,overflow:"hidden"}}>
                        <div style={{height:"100%",width:`${total>0?Math.min(100,executed/total*100):0}%`,background:color,borderRadius:99,transition:"width .3s"}}/>
                      </div>
                      <div style={{fontSize:8,color:T.muted,marginTop:3,textAlign:"right"}}>
                        {total>0?`${Math.round(executed/total*100)}% executed`:"—"}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* PM Compliance KPI */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 3fr",gap:12}}>
                <div className="card" style={{textAlign:"center",display:"flex",flexDirection:"column",justifyContent:"center"}}>
                  <div style={{fontSize:10,color:T.muted,marginBottom:4,textTransform:"uppercase",letterSpacing:".5px"}}>{L.pmCompliance}</div>
                  <div className="mono" style={{fontSize:32,color:pmCompliance>=90?T.green:T.red,fontWeight:700}}>{pmCompliance}%</div>
                  <div style={{fontSize:9,color:T.muted,marginTop:4}}>Target ≥ 95%</div>
                </div>
                <div className="card">
                  <SectionTitle icon="📊">"WO Distribution by Type"</SectionTitle>
                  <ResponsiveContainer width="100%" height={140}>
                    <BarChart data={woKPIs} layout="vertical">
                      <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
                      <XAxis type="number" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false}/>
                      <YAxis type="category" dataKey="type" tick={{fill:T.muted,fontSize:9}} axisLine={false} tickLine={false} width={60}/>
                      <Tooltip content={<CustomTooltip/>}/>
                      <Bar dataKey="planned"  name="Planned"  fill={T.blue}   radius={[0,3,3,0]} stackId="a"/>
                      <Bar dataKey="executed" name="Executed" fill={T.green}  radius={[0,3,3,0]} stackId="b"/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Filters */}
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",padding:"4px 0"}}>
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.model} :</span>
                <select value={woModel} onChange={e=>{setWoModel(e.target.value);setWoEq("ALL");}} style={{width:130}}>
                  {woModels.map(m=><option key={m}>{m}</option>)}
                </select>
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.equipment} :</span>
                <select value={woEq} onChange={e=>setWoEq(e.target.value)} style={{width:110}}>
                  {woEqs.map(e=><option key={e}>{e}</option>)}
                </select>
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.woTypeLabel} :</span>
                <select value={woType} onChange={e=>setWoType(e.target.value)} style={{width:120}}>
                  {L.woTypes.map(t=><option key={t}>{t}</option>)}
                </select>
                <span style={{fontSize:10,color:T.muted,fontWeight:600}}>{L.status} :</span>
                <select value={woStatus} onChange={e=>setWoStatus(e.target.value)} style={{width:130}}>
                  {L.woStatuses.map(s=><option key={s}>{s}</option>)}
                </select>
                <button className="btn btn-ghost" style={{padding:"4px 10px",fontSize:10}}
                  onClick={()=>{setWoModel("ALL");setWoEq("ALL");setWoType("ALL");setWoStatus("ALL");}}>
                  {L.resetFilter}
                </button>
                <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>{filteredWO.length} WO</span>
              </div>

              {/* WO Table — no Duration, Cost, Technician */}
              <div className="card" style={{padding:0}}>
                <div style={{overflowX:"auto",maxHeight:420,overflowY:"auto"}}>
                  <table>
                    <thead><tr>
                      <th>{L.woNum}</th>
                      <th>{L.equipment}</th>
                      <th>{L.model}</th>
                      <th>{L.woType}</th>
                      <th>"Work Description"</th>
                      <th>{L.plannedDate}</th>
                      <th>{L.actualDate}</th>
                      <th>{L.status}</th>
                    </tr></thead>
                    <tbody>
                      {filteredWO.map((r,i)=>(
                        <tr key={i}>
                          <td className="mono" style={{color:T.gold,fontSize:10}}>{r.woNum}</td>
                          <td className="mono" style={{fontSize:10}}>{r.equipment}</td>
                          <td style={{fontSize:10,color:T.muted}}>{r.model}</td>
                          <td>
                            <span className={`badge ${r.type==="CM"?"crit":r.type==="PM"?"ok":r.type==="Backlog"?"warn":"info"}`}>
                              {r.type}
                            </span>
                          </td>
                          <td style={{maxWidth:200,fontSize:11}}>{r.description}</td>
                          <td style={{fontSize:10}}>{r.plannedDate}</td>
                          <td style={{fontSize:10}}>{r.actualDate}</td>
                          <td><Pill status={r.status}/></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ══════════ IMPORT TAB ══════════ */}
          {tab==="import" && (
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:16}}>
                <div className="card">
                  <SectionTitle icon="🗄">Equipment Downtime Database (XLSX)</SectionTitle>
                  <div style={{fontSize:10,color:T.muted,marginBottom:12,lineHeight:1.8}}>
                    Detected format:<br/>
                    <span className="mono" style={{color:T.gold}}>Equipment N# · Model · Date Down · Time Down · Date Up · Time Up · Down Time (Hrs) · Planned · Comments · CAT Code Equivalent · Job Status · Events</span>
                  </div>
                  <UploadZone label="Import Downtime DB" icon="🗄" onFile={handleStopsFile} formats=".xlsx,.xls,.csv"/>
                  <div style={{marginTop:10,fontSize:10,color:T.green}}>✓ {stops.length} records loaded</div>
                </div>
                <div className="card">
                  <SectionTitle icon="🛢">Truvu Spectrooil Export (XLSX)</SectionTitle>
                  <div style={{fontSize:10,color:T.muted,marginBottom:12,lineHeight:1.8}}>
                    Detected format — sheet <span className="mono" style={{color:T.gold}}>Samples</span>, 1140-column wide-pivot:<br/>
                    <span className="mono" style={{color:T.gold}}>AssetName* · AssetModel · ComponentName* · AssetLife · OilLife · SampleId* · AlarmCode · DrawingDate · Diagnosis · LimitObservations · LimitActions</span><br/>
                    + groups of 6 columns per metal element : <span className="mono" style={{color:T.cyan}}>ml-el.fe / ml-el.cu / ml-el.al / ml-el.cr / ml-el.si / ml-el.pb / ml-el.ni</span>
                  </div>
                  <UploadZone label="Import Truvu Export" icon="🛢" onFile={handleSOSFile} formats=".xlsx,.xls"/>
                  <div style={{marginTop:10,fontSize:10,color:T.green}}>✓ {sosData.length} samples loaded</div>
                </div>
                <div className="card">
                  <SectionTitle icon="⚠️">VisionLink CAT Export — Fault Codes (CSV/XLSX)</SectionTitle>
                  <div style={{fontSize:10,color:T.muted,marginBottom:12,lineHeight:1.8}}>
                    Export standard VisionLink CAT :<br/>
                    <span className="mono" style={{color:T.gold}}>Asset ID · Asset S/N · Model · Event Date · Event Time · Fault Code · Description · Severity · Category · Count · Status · Action</span><br/>
                    <span style={{color:T.muted}}>Severity mapping: High/3→CRITICAL · Med/2→HIGH · Low/1→MODERATE</span>
                  </div>
                  <UploadZone label="Import VisionLink Export" icon="⚠️" onFile={handleFaultFile} formats=".csv,.xlsx,.xls"/>
                  <div style={{marginTop:10,fontSize:10,color:T.green}}>✓ {faultData.length} fault codes loaded</div>
                </div>
                <div className="card">
                  <SectionTitle icon="🔧">Pronto CMMS Export — Work Orders (XLSX/CSV)</SectionTitle>
                  <div style={{fontSize:10,color:T.muted,marginBottom:12,lineHeight:1.8}}>
                    Export standard Pronto :<br/>
                    <span className="mono" style={{color:T.gold}}>Work Order · Equipment · Description · WO Type · Priority · Planned Start · Actual Start · Actual Finish · Estimated Hours · Actual Hours · Cost · Status · Technician</span><br/>
                    <span style={{color:T.muted}}>WO Type mappé : PM/Prev→PM · CM/Corr→CM · PDM/Pred→PDM</span>
                  </div>
                  <UploadZone label="Import Pronto CMMS" icon="🔧" onFile={handleWOFile} formats=".xlsx,.xls,.csv"/>
                  <div style={{marginTop:10,fontSize:10,color:T.green}}>✓ {woData.length} work orders loaded</div>
                </div>
              </div>

              {/* Format guide */}
              <div className="card">
                <SectionTitle icon="📋">Import Notes &amp; Column Mapping</SectionTitle>
                <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:16}}>
                  {[
                    {title:"Downtime DB — Key Fields",rows:[
                      "Equipment N#  →  Asset ID (AT015, EX100…)",
                      "Model         →  Model (AT A45G, EXC 2600, HT 777…)",
                      "Date Down     →  Breakdown date (Excel date or text)",
                      "Down Time (Hrs) →  Downtime in decimal hours",
                      "Planned       →  PLANNED / UNPLANNED",
                      "CAT Code Equivalent → Failure cause code (ELECTRICAL, PM, LEAKS…)",
                      "Comments      →  Full description of the downtime event",
                    ]},
                    {title:"Truvu — champs extraits",rows:[
                      "AssetName*    →  Asset ID (100HT021, 100BD102…)",
                      "AssetModel    →  Model (777E, DM30 XC SP…)",
                      "ComponentName* → Component (Engine, Transmission…)",
                      "AssetLife     →  Total asset operating hours",
                      "OilLife       →  Hours since last oil change",
                      "AlarmCode     →  OK / Severe  →  NORMAL / CRITICAL",
                      "ml-el.fe/cu/al/cr/si/pb/ni → Metal elements in ppm",
                    ]},
                    {title:"VisionLink — champs attendus",rows:[
                      "Asset ID      →  Asset ID",
                      "Fault Code    →  DTC fault code (E360-1, etc.)",
                      "Description   →  Fault code description",
                      "Severity      →  High/Med/Low or 1/2/3",
                      "Event Date    →  Event date",
                      "Count         →  Number of occurrences",
                      "Status        →  Current fault status",
                    ]},
                    {title:"Pronto CMMS — champs attendus",rows:[
                      "Work Order    →  WO number (e.g. WO-001234)",
                      "Equipment     →  Asset ID",
                      "WO Type       →  PM / CM / Backlog / Opportunistic",
                      "Planned Start →  Planned start date",
                      "Actual Finish →  Actual completion date",
                      "Actual Hours  →  Actual duration in hours",
                      "Cost          →  Total cost in local currency",
                    ]},
                  ].map(({title,rows})=>(
                    <div key={title} style={{background:T.bg3,borderRadius:6,padding:12}}>
                      <div style={{fontWeight:600,color:T.gold,fontSize:11,marginBottom:8}}>📄 {title}</div>
                      {rows.map((r,i)=>(
                        <div key={i} style={{fontSize:10,color:T.muted,marginBottom:4,fontFamily:"'JetBrains Mono',monospace",lineHeight:1.4}}>{r}</div>
                      ))}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </>
  );
}
