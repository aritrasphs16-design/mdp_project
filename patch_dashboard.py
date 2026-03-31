import re

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Inject the generator call
inject_gen = """            import advanced_columns_generator as acg
            
            # Base data columns
            for c in base_data:
                if "_ext" not in c: c["_ext"] = {}
                c["_ext"]["_adv"] = acg.generate_ultra_advanced_columns(c)
"""

content = content.replace("            if base_data:", inject_gen + "            if base_data:")

# Also inject into the synthetic generation output
inject_synth = """                        ext["Synthetic_Difficulty"] = random.choice(["Easy", "Moderate", "Hard"])
                        
                    ext["_adv"] = acg.generate_ultra_advanced_columns(new_c)
"""
content = content.replace('ext["Synthetic_Difficulty"] = random.choice(["Easy", "Moderate", "Hard"])', inject_synth)


# 2. Add the `_adv` dict into the JS output array
inject_js_row = """            "ppb": ext.get("Plasma_Protein_Binding", "N/A"),
            "clearance": ext.get("Clearance", "N/A"),
            "half_life": ext.get("Half_Life", "N/A"),
            **ext.get("_adv", {})
        })"""
content = content.replace('            "ppb": ext.get("Plasma_Protein_Binding", "N/A"),\n            "clearance": ext.get("Clearance", "N/A"),\n            "half_life": ext.get("Half_Life", "N/A")\n        })', inject_js_row)

# 3. Add column group buttons to the UI
ui_inject = """  </div>
  <div id="col-groups" style="display:flex;gap:6px;padding:6px 14px;background:#090d18;border-top:1px solid rgba(232,160,32,0.1);flex-wrap:wrap;">
    <span style="color:rgba(200,222,255,0.4);font-size:0.5rem;letter-spacing:1px;align-self:center;margin-right:10px;">TOGGLE GROUPS:</span>
    <button class="grp-btn active" onclick="toggleGrp('core')">Core Metrics</button>
    <button class="grp-btn" onclick="toggleGrp('phys')">Physicochem</button>
    <button class="grp-btn" onclick="toggleGrp('adme')">ADME</button>
    <button class="grp-btn" onclick="toggleGrp('metab')">Metabolism</button>
    <button class="grp-btn" onclick="toggleGrp('tox')">Toxicity</button>
    <button class="grp-btn" onclick="toggleGrp('synth')">Synthesis & Scale</button>
    <button class="grp-btn" onclick="toggleGrp('bio')">BioActivity</button>
    <button class="grp-btn" onclick="toggleGrp('ai')">AI Models</button>
  </div>
</div>

<style>
.grp-btn { background:rgba(232,160,32,0.05); border:1px solid rgba(232,160,32,0.2); color:#e8a020; padding:3px 10px; border-radius:4px; font-size:0.5rem; cursor:pointer; font-family:'JetBrains Mono',monospace; letter-spacing:0.5px; transition:all 0.2s;}
.grp-btn:hover { background:rgba(232,160,32,0.1); }
.grp-btn.active { background:rgba(232,160,32,0.15); box-shadow:0 0 8px rgba(232,160,32,0.3); border-color:rgba(232,160,32,0.5); }
</style>
"""

content = content.replace('    </div>\n  </div>\n</div>\n\n<script>', ui_inject + '\n<script>')

# 4. Modify COLS and buildBody in JS
js_rewrite = """
var ADV_DEF = {
  "phys": ["Lipinski_Score","Veber_Rule","Ghose_Filter","Muegge_Drug_Likeness","Lead_Likeness_Index","Fragment_Like_Score","Molecular_Flexibility","Polar_Surface_Bal","Hydrophobicity_Bal","Arom_Ring_Count","Aliphatic_Ring_Count","Rot_Bond_Stress","H_Bond_Saturation","Heteroatom_Density","Carbon_Fraction","Molec_Shape_Index","Chirality_Count","Polarity_Index","Structural_Diversity","Scaffold_Novelty"],
  "adme": ["Human_Intest_Absorp","Caco2_Perm","MDCK_Perm","BBB_Penetration","Plasma_Prot_Binding","Oral_Bioavail_Pred","Bioavail_Radar","Hepatic_Uptake","Renal_Clearance","GI_Absorption","Tissue_Dist_Index","Skin_Perm","Lung_Penetration","CNS_Exposure_Prob","Absorp_Rate_Est","Dist_Vol_Pred","Membrane_Diff","Passive_Perm_Score","Active_Transport","Drug_Transporter_Int"],
  "metab": ["CYP1A2_Inhib","CYP2C9_Inhib","CYP2C19_Inhib","CYP2D6_Inhib","CYP3A4_Inhib","CYP_Enz_Stability","Microsomal_Stab","Phase_I_Metab","Phase_II_Metab","Metabolic_Hotspots","Metab_Half_Life","Liver_Clearance_Risk","Enzyme_Bind_Strength","Oxidation_Suscept","Hydrolysis_Suscept","Glucuronidation_Pot","Sulfation_Pot","Metabolite_Tox","Reactive_Metabolite","Enzyme_Interact_Idx"],
  "tox": ["hERG_Cardiotox","Mutagenicity","Carcinogenicity","Hepatotoxicity","Nephrotoxicity","Neurotoxicity","Skin_Sensitization","Resp_Tox","Repro_Tox","Devel_Tox","Cytotoxicity","LD50_Estimate","DILI_Risk","Genotoxicity","Teratogenicity","Reactive_Func_Grp","PAINS_Alert","Toxicophore_Alert","Off_Target_Tox","Safety_Margin"],
  "synth": ["Synth_Access_Score","Reaction_Complexity","Synth_Route_Steps","BB_Availability","Scaffold_Complexity","Func_Grp_Diversity","Protecting_Grp_Req","Stereochem_Diff","Reaction_Yield_Est","Ind_Scalability","Reagent_Cost_Est","Lab_Feasibility","Synth_Time_Est","Automation_Compat","Retrosynth_Conf","Reaction_Risk","Process_Chem_Diff","Chem_Stability","Shelf_Life_Pred","Degradation_Risk"],
  "bio": ["Tgt_Bind_Prob","Docking_Affinity","Binding_Pocket_Fit","Ligand_Efficiency","Lipophilic_Lig_Eff","Binding_Selectivity","Protein_Interact_Sc","Binding_Stability","Off_Target_Bind","Pharmacophore_Match","Binding_Pose_Conf","Mol_Interact_Count","H_Bond_Interact","Hydrophobic_Interact","Electrostatic_Inter"],
  "ai": ["AI_Druglikeness_Conf","AI_Tox_Probability","AI_Metabolism_Pred","AI_Target_Affinity","AI_Opt_Potential","AI_Novelty_Score","AI_Synthesizability","AI_Selectivity_Pred","AI_Property_Fit","AI_Clinical_Risk"]
};

// Auto-build the remaining COLS
Object.keys(ADV_DEF).forEach(k => {
  ADV_DEF[k].forEach(col => {
    COLS.push({k:col, l:col.replace(/_/g," "), grp:k, hide:true});
  });
});

var activeGrp = {core:true};

function toggleGrp(grp) {
  if (grp === 'core') return; // Cannot hide core
  activeGrp[grp] = !activeGrp[grp];
  document.querySelectorAll('.grp-btn').forEach(b => {
    var txt = b.textContent.toLowerCase();
    if (txt.includes(grp) || txt.includes(grp.substring(0,3))) {
       if (activeGrp[grp]) b.classList.add('active'); else b.classList.remove('active');
    }
  });
  
  // update hides
  COLS.forEach(c => {
    if(c.grp) { c.hide = !activeGrp[c.grp]; }
  });
  
  // redraw headers and body
  var hm = document.getElementById("main-hr");
  var bm = document.getElementById("main-body");
  if(hm) buildHdr(hm);
  if(bm) buildBody(bm, cur);
  
  var fh = document.getElementById("fs-hr");
  var fb = document.getElementById("fs-body-rows");
  if(fh && fh.children.length>0) { buildHdr(fh); buildBody(fb, cur); }
}

function genericColor(v) {
  var s = String(v).toLowerCase();
  if(s.includes("fail") || s.includes("high risk") || s.includes("positive") || s.includes("poor") || s.includes("tox") || s.includes("warning") || s.includes("hazard") || s.includes("alert")) return "#f87171";
  if(s.includes("pass") || s.includes("safe") || s.includes("low risk") || s.includes("excellent") || s.includes("optimal") || s.includes("negative") || s.includes("good")) return "#34d399";
  if(s.includes("moderate") || s.includes("borderline") || s.includes("caution") || s.includes("weak")) return "#fbbf24";
  return "rgba(200,222,255,0.7)";
}

// Modify existing buildHdr
function buildHdr(tr){
  tr.innerHTML="";
  COLS.forEach(function(c){
    if(c.hide) return;
    var th=document.createElement("th");
    th.setAttribute("data-k",c.k);
    th.textContent=c.l;
    if(c.k===sortCol) th.classList.add(sortDir===-1?"sdsc":"sasc");
    th.addEventListener("click",function(){doSort(c.k);});
    tr.appendChild(th);
  });
}
"""

content = re.sub(r'var cur = ROWS\.slice\(\);\s*var COLS = \[.*?\];', 
                 r'var cur = ROWS.slice();\nvar COLS = [' + 
                 content.split('var COLS = [')[1].split('];')[0] + '];\n' + js_rewrite, 
                 content, flags=re.DOTALL)

# Modify buildBody to render dynamic columns smoothly
body_rewrite = """  rows.forEach(function(d,i){
    var dl=(i*0.005).toFixed(3); // Fast staggered animation for 200 rows
    h+='<tr style="animation-delay:'+dl+'s">';
    
    COLS.forEach(function(c) {
      if(c.hide) return;
      var v = d[c.k];
      
      // Core specific renders
      if(c.k==="idx") { h+='<td style="color:rgba(200,222,255,.2);font-size:.6rem">'+(i+1)+'</td>'; return; }
      if(c.k==="id") { h+='<td style="color:#e8a020;font-weight:500">'+v+'</td>'; return; }
      if(c.k==="grade") { h+='<td><span class="gr gr'+v+'">'+v+'</span></td>'; return; }
      if(c.k==="dl_badge") { h+='<td style="color:'+d.dl_color+'"><span style="border:1px solid '+d.dl_color+'40;background:'+d.dl_color+'11;padding:2px 6px;border-radius:10px;font-size:0.55rem;letter-spacing:0.5px">'+v+'</span></td>'; return; }
      
      if(["lead","oral","qed","np","stress","prom","lig_eff","frag_eff","lip_eff","bio_score"].includes(c.k)) { h+='<td>'+bar(v,c.k)+'</td>'; return; }
      if(["mw","logp","tpsa","fsp3","cyp","sim","logs","ext_logs","cns"].includes(c.k)) { h+='<td style="color:'+txtC(v,c.k)+'">'+v+'</td>'; return; }
      if(c.k==="sa") { h+='<td style="color:'+txtC(v,"sa")+'">'+v+' <span style="font-size:.58rem;opacity:.5">('+d.sa_lbl+')</span></td>'; return; }
      if(c.k==="herg") { h+='<td style="'+hS(v)+'">'+v+'</td>'; return; }
      if(c.k==="ames") { h+='<td style="color:'+aC(v)+';font-size:.64rem">'+v+'</td>'; return; }
      if(["hia","bbb_p"].includes(c.k)) { 
          if(typeof v === "boolean") { h+='<td class="'+(v?"bok":"bno")+'">'+(v?"✓":"✗")+'</td>'; return; }
      }
      
      // Advanced columns auto-color
      if(c.grp) {
          var colHex = genericColor(v);
          h+='<td style="color:'+colHex+'">'+v+'</td>';
          return;
      }
      
      // Fallback
      h+='<td style="color:rgba(200,222,255,0.7)">'+v+'</td>';
    });
    
    h+='</tr>';
  });"""

content = re.sub(r'rows\.forEach\(function\(d\,i\)\{\s*var dl\=\(i\*0\.005\)\.toFixed\(3\);\s*h\+\=\'\<tr style\=\".*?\}\)\;', 
                 body_rewrite, 
                 content, flags=re.DOTALL)


with open("app.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Path applied successfully!")
