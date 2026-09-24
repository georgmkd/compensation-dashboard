(() => {
  const years = [2017,2018,2019,2020,2021,2022,2023,2024,2025,2026];
  const cpi = {
    "North Macedonia":[1.4,1.5,.8,1.2,3.2,14.2,9.4,3.5,4.1,3.0],
    "Uruguay":[6.2,7.6,7.9,9.8,7.7,9.1,5.9,4.8,4.5,4.5],
    "Argentina":[25.7,34.3,53.5,42,48.4,72.4,133.5,219.9,41.3,17.5],
    "Brazil":[3.4,3.7,3.7,3.2,8.3,9.3,4.6,4.4,5.2,4.2],
    "Netherlands":[1.4,1.7,2.6,1.3,2.7,10,3.8,3.3,3.3,2.3],
    "USA":[2.1,2.4,1.8,1.2,4.7,8,4.1,2.9,2.7,2.4]
  };

  let raw = [], people = [];
  const $ = id => document.getElementById(id);

  function normCountry(v) {
    const s = String(v || "").trim(), l = s.toLowerCase();
    if (["north macedonia","macedonia","mk","mkd"].includes(l)) return "North Macedonia";
    if (["usa","us","united states","united states of america"].includes(l)) return "USA";
    if (l === "uruguay") return "Uruguay";
    if (l === "argentina") return "Argentina";
    if (l === "brazil" || l === "brasil") return "Brazil";
    if (["netherlands","the netherlands","holland"].includes(l)) return "Netherlands";
    return s;
  }

  function factor(country, from, to) {
    const arr = cpi[country];
    if (!arr) return 1;
    let f = 1;
    for (let y = from + 1; y <= to; y++) {
      const idx = years.indexOf(y);
      if (idx >= 0) f *= 1 + arr[idx] / 100;
    }
    return f;
  }

  function money(v, c) {
    try {
      return new Intl.NumberFormat("en-US", {
        style:"currency", currency:c || "USD", maximumFractionDigits:0
      }).format(v);
    } catch {
      return `${c || ""} ${Math.round(v).toLocaleString()}`;
    }
  }

  function esc(v) {
    return String(v ?? "").replace(/[&<>"']/g, ch => ({
      "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
    }[ch]));
  }

  function parseInputDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }
    if (typeof value === "number" && window.XLSX && XLSX.SSF) {
      const d = XLSX.SSF.parse_date_code(value);
      if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d));
    }
    const s = String(value ?? "").trim();
    if (!s) return null;
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
    m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/);
    if (m) return new Date(Date.UTC(+m[3], +m[2]-1, +m[1]));
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? null : new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }

  function formatDate(iso) {
    const [y,m,d] = iso.split("-").map(Number);
    return new Intl.DateTimeFormat("en-GB",{day:"2-digit",month:"short",year:"numeric",timeZone:"UTC"})
      .format(new Date(Date.UTC(y,m-1,d)));
  }

  function analyze() {
    const map = new Map();
    raw.forEach(r => {
      const name = String(r.name || "").trim();
      const country = normCountry(r.country);
      const dateRaw = r.date;
      const parsedDate = parseInputDate(dateRaw);
      const year = parsedDate ? parsedDate.getUTCFullYear() : NaN;
      const comp = Number(String(r.compensation).replace(/[^0-9.-]/g,""));
      const currency = String(r.currency || "USD").trim().toUpperCase();
      if (!name || !country || !parsedDate || !Number.isFinite(year) || !Number.isFinite(comp) || !currency) return;

      // Name + country + currency is treated as one employee history.
      const key = `${name.toLowerCase()}|${country.toLowerCase()}|${currency}`;
      if (!map.has(key)) map.set(key, {name,country,currency,entries:[]});
      map.get(key).entries.push({year,date:parsedDate.toISOString().slice(0,10),comp});
    });

    people = [...map.values()].map(p => {
      p.entries.sort((a,b) => a.date.localeCompare(b.date));
      const latest = p.entries[p.entries.length-1];
      const latestYear = latest.year;
      const today = new Date();
      const currentYear = today.getFullYear();
      const comparisonYear = Math.min(currentYear, years[years.length - 1]);
      const normalized = p.entries.map(e => ({
        ...e, equiv:e.comp * factor(p.country,e.year,comparisonYear)
      }));
      const strongest = [...normalized].sort((a,b) =>
        (b.equiv-a.equiv) || b.date.localeCompare(a.date)
      )[0];
      const target = strongest.equiv;
      const currentEquivalent = latest.comp * factor(p.country,latest.year,comparisonYear);
      const gap = Math.max(0,target-currentEquivalent);
      const raise = currentEquivalent > 0 ? gap/currentEquivalent*100 : 0;
      const latestDate = new Date(latest.date + "T00:00:00");
      const ageDays = Math.floor((today - latestDate) / 86400000);
      const status = latestDate > today ? "Future entry" :
        latest.year < currentYear ? "Update due" :
        latest.year === currentYear ? "Current" : "Future entry";
      return {...p,latest,latestYear,currentYear,comparisonYear,currentEquivalent,
        normalized,strongest,target,gap,raise,needs:gap>.5,status,ageDays};
    }).sort((a,b) => b.raise-a.raise);

    refreshFilters();
    render();
  }

  function refreshFilters() {
    const old = $("countryFilter").value;
    const countries = [...new Set(people.map(p=>p.country))].sort();
    $("countryFilter").innerHTML =
      '<option value="ALL">All countries</option>' +
      countries.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join("");
    if (countries.includes(old)) $("countryFilter").value = old;
  }

  function filtered() {
    const q = $("search").value.trim().toLowerCase();
    const cf = $("countryFilter").value, nf = $("needFilter").value;
    return people.filter(p =>
      (!q || p.name.toLowerCase().includes(q)) &&
      (cf==="ALL" || p.country===cf) &&
      (nf==="ALL" || (nf==="NEED" && p.needs) || (nf==="OK" && !p.needs))
    );
  }

  function render() {
    const list = filtered(), need = list.filter(p=>p.needs);
    $("employeeCount").textContent = list.length;
    $("needCount").textContent = need.length;
    $("avgRaise").textContent = need.length
      ? (need.reduce((s,p)=>s+p.raise,0)/need.length).toFixed(1)+"%" : "—";
    $("largestGap").textContent = need.length ? need[0].raise.toFixed(1)+"%" : "—";

    $("tbody").innerHTML = list.length ? list.map(p => `
      <tr>
        <td><button class="person" data-index="${people.indexOf(p)}">${esc(p.name)}</button>
          <small>${p.entries.length} entr${p.entries.length===1?"y":"ies"}</small></td>
        <td>${esc(p.country)}</td>
        <td>${money(p.latest.comp,p.currency)}<small>${formatDate(p.latest.date)}</small></td>
        <td>${formatDate(p.strongest.date)}<small>${money(p.strongest.comp,p.currency)}</small></td>
        <td><strong>${money(p.target,p.currency)}</strong></td>
        <td>${p.needs ? money(p.gap,p.currency) : "—"}</td>
        <td><strong>${p.needs ? "+"+p.raise.toFixed(1)+"%" : "0%"}</strong></td>
        <td><span class="pill ${p.status==="Current"?"ok":"warning"}">${p.status}</span>
          <small>${p.status==="Update due" ? "Latest: "+formatDate(p.latest.date) : (p.needs ? "Raise needed" : "At / above target")}</small></td>
      </tr>`).join("") :
      '<tr><td colspan="8" class="empty">Upload a spreadsheet or load demo data.</td></tr>';

    document.querySelectorAll(".person").forEach(b =>
      b.addEventListener("click",()=>showDetail(people[+b.dataset.index]))
    );
  }

  function showDetail(p) {
    $("detail").classList.remove("hidden");
    $("detailName").textContent = p.name;
    $("detailMeta").textContent = `${p.country} • ${p.currency} • ${p.entries.length} compensation entries`;
    $("detailCards").innerHTML = `
      <article class="metric"><span>Current salary</span><strong>${money(p.latest.comp,p.currency)}</strong></article>
      <article class="metric target"><span>Equal-comp target</span><strong>${money(p.target,p.currency)}</strong></article>
      <article class="metric"><span>Required adjustment</span><strong>${p.needs ? "+"+money(p.gap,p.currency)+" / +"+p.raise.toFixed(1)+"%" : "No raise required"}</strong></article>`;
    const max = Math.max(...p.normalized.map(e=>e.equiv));
    $("history").innerHTML = p.normalized.map(e => {
      const strength = max ? e.equiv/max*100 : 0;
      return `<tr><td>${formatDate(e.date)}${e.date===p.strongest.date?" ★":""}</td>
        <td>${money(e.comp,p.currency)}</td><td>${money(e.equiv,p.currency)}</td>
        <td><div class="strength"><div class="bar"><span style="width:${strength}%"></span></div>${strength.toFixed(1)}%</div></td></tr>`;
    }).join("");
    $("detail").scrollIntoView({behavior:"smooth",block:"start"});
  }

  function parseCSV(text) {
    const lines = text.replace(/^\uFEFF/,"").split(/\r?\n/).filter(x=>x.trim());
    if (!lines.length) return [];
    function cols(line) {
      const out=[]; let s="", q=false;
      for (let i=0;i<line.length;i++) {
        const ch=line[i];
        if (ch==='"') {
          if (q && line[i+1]==='"') { s+='"'; i++; } else q=!q;
        } else if (ch==="," && !q) { out.push(s); s=""; }
        else s+=ch;
      }
      out.push(s); return out;
    }
    const h=cols(lines[0]);
    return lines.slice(1).map(l=>{
      const v=cols(l),o={}; h.forEach((k,i)=>o[k]=v[i]??""); return o;
    });
  }

  async function parseFile(file) {
    $("status").textContent = `Reading ${file.name}…`;
    try {
      let rows;
      if (file.name.toLowerCase().endsWith(".csv")) {
        rows=parseCSV(await file.text());
      } else {
        if (!window.XLSX) throw new Error("Excel reader failed to load. Try CSV instead.");
        const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        rows=XLSX.utils.sheet_to_json(ws,{defval:""});
      }
      const normalized=rows.map(row=>{
        const o={}; Object.keys(row).forEach(k=>o[String(k).trim().toLowerCase()]=row[k]); return o;
      });
      const required=["name","country","date","compensation","currency"];
      const missing=required.filter(k=>!normalized.length || !(k in normalized[0]));
      if (missing.length) throw new Error("Missing columns: "+missing.join(", "));
      raw=normalized; analyze();
      $("status").textContent=`Loaded ${raw.length} rows and matched ${people.length} employee histories.`;
    } catch(e) {
      $("status").textContent=e.message || "Could not read file.";
    }
  }

  $("choose").addEventListener("click",()=>$("file").click());
  $("file").addEventListener("change",e=>e.target.files[0]&&parseFile(e.target.files[0]));
  const drop=$("drop");
  ["dragenter","dragover"].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.add("drag")}));
  ["dragleave","drop"].forEach(t=>drop.addEventListener(t,e=>{e.preventDefault();drop.classList.remove("drag")}));
  drop.addEventListener("drop",e=>{const f=e.dataTransfer.files[0];if(f)parseFile(f)});
  ["countryFilter","needFilter"].forEach(id=>$(id).addEventListener("change",render));
  $("search").addEventListener("input",render);
  $("closeDetail").addEventListener("click",()=>$("detail").classList.add("hidden"));
  $("demo").addEventListener("click",()=>{
    raw=[
      {name:"Michael Anderson",country:"North Macedonia",date:"2022-01-01",compensation:4100,currency:"USD"},
      {name:"Michael Anderson",country:"North Macedonia",date:"2024-01-01",compensation:4800,currency:"USD"},
      {name:"Michael Anderson",country:"North Macedonia",date:"2026-01-01",compensation:4800,currency:"USD"},
      {name:"John Smith",country:"USA",date:"2021-01-01",compensation:6000,currency:"USD"},
      {name:"John Smith",country:"USA",date:"2024-01-01",compensation:6900,currency:"USD"},
      {name:"John Smith",country:"USA",date:"2026-01-01",compensation:7000,currency:"USD"},
      {name:"Anna de Vries",country:"Netherlands",date:"2022-01-01",compensation:4800,currency:"EUR"},
      {name:"Anna de Vries",country:"Netherlands",date:"2024-01-01",compensation:5400,currency:"EUR"},
      {name:"Anna de Vries",country:"Netherlands",date:"2026-01-01",compensation:5600,currency:"EUR"}
    ];
    $("status").textContent="Demo data loaded.";
    analyze();
  });
  render();
})();