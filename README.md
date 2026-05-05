<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <title>SAARC KPIs (Tanmoy Ghosh)</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f7f6; margin: 0; padding: 20px; }
        .container { max-width: 1900px; margin: auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 5px 20px rgba(0,0,0,0.1); }
        h2 { text-align: center; color: #333; margin-bottom: 20px; text-transform: uppercase; letter-spacing: 1px; }
        .upload-area { border: 2px dashed #d32f2f; padding: 25px; text-align: center; border-radius: 10px; background: #fff5f5; cursor: pointer; margin-bottom: 25px; }
        
        .action-bar { display: flex; justify-content: flex-end; margin-bottom: 15px; }
        .btn-download { background: #27ae60; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; }

        .stats-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 10px; margin-bottom: 25px; }
        .card { padding: 12px; border-radius: 8px; color: white; text-align: center; }
        .bg-1 { background: #2c3e50; } .bg-2 { background: #e67e22; } .bg-3 { background: #27ae60; } .bg-4 { background: #2980b9; } .bg-6 { background: #16a085; } .bg-7 { background: #d35400; }

        .table-wrapper { overflow-x: auto; border-radius: 8px; border: 1px solid #ddd; }
        table { width: 100%; border-collapse: collapse; background: white; min-width: 1850px; }
        th, td { padding: 10px 6px; border: 1px solid #eee; text-align: center; font-size: 11px; }
        th { background: #d32f2f; color: white; position: sticky; top: 0; z-index: 10; }
        .area-header { background: #222 !important; color: white; text-align: left; padding-left: 15px; font-weight: bold; font-size: 13px; }
        
        .score-box { font-weight: bold; padding: 6px 10px; border-radius: 4px; color: white; font-size: 13px; min-width: 45px; display: inline-block; }
        .score-green { background: #27ae60; border: 1px solid #1e8449; }
        .score-red { background: #c0392b; border: 1px solid #922b21; }
        
        .gap-text { font-size: 10px; text-align: left; font-weight: bold; min-width: 180px; }
        .gap-red { color: #d32f2f; }
        .perc-col { background: #fff9db; font-weight: bold; color: #856404; }
        .im-col { background: #f3e5f5; font-weight: bold; color: #4a148c; }
        .dealer-name { text-align: left; font-weight: bold; }
    </style>
</head>
<body>

<div class="container">
    <h2>SAARC KPIs (Tanmoy Ghosh)</h2>
    <div class="upload-area" onclick="document.getElementById('fileInput').click()">
        <strong>Click to Upload Files (Enquiry, Booking, Retail, Finance, IM, etc.)</strong>
        <input type="file" id="fileInput" multiple accept=".xlsx, .xls" style="display:none">
    </div>

    <div id="mainDashboard" style="display:none;">
        <div class="action-bar">
            <button class="btn-download" onclick="exportTableToExcel()">Download Full Report (Excel)</button>
        </div>

        <div class="stats-container">
            <div class="card bg-1"><h4>Total Walk-in Enquiry</h4><h2 id="total-enq">0</h2></div>
            <div class="card bg-2"><h4>Total Booking</h4><h2 id="total-book">0</h2></div>
            <div class="card bg-3"><h4>Total Retail</h4><h2 id="total-retail">0</h2></div>
            <div class="card bg-4"><h4>Total Finance</h4><h2 id="total-fin">0</h2></div>
            <div class="card bg-6"><h4>Total Exchange</h4><h2 id="total-exch">0</h2></div>
        </div>

        <div class="table-wrapper">
            <table id="kpiTable">
                <thead>
                    <tr>
                        <th rowspan="2">Dealer Name</th>
                        <th rowspan="2" style="background:#2c3e50">Perf. Score</th>
                        <th rowspan="2">Gap Analysis (Enquiry & TR vs Avg)</th>
                        <th colspan="9">Volume Parameters</th>
                        <th colspan="2">Avg IM Scores</th>
                        <th colspan="5">Efficiency (%)</th>
                    </tr>
                    <tr>
                        <th>Enq</th><th>TR</th><th>Book</th><th>Retail</th><th>P.Book</th><th>Exch</th><th>Fin</th><th>Drop</th><th>Overdue</th>
                        <th class="im-col">Booking IM</th><th class="im-col">Delivery IM</th>
                        <th>E2B%</th><th>E2R%</th><th>E2D%</th><th>R2F%</th><th class="perc-col">TR Ratio%</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script>
    const config = {
        "Alpha Automotive - Teku": { area: "Central", codes: ["100011"] },
        "Company Store Koteshwor": { area: "Central", codes: ["100005"] },
        "Company Store Naxal": { area: "Central", codes: ["100006"] },
        "Vintage Moto Café": { area: "Central", codes: ["100017"] },
        "Om Automobiles": { area: "East", codes: ["100008"] },
        "Narayani Motors": { area: "East", codes: ["100003"] },
        "NOVUS BROTHERS": { area: "East", codes: ["100013"] },
        "Premium Autohub": { area: "East", codes: ["100021"] },
        "RAJBIRAJ AUTOMOBILES": { area: "East", codes: ["100024"] },
        "Sun Automobiles": { area: "East", codes: ["100007"] },
        "Yeti Automobiles": { area: "East", codes: ["100010"] },
        "ROYAL MOTORS - BIRTANAGAR": { area: "East", codes: ["100018"] },
        "Royal Motors- Birtamode": { area: "East", codes: ["100019"] },
        "Royal Motors- Dharan": { area: "East", codes: ["100022"] },
        "Royal Motors- Itahari": { area: "East", codes: ["100027"] },
        "Bhoomi auto Hub": { area: "West", codes: ["100020"] },
        "Bullet Traders - Sales / Pokhara": { area: "West", codes: ["100004", "100003"], aliases: ["Bullet Traders", "Royal Enfield Showroom - Pokhara"] },
        "Gaire Pratisthan": { area: "West", codes: ["100014", "100016", "100025"] },
        "Metro Auto World": { area: "West", codes: ["100012"] },
        "Na Bha Enterprises": { area: "West", codes: ["100009"] },
        "Sabriti Auto Mobile": { area: "West", codes: ["100015"] },
        "Suvariti Moto": { area: "West", codes: ["100023"] },
        "The Bullet Zone": { area: "West", codes: ["100025"] },
        "West End Nepal": { area: "West", codes: ["100026"] }
    };

    let dataStore = {};

    document.getElementById('fileInput').addEventListener('change', async (e) => {
        const files = e.target.files;
        dataStore = {}; 
        for (let file of files) {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            workbook.SheetNames.forEach(sheetName => {
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                if (json.length > 0) processData(file.name.toLowerCase(), json);
            });
        }
        renderReport();
    });

    function getDealerKey(val) {
        if (!val) return null;
        let sVal = val.toString().trim().toUpperCase();
        for (let name in config) {
            if (sVal.includes(name.toUpperCase()) || config[name].codes.includes(sVal)) return name;
            if (config[name].aliases) {
                for (let alias of config[name].aliases) { if (sVal.includes(alias.toUpperCase())) return name; }
            }
        }
        return null;
    }

    function processData(file, rows) {
        rows.forEach(row => {
            let rawDealer = row['Dealer Name'] || row['DEALER NAME'] || row['Dealer name'] || row['Dealer'] || row['Name'];
            let dealer = getDealerKey(rawDealer);
            if (!dealer) return;

            if (!dataStore[dealer]) {
                dataStore[dealer] = { 
                    name: dealer, area: config[dealer].area, enq:0, tr:0, book:0, retail:0, pbook:0, exch:0, fin:0, drop:0, over:0,
                    bookIM_sum:0, bookIM_count:0, delIM_sum:0, delIM_count:0
                };
            }

            let d = dataStore[dealer];
            let imScore = parseFloat(row['IM Score'] || row['IM SCORE'] || row['Score'] || 0);
            
            if (file.includes('enquir')) {
                d.enq++;
                if (row['Test Ride Taken'] === 'Yes' || row['TR Taken'] === 'Yes') d.tr++;
            }
            if (file.includes('booking')) {
                d.book++;
                if (imScore > 0) { d.bookIM_sum += imScore; d.bookIM_count++; }
                if (row['Lead 2 Wheeler Exchange'] === 'Yes' || row['Exchange'] === 'Yes') d.exch++;
            }
            if (file.includes('deliveries') || file.includes('retail')) {
                d.retail++;
                if (imScore > 0) { d.delIM_sum += imScore; d.delIM_count++; }
            }
            if (file.includes('finance')) d.fin++;
            if (file.includes('overdue')) d.over++;
            if (file.includes('exchange')) d.exch++;
            if (file.includes('dropped')) d.drop++;
            if (file.includes('pending')) d.pbook++;
        });
    }

    function renderReport() {
        document.getElementById('mainDashboard').style.display = 'block';
        let html = "";
        let totals = { enq:0, tr:0, book:0, retail:0, pbook:0, exch:0, fin:0, drop:0, over:0 };
        const dealersList = Object.values(dataStore);
        
        const totalDealers = dealersList.length || 1;
        const netTotalEnq = dealersList.reduce((sum, d) => sum + d.enq, 0);
        const netAvgEnq = netTotalEnq / totalDealers;
        const netAvgTRRatio = dealersList.reduce((sum, d) => sum + (d.tr/d.enq || 0), 0) / totalDealers;

        const areas = ["Central", "East", "West"];
        areas.forEach(area => {
            let areaDealers = dealersList.filter(d => d.area === area);
            if (areaDealers.length > 0) {
                html += `<tr class="area-header"><td colspan="19">${area} Area</td></tr>`;
                areaDealers.forEach(d => {
                    let e2r = (d.retail / d.enq) || 0;
                    let trRatio = (d.tr / d.enq) || 0;
                    
                    let score = (Math.min(60, (e2r/0.12)*60) + Math.min(40, (trRatio/0.50)*40)).toFixed(1);
                    let scoreClass = score >= 50 ? 'score-green' : 'score-red';

                    let gapMsgs = [];
                    if (d.enq < netAvgEnq) gapMsgs.push(`<span class="gap-red">Enquiry < Avg (${netAvgEnq.toFixed(0)})</span>`);
                    if (trRatio < netAvgTRRatio) gapMsgs.push(`<span class="gap-red">TR Ratio < Avg</span>`);

                    let avgB_IM = d.bookIM_count > 0 ? (d.bookIM_sum / d.bookIM_count).toFixed(2) : "0.00";
                    let avgD_IM = d.delIM_count > 0 ? (d.delIM_sum / d.delIM_count).toFixed(2) : "0.00";

                    html += `<tr>
                        <td class="dealer-name">${d.name}</td>
                        <td><span class="score-box ${scoreClass}">${score}</span></td>
                        <td class="gap-text">${gapMsgs.length > 0 ? gapMsgs.join("<br>") : "✅ Healthy"}</td>
                        <td>${d.enq}</td><td>${d.tr}</td><td>${d.book}</td><td>${d.retail}</td><td>${d.pbook}</td><td>${d.exch}</td><td>${d.fin}</td><td>${d.drop}</td><td>${d.over}</td>
                        <td class="im-col">${avgB_IM}</td><td class="im-col">${avgD_IM}</td>
                        <td class="perc-col">${((d.book/d.enq||0)*100).toFixed(1)}%</td>
                        <td class="perc-col">${(e2r*100).toFixed(1)}%</td>
                        <td class="perc-col">${((d.drop/d.enq||0)*100).toFixed(1)}%</td>
                        <td class="perc-col">${((d.fin/d.retail||0)*100).toFixed(1)}%</td>
                        <td class="perc-col">${(trRatio*100).toFixed(1)}%</td>
                    </tr>`;
                    
                    Object.keys(totals).forEach(k => { totals[k] += d[k]; });
                });
            }
        });

        document.getElementById('tableBody').innerHTML = html;
        document.getElementById('total-enq').innerText = totals.enq;
        document.getElementById('total-book').innerText = totals.book;
        document.getElementById('total-retail').innerText = totals.retail;
        document.getElementById('total-fin').innerText = totals.fin; 
        document.getElementById('total-exch').innerText = totals.exch;
    }

    function exportTableToExcel() {
        XLSX.writeFile(XLSX.utils.table_to_book(document.getElementById("kpiTable")), "SAARC_Performance_Report.xlsx");
    }
</script>
</body>
</html>
