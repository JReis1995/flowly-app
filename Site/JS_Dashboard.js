(function() {
    'use strict';

    // ==========================================
    // 1. VARIÁVEIS DE ESTADO DO DASHBOARD
    // ==========================================
    window._dashStart = '';
    window._dashEnd = '';
    window._dashGranularity = 'month';
    window._chartStart = '';
    window._chartEnd = '';

    // ==========================================
    // 2. NAVEGAÇÃO INTERNA (TABS E PERÍODOS)
    // ==========================================

    function setDashSubTab(tab) {
        const tabs = ['Fin', 'Stock', 'RH', 'Frota'];
        tabs.forEach(t => {
            const content = document.getElementById('d' + t);
            const btn = document.getElementById('t' + t);
            if (content) content.classList.add('hidden');
            if (btn) btn.className = "flex-1 py-3 rounded-xl text-[10px] font-bold text-slate-400 border border-slate-50 transition-all";
        });
        const activeContent = document.getElementById('d' + tab);
        const activeBtn = document.getElementById('t' + tab);
        if (activeContent) activeContent.classList.remove('hidden');
        if (activeBtn) activeBtn.className = "flex-1 py-3 rounded-xl text-[10px] font-bold bg-flowly-midnight text-white shadow-xl scale-[1.02] transition-all";

        try {
            if (typeof loadAlertsUI === 'function') loadAlertsUI(window.currentAiInsight || null);
        } catch (e) {
            console.warn("Aviso: Falha ao carregar Alertas IA na mudança de aba.", e);
        }

        if (typeof lucide !== 'undefined' && lucide.createIcons) lucide.createIcons();

        setTimeout(function () {
            if (typeof Chart !== 'undefined') {
                for (var id in Chart.instances) {
                    try { Chart.instances[id].resize(); } catch (err) { }
                }
            }
        }, 50);
    }

    function setDashPeriod(period) {
        const now = new Date();
        const pad = n => String(n).padStart(2, '0');
        const fmt = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

        let start = '', end = fmt(now);
        let chartStart = '', chartEnd = fmt(now);

        if (period === 'day') {
            start = fmt(now);
            window._dashGranularity = 'day';
            const cs = new Date(now); cs.setDate(now.getDate() - 6);
            chartStart = fmt(cs);
            chartEnd = fmt(now);
        } else if (period === 'week') {
            const mon = new Date(now); mon.setDate(now.getDate() - ((now.getDay() + 6) % 7));
            start = fmt(mon);
            window._dashGranularity = 'week';
            const cs = new Date(now); cs.setDate(now.getDate() - 7 * 7);
            chartStart = fmt(cs);
            chartEnd = fmt(now);
        } else if (period === 'month') {
            start = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`;
            end = fmt(now);
            window._dashGranularity = 'month';
            const cs = new Date(now.getFullYear(), now.getMonth() - 6, 1);
            chartStart = fmt(cs);
            chartEnd = fmt(now);
        } else if (period === 'year') {
            start = `${now.getFullYear()}-01-01`;
            window._dashGranularity = 'year';
            const cs = new Date(now.getFullYear() - 6, 0, 1);
            chartStart = fmt(cs);
            chartEnd = fmt(now);
        } else if (period === 'custom') {
            document.getElementById('customDateRow').classList.remove('hidden');
            document.querySelectorAll('.dash-period-btn').forEach(b => b.classList.remove('active'));
            document.getElementById('periodCustom').classList.add('active');
            window._dashGranularity = 'custom';
            const sEl = document.getElementById('dStart');
            const eEl = document.getElementById('dEnd');
            const defStart = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`;
            const defEnd = fmt(now);
            if (sEl && !sEl.value) sEl.value = defStart;
            if (eEl && !eEl.value) eEl.value = defEnd;
            window._dashStart = (sEl && sEl.value) ? sEl.value : defStart;
            window._dashEnd = (eEl && eEl.value) ? eEl.value : defEnd;
            window._chartStart = window._dashStart;
            window._chartEnd = window._dashEnd;
            updateDash();
            return;
        }

        document.getElementById('customDateRow').classList.add('hidden');
        document.querySelectorAll('.dash-period-btn').forEach(b => b.classList.remove('active'));
        const map = { day: 'periodDay', week: 'periodWeek', month: 'periodMonth', year: 'periodYear' };
        document.getElementById(map[period]).classList.add('active');

        const sEl = document.getElementById('dStart'); const eEl = document.getElementById('dEnd');
        if (sEl) sEl.value = start; if (eEl) eEl.value = end;
        window._dashStart = start; window._dashEnd = end;
        window._chartStart = chartStart; window._chartEnd = chartEnd;
        updateDash();
    }

    function toggleWidget(widgetId) {
        const el = document.getElementById(widgetId);
        if (el) el.classList.toggle('hidden');
    }

    // ==========================================
    // 3. ATUALIZAÇÃO E RENDERING (KPIs E GRÁFICOS)
    // ==========================================

    function aggregateChartData(labels, data, granularity) {
        if (!labels || !data || labels.length === 0) return { labels: [], data: [] };
        const dayNames = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
        const mNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

        if (granularity === 'day') {
            return {
                labels: labels.map(l => { const p = l.split('/'); return p.length >= 2 ? `${p[0]}/${p[1]}` : l; }),
                data
            };
        }

        const agg = {}, aggOrder = [];
        const addKey = (key, val) => { if (!agg.hasOwnProperty(key)) { agg[key] = 0; aggOrder.push(key); } agg[key] += (parseFloat(val) || 0); };

        labels.forEach((lbl, i) => {
            const parts = lbl.split('/');
            if (parts.length < 3) return;
            const d = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
            if (isNaN(d)) return;

            if (granularity === 'week') {
                addKey(`${dayNames[d.getDay()]} ${parts[0]}/${parts[1]}`, data[i]);
            } else if (granularity === 'month') {
                addKey(`S${Math.ceil(d.getDate() / 7)} ${parts[1]}/${parts[2]}`, data[i]);
            } else if (granularity === 'year') {
                addKey(mNames[d.getMonth()], data[i]);
            } else {
                let rangeDays = 365;
                if (window._dashStart && window._dashEnd) {
                    rangeDays = (new Date(window._dashEnd) - new Date(window._dashStart)) / 86400000;
                }
                if (rangeDays <= 60) {
                    addKey(`S${Math.ceil(d.getDate() / 7)} ${parts[1]}/${parts[2]}`, data[i]);
                } else {
                    addKey(`${mNames[d.getMonth()]} ${parts[2]}`, data[i]);
                }
            }
        });
        return { labels: aggOrder, data: aggOrder.map(k => agg[k]) };
    }

    function updateDash(forceAI) {
        const displayEl = document.getElementById('displayAutoOpEx');
        if (displayEl) {
            displayEl.innerText = 'Taxa de Absorção: --%';
            displayEl.classList.add('animate-pulse');
        }
        const kpIds = ['kpFat', 'kpFatSublabel', 'kpLucro', 'kpMargem', 'kpDespesas', 'kpDespesasSublabel', 'kpCompromissosRH', 'kpRHSublabel', 'kpIVAIRC', 'kpTicket', 'kpTrans', 'kpFaturasAPagamento', 'kpCaixaLivre', 'kpAReceber', 'kpStockVal', 'kpRot', 'kpCompras', 'kpDeadStock', 'kpErosion', 'kpDiasCobertura', 'kpRHCost', 'kpRHCostEstimado', 'kpRHCount', 'kpRHMedio', 'kpRHIdeal', 'kpFrotaCusto', 'kpFrotaCustoKm', 'kpFrotaKmTotais', 'kpFrotaConsumo', 'kpFrotaAlertas', 'kpFrotaManutencao', 'kpFrotaPortagens', 'kpFrotaAbastecimento'];
        kpIds.forEach(id => { const el = document.getElementById(id); if (el) el.innerText = '--'; });

        if (!window._dashStart || !window._dashEnd) {
            setDashPeriod('month');
            return;
        }

        if (document.getElementById('periodCustom')?.classList.contains('active')) {
            const sEl = document.getElementById('dStart');
            const eEl = document.getElementById('dEnd');
            if (sEl?.value && eEl?.value) {
                window._dashStart = sEl.value;
                window._dashEnd = eEl.value;
                window._chartStart = sEl.value;
                window._chartEnd = eEl.value;
                window._dashGranularity = 'custom';
            }
        }

        google.script.run.withSuccessHandler(renderDashboardCompleto).withFailureHandler(() => {
            kpIds.forEach(id => { const el = document.getElementById(id); if (el) el.innerText = 'Erro'; });
        }).getUnifiedDashboardData(_ctxEmail(), window._dashStart || '', window._dashEnd || '', window._chartStart || '', window._chartEnd || '', forceAI === true);
    }

    function renderDashboardCompleto(res) {
        if (typeof Chart === 'undefined') {
            console.warn('Chart.js não detetado. Gráficos ignorados.');
        }

        if (!res || !res.success) return;
        window._dashboardResForCFO = res;
        window.currentAiInsight = res;

        var clientCfg = res.clientConfig || (currentUser && (currentUser._impersonatedPlanConfig || currentUser.planConfig));
        var cfg = clientCfg || {};

        var canFin = typeof resolveFeaturePermission === 'function' && resolveFeaturePermission(cfg, 'dashFinanceiro');
        var canStock = typeof resolveFeaturePermission === 'function' && resolveFeaturePermission(cfg, 'dashStocks');
        var canRH = typeof resolveFeaturePermission === 'function' && resolveFeaturePermission(cfg, 'dashRH');
        var canCaixa = typeof resolveFeaturePermission === 'function' && resolveFeaturePermission(cfg, 'caixaLivre');

        if (typeof syncAIAutoToggleFromRes === 'function') syncAIAutoToggleFromRes(res);

        const fin = res.financeiro || {};
        const st = res.stock || {};
        const rh = res.rh || {};
        const set = (id, text) => { const el = document.getElementById(id); if (el) el.innerText = text; };

        if (canFin) {
            set('kpFat', (parseFloat(fin.faturacao) || 0).toFixed(2) + "€");
            set('kpFatSublabel', (fin.fatRecebido != null ? "Recebido: " + (parseFloat(fin.fatRecebido) || 0).toFixed(2) + "€" : "Recebido: --€") + " | " + (fin.fatPendente != null ? "Pendente: " + (parseFloat(fin.fatPendente) || 0).toFixed(2) + "€" : "Pendente: --€"));
            set('kpLucro', (parseFloat(fin.lucro_liq) || 0).toFixed(2) + "€");
            set('kpMargem', (parseFloat(fin.margem_perc) || 0) + "%");
            set('kpDespesas', (parseFloat(fin.despesas) || 0).toFixed(2) + "€");
            set('kpDespesasSublabel', (fin.despesasPago != null ? "Pago: " + (parseFloat(fin.despesasPago) || 0).toFixed(2) + "€" : "Pago: --€") + " | " + (fin.despesasPorPagar != null ? "Por Pagar: " + (parseFloat(fin.despesasPorPagar) || 0).toFixed(2) + "€" : "Por Pagar: --€"));
            set('kpCompromissosRH', (parseFloat(fin.rhProcessado) || 0).toFixed(2) + "€");
            set('kpRHSublabel', "Impacto no Lucro: " + (parseFloat(fin.rhProcessado) || 0).toFixed(2) + "€ | Pago: " + (parseFloat(fin.rhPago) || 0).toFixed(2) + "€");
            set('kpIVAIRC', ((parseFloat(fin.iva_pagar) || 0) + (parseFloat(fin.irc_est) || 0)).toFixed(2) + "€");
            set('kpTicket', (parseFloat(fin.ticket_medio) || 0).toFixed(2) + "€");
            set('kpTrans', String(fin.transacoes != null ? fin.transacoes : "--"));
            set('kpFaturasAPagamento', (parseFloat(fin.faturasAPagamento) || 0).toFixed(2) + "€");
            set('kpAReceber', (parseFloat(fin.contasAReceber) || 0).toFixed(2) + "€");

            var topDevedoresEl = document.getElementById('topDevedoresList');
            if (topDevedoresEl) {
                var lista = fin.listaDevedores || [];
                if (lista.length === 0) topDevedoresEl.innerHTML = '<p class="text-slate-400 dark:text-slate-400 text-[9px]">Sem devedores em aberto.</p>';
                else topDevedoresEl.innerHTML = lista.map((d, i) => '<div class="flex justify-between items-center py-2 px-3 bg-slate-100 dark:bg-slate-700/50 rounded-lg"><span class="font-bold text-slate-800 dark:text-slate-200 truncate">' + (i + 1) + '. ' + (d.nome || '').toString().replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</span><span class="text-cyan-600 dark:text-cyan-400 font-black ml-2">' + (parseFloat(d.valor) || 0).toFixed(2) + ' €</span></div>').join('');
            }
        }

        if (canCaixa) {
            var caixaLivreVal = fin.caixaLivre != null ? (parseFloat(fin.caixaLivre) || 0).toFixed(2) + "€" : "0.00€";
            var kpCaixaEl = document.getElementById('kpCaixaLivre');
            if (kpCaixaEl) {
                kpCaixaEl.innerText = caixaLivreVal;
                kpCaixaEl.classList.remove('text-red-500', 'text-cyan-600', 'text-flowly-success', 'dark:text-white', 'dark:text-red-400');

                if ((parseFloat(fin.caixaLivre) || 0) < 0) {
                    kpCaixaEl.classList.add('text-red-500', 'dark:text-red-400');
                } else {
                    kpCaixaEl.classList.add('text-cyan-600', 'dark:text-white');
                }
            }
        }

        if (canStock) {
            set('kpStockVal', (parseFloat(st.valor) || 0).toFixed(2) + "€");
            set('kpRot', st.rotatividade != null ? st.rotatividade : "--");
            set('kpCompras', (parseFloat(st.compras) || 0).toFixed(2) + "€");
            set('kpDeadStock', (parseFloat(res.deadStockValue) || 0).toFixed(2) + "€");
            set('kpErosion', String(res.erosionCount != null ? res.erosionCount : "--"));
            set('kpDiasCobertura', res.stock?.diasCobertura != null ? res.stock.diasCobertura + " dias" : "--");

            const stockList = document.getElementById('dStockPorArtigo');
            if (stockList && Array.isArray(st.stockPorArtigo)) {
                if (st.stockPorArtigo.length === 0) stockList.innerHTML = '<p class="text-slate-400 dark:text-slate-100 text-center py-4">Sem dados de stock por artigo.</p>';
                else stockList.innerHTML = st.stockPorArtigo.map(x => `<div class="flex justify-between items-center p-3 bg-slate-50 dark:bg-slate-700/50 rounded-xl"><span class="font-bold text-flowly-midnight dark:text-slate-100 truncate">${(x.artigo || "").toString()}</span><span class="text-flowly-warning dark:text-amber-400 font-black whitespace-nowrap ml-2">${Number(x.quantidadeEmStock)} un · ${(parseFloat(x.valorEmStock) || 0).toFixed(2)}€</span></div>`).join('');
            }

            var vasilhame = st.vasilhame || {};
            var vasilhameSection = document.getElementById('dStockVasilhame');
            if (vasilhameSection) {
                var valorNaRua = parseFloat(vasilhame.valorTotalNaRua) || 0;
                var saldos = vasilhame.saldos || [];
                vasilhameSection.innerHTML = '<p class="text-[10px] font-black uppercase mb-2 text-slate-400 dark:text-cyan-400 tracking-wider">Controlo de Vasilhame</p><p class="text-lg font-black text-flowly-midnight dark:text-slate-100">Capital na rua: <span class="text-flowly-warning dark:text-amber-400">' + valorNaRua.toFixed(2) + '€</span></p>' + (saldos.length > 0 ? '<div class="mt-3 space-y-2 max-h-[120px] overflow-y-auto">' + saldos.slice(0, 10).map(s => '<div class="flex justify-between text-xs p-2 bg-slate-50 dark:bg-slate-700/50 rounded-lg"><span class="text-slate-600 dark:text-slate-100">' + (s.cliente || '') + ' · ' + (s.artigo || '') + '</span><span class="font-bold text-slate-700 dark:text-slate-100">' + Number(s.saldo) + ' un</span></div>').join('') + '</div>' : '<p class="text-slate-400 dark:text-slate-400 text-sm mt-2">Sem saldos de vasilhame.</p>');
            }
        }

        var canFrota = typeof resolveFeaturePermission === 'function' && resolveFeaturePermission(cfg, 'dashFrota');
        if (canFrota && res.frota) {
            var frota = res.frota;
            set('kpFrotaCusto', (parseFloat(frota.custoTotalMes) || 0).toFixed(2) + "€");
            set('kpFrotaCustoKm', frota.custoPorKm != null ? (parseFloat(frota.custoPorKm) || 0).toFixed(3) + " €/km" : "N/A");
            set('kpFrotaKmTotais', (frota.kmTotais != null && frota.kmTotais !== undefined ? Math.round(frota.kmTotais) : 0) + " km");
            set('kpFrotaConsumo', frota.consumoMedioGlobal != null ? frota.consumoMedioGlobal + " L/100km" : "N/A");
            set('kpFrotaAlertas', String(frota.alertasCount != null ? frota.alertasCount : "--"));

            var gastos = frota.gastosPorCategoria || {};
            set('kpFrotaManutencao', (parseFloat(gastos.manutencao) || 0).toFixed(2) + "€");
            set('kpFrotaPortagens', (parseFloat(gastos.portagens) || 0).toFixed(2) + "€");
            set('kpFrotaAbastecimento', (parseFloat(gastos.combustivel) || 0).toFixed(2) + "€");

            var rankingEl = document.getElementById('dFrotaRanking');
            if (rankingEl && Array.isArray(frota.rankingConsumo)) {
                if (frota.rankingConsumo.length === 0) rankingEl.innerHTML = '<p class="text-slate-400 dark:text-slate-100 text-center py-4">Sem dados de consumo.</p>';
                else rankingEl.innerHTML = frota.rankingConsumo.map((r, i) => '<div class="flex justify-between items-center p-3 bg-slate-50 dark:bg-slate-800/50 rounded-xl"><span class="font-bold text-flowly-midnight dark:text-slate-100">' + (i + 1) + '. ' + (r.matricula || '') + '</span><span class="text-flowly-warning dark:text-amber-400 font-black">' + (r.consumoL100km != null ? r.consumoL100km + ' L/100km' : 'N/A') + '</span></div>').join('');
            }

            var alertasEl = document.getElementById('dFrotaAlertas');
            if (alertasEl && Array.isArray(frota.veiculosComAlertas)) {
                if (frota.veiculosComAlertas.length === 0) alertasEl.innerHTML = '<p class="text-flowly-success text-sm font-bold text-center py-2 text-slate-100">Sem alertas ativos.</p>';
                else {
                    var semaforoIcon = function (s) { return s === 'vermelho' ? '<span class="inline-flex w-4 h-4 rounded-full bg-flowly-danger" title="IPO/Seguro expirado ou &lt; 7 dias"></span>' : s === 'laranja' ? '<span class="inline-flex w-4 h-4 rounded-full bg-orange-500" title="Lavagem &gt; 7 dias ou Revisão próxima"></span>' : '<span class="inline-flex w-4 h-4 rounded-full bg-emerald-500" title="Tudo em dia"></span>'; };
                    alertasEl.innerHTML = frota.veiculosComAlertas.map(v => {
                        var sem = v.statusSemaforo || 'laranja';
                        var badges = [];
                        if (v.alertaIPO) badges.push('<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-flowly-danger/20 text-red-400" title="Margem 30 dias">IPO' + (v.diasIPO != null ? ': ' + v.diasIPO + ' dias' : '') + '</span>');
                        if (v.alertaSeguro) badges.push('<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-flowly-danger/20 text-red-400" title="Margem 30 dias">Seguro' + (v.diasSeguro != null ? ': ' + v.diasSeguro + ' dias' : '') + '</span>');
                        if (v.alertaLavagem) badges.push('<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-orange-500/20 text-orange-400">Lavagem</span>');
                        return '<div class="flex items-center justify-between gap-2 p-3 bg-slate-800/50 rounded-xl text-xs"><div class="flex items-center gap-2"><span>' + semaforoIcon(sem) + '</span><span class="font-bold text-slate-100">' + (v.matricula || '') + ' ' + (v.marcaModelo || '') + '</span></div><div class="flex gap-1.5 flex-wrap justify-end">' + badges.join('') + '</div></div>';
                    }).join('');
                }
            }
        }

        if (canRH) {
            set('kpRHCost', (parseFloat(rh.custo_total) || 0).toFixed(2) + "€");
            set('kpRHCostEstimado', (parseFloat(rh.custo_estimado) || 0).toFixed(2) + "€");
            set('kpRHCount', String(rh.ativos != null ? rh.ativos : "--"));
            set('kpRHMedio', String(rh.custo_medio != null ? rh.custo_medio : "--"));
            set('kpRHIdeal', String(rh.ideal != null ? rh.ideal : "--"));
        }

        if (clientCfg && typeof applyClientConfig === 'function') applyClientConfig(clientCfg);
        setDashSubTab('Fin');

        const gran = window._dashGranularity || 'month';
        const granLabels = { day: 'Últimos 7 Dias', week: 'Últimas 7 Semanas', month: 'Últimos 7 Meses', year: 'Últimos 7 Anos', custom: 'Período Custom' };

        if (canFin) {
            const lbl1 = document.getElementById('chartProfitLabel');
            if (lbl1) lbl1.textContent = `Evolução do Lucro — ${granLabels[gran] || ''}`;
            const chartFin = res.chartFinanceiro || fin;
            const aggProfit = aggregateChartData(chartFin.chartLabels || [], (chartFin.chartData || []).map(Number), gran);
            const cProfit = document.getElementById('chartProfit');

            if (cProfit && typeof Chart !== 'undefined') {
                var existingProfit = Chart.getChart(cProfit);
                if (existingProfit) existingProfit.destroy();
                new Chart(cProfit, {
                    type: 'line',
                    data: { labels: aggProfit.labels, datasets: [{ label: 'Lucro (€)', data: aggProfit.data, borderColor: '#06B6D4', backgroundColor: 'rgba(6,182,212,0.08)', borderWidth: 2.5, pointRadius: aggProfit.labels.length > 20 ? 0 : 3, tension: 0.4, fill: true }] },
                    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { font: { size: 9, weight: 'bold' }, maxRotation: 45, maxTicksLimit: 10 }, grid: { display: false } }, y: { ticks: { font: { size: 9 }, callback: v => v.toFixed(0) + '€' }, grid: { color: '#f1f5f9' } } } }
                });
            }

            const aggFat = aggregateChartData(chartFin.chartLabels || [], (chartFin.chartFatData || chartFin.chartData || []).map(v => parseFloat(v) || 0), gran);
            const aggDesp = aggregateChartData(chartFin.chartLabels || [], (chartFin.chartDespData || []).map(v => parseFloat(v) || 0), gran);
            const cFatDesp = document.getElementById('chartFatDesp');

            if (cFatDesp && typeof Chart !== 'undefined') {
                var existingFatDesp = Chart.getChart(cFatDesp);
                if (existingFatDesp) existingFatDesp.destroy();
                new Chart(cFatDesp, {
                    type: 'bar',
                    data: {
                        labels: aggFat.labels.length ? aggFat.labels : [granLabels[gran] || 'Período'],
                        datasets: [
                            { label: 'Faturação', data: aggFat.data, backgroundColor: 'rgba(16,185,129,0.75)', borderRadius: 6, borderSkipped: false },
                            { label: 'Despesas', data: aggDesp.data.length ? aggDesp.data : aggFat.labels.map(() => 0), backgroundColor: 'rgba(239,68,68,0.65)', borderRadius: 6, borderSkipped: false }
                        ]
                    },
                    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { font: { size: 9, weight: 'bold' } } } }, scales: { x: { ticks: { font: { size: 9 }, maxRotation: 45, maxTicksLimit: 10 }, grid: { display: false } }, y: { ticks: { font: { size: 9 }, callback: v => v.toFixed(0) + '€' }, grid: { color: '#f1f5f9' } } } }
                });
            }
        }

        if (canStock) {
            const cTop = document.getElementById('chartTopProd');
            if (cTop && Array.isArray(st.topLabels) && st.topLabels.length && typeof Chart !== 'undefined') {
                var existingTop = Chart.getChart(cTop);
                if (existingTop) existingTop.destroy();
                new Chart(cTop, {
                    type: 'bar',
                    data: { labels: st.topLabels, datasets: [{ label: 'Unidades', data: st.topData || [], backgroundColor: 'rgba(6,182,212,0.7)', borderRadius: 6, borderSkipped: false }] },
                    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { font: { size: 9 } }, grid: { color: '#f1f5f9' } }, y: { ticks: { font: { size: 9, weight: 'bold' } }, grid: { display: false } } } }
                });
            }
        }

        if (typeof loadMarginSettingsUI === 'function') loadMarginSettingsUI();

        var ai = res && res.aiInsight;
        var hasDisplayable = ai && (getText(ai, 'financeiro.summary') || getText(ai, 'stocks.summary') || getText(ai, 'rh.summary'));
        if (!hasDisplayable && (ai && ai.atRest)) {
            var fallback = null;
            try {
                var c = localStorage.getItem(_aiCacheKey());
                if (c) {
                    var parsed = JSON.parse(c);
                    if (parsed && (!parsed._ownerEmail || parsed._ownerEmail === _ctxEmail())) fallback = parsed;
                }
            } catch (e) { }
            if (fallback && fallback.aiInsight && (fallback.aiInsight.financeiro || fallback.aiInsight.stocks || fallback.aiInsight.rh)) {
                res = Object.assign({}, res, { aiInsight: fallback.aiInsight });
            }
        }

        if (typeof loadAlertsUI === 'function') loadAlertsUI(res);
    }

    function loadMarginSettingsUI() {
        google.script.run.withSuccessHandler(function (s) {
            const m = (s && s.margemDesejada != null) ? parseFloat(s.margemDesejada) : 30;
            const i = (s && s.ircEstimado != null) ? parseFloat(s.ircEstimado) : 20;
            const sel = document.getElementById('marginTarget');
            const customRow = document.getElementById('marginCustomRow');
            const customInp = document.getElementById('marginCustom');
            const ircInp = document.getElementById('marginIRC');
            const methodSel = document.getElementById('marginMethod');

            if (ircInp) ircInp.value = i;
            if (methodSel && s && s.metodoCalculo) methodSel.value = s.metodoCalculo;
            const opts = ['10', '20', '30', '40', '50'];

            if (opts.indexOf(String(m)) >= 0 && sel) {
                sel.value = String(m);
                if (customRow) customRow.classList.add('hidden');
            } else if (sel) {
                sel.value = 'Personalizado';
                if (customRow) customRow.classList.remove('hidden');
                if (customInp) customInp.value = m;
            }

            google.script.run.withSuccessHandler(function (res) {
                const el = document.getElementById('displayAutoOpEx');
                if (el) {
                    el.classList.remove('animate-pulse');
                    const r = (res && typeof res === 'object') ? res : { rate: res, label: '365d' };
                    el.innerText = 'Taxa de Absorção (' + (r.label || '365d') + '): ' + (r.rate != null ? r.rate + '%' : 'Erro');
                }
            }).calculateDynamicOpEx(_ctxEmail(), window._dashStart || '', window._dashEnd || '');
        }).withFailureHandler(function () { }).getMarginSettings();
    }

    function saveMarginUI() {
        const targetEl = document.getElementById('marginTarget');
        const ircEl = document.getElementById('marginIRC');
        const customEl = document.getElementById('marginCustom');
        const updateHistEl = document.getElementById('marginUpdateHistory');
        const methodEl = document.getElementById('marginMethod');
        if (!targetEl || !ircEl || !updateHistEl) return;

        let margem = parseFloat(targetEl.value);
        if (targetEl.value === 'Personalizado' && customEl) {
            margem = parseFloat(customEl.value);
        }
        if (isNaN(margem) || margem < 0 || margem > 100) {
            notify('Margem inválida. Introduza um valor entre 0 e 100.', 'error');
            return;
        }

        const irc = parseFloat(ircEl.value);
        if (isNaN(irc) || irc < 0 || irc > 100) {
            notify('IRC inválido. Introduza um valor entre 0 e 100.', 'error');
            return;
        }

        const metodoCalculo = (methodEl && methodEl.value) ? methodEl.value : 'markup';
        const updateHistory = updateHistEl.checked === true;

        if (updateHistory) {
            var dashStart = window._dashStart || '';
            var dashEnd = window._dashEnd || '';
            if (dashStart && dashEnd) {
                var start = new Date(dashStart);
                var end = new Date(dashEnd);
                var dias = Math.round((end - start) / 86400000) + 1;
                if (dias < 30) {
                    notify('Selecione um período superior a 30 dias para aplicar a estratégia.', 'error');
                    return;
                }
            }
            if (typeof showBulkLoading === 'function') showBulkLoading(true);
        }

        google.script.run.withSuccessHandler(function (res) {
            if (typeof showBulkLoading === 'function') showBulkLoading(false);
            if (res && res.success) {
                notify(res.message || 'Estratégia aplicada.', 'success');
                if (res.ratePersisted === false) {
                    notify('Cliente novo: a taxa foi aplicada ao histórico. Os próximos artigos não terão esta taxa até selecionar um período de 30+ dias com dados.', 'info');
                }
                google.script.run.withSuccessHandler(function (res) {
                    var el = document.getElementById('displayAutoOpEx');
                    if (el) {
                        el.classList.remove('animate-pulse');
                        var r = (res && typeof res === 'object') ? res : { rate: res, label: '365d' };
                        el.innerText = 'Taxa de Absorção (' + (r.label || '365d') + '): ' + (r.rate != null ? r.rate + '%' : 'Erro');
                    }
                }).calculateDynamicOpEx(_ctxEmail(), window._dashStart || '', window._dashEnd || '');
            } else {
                notify(res && res.error ? res.error : 'Erro ao guardar.', 'error');
            }
        }).withFailureHandler(function (err) {
            if (typeof showBulkLoading === 'function') showBulkLoading(false);
            notify('Erro: ' + (err || 'Falha ao aplicar.'), 'error');
        }).saveMarginSettingsAndHistory(margem, irc, updateHistory, _ctxEmail(), metodoCalculo, window._dashStart || '', window._dashEnd || '');
    }

    function toggleTask(event, element) {
        if (event) event.stopPropagation();
        const label = element && element.closest ? element.closest('label') : (element || null);
        if (!label) return;
        const checkbox = label.querySelector('input[type="checkbox"]');
        const span = label.querySelector('span');
        if (!checkbox || !span) return;
        checkbox.checked = !checkbox.checked;
        span.classList.toggle('line-through', checkbox.checked);
        span.classList.toggle('opacity-50', checkbox.checked);
        const key = label.getAttribute('data-todo-key') || _todoStorageKey(span.textContent, 'fin');
        try { localStorage.setItem(key, checkbox.checked ? '1' : '0'); } catch (e) { }
    }

    function openAIHistoryModal(category) {
        var listEl = document.getElementById('modalAIHistoryList');
        if (listEl) listEl.innerHTML = '<p class="text-slate-500 dark:text-slate-400 text-center py-4">A carregar histórico...</p>';
        var modal = document.getElementById('modalAIHistory');
        if (modal) modal.classList.remove('hidden');

        if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run
                .withSuccessHandler(function (entries) {
                    if (typeof renderAIHistoryCards === 'function') renderAIHistoryCards(entries);
                })
                .withFailureHandler(function () {
                    if (listEl) listEl.innerHTML = '<p class="text-slate-500 dark:text-slate-400 text-center py-4">Não foi possível carregar o histórico.</p>';
                })
                .getAIHistory(typeof _ctxEmail === 'function' ? _ctxEmail() : '', category || '');
        }
    }

    function fetchCategoryAI(category) {
        var res = window._dashboardResForCFO;
        if (!res) {
            updateDash(false);
            return;
        }
        var textIdMap = { financeiro: 'aiText_financeiro', stocks: 'aiText_stocks', rh: 'aiText_rh', fleet: 'aiText_fleet' };
        var el = document.getElementById(textIdMap[category]);
        if (el) el.innerHTML = 'A analisar...';
        var label = (category === 'financeiro') ? 'Financeiro' : (category === 'stocks') ? 'Stocks' : (category === 'fleet') ? 'Frota' : 'RH';

        google.script.run.withSuccessHandler(function (responseInsight) {
            if (responseInsight && responseInsight.success !== false) {
                if (!res.aiInsight) res.aiInsight = {};

                res.aiInsight[category] = responseInsight[category] ? responseInsight[category] : responseInsight;
                if (responseInsight.currentCredits != null) res.aiInsight.currentCredits = responseInsight.currentCredits;

                window._dashboardResForCFO = res;
                window.currentAiInsight = res;

                if (typeof refreshSingleCategoryCard === 'function') {
                    refreshSingleCategoryCard(category, res.aiInsight);
                }

                if (window.notify) window.notify('Análise ' + label + ' atualizada.', 'success');
            } else {
                if (el) el.innerHTML = (responseInsight && responseInsight.error) ? responseInsight.error : 'Erro ao atualizar.';
                if (window.notify) window.notify(responseInsight && responseInsight.error ? responseInsight.error : 'Erro.', 'error');
            }
        }).withFailureHandler(function () {
            if (el) el.innerHTML = 'Erro ao atualizar.';
            if (window.notify) window.notify('Erro de rede.', 'error');
        }).getFlowlyAIInsightCategory(_ctxEmail(), category, res);
    }

    function _restoreTodoStates(labelSelector) {
        try {
            const labels = document.querySelectorAll(labelSelector || 'label[data-todo-key]');
            labels.forEach(function (label) {
                const key = label.getAttribute('data-todo-key');
                const checkbox = label.querySelector('input[type="checkbox"]');
                const span = label.querySelector('span');
                if (!key || !checkbox || !span) return;
                const stored = localStorage.getItem(key);
                if (stored === '1') {
                    checkbox.checked = true;
                    span.classList.add('line-through', 'opacity-50');
                }
            });
        } catch (e) { console.warn("Erro ao restaurar tarefas:", e); }
    }

    window.setDashSubTab = setDashSubTab;
    window.setDashPeriod = setDashPeriod;
    window.toggleWidget = toggleWidget;
    window.aggregateChartData = aggregateChartData;
    window.updateDash = updateDash;
    window.renderDashboardCompleto = renderDashboardCompleto;
    window.loadMarginSettingsUI = loadMarginSettingsUI;
    window.saveMarginUI = saveMarginUI;
    window.toggleTask = toggleTask;
    window.openAIHistoryModal = openAIHistoryModal;
    window.fetchCategoryAI = fetchCategoryAI;
    window._restoreTodoStates = _restoreTodoStates;

})();
