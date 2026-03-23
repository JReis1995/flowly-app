(function() {
    'use strict';
    
    // ==========================================
    // 1. VARIÁVEIS GLOBAIS (CC E LOGÍSTICA)
    // ==========================================
    window.selectedDocIds = [];
    window.ccCurrentGroup = null;

    let pendingInvoiceB64 = null;
    let pendingInvoiceName = "";
    let pendingTipoDocumento = "FaturaCompra";

    // ==========================================
    // 2. CONTA CORRENTE E TESOURARIA
    // ==========================================

    function setCcFlow(flow) {
        currentCcFlow = flow;
        const btnAvisos = document.getElementById('ccBulkAvisosBtn');
        if (btnAvisos) btnAvisos.classList.toggle('hidden', flow !== 'Saidas');

        const topBarAvisos = document.getElementById('ccAvisosTopBar');
        if (topBarAvisos) topBarAvisos.classList.toggle('hidden', flow !== 'Saidas');

        const btnEnt = document.getElementById('tabCcFlowEntradas');
        const btnSai = document.getElementById('tabCcFlowSaidas');

        if (btnEnt) btnEnt.className = flow === 'Entradas' ? "flex-1 py-3 rounded-xl text-[10px] font-black bg-flowly-midnight text-white shadow-md transition-all" : "flex-1 py-3 rounded-xl text-[10px] font-black text-slate-400 hover:bg-white transition-all";
        if (btnSai) btnSai.className = flow === 'Saidas' ? "flex-1 py-3 rounded-xl text-[10px] font-black bg-flowly-midnight text-white shadow-md transition-all" : "flex-1 py-3 rounded-xl text-[10px] font-black text-slate-400 hover:bg-white transition-all";

        loadCC();
    }

    function filterCcTab(status) {
        currentCcTab = status;
        const tabs = ['Aberto', 'Pago', 'Todas'];
        tabs.forEach(t => {
            const btn = document.getElementById('tabCc' + t);
            if (btn) {
                btn.className = t === status
                    ? "flex-1 py-3 rounded-xl text-[10px] font-black bg-flowly-midnight text-white shadow-md transition-all"
                    : "flex-1 py-3 rounded-xl text-[10px] font-black text-slate-400 hover:bg-white transition-all";
            }
        });
        loadCC();
    }

    function loadCC() {
        const list = document.getElementById('ccList');
        if (!list) return;
        list.innerHTML = '<div class="text-center py-20 opacity-50 text-[10px] font-black uppercase tracking-widest animate-pulse flex flex-col items-center gap-3"><div class="w-10 h-10 border-2 border-flowly-border border-t-flowly-primary rounded-full animate-spin"></div>A processar dados...</div>';

        const ctxBanner = document.getElementById('ccContextBanner');
        const ctxName = document.getElementById('ccContextName');
        if (ctxBanner) {
            if (currentImpersonate) {
                ctxBanner.classList.remove('hidden');
                if (ctxName) ctxName.textContent = document.getElementById('impName') ? document.getElementById('impName').innerText : currentImpersonate;
            } else {
                ctxBanner.classList.add('hidden');
            }
        }

        const setEl = (id, text) => { const el = document.getElementById(id); if (el) el.innerText = text; };

        google.script.run
            .withFailureHandler(err => {
                list.innerHTML = `<div class="bg-rose-50 text-rose-600 p-8 rounded-2xl text-center text-xs font-black uppercase border border-rose-100 shadow-inner">
                Erro ao carregar tesouraria.<br><span class="font-normal normal-case text-rose-400 mt-1 block">${err && err.message ? err.message : 'Verifique o Apps Script'}</span>
                <button onclick="loadCC()" class="mt-4 bg-rose-600 text-white px-5 py-2 rounded-2xl text-[9px] font-black uppercase shadow hover:bg-rose-700 active:scale-95 transition">Tentar de novo</button>
            </div>`;
            })
            .withSuccessHandler(res => {
                list.innerHTML = '';
                let totalPagar = 0;
                let totalReceber = 0;

                if (!res || !res.cc || res.cc.length === 0) {
                    const emptyMsg = currentImpersonate
                        ? `<div class="bg-flowly-warning/10 text-amber-700 p-10 rounded-[3rem] text-center text-xs font-black uppercase border border-amber-100 shadow-inner">Sem registos para este cliente.<br><span class="font-normal normal-case text-amber-500 text-[10px] mt-2 block">Os dados são lidos diretamente do Google Sheet do cliente.<br>Contexto: ${currentImpersonate}</span></div>`
                        : `<div class="bg-[#10B981]/20 text-[#10B981] p-10 rounded-2xl text-center text-xs font-black uppercase border border-[#10B981]/40 shadow-inner">Tesouraria Vazia!</div>`;
                    list.innerHTML = emptyMsg;
                    setEl('ccTotalPagar', '0,00 €');
                    setEl('ccTotalReceber', '0,00 €');
                    if (window.lucide) lucide.createIcons();
                    return;
                }

                const diasComSaida = res.diasComSaida || [];
                res.cc.forEach(i => {
                    if (i.estado !== 'Aberto') return;
                    const contaStock = (i.contaStock != null && String(i.contaStock).trim() !== '') ? String(i.contaStock).trim() : 'Sim';
                    if (contaStock === 'Não') return;
                    const val = parseFloat(i.totalPendente) || 0;
                    const tipo = (i.tipo || '').toString().trim().toLowerCase();

                    if (tipo === 'entrada') totalPagar += val;
                    else if (tipo === 'despesa' || tipo === 'despesas' || tipo === 'consumo') totalPagar += val;
                    else if (tipo === 'saida') totalReceber += val;
                    else if (tipo === 'fechocaixa') {
                        const dayKey = i.dataNorm || (i.data != null ? String(i.data) : '');
                        if (diasComSaida.indexOf(dayKey) === -1) totalReceber += val;
                    }
                });

                setEl('ccTotalPagar', totalPagar.toFixed(2).replace('.', ',') + ' €');
                setEl('ccTotalReceber', totalReceber.toFixed(2).replace('.', ',') + ' €');

                const isEntradas = currentCcFlow === 'Entradas';
                let filteredData = res.cc.filter(i => {
                    const t = (i.tipo || '').toString().trim().toLowerCase();
                    if (isEntradas) return t === 'entrada' || t === 'despesa' || t === 'despesas' || t === 'consumo';
                    return t === 'saida' || t === 'fechocaixa';
                });

                if (currentCcTab !== 'Todas') {
                    filteredData = filteredData.filter(i => i.estado === currentCcTab);
                }

                if (filteredData.length === 0) {
                    const msgAberto = (currentCcTab === 'Aberto') ? "Nenhum pendente. Verifique 'Liquidadas' ou 'Visão Total'." : `Sem documentos ${isEntradas ? 'a pagar' : 'a receber'} em estado: ${currentCcTab}`;
                    list.innerHTML = `<div class="bg-slate-50 dark:bg-slate-800 text-slate-400 p-10 rounded-[3rem] text-center text-xs font-black uppercase border border-slate-100 dark:border-slate-700 shadow-inner">${msgAberto}</div>`;
                    return;
                }

                const payLabel = isEntradas ? 'Registar Pag.' : 'Registar pag. cliente';
                window.selectedDocIds = [];
                updateBulkBar();
                renderCCList(filteredData, payLabel);
            }).getMasterData(_ctxEmail());
    }

    function renderCCList(filteredData, payLabel) {
        const list = document.getElementById('ccList');
        if (!list) return;
        const groups = groupCCByDoc(filteredData);
        window.ccGroups = groups;
        window.selectedDocIds = [];

        list.innerHTML = groups.map(function (group, idx) {
            const totalDoc = group.totalDoc.toFixed(2);
            const totalPendente = group.totalPendente.toFixed(2);
            const diffTime = new Date() - parsePTDate(group.data);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            let badgeColor = 'bg-amber-100 text-flowly-warning border border-amber-200';
            let estadoLabel = 'Dentro do prazo';
            if (group.estado === 'Pago') {
                badgeColor = 'bg-[#10B981]/20 text-[#10B981] border border-[#10B981]/40';
                estadoLabel = 'Pago';
            } else if (diffDays > 30) {
                badgeColor = 'bg-rose-100 text-rose-600 border border-rose-200';
                estadoLabel = 'Em Atraso';
            }

            const esc = function (s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
            const forn = esc(group.fornecedor);
            const dataStr = esc(group.dataNorm || group.data);
            const docIdVal = (group.docIds && group.docIds[0]) || group.docId || group.key || '';
            const canSelect = group.canBulkPay && group.estado === 'Aberto';
            const linkFotoRaw = group.linkFoto || (group.items && group.items[0] && group.items[0].linkFoto) || '';
            const linkFoto = typeof linkFotoRaw === 'string' ? linkFotoRaw.trim() : '';
            const linkHref = linkFoto && /^https?:\/\//.test(linkFoto) ? linkFoto : (linkFoto && /^[a-zA-Z0-9.-]+\.[a-z]{2,}/.test(linkFoto) ? 'https://' + linkFoto : '');
            const linkFotoValido = !!linkHref;
            const verFaturaLink = linkFotoValido ? '<a href="' + esc(linkHref) + '" target="_blank" rel="noopener" class="ml-auto flex items-center gap-1 px-3 py-1 bg-slate-700 text-cyan-400 rounded hover:bg-slate-600 transition-colors text-sm"><span>📎</span> Ver Fatura</a>' : '';
            const verItensBtn = '<button onclick="openCCItensModalByIndex(' + idx + ')" class="flex items-center gap-1.5 text-[#06B6D4] hover:underline text-[9px] font-bold uppercase mt-2"><i data-lucide="search" class="w-3.5 h-3.5"></i> Ver Itens</button>';
            const firstItem = group.items[0];
            const payPendente = firstItem ? parseFloat(firstItem.totalPendente || 0).toFixed(2) : totalPendente;

            const payBtn = group.estado === 'Aberto'
                ? '<button onclick="pay(' + (firstItem ? firstItem.rowIndex : 0) + ',' + payPendente + ')" class="bg-[#06B6D4] text-white px-5 py-2.5 rounded-2xl text-[9px] font-black uppercase mt-3 shadow-lg hover:bg-cyan-500 active:scale-95 transition">' + payLabel + '</button>'
                : '';

            const isVendas = payLabel === 'Registar pag. cliente';
            const idEntVal = (group.idEntidade || (firstItem && firstItem.idEntidade) || '').toString().trim();
            const avisoBtn = (isVendas && group.estado === 'Aberto' && idEntVal)
                ? '<button onclick="enviarAvisoIndividual(\'' + esc(idEntVal) + '\')" class="text-[#020617] px-4 py-2.5 rounded-2xl text-[9px] font-black uppercase shadow-lg active:scale-95 transition inline-flex items-center gap-1.5 hover:brightness-95" style="background-color:#FACC15"><i class="fas fa-paper-plane w-3.5 h-3.5"></i> Aviso Cobrança</button>'
                : '';

            const btnsHtml = (payBtn || avisoBtn) ? '<div class="flex flex-wrap gap-2 mt-3 justify-end">' + payBtn + avisoBtn + '</div>' : '';

            return '<div class="flex items-start gap-3 mb-4"><div class="flex items-center pr-3">' +
                '<input type="checkbox" class="cc-selector w-6 h-6 rounded-md border-flowly-border cursor-pointer" style="accent-color: #06B6D4" data-docid="' + esc(docIdVal) + '" data-valor="' + esc(String(group.totalPendente)) + '" onchange="updateBulkAction()"' + (canSelect ? '' : ' disabled') + '></div>' +
                '<div class="flex-1 bg-white dark:bg-slate-800 p-6 rounded-2xl border border-slate-100 dark:border-slate-700 shadow-sm flex gap-4 justify-between items-start hover:border-flowly-primary transition-all">' +
                '<div class="flex gap-3 flex-1 min-w-0">' +
                '<div class="flex flex-col min-w-0">' +
                '<span class="font-black text-xs text-flowly-midnight dark:text-slate-50 uppercase tracking-tighter leading-none">' + forn + '</span>' +
                '<div class="flex items-center gap-2 mt-2 flex-wrap"><span class="text-[9px] text-slate-400 font-black uppercase tracking-widest">' + dataStr + '</span><span class="px-2 py-0.5 rounded-full text-[8px] font-black uppercase tracking-widest ' + badgeColor + '">' + estadoLabel + '</span></div>' +
                '<span class="text-[9px] text-slate-400 font-bold mt-1">Total Doc: ' + totalDoc + ' €</span>' +
                '<div class="flex items-center gap-2 mt-2 flex-wrap">' + verItensBtn + verFaturaLink + '</div></div></div>' +
                '<div class="text-right flex flex-col items-end flex-shrink-0">' +
                '<span class="font-black text-flowly-primary text-xl tracking-tighter leading-none">' + totalPendente + ' €</span>' +
                '<span class="text-[8px] font-black text-slate-400 uppercase tracking-widest mt-1">Pendente</span>' +
                btnsHtml + '</div></div></div></div>';
        }).join('');
        updateBulkAction();
        if (window.lucide) lucide.createIcons();
    }

    function groupCCByDoc(ccData) {
        const map = {};
        ccData.forEach(function (item) {
            const key = (item.docId && item.docId !== 'sem-id') ? item.docId : ((item.dataNorm || '') + '|' + (item.fornecedor || ''));
            if (!map[key]) {
                map[key] = {
                    key: key,
                    docId: item.docId,
                    docIds: [],
                    fornecedor: item.fornecedor || 'N/A',
                    data: item.data,
                    dataNorm: item.dataNorm || '',
                    estado: item.estado || 'Aberto',
                    totalDoc: 0,
                    totalPendente: 0,
                    items: [],
                    canBulkPay: item.docId && item.docId !== 'sem-id',
                    linkFoto: item.linkFoto || ''
                };
            }
            const g = map[key];
            g.totalDoc += parseFloat(item.total) || 0;
            g.totalPendente += parseFloat(item.totalPendente) || 0;
            g.items.push(item);
            if (item.estado === 'Aberto') g.estado = 'Aberto';
            if (item.docId && item.docId !== 'sem-id' && g.docIds.indexOf(item.docId) === -1) g.docIds.push(item.docId);
            if (item.linkFoto && !g.linkFoto) g.linkFoto = item.linkFoto;
            if (item.idEntidade && !g.idEntidade) g.idEntidade = item.idEntidade;
        });
        const groups = Object.values(map);
        return groups;
    }

    function openCCItensModalByIndex(idx) {
        const groups = window.ccGroups || [];
        const group = groups[idx];
        if (!group) return;
        const listEl = document.getElementById('modalCCItensList');
        const modal = document.getElementById('modalCCItens');
        if (!listEl || !modal) return;
        const headerRow = '<div class="grid grid-cols-4 gap-2 py-2 px-3 text-[9px] font-black text-slate-400 uppercase tracking-widest border-b border-flowly-border dark:border-slate-600"><span>Artigo</span><span class="text-center">Qtd</span><span class="text-center">IVA (%)</span><span class="text-right">Total</span></div>';
        listEl.innerHTML = headerRow + (group.items || []).map(function (it) {
            const art = (it.artigo || 'Item').toString().replace(/</g, '&lt;').replace(/>/g, '&gt;');
            const qtdVal = it.qtd != null ? (parseFloat(it.qtd) || 0) : 0;
            const qtdStr = (qtdVal === Math.floor(qtdVal)) ? qtdVal : qtdVal.toFixed(2);
            const ivaStr = (it.taxaIva != null && it.taxaIva !== '') ? (it.taxaIva + '%') : '-';
            const tot = parseFloat(it.total || 0).toFixed(2);
            return '<div class="grid grid-cols-4 gap-2 py-2 px-3 bg-slate-50 dark:bg-slate-800 rounded-2xl items-center"><span class="font-bold text-flowly-midnight dark:text-slate-50 truncate">' + art + '</span><span class="text-center text-sm">' + qtdStr + '</span><span class="text-center text-sm">' + ivaStr + '</span><span class="text-[#06B6D4] font-black text-right">' + tot + ' €</span></div>';
        }).join('');
        modal.classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    }

    // ==========================================
    // 3. AÇÕES EM MASSA (TESOURARIA)
    // ==========================================

    function toggleCCSelection(docId, checked) {
        if (!docId || docId === 'sem-id') return;
        const idx = window.selectedDocIds.indexOf(docId);
        if (checked && idx === -1) window.selectedDocIds.push(docId);
        else if (!checked && idx >= 0) window.selectedDocIds.splice(idx, 1);
        updateBulkBar();
    }

    function toggleBulkAction() {
        updateBulkAction();
    }

    window.updateBulkAction = function updateBulkAction() {
        const checked = document.querySelectorAll('.cc-selector:checked:not([disabled])');
        const selectedIds = [];
        let totalSoma = 0;
        checked.forEach(function (cb) {
            const docId = cb.getAttribute('data-docid');
            const valor = parseFloat(cb.getAttribute('data-valor') || 0) || 0;
            if (docId && docId !== 'sem-id') selectedIds.push(docId);
            totalSoma += valor;
        });
        window.selectedDocIds = selectedIds;
        const bar = document.getElementById('ccBulkBar');
        const countEl = document.getElementById('ccBulkCount');
        const totalEl = document.getElementById('ccBulkTotal');
        if (bar) {
            if (selectedIds.length > 0) bar.classList.remove('hidden');
            else bar.classList.add('hidden');
        }
        const btnAvisos = document.getElementById('ccBulkAvisosBtn');
        if (btnAvisos) btnAvisos.classList.toggle('hidden', currentCcFlow !== 'Saidas' || selectedIds.length === 0);
        if (countEl) countEl.textContent = selectedIds.length + ' faturas selecionadas';
        if (totalEl) totalEl.textContent = 'Total: ' + totalSoma.toFixed(2).replace('.', ',') + ' €';
    };

    function toggleCCSelectionByIndex(groupIdx, checked) {
        const groups = window.ccGroups || [];
        const group = groups[groupIdx];
        if (!group || !group.docIds || group.docIds.length === 0) return;
        const docId = group.docIds[0];
        toggleCCSelection(docId, checked);
    }

    function updateBulkBar() {
        const bar = document.getElementById('ccBulkBar');
        const countEl = document.getElementById('ccBulkCount');
        const totalEl = document.getElementById('ccBulkTotal');
        if (!bar || !countEl) return;
        const docIds = window.selectedDocIds || [];
        const n = docIds.length;
        if (n > 0) {
            let totalPendente = 0;
            const groups = window.ccGroups || [];
            docIds.forEach(function (docId) {
                const g = groups.find(function (gr) { return gr.docId === docId || (gr.docIds && gr.docIds.indexOf(docId) >= 0); });
                if (g) totalPendente += parseFloat(g.totalPendente || 0) || 0;
            });
            bar.classList.remove('hidden');
            countEl.textContent = n + ' faturas selecionadas';
            if (totalEl) totalEl.textContent = 'Total: ' + totalPendente.toFixed(2).replace('.', ',') + ' €';
        } else {
            bar.classList.add('hidden');
            if (totalEl) totalEl.textContent = 'Total: 0,00 €';
        }
    }

    window.openBulkPaymentModal = function openBulkPaymentModal() {
        const selectedIds = window.selectedDocIds || [];
        if (selectedIds.length === 0) return;
        const checked = document.querySelectorAll('.cc-selector:checked:not([disabled])');
        let totalSoma = 0;
        checked.forEach(function (cb) {
            totalSoma += parseFloat(cb.getAttribute('data-valor') || 0) || 0;
        });
        const valorStr = totalSoma.toFixed(2).replace('.', ',') + ' €';
        const textEl = document.getElementById('modalBulkPaymentText');
        const totalEl = document.getElementById('modalBulkTotalValue');
        if (textEl) textEl.textContent = 'A registar pagamento de ' + selectedIds.length + ' faturas num total de ' + valorStr + '.';
        if (totalEl) totalEl.textContent = valorStr;
        const d = new Date();
        const dateInput = document.getElementById('bulkPay_data');
        if (dateInput) dateInput.value = d.getFullYear() + '-' + (d.getMonth() + 1).toString().padStart(2, '0') + '-' + d.getDate().toString().padStart(2, '0');
        document.getElementById('modalBulkPayment').classList.remove('hidden');
    };

    function closeBulkPaymentModal() {
        document.getElementById('modalBulkPayment').classList.add('hidden');
    }

    function confirmBulkPayment() {
        const dateInput = document.getElementById('bulkPay_data');
        const dateVal = (dateInput && dateInput.value) ? dateInput.value.trim() : '';
        if (!dateVal) {
            notify('Indique a data do pagamento.', 'error');
            return;
        }
        var dateStr;
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateVal)) {
            var p = dateVal.split('-');
            dateStr = p[2] + '/' + p[1] + '/' + p[0];
        } else if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateVal)) {
            dateStr = dateVal;
        } else {
            notify('Data inválida. Use o formato DD/MM/AAAA.', 'error');
            return;
        }
        const docIds = window.selectedDocIds.slice();
        if (docIds.length === 0) return;
        closeBulkPaymentModal();
        const bar = document.getElementById('ccBulkBar');
        if (bar) bar.classList.add('hidden');
        window.selectedDocIds = [];
        notify('A processar pagamentos...', 'info');

        google.script.run.withFailureHandler(function (err) {
            notify('Erro: ' + (err || ''), 'error');
            loadCC();
        }).withSuccessHandler(function (res) {
            if (res && res.success) {
                notify('Sucesso', 'success');
            } else {
                notify('Erro: ' + (res ? res.error : 'Resposta inválida'), 'error');
            }
            loadCC();
        }).processBulkPayments(docIds, { dataPagamento: dateStr, impersonateTarget: _ctxEmail() });
    }

    function enviarAvisoIndividual(idEntidade) {
        if (!idEntidade || !idEntidade.trim()) {
            notify('Cliente sem ID de entidade associado. Adicione o cliente na Clientes_DB.', 'error');
            return;
        }
        if (!confirm('Deseja enviar aviso de cobrança a este cliente?')) return;
        notify('A enviar aviso...', 'info');
        google.script.run.withFailureHandler(function (err) {
            notify('Erro: ' + (err || ''), 'error');
        }).withSuccessHandler(function (res) {
            if (res && res.success) {
                if (res.enviados > 0) notify('Aviso enviado com sucesso.', 'success');
                else if (res.falhas > 0 && res.erros && res.erros.length > 0) notify(res.erros[0] || 'Não foi possível enviar.', 'error');
                else notify('Nenhum aviso enviado.', 'warning');
            } else {
                notify(res ? (res.error || 'Erro desconhecido') : 'Erro ao processar', 'error');
            }
        }).processarEmailsCobranca([idEntidade], _ctxEmail());
    }

    function enviarAvisosMassivos() {
        const docIds = window.selectedDocIds || [];
        if (docIds.length === 0) {
            notify('Selecione pelo menos um cliente.', 'warning');
            return;
        }
        const groups = window.ccGroups || [];
        const idsEntidade = [];
        const seen = {};
        docIds.forEach(function (docId) {
            const g = groups.find(function (gr) { return gr.docId === docId || (gr.docIds && gr.docIds.indexOf(docId) >= 0); });
            if (g && g.idEntidade && g.idEntidade.trim() && !seen[g.idEntidade]) {
                seen[g.idEntidade] = true;
                idsEntidade.push(g.idEntidade);
            }
        });
        if (idsEntidade.length === 0) {
            notify('Nenhum dos clientes selecionados tem ID de entidade. Adicione os clientes na Clientes_DB.', 'error');
            return;
        }
        if (!confirm('Deseja enviar aviso de cobrança a ' + idsEntidade.length + ' cliente(s) selecionado(s)?')) return;
        notify('A enviar avisos...', 'info');
        google.script.run.withFailureHandler(function (err) {
            notify('Erro: ' + (err || ''), 'error');
        }).withSuccessHandler(function (res) {
            if (res && res.success) {
                var msg = res.enviados + ' aviso(s) enviado(s).';
                if (res.falhas > 0 && res.erros && res.erros.length > 0) msg += ' ' + res.falhas + ' falha(s): ' + (res.erros[0] || '');
                notify(msg, res.enviados > 0 ? 'success' : (res.falhas > 0 ? 'warning' : 'info'));
            } else {
                notify(res ? (res.error || 'Erro desconhecido') : 'Erro ao processar', 'error');
            }
        }).processarEmailsCobranca(idsEntidade, _ctxEmail());
    }

    function pay(rowIndex, pendenteAtual) {
        document.getElementById('pay_rowIndex').value = rowIndex;
        document.getElementById('pay_maxPendente').value = pendenteAtual;
        document.getElementById('pay_valor').value = pendenteAtual;

        const d = new Date();
        document.getElementById('pay_data').value = `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getFullYear()}`;

        document.getElementById('pay_warning').classList.add('hidden');
        document.getElementById('pay_error').classList.add('hidden');
        document.getElementById('btnConfirmPay').disabled = false;
        document.getElementById('btnConfirmPay').classList.remove('opacity-50', 'cursor-not-allowed');

        document.getElementById('modalPayment').classList.remove('hidden');
    }

    function validatePaymentAmount() {
        const val = parseFloat(document.getElementById('pay_valor').value);
        const maxPendente = parseFloat(document.getElementById('pay_maxPendente').value);
        const btn = document.getElementById('btnConfirmPay');
        const warnEl = document.getElementById('pay_warning');
        const errEl = document.getElementById('pay_error');

        if (isNaN(val) || val <= 0) {
            btn.disabled = true;
            btn.classList.add('opacity-50', 'cursor-not-allowed');
            warnEl.classList.add('hidden');
            errEl.classList.add('hidden');
            return;
        }
        if (val > maxPendente) {
            errEl.classList.remove('hidden');
            warnEl.classList.add('hidden');
            btn.disabled = true;
            btn.classList.add('opacity-50', 'cursor-not-allowed');
        } else if (val < maxPendente) {
            document.getElementById('pay_restante').innerText = (maxPendente - val).toFixed(2);
            warnEl.classList.remove('hidden');
            errEl.classList.add('hidden');
            btn.disabled = false;
            btn.classList.remove('opacity-50', 'cursor-not-allowed');
        } else {
            warnEl.classList.add('hidden');
            errEl.classList.add('hidden');
            btn.disabled = false;
            btn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
    }

    function confirmPayment() {
        const rowIndex = document.getElementById('pay_rowIndex').value;
        const val = parseFloat(document.getElementById('pay_valor').value);
        const dateStr = document.getElementById('pay_data').value;
        const maxPendente = parseFloat(document.getElementById('pay_maxPendente').value);

        if (!rowIndex || isNaN(val) || val <= 0 || !dateStr) return notify('Preencha os dados corretamente.', 'warning');
        if (val > maxPendente) return notify('O valor excede o pendente!', 'error');

        document.getElementById('modalPayment').classList.add('hidden');
        notify('A processar pagamento...', 'info');

        google.script.run
            .withFailureHandler(err => notify('Erro: ' + err, 'error'))
            .withSuccessHandler(res => {
                if (res && res.success) { notify('Pagamento registado!', 'success'); loadCC(); }
                else notify('Erro: ' + (res ? res.error : 'Resposta inválida'), 'error');
            }).processPayment(rowIndex, val, dateStr, _ctxEmail());
    }

    // ==========================================
    // 4. LOGÍSTICA - SCANNER OCR E MANUAL
    // ==========================================

    function setMode(mode) {
        const sections = ['docModule', 'multiModule', 'manualModule'];
        const buttons = ['btnModoDoc', 'btnModoMulti', 'btnModoManual'];
        sections.forEach(s => {
            const el = document.getElementById(s);
            if (el) el.classList.add('hidden');
        });
        buttons.forEach(b => {
            const btn = document.getElementById(b);
            if (btn) btn.className = "flex-1 py-3 rounded-xl text-[10px] font-bold bg-white text-slate-500 border border-slate-50 transition-all";
        });
        const activeSec = mode === 'doc' ? 'docModule' : (mode === 'multi' ? 'multiModule' : 'manualModule');
        if (document.getElementById(activeSec)) document.getElementById(activeSec).classList.remove('hidden');
        const bAct = document.getElementById(mode === 'doc' ? 'btnModoDoc' : (mode === 'multi' ? 'btnModoMulti' : 'btnModoManual'));
        if (bAct) bAct.className = "flex-1 py-3 rounded-xl text-[10px] font-bold bg-flowly-midnight text-white shadow-xl scale-[1.05] transition-all";
        if (mode === 'multi') renderMultiModule();
        if (mode === 'manual') renderManualForm();
        lucide.createIcons();
    }

    function handleMultiUpload(e) {
        const files = e.target.files;
        if (!files.length) return;

        const maxSize = 10 * 1024 * 1024;
        const tooBig = Array.from(files).find(f => f.size > maxSize);
        if (tooBig) {
            notify("⚠️ Ficheiro demasiado pesado! Por favor, carrega imagens menores para garantir o processamento.", "error");
            e.target.value = '';
            return;
        }

        const imageFiles = Array.from(files).filter(f => (f.type || '').startsWith('image/'));
        const pdfCount = Array.from(files).length - imageFiles.length;
        if (pdfCount > 0) {
            notify("Apenas imagens (JPG/PNG) são processadas por agora. PDFs serão suportados em breve.", "info");
        }
        if (!imageFiles.length) {
            notify("Seleciona pelo menos uma imagem (JPG/PNG).", "warning");
            e.target.value = '';
            return;
        }

        document.getElementById('artigosContainer').innerHTML = '';
        document.getElementById('validationForm').classList.remove('hidden');
        document.getElementById('loadingUI').classList.remove('hidden');
        const total = imageFiles.length;
        const counterEl = document.getElementById('loadingCounter');
        if (counterEl) counterEl.textContent = 'A processar dados (1/' + total + ')...';

        let firstDataFilled = false;
        const processNext = (idx) => {
            if (idx >= total) {
                document.getElementById('loadingUI').classList.add('hidden');
                notify("Leitura Concluída!", "success");
                return;
            }
            if (counterEl) counterEl.textContent = 'A processar dados (' + (idx + 1) + '/' + total + ')...';

            const file = imageFiles[idx];
            const reader = new FileReader();
            reader.onload = (ev) => {
            compressImageForOCR(ev.target.result, 1000, 0.75).then(compressed => {
                    google.script.run
                        .withFailureHandler(() => {
                            notify("Foto " + (idx + 1) + " não foi lida. A continuar com as restantes...", "warning");
                            processNext(idx + 1);
                        })
                        .withSuccessHandler(res => {
                            if (res.success && res.data && Array.isArray(res.data)) {
                                res.data.forEach(item => {
                                    addArtigoRow(item.artigo, item.quantidade, item.preco_custo, item.taxa_iva);
                                    if (!firstDataFilled && (item.data_doc || item.fornecedor)) {
                                        const fData = document.getElementById('f_data');
                                        const fForn = document.getElementById('f_fornecedor');
                                        if (item.data_doc && fData) {
                                            let d = String(item.data_doc || '').trim();
                                            if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
                                                const p = d.split('-');
                                                d = p[2] + '/' + p[1] + '/' + p[0];
                                            }
                                            if (d && /^\d{2}\/\d{2}\/\d{4}$/.test(d)) { fData.value = d; }
                                        }
                                        if (item.fornecedor && fForn) fForn.value = item.fornecedor;
                                        firstDataFilled = true;
                                    }
                                });
                            } else if (!res.success) {
                                notify("Foto " + (idx + 1) + " não foi lida. A continuar com as restantes...", "warning");
                            }
                            processNext(idx + 1);
                        })
                        .extractDataWithAI([compressed]);
                });
            };
            reader.readAsDataURL(file);
        };
        processNext(0);
    }

    function renderMultiModule() {
        document.getElementById('multiModule').innerHTML = `<div class="bg-white p-8 rounded-[3rem] shadow-xl border border-slate-50 animate-slide-up relative overflow-hidden"><div class="relative z-10 aspect-[4/3] bg-flowly-midnight rounded-[2.5rem] flex flex-col items-center justify-center overflow-hidden mb-6"><input type="file" id="multiOCRInput" multiple accept="image/*,application/pdf" class="absolute inset-0 opacity-0 z-50 cursor-pointer" onchange="handleMultiUpload(event)"><div class="bg-white/10 p-8 rounded-full border border-white/20 scanner-pulse"><i data-lucide="layers" class="w-12 h-12 text-white"></i></div><p class="text-white font-black uppercase text-xs mt-6 tracking-widest">Multi-Scanner AI</p><p class="text-[9px] text-slate-400 mt-1">Carregue várias fotos de uma vez</p></div></div>`;
        lucide.createIcons();
    }

    function renderManualForm() {
        const container = document.getElementById('manualModule');
        container.innerHTML = `<div class="bg-white p-6 rounded-[2.5rem] shadow-2xl border border-slate-50 space-y-5 animate-slide-up">
        <div class="grid grid-cols-2 gap-4">
            <div class="space-y-1.5">
                <p class="text-[9px] font-black text-slate-400 uppercase ml-2 tracking-widest">Data (DD/MM/AAAA)</p>
                <input type="text" id="m_date" inputmode="numeric" maxlength="10" placeholder="DD/MM/AAAA"
                    class="w-full bg-slate-50 p-4 rounded-2xl font-bold text-xs outline-none border-2 border-transparent focus:border-flowly-primary transition"
                    oninput="this.value=this.value.replace(/[^\\d\\/]/g,'').replace(/^(\\d{2})(\\d)$/,'$1/$2').replace(/^(\\d{2}\\/\\d{2})(\\d)$/,'$1/$2')"
                    onblur="validateDate(this)">
            </div>
            <div id="m_blocoFornecedorCliente" class="space-y-1.5">
                <p id="m_labelFornecedorCliente" class="text-[9px] font-black text-slate-400 uppercase ml-2 tracking-widest">Fornecedor</p>
                <div class="flex gap-2">
                    <input id="m_forn" list="listaFornecedores"
                        class="flex-1 bg-slate-50 p-4 rounded-2xl font-bold text-xs outline-none border-2 border-transparent focus:border-flowly-primary transition"
                        placeholder="Fornecedor" onblur="validateDatalistFornecedorCliente('m')">
                    <button type="button" onclick="showPrompt('Novo Fornecedor','listaFornecedores','m_forn')" class="bg-flowly-primary text-white px-4 rounded-2xl font-bold shadow-md active:scale-95 transition text-lg">+</button>
                </div>
            </div>
            <div class="space-y-1.5">
                <p class="text-[9px] font-black text-slate-400 uppercase ml-2 tracking-widest">Tipo de Lançamento</p>
                <select id="m_tipo" class="w-full bg-slate-50 p-4 rounded-2xl font-bold text-xs outline-none border-2 border-transparent focus:border-flowly-primary transition text-flowly-midnight" onchange="updateTipoLancamento('m'); updateValidationTotal()">
                    <option value="Entrada">Entrada</option>
                    <option value="Saída">Saída</option>
                    <option value="Quebra">Quebra</option>
                    <option value="Oferta">Oferta</option>
                    <option value="Despesas">Despesas</option>
                    <option value="Fecho de caixa/Relatório">Fecho de caixa/Relatório</option>
                </select>
            </div>
            <div class="space-y-1.5">
                <p class="text-[9px] font-black text-slate-400 uppercase ml-2 tracking-widest">Fatura Paga?</p>
                <select id="m_ja_paga" class="w-full bg-slate-50 p-4 rounded-2xl font-bold text-xs outline-none border-2 border-transparent focus:border-flowly-primary transition text-flowly-midnight" onchange="toggleMDataPagamento()">
                    <option value="0">Não (Aberto)</option>
                    <option value="1">Sim (Pago)</option>
                </select>
            </div>
            <div class="space-y-1.5">
                <label class="flex items-center gap-3 p-4 bg-slate-50 rounded-2xl cursor-pointer border border-slate-100">
                    <input type="checkbox" id="m_stock" class="w-5 h-5 accent-flowly-primary rounded" checked>
                    <span class="text-sm font-bold text-flowly-midnight">Contar para Stock?</span>
                </label>
            </div>
        </div>
        <div id="m_bloco_data_pagamento" class="hidden grid grid-cols-2 gap-4">
            <div class="space-y-1.5">
                <p class="text-[9px] font-black text-slate-400 uppercase ml-2 tracking-widest">Data do pagamento (DD/MM/AAAA)</p>
                <input type="text" id="m_data_pagamento" maxlength="10" placeholder="DD/MM/AAAA" class="w-full bg-slate-50 p-4 rounded-2xl font-bold text-xs outline-none border-2 border-transparent focus:border-flowly-primary transition" oninput="this.value=this.value.replace(/[^\\d\\/]/g,'').replace(/^(\\d{2})(\\d)$/,'$1/$2').replace(/^(\\d{2}\\/\\d{2})(\\d)$/,'$1/$2')">
            </div>
        </div>
        <div id="m_rows" class="space-y-3 pt-4 border-t border-slate-50">
            <div class="flex justify-between items-center px-1">
                <p class="text-[10px] font-black text-flowly-midnight uppercase tracking-widest">Artigos</p>
                <button onclick="addManualRow()" class="text-flowly-primary font-black text-[9px] uppercase border px-4 py-1.5 rounded-full">+ Nova</button>
            </div>
        </div>
        <div id="m_validationSummary" class="mt-4 p-4 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
            <div class="flex justify-between items-center text-sm font-bold">
                <span class="text-slate-500">Soma dos Itens (com IVA):</span>
                <span id="m_somaArtigos" class="text-flowly-midnight">0.00 €</span>
            </div>
            <div class="flex justify-between items-center gap-2">
                <span class="text-slate-500 text-sm font-bold">Total da Fatura:</span>
                <input id="m_totalDoc" type="number" step="0.01" placeholder="0.00" class="w-24 p-2 rounded-lg text-sm font-bold border-2 text-right" oninput="updateValidationTotal()">
                <span class="text-sm font-bold">€</span>
            </div>
            <p id="m_totalStatus" class="text-[10px] font-bold hidden"></p>
        </div>
        <button id="btnSubmitManual" onclick="submitManualLog()" class="w-full bg-flowly-midnight text-white py-6 rounded-[2rem] font-black uppercase text-xs tracking-widest shadow-2xl active:scale-95 transition-all mt-6">Guardar</button>
    </div>`;
        addManualRow();
        lucide.createIcons();
    }

    function handleInvoiceUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        pendingInvoiceName = `Fatura_${Date.now()}_${file.name}`;
        e.target.value = '';

        notify(`A analisar documento com IA...`, "info");
        document.getElementById('loadingUI').classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = (ev) => {
            compressImageForOCR(ev.target.result, 1000, 0.75).then(compressed => {
                pendingInvoiceB64 = compressed;
                google.script.run
                    .withFailureHandler(err => { document.getElementById('loadingUI').classList.add('hidden'); notify("Erro OCR: " + err, "error"); })
                    .withSuccessHandler(ocrRes => {
                        document.getElementById('loadingUI').classList.add('hidden');
                        if (!ocrRes.success) return notify("Erro na AI: " + (ocrRes.error || ""), "error");

                        pendingTipoDocumento = ocrRes.tipoDocumento || "FaturaCompra";
                        const data = ocrRes.data || {};
                        const cabecalho = data.cabecalho || { nif: data.nif || "", fornecedor: data.fornecedor || "", data: data.data || "" };
                        let linhas = data.linhas;
                        if (!linhas || !Array.isArray(linhas) || linhas.length === 0) {
                            linhas = [{ artigo: data.artigo || "Resumo", quantidade: 1, preco_custo: parseFloat(data.valor_base) || parseFloat(data.valor_total) || 0, preco_venda: 0, taxa_iva: parseInt(data.taxa_iva) || 23, valor_iva: parseFloat(data.valor_iva) || 0 }];
                        }

                        document.getElementById('iv_forn').value = cabecalho.fornecedor || "";
                        const nifInput = document.getElementById('iv_nif');
                        const nifValue = cabecalho.nif || "";
                        nifInput.value = nifValue;
                        if (nifValue && typeof isValidNIFPT === 'function' && !isValidNIFPT(nifValue)) {
                            nifInput.classList.remove('border-transparent', 'border-flowly-border');
                            nifInput.classList.add('border-rose-500');
                            notify("Aviso: o NIF detetado parece ser inválido.", "warning");
                        } else {
                            nifInput.classList.remove('border-rose-500');
                        }
                        document.getElementById('iv_data').value = cabecalho.data || "";
                        
                        document.getElementById('modalInvoiceValidation').classList.remove('hidden');
                        if (window.lucide) lucide.createIcons();

                        renderInvoiceLines(linhas);

                        const tipoSelect = document.getElementById('iv_tipo');
                        if (tipoSelect) {
                            if (pendingTipoDocumento === "FechoCaixa") tipoSelect.value = "Fecho de caixa/Relatório";
                            else if (pendingTipoDocumento === "RelatorioSaidas") tipoSelect.value = "Saída";
                            else tipoSelect.value = "Entrada";
                        }
                        const jaPagaSelect = document.getElementById('iv_ja_paga');
                        if (jaPagaSelect) jaPagaSelect.value = "0";
                        const contaStockEl = document.getElementById('iv_stock');
                        if (contaStockEl) contaStockEl.checked = true;

                        document.getElementById('iv_erro_valor_pago').classList.add('hidden');
                        toggleIvDataPagamento();

                        const titleEl = document.getElementById('iv_modal_title');
                        const subEl = document.getElementById('iv_modal_subtitle');
                        if (titleEl) titleEl.textContent = pendingTipoDocumento === "FechoCaixa" ? "Validar Fecho de Caixa" : (pendingTipoDocumento === "RelatorioSaidas" ? "Validar Relatório de Saídas" : "Validar Fatura");
                        if (subEl) subEl.textContent = "Confirme e corrija os dados antes de gravar. Pode editar todos os campos.";
                    }).extractDocument(pendingInvoiceB64);
            });
        };
        reader.readAsDataURL(file);
    }

    function compressImageForOCR(dataUrl, maxSize, quality) {
        maxSize = maxSize || 1200; 
        quality = 0.6; // Valor estrito solicitado
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                let w = img.width, h = img.height;
                if (w > maxSize || h > maxSize) {
                    if (w > h) { h = Math.round(h * maxSize / w); w = maxSize; }
                    else { w = Math.round(w * maxSize / h); h = maxSize; }
                }
                const canvas = document.createElement('canvas');
                canvas.width = w; canvas.height = h;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, w, h);
                // Força JPEG 0.6 para garantir ficheiros leves < 800kb
                resolve(canvas.toDataURL('image/jpeg', quality));
            };
            img.onerror = () => resolve(dataUrl);
            img.src = dataUrl;
        });
    }

    // ==========================================
    // 5. GESTÃO DE LINHAS (DOCUMENTOS E MANUAL)
    // ==========================================

    function addInvoiceLineRow(line) {
        line = line || { artigo: "", quantidade: 1, preco_custo: 0, preco_venda: 0, taxa_iva: 23, valor_iva: 0 };
        const id = "iv_line_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6);
        const container = document.getElementById('iv_linhas_container');
        const div = document.createElement('div');
        div.className = 'grid grid-cols-[3fr_1fr_2fr_2fr_1.5fr_1.5fr_auto] gap-2 items-end p-3 bg-slate-50 rounded-xl border border-slate-100';
        div.dataset.lineId = id;
        div.innerHTML = `
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">Artigo</label><input type="text" class="iv_line_artigo w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${(line.artigo || "").toString().replace(/"/g, '&quot;')}" placeholder="Nome item" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">Qtd</label><input type="number" step="0.01" class="iv_line_qty w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${line.quantidade != null ? line.quantidade : 1}" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">Custo (€)</label><input type="text" inputmode="decimal" class="iv_line_preco w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${line.preco_custo != null ? line.preco_custo : 0}" placeholder="0,00" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">Venda (€)</label><input type="text" inputmode="decimal" class="iv_line_venda w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${line.preco_venda != null ? line.preco_venda : 0}" placeholder="0,00" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">IVA (%)</label><input type="number" class="iv_line_taxa w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${line.taxa_iva != null ? line.taxa_iva : 23}" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">Val IVA (€)</label><input type="text" inputmode="decimal" class="iv_line_iva w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${line.valor_iva != null ? line.valor_iva : 0}" placeholder="0,00" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase block mb-1">&nbsp;</label><button type="button" data-action="remove-line" class="p-2 rounded-lg bg-rose-50 text-rose-500 hover:bg-rose-100 flex items-center justify-center" title="Remover linha"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div>
    `;
        container.appendChild(div);
        if (window.lucide) lucide.createIcons();
    }

    function removeInvoiceLineRow(btn) {
        const row = btn.closest('[data-line-id]');
        if (row && document.getElementById('iv_linhas_container').children.length > 1) {
            row.remove();
            updateValidationTotal();
        }
    }

    (function initInvoiceLineDelegation() {
        const container = document.getElementById('iv_linhas_container');
        if (container) {
            container.addEventListener('click', function (e) {
                const btn = e.target.closest('[data-action="remove-line"]');
                if (btn) removeInvoiceLineRow(btn);
            });
            // Event Delegation para recálculo instantâneo
            container.addEventListener('input', function (e) {
                updateValidationTotal();
            });
        }
    })();

    function renderInvoiceLines(linhas) {
        const container = document.getElementById('iv_linhas_container');
        if (!container) return;
        container.innerHTML = '';
        if (linhas && linhas.length) linhas.forEach(l => addInvoiceLineRow(l));
        else addInvoiceLineRow({ artigo: "", quantidade: 1, preco_custo: 0, taxa_iva: 23, valor_iva: 0 });
        updateValidationTotal();

        container.querySelectorAll('.iv_line_qty, .iv_line_preco, .iv_line_venda').forEach(el => {
            el.dispatchEvent(new Event('input', { bubbles: true }));
        });
    }

    function getInvoiceLinesFromModal() {
        const container = document.getElementById('iv_linhas_container');
        if (!container) return [];
        const rows = container.querySelectorAll('[data-line-id]');
        const _sn = (v) => standardizeNumeric(v);
        return Array.from(rows).map(row => ({
            artigo: (row.querySelector('.iv_line_artigo') || {}).value || "",
            quantidade: _sn((row.querySelector('.iv_line_qty') || {}).value) || 1,
            preco_custo: _sn((row.querySelector('.iv_line_preco') || {}).value),
            preco_venda: _sn((row.querySelector('.iv_line_venda') || {}).value),
            taxa_iva: parseInt((row.querySelector('.iv_line_taxa') || {}).value) || 23,
            valor_iva: _sn((row.querySelector('.iv_line_iva') || {}).value)
        })).filter(l => (l.artigo || "").toString().trim() !== "" || l.preco_custo > 0 || l.preco_venda > 0);
    }

    function getInvoiceModalTotal() {
        const linhas = getInvoiceLinesFromModal();
        const tipo = (document.getElementById('iv_tipo') || {}).value || 'Entrada';
        const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
        return linhas.reduce((sum, l) => {
            const qty = parseFloat(l.quantidade) || 0;
            const val = isSaida ? (parseFloat(l.preco_venda) || 0) : (parseFloat(l.preco_custo) || 0);
            return sum + qty * val;
        }, 0);
    }

    function onIvTipoChange() {
        const tipo = (document.getElementById('iv_tipo') || {}).value || 'Entrada';
        const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
        if (!isSaida) return;
        const container = document.getElementById('iv_linhas_container');
        if (!container) return;
        container.querySelectorAll('[data-line-id]').forEach(row => {
            const custoInput = row.querySelector('.iv_line_preco');
            const vendaInput = row.querySelector('.iv_line_venda');
            if (!custoInput || !vendaInput) return;
            const custoVal = parseFloat(custoInput.value) || 0;
            custoInput.value = '0';
            vendaInput.value = custoVal;
        });
    }

    function toggleIvDataPagamento() {
        const jaPaga = document.getElementById('iv_ja_paga');
        const bloco = document.getElementById('iv_bloco_data_pagamento');
        const dataPag = document.getElementById('iv_data_pagamento');
        const valorPago = document.getElementById('iv_valor_pago');
        const errEl = document.getElementById('iv_erro_valor_pago');
        if (!bloco || !jaPaga) return;
        if (jaPaga.value === "1") {
            bloco.classList.remove('hidden');
            const today = new Date();
            if (dataPag) dataPag.value = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
            
            const linhas = getInvoiceLinesFromModal();
            const tipo = (document.getElementById('iv_tipo') || {}).value || 'Entrada';
            const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
            
            let soma = linhas.reduce((acc, l) => {
                const q = standardizeNumeric(l.quantidade);
                const val = isSaida ? standardizeNumeric(l.preco_venda) : standardizeNumeric(l.preco_custo);
                const t = parseInt(l.taxa_iva) || 23;
                return acc + q * val * (1 + t / 100);
            }, 0);
            soma = Math.round(soma * 100) / 100;
            if (valorPago) {
                valorPago.value = soma.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                valorPago.readOnly = false;
            }
        } else {
            bloco.classList.add('hidden');
            if (dataPag) dataPag.value = '';
            if (valorPago) {
                valorPago.value = '';
                valorPago.readOnly = false;
            }
        }
        if (errEl) errEl.classList.add('hidden');
    }

    function toggleFDataPagamento() {
        const jaPaga = document.getElementById('f_ja_paga');
        const bloco = document.getElementById('f_bloco_data_pagamento');
        const dataPag = document.getElementById('f_data_pagamento');
        if (!bloco || !jaPaga) return;
        if (jaPaga.value === "1") {
            bloco.classList.remove('hidden');
            const today = new Date();
            if (dataPag) dataPag.value = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
        } else {
            bloco.classList.add('hidden');
            if (dataPag) dataPag.value = '';
        }
    }

    function toggleMDataPagamento() {
        const jaPaga = document.getElementById('m_ja_paga');
        const bloco = document.getElementById('m_bloco_data_pagamento');
        const dataPag = document.getElementById('m_data_pagamento');
        if (!bloco || !jaPaga) return;
        if (jaPaga.value === "1") {
            bloco.classList.remove('hidden');
            const today = new Date();
            if (dataPag) dataPag.value = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
        } else {
            bloco.classList.add('hidden');
            if (dataPag) dataPag.value = '';
        }
    }

    function getManualFormArtigos() {
        const tipo = (document.getElementById('m_tipo') || {}).value || 'Entrada';
        const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
        const artigosInputs = document.querySelectorAll('#m_rows .i-artigo');
        return Array.from(artigosInputs).map(artInput => {
            const row = artInput.closest('.bg-slate-50');
            if (!row) return null;
            const artigo = (row.querySelector('.i-artigo') || {}).value || "";
            const qtd = parseFloat((row.querySelector('.i-qtd') || {}).value) || 0;
            const preco = parseFloat((row.querySelector('.i-preco') || {}).value) || 0;
            const taxa = parseInt((row.querySelector('.i-taxa') || {}).value) || 23;
            const val = isSaida ? preco : preco;
            const valorIva = qtd * val * (taxa / 100);
            return {
                artigo,
                quantidade: qtd,
                preco_custo: isSaida ? 0 : preco,
                preco_venda: isSaida ? preco : 0,
                taxa_iva: taxa,
                valor_iva: valorIva
            };
        }).filter(l => l && ((l.artigo && l.artigo.trim()) || l.preco_custo > 0 || l.preco_venda > 0 || l.quantidade > 0));
    }

    function addManualRow(artigo, quantidade, preco_custo, taxa_iva) {
        artigo = artigo != null ? String(artigo) : "";
        quantidade = parseFloat(quantidade) || 1;
        preco_custo = parseFloat(preco_custo) || 0;
        taxa_iva = parseInt(taxa_iva) || 23;
        const row = document.createElement('div');
        row.className = 'bg-slate-50 p-4 rounded-2xl border border-slate-100 flex flex-col gap-3 mb-2 animate-slide-up';
        row.innerHTML = `<div class="flex gap-2">
        <input class="flex-1 bg-white p-4 rounded-xl text-xs font-bold outline-none border-2 border-transparent focus:border-flowly-primary transition i-artigo" placeholder="Artigo" list="listaArtigos" value="${(artigo || "").replace(/"/g, '&quot;')}" onblur="validateDatalist(this,'listaArtigos')" oninput="updateValidationTotal()">
        <button type="button" onclick="showPrompt('Novo Artigo','listaArtigos',this.previousElementSibling)" class="bg-flowly-primary text-white px-4 rounded-xl font-bold shadow-md active:scale-95 transition text-lg">+</button>
    </div>
    <div class="grid grid-cols-3 gap-3">
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Qtd</label><input class="w-full bg-white p-4 rounded-xl text-xs font-bold outline-none border-2 border-transparent focus:border-flowly-primary i-qtd" placeholder="Qtd" inputmode="decimal" value="${quantidade}" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Preço</label><input class="w-full bg-white p-4 rounded-xl text-xs font-bold outline-none border-2 border-transparent focus:border-flowly-primary i-preco" placeholder="Preço" inputmode="decimal" value="${preco_custo}" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Taxa IVA (%)</label><input type="number" class="w-full bg-white p-4 rounded-xl text-xs font-bold outline-none border-2 border-transparent focus:border-flowly-primary i-taxa" placeholder="23" value="${taxa_iva}" oninput="updateValidationTotal()"></div>
    </div>`;
        document.getElementById('m_rows').appendChild(row);
        updateValidationTotal();
    }

    function addArtigoRow(artigo, quantidade, preco_custo, taxa_iva) {
        artigo = artigo != null ? String(artigo) : "";
        quantidade = parseFloat(quantidade) || 1;
        preco_custo = parseFloat(preco_custo) || 0;
        taxa_iva = parseInt(taxa_iva) || 23;
        const row = document.createElement('div');
        row.className = 'artigo-row grid grid-cols-[repeat(6,minmax(0,1fr))] gap-2 items-end p-3 bg-slate-50 rounded-xl border border-slate-100 mb-2 animate-slide-up';
        const valorIvaInicial = (quantidade * preco_custo * (taxa_iva / 100)).toFixed(2);
        row.innerHTML = `
        <div class="col-span-2"><label class="text-[9px] font-bold text-slate-400 uppercase">Artigo</label><input type="text" class="i-artigo w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${(artigo || "").replace(/"/g, '&quot;')}" placeholder="Nome item" list="listaArtigos" onblur="validateDatalist(this,'listaArtigos')" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Qtd</label><input type="number" step="0.01" class="i-qtd w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${quantidade}" placeholder="Qtd" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Preço Custo</label><input type="number" step="0.01" class="i-preco w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${preco_custo}" placeholder="Preço" oninput="updateValidationTotal()"></div>
        <div><label class="text-[9px] font-bold text-slate-400 uppercase">Taxa IVA (%)</label><input type="number" class="i-taxa w-full bg-white p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${taxa_iva}" placeholder="23" oninput="updateValidationTotal()"></div>
        <div class="flex gap-1 items-end"><div class="flex-1"><label class="text-[9px] font-bold text-slate-400 uppercase">Valor IVA (€)</label><input type="text" readonly class="i-valoriva w-full bg-slate-100 p-2 rounded-lg text-xs font-bold border border-flowly-border" value="${valorIvaInicial}" placeholder="0.00"></div><button type="button" data-action="remove-artigo-row" class="p-2 rounded-lg bg-rose-50 text-rose-500 hover:bg-rose-100 text-xs self-end" title="Remover linha"><i data-lucide="trash-2" class="w-4 h-4"></i></button></div>
    `;
        document.getElementById('artigosContainer').appendChild(row);
        if (window.lucide) lucide.createIcons();
        updateValidationTotal();
    }

    (function initArtigoRowDelegation() {
        const container = document.getElementById('artigosContainer');
        if (container) container.addEventListener('click', function (e) {
            const btn = e.target.closest('[data-action="remove-artigo-row"]');
            if (btn) {
                const row = btn.closest('.artigo-row');
                if (row && document.querySelectorAll('#artigosContainer .artigo-row').length > 1) {
                    row.remove();
                    updateValidationTotal();
                }
            }
        });
    })();

    function getValidationFormArtigos() {
        const rows = document.querySelectorAll('#artigosContainer .artigo-row');
        return Array.from(rows).map(row => {
            const artigo = (row.querySelector('.i-artigo') || {}).value || "";
            const qtd = parseFloat((row.querySelector('.i-qtd') || {}).value) || 0;
            const preco = parseFloat((row.querySelector('.i-preco') || {}).value) || 0;
            const taxa = parseInt((row.querySelector('.i-taxa') || {}).value) || 23;
            const valorIva = parseFloat((row.querySelector('.i-valoriva') || {}).value) || (qtd * preco * (taxa / 100));
            return { artigo, quantidade: qtd, preco_custo: preco, taxa_iva: taxa, valor_iva: valorIva };
        }).filter(l => (l.artigo && l.artigo.trim()) || l.quantidade > 0 || l.preco_custo > 0);
    }

    function updateValidationTotal() {
        if (typeof standardizeNumeric !== 'function') return;
        const round2 = v => Math.round((v || 0) * 100) / 100;
        const fmt = v => {
            if (typeof standardizeNumeric !== 'function') return;
            return standardizeNumeric(v, true) + ' €';
        };
        const _sn = (v) => standardizeNumeric(v);
        const TOLERANCIA_AVISO = 0.10;

        const applyStatus = (statusEl, totalInput, soma, saveBtn) => {
            if (!statusEl || !totalInput) return false;
            const totalVal = _sn(totalInput.value);

            if (totalVal <= 0) {
                statusEl.classList.add('hidden');
                totalInput.classList.remove('border-rose-500', 'border-amber-400', 'border-emerald-500', 'bg-rose-50', 'bg-amber-50', 'bg-emerald-50');
                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-amber-500');
                    saveBtn.classList.add('bg-flowly-primary');
                    saveBtn.dataset.overrideDiff = '';
                    saveBtn.innerHTML = 'GUARDAR';
                }
                return true;
            }

            const diff = Math.abs(round2(soma) - round2(totalVal));
            statusEl.classList.remove('hidden');
            totalInput.classList.remove('border-rose-500', 'border-amber-400', 'border-emerald-500', 'bg-rose-50', 'bg-amber-50', 'bg-emerald-50');

            if (diff <= TOLERANCIA_AVISO) {
                statusEl.textContent = '✓ Totais conferem (dif. ' + diff.toFixed(2) + '€)';
                statusEl.className = 'text-[10px] font-bold text-emerald-600';
                totalInput.classList.add('border-emerald-500', 'bg-emerald-50');
                if (saveBtn) {
                    saveBtn.disabled = false;
                    saveBtn.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-amber-500');
                    saveBtn.classList.add('bg-flowly-primary');
                    saveBtn.dataset.overrideDiff = '';
                    saveBtn.innerHTML = 'GUARDAR';
                }
                return true;
            } else {
                statusEl.textContent = '⚠ Diferença: ' + diff.toFixed(2) + '€ entre soma e total. Verifique antes de guardar.';
                statusEl.className = 'text-[10px] md:text-sm font-bold text-amber-600';
                totalInput.classList.add('border-amber-400', 'bg-amber-50');
                if (saveBtn) {
                    saveBtn.classList.remove('bg-flowly-primary');
                    saveBtn.classList.add('bg-amber-500');
                    saveBtn.dataset.overrideDiff = 'warned';
                    saveBtn.innerHTML = '⚠ Confirmar Registo';
                }
                return false;
            }
        };

        // 1. validationForm (Scanner Único/Multi)
        if (document.getElementById('validationSummary') && document.getElementById('artigosContainer')) {
            document.querySelectorAll('#artigosContainer .artigo-row').forEach(row => {
                const q = _sn(row.querySelector('.i-qtd').value);
                const p = _sn(row.querySelector('.i-preco').value);
                const t = parseInt((row.querySelector('.i-taxa') || {}).value) || 23;
                const valIva = round2(q * p * (t / 100));
                const el = row.querySelector('.i-valoriva');
                if (el) el.value = valIva.toFixed(2);
            });
            const artigos = getValidationFormArtigos();
            let soma = artigos.reduce((acc, l) => {
                const q = _sn(l.quantidade);
                const p = _sn(l.preco_custo);
                const t = parseInt(l.taxa_iva) || 23;
                return acc + q * p * (1 + t / 100);
            }, 0);
            soma = round2(soma);
            const somaEl = document.getElementById('validationSomaArtigos');
            const totalInput = document.getElementById('validationTotalDoc');
            const statusEl = document.getElementById('validationTotalStatus');
            const saveBtn = document.getElementById('saveBtn');
            if (somaEl) somaEl.textContent = fmt(soma);
            applyStatus(statusEl, totalInput, soma, saveBtn);
        }

        // 2. modalInvoiceValidation (iv_linhas_container)
        if (document.getElementById('iv_validationSummary') && document.getElementById('modalInvoiceValidation') && !document.getElementById('modalInvoiceValidation').classList.contains('hidden')) {
            console.log("Iniciando soma das linhas...");
            
            const container = document.getElementById('iv_linhas_container');
            const rows = container ? container.querySelectorAll('[data-line-id]') : [];
            const tipo = (document.getElementById('iv_tipo') || {}).value || 'Entrada';
            const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
            
            let soma = 0;
            rows.forEach(row => {
                const inputQty = row.querySelector('.iv_line_qty');
                const inputPreco = isSaida ? row.querySelector('.iv_line_venda') : row.querySelector('.iv_line_preco');
                const rawQty = inputQty ? inputQty.value : "0";
                const rawPreco = inputPreco ? inputPreco.value : "0";
                
                console.log("Validação de captura de seletores:", { inputQty, inputPreco, rawQty, rawPreco });

                const q = _sn(rawQty);
                const val = _sn(rawPreco);
                const t = parseInt((row.querySelector('.iv_line_taxa') || {}).value) || 23;
                
                console.log(`Linha: Qtd Bruto=${rawQty}, Preco Bruto=${rawPreco} | Qtd Final=${q}, Preco Final=${val}`);
                
                soma += (q * val * (1 + t / 100));
            });
            
            soma = round2(soma);
            const somaEl = document.getElementById('iv_somaArtigos');
            const totalInput = document.getElementById('iv_totalDoc');
            const statusEl = document.getElementById('iv_totalStatus');
            const saveBtn = document.getElementById('btnConfirmInvoiceValidation');
            
            if (somaEl) somaEl.textContent = fmt(soma);
            
            // SINCRONIZAÇÃO AUTOMÁTICA: O total do documento segue a soma das linhas
            if (totalInput) {
                totalInput.value = soma.toFixed(2);
            }
            
            applyStatus(statusEl, totalInput, soma, saveBtn);

            const jaPagaEl = document.getElementById('iv_ja_paga');
            const valorPagoEl = document.getElementById('iv_valor_pago');
            if (jaPagaEl && jaPagaEl.value === '1' && valorPagoEl) {
                // Quando Pago, o valor pago é a soma calculada, mas permanece editável
                valorPagoEl.value = soma.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                valorPagoEl.readOnly = false;
            }
        }

        // 3. manualModule (m_rows)
        if (document.getElementById('m_validationSummary')) {
            const artigos = getManualFormArtigos();
            const tipo = (document.getElementById('m_tipo') || {}).value || 'Entrada';
            const isSaida = (tipo === 'Saída' || tipo === 'Saida' || tipo === 'Quebra' || tipo === 'Oferta' || tipo === 'Despesas' || tipo === 'Fecho de caixa/Relatório');
            let soma = artigos.reduce((acc, l) => {
                const q = _sn(l.quantidade);
                const val = isSaida ? _sn(l.preco_venda) : _sn(l.preco_custo);
                const t = parseInt(l.taxa_iva) || 23;
                return acc + q * val * (1 + t / 100);
            }, 0);
            soma = round2(soma);
            const somaEl = document.getElementById('m_somaArtigos');
            const totalInput = document.getElementById('m_totalDoc');
            const statusEl = document.getElementById('m_totalStatus');
            const saveBtn = document.getElementById('btnSubmitManual');
            if (somaEl) somaEl.textContent = fmt(soma);
            applyStatus(statusEl, totalInput, soma, saveBtn);
        }
    }

    // ==========================================
    // 6. VALIDAÇÕES UTILITÁRIAS
    // ==========================================

    function validateDate(inputEl) {
        if (!inputEl.value.trim()) return;
        var reDDMMYYYY = new RegExp('^\\d{2}\\/\\d{2}\\/\\d{4}$');
        if (!reDDMMYYYY.test(inputEl.value.trim())) {
            notify('Data inválida. Use o formato DD/MM/AAAA', 'error');
            inputEl.classList.add('input-error');
            inputEl.value = '';
            setTimeout(() => inputEl.classList.remove('input-error'), 2000);
        }
    }

    function validateDatalist(inputEl, datalistId) {
        const dl = document.getElementById(datalistId);
        if (!dl || !inputEl.value.trim()) return;
        const opts = Array.from(dl.options).map(o => o.value.toLowerCase());
        if (!opts.includes(inputEl.value.trim().toLowerCase())) {
            notify('Valor "' + inputEl.value + '" não existe na lista. Use (+) para adicionar.', 'warning');
            inputEl.classList.add('input-error');
            inputEl.value = '';
            setTimeout(() => inputEl.classList.remove('input-error'), 2000);
        }
    }

    function validateDatalistFornecedorCliente(formPrefix) {
        var tipoEl = document.getElementById(formPrefix === 'f' ? 'f_tipo' : 'm_tipo');
        var listId = (tipoEl && tipoEl.value === 'Saída') ? 'listaClientes' : 'listaFornecedores';
        var input = document.getElementById(formPrefix === 'f' ? 'f_fornecedor' : 'm_forn');
        if (input) validateDatalist(input, listId);
    }

    function updateTipoLancamento(formPrefix) {
        var tipoEl = document.getElementById(formPrefix === 'f' ? 'f_tipo' : 'm_tipo');
        var tipoSelecionado = (tipoEl && tipoEl.value) ? tipoEl.value : 'Entrada';
        var bloco = document.getElementById(formPrefix === 'f' ? 'f_blocoFornecedorCliente' : 'm_blocoFornecedorCliente');
        var label = document.getElementById(formPrefix === 'f' ? 'f_labelFornecedorCliente' : 'm_labelFornecedorCliente');
        var input = document.getElementById(formPrefix === 'f' ? 'f_fornecedor' : 'm_forn');
        var btn = (formPrefix === 'f') ? document.getElementById('f_btnAddFornecedorCliente') : document.querySelector('#m_blocoFornecedorCliente button');
        var listId = (tipoSelecionado === 'Saída') ? 'listaClientes' : 'listaFornecedores';
        var promptLabel = (tipoSelecionado === 'Saída') ? 'Novo Cliente' : 'Novo Fornecedor';
        var inputId = (formPrefix === 'f') ? 'f_fornecedor' : 'm_forn';

        if (tipoSelecionado === 'Fecho de caixa/Relatório') {
            if (bloco) bloco.classList.add('hidden');
        } else {
            if (bloco) bloco.classList.remove('hidden');
            if (label) label.textContent = (tipoSelecionado === 'Saída') ? 'Cliente' : 'Fornecedor';
            if (input) {
                input.setAttribute('list', listId);
                input.placeholder = (tipoSelecionado === 'Saída') ? 'Cliente' : 'Fornecedor';
            }
            if (btn) {
                btn.setAttribute('onclick', "showPrompt('" + promptLabel + "', '" + listId + "', '" + inputId + "')");
                btn.textContent = '+';
            }
        }
    }

    // ==========================================
    // 7. DATALISTS
    // ==========================================

    function loadFornecedoresDatalist() {
        google.script.run.withSuccessHandler(function (list) {
            var dl = document.getElementById('listaFornecedores');
            if (dl && list && list.length > 0) {
                dl.innerHTML = list.map(function (v) { return '<option value="' + (v || '').replace(/"/g, '&quot;') + '">'; }).join('');
            }
        }).withFailureHandler(function () { }).getFornecedoresForDatalist(_ctxEmail());
    }

    function loadClientesDatalist() {
        google.script.run.withSuccessHandler(function (list) {
            var dl = document.getElementById('listaClientes');
            if (dl && list && list.length > 0) {
                dl.innerHTML = list.map(function (v) { return '<option value="' + (v || '').replace(/"/g, '&quot;') + '">'; }).join('');
            }
        }).withFailureHandler(function () { }).getClientesForDatalist(_ctxEmail());
    }

    // ==========================================
    // 8. SUBMISSÕES E GRAVAÇÃO
    // ==========================================

    function submitManualLog() {
        const saveBtn = document.getElementById('btnSubmitManual');
        if (saveBtn && saveBtn.disabled) {
            notify("Corrija a diferença entre a soma e o total da fatura (máx. 0.10€).", "error");
            return;
        }

        const data = (document.getElementById('m_date') || {}).value || "";
        const fornecedor = (document.getElementById('m_forn') || {}).value || "";
        const tipo = (document.getElementById('m_tipo') || {}).value || "Entrada";
        const jaPago = (document.getElementById('m_ja_paga') || {}).value === "1";
        const dataPagamento = (document.getElementById('m_data_pagamento') || {}).value || "";
        const artigos = getManualFormArtigos();

        if (!data || !/^\d{2}\/\d{2}\/\d{4}$/.test(data.trim())) {
            notify("Indique a data do documento (DD/MM/AAAA).", "error");
            return;
        }
        if (!artigos.length) {
            notify("Adicione pelo menos um artigo com quantidade e preço.", "error");
            return;
        }
        if (jaPago && (!dataPagamento || !/^\d{2}\/\d{2}\/\d{4}$/.test(dataPagamento.trim()))) {
            notify("Indique a data do pagamento (DD/MM/AAAA).", "error");
            return;
        }

        const mStockEl = document.getElementById('m_stock');
        const contaStock = (mStockEl && mStockEl.checked) ? 'Sim' : 'Não';

        const p = {
            data: data.trim(),
            tipo: tipo,
            metodo: "Manual",
            fornecedor: fornecedor.trim(),
            observacoes: "",
            artigos: artigos,
            impersonateEmail: _ctxEmail(),
            jaPago: jaPago,
            dataPagamento: jaPago ? dataPagamento.trim() : "",
            contaStock: contaStock
        };

        notify("A registar...", "info");
        google.script.run
            .withFailureHandler(err => notify("Erro: " + err, "error"))
            .withSuccessHandler(res => {
                if (res && res.success) {
                    notify("Registo manual guardado!", "success");
                    nav('cc');
                    loadCC();
                } else {
                    notify("Erro ao gravar: " + (res ? res.error : "Resposta inválida"), "error");
                }
            }).saveValidatedData(p);
    }

    function confirmInvoiceValidation() {
        const saveBtn = document.getElementById('btnConfirmInvoiceValidation');
        if (saveBtn && saveBtn.dataset.overrideDiff === 'warned') {
            saveBtn.dataset.overrideDiff = 'confirmed';
            saveBtn.innerHTML = 'GRAVAR MESMO ASSIM';
            saveBtn.classList.remove('bg-amber-500');
            saveBtn.classList.add('bg-rose-600');
            notify("Atenção: Diferença superior a 0.10€ detetada. Clique novamente para confirmar.", "warning");
            return;
        }

        const cabecalho = {
            fornecedor: (document.getElementById('iv_forn') || {}).value || "",
            nif: (document.getElementById('iv_nif') || {}).value || "",
            data: (document.getElementById('iv_data') || {}).value || ""
        };

        const linhas = getInvoiceLinesFromModal();
        if (!linhas.length) {
            notify("Adicione pelo menos uma linha (artigo e/ou preço).", "error");
            return;
        }

        const tipo = (document.getElementById('iv_tipo') || {}).value || "Entrada";
        const jaPago = (document.getElementById('iv_ja_paga') || {}).value === "1";
        const dataPagamento = (document.getElementById('iv_data_pagamento') || {}).value || "";
        const errEl = document.getElementById('iv_erro_valor_pago');
        if (errEl) errEl.classList.add('hidden');

        // Calcula soma real das linhas para usar como valor pago (nunca OCR)
        const linhasConfirm = getInvoiceLinesFromModal();
        const tipoConfirm = tipo;
        const isSaidaConfirm = (tipoConfirm === 'Saída' || tipoConfirm === 'Saida' || tipoConfirm === 'Quebra' || tipoConfirm === 'Oferta' || tipoConfirm === 'Despesas' || tipoConfirm === 'Fecho de caixa/Relatório');
        const _pfC = (v) => { if (!v && v !== 0) return 0; const s = String(v).replace(/\s/g,''); const ep = /^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s); return ep ? parseFloat(s.replace(/\./g,'').replace(',','.')) || 0 : parseFloat(s.replace(',','.')) || 0; };
        const somaConfirm = Math.round(linhasConfirm.reduce((acc, l) => {
            const q = _pfC(l.quantidade); const val = isSaidaConfirm ? _pfC(l.preco_venda) : _pfC(l.preco_custo); const t = parseInt(l.taxa_iva) || 23;
            return acc + q * val * (1 + t / 100);
        }, 0) * 100) / 100;

        if (jaPago) {
            if (!dataPagamento || !/^\d{2}\/\d{2}\/\d{4}$/.test(dataPagamento.trim())) {
                notify("Indique a data do pagamento (DD/MM/AAAA).", "error");
                return;
            }
        }

        document.getElementById('modalInvoiceValidation').classList.add('hidden');
        document.getElementById('loadingUI').classList.remove('hidden');
        notify(`A enviar ficheiro para a Drive...`, "info");

        const contaStockEl = document.getElementById('iv_stock');
        const contaStock = (contaStockEl && contaStockEl.checked) ? 'Sim' : 'Não';
        // Valor pago = sempre a soma calculada das linhas
        const opcoes = jaPago ? { jaPago: true, dataPagamento: dataPagamento.trim(), valorPago: somaConfirm, contaStock: contaStock } : { contaStock: contaStock };

        google.script.run
            .withFailureHandler(err => { document.getElementById('loadingUI').classList.add('hidden'); notify("Erro Drive: " + err, "error"); })
            .withSuccessHandler(driveRes => {
                if (!driveRes.success) { document.getElementById('loadingUI').classList.add('hidden'); return notify("Erro Upload: " + driveRes.error, "error"); }
                notify(`A registar...`, "info");
                google.script.run
                    .withFailureHandler(err => { document.getElementById('loadingUI').classList.add('hidden'); notify("Erro DB: " + err, "error"); })
                    .withSuccessHandler(dbRes => {
                        document.getElementById('loadingUI').classList.add('hidden');
                        if (dbRes.success) {
                            notify("Documento validado e registado!", "success");
                            pendingInvoiceB64 = null;
                            pendingInvoiceName = "";
                            pendingTipoDocumento = "FaturaCompra";
                            loadCC();
                        } else {
                            notify("Erro ao gravar: " + dbRes.error, "error");
                        }
                    })
                    .saveDocumentToDB(tipo, cabecalho, linhas, driveRes.url, _ctxEmail(), opcoes);
            }).uploadInvoiceToDrive(pendingInvoiceB64, pendingInvoiceName, _ctxEmail());
    }

    function cancelInvoiceValidation() {
        document.getElementById('modalInvoiceValidation').classList.add('hidden');
        const ivTotal = document.getElementById('iv_totalDoc');
        if (ivTotal) ivTotal.value = '';
        pendingInvoiceB64 = null;
        pendingInvoiceName = "";
        pendingTipoDocumento = "FaturaCompra";
        notify("Registo cancelado pelo utilizador.", "info");
    }

    function submitLogistica() {
        const saveBtn = document.getElementById('saveBtn');
        if (saveBtn && saveBtn.dataset.overrideDiff === 'warned') {
            saveBtn.dataset.overrideDiff = 'confirmed';
            saveBtn.innerHTML = 'CONFIRMAR GRAVAÇÃO (VIÉS)';
            saveBtn.classList.remove('bg-amber-500');
            saveBtn.classList.add('bg-rose-600');
            notify("Atenção: Diferença superior a 0.10€ detetada. Clique novamente para confirmar.", "warning");
            return;
        }

        const tipo = (document.getElementById('f_tipo') || {}).value || "Entrada";
        const data = (document.getElementById('f_data') || {}).value || "";
        const fornecedor = (document.getElementById('f_fornecedor') || {}).value || "";
        const stockEl = document.getElementById('f_stock');
        const stock = (stockEl && stockEl.checked) ? "Sim" : "Não";
        const jaPago = (document.getElementById('f_ja_paga') || {}).value === "1";
        const dataPagamento = (document.getElementById('f_data_pagamento') || {}).value || "";
        const artigos = getValidationFormArtigos();

        if (!data || !/^\d{2}\/\d{2}\/\d{4}$/.test(data.trim())) {
            notify("Indique a data do documento (DD/MM/AAAA).", "error");
            return;
        }
        if (!artigos.length) {
            notify("Adicione pelo menos um artigo com quantidade ou preço.", "error");
            return;
        }
        if (jaPago && (!dataPagamento || !/^\d{2}\/\d{2}\/\d{4}$/.test(dataPagamento.trim()))) {
            notify("Indique a data do pagamento (DD/MM/AAAA).", "error");
            return;
        }

        const dados = {
            tipo: tipo,
            data: data.trim(),
            fornecedor: fornecedor.trim(),
            stock: stock,
            jaPago: jaPago,
            dataPagamento: jaPago ? dataPagamento.trim() : "",
            artigos: artigos,
            impersonateEmail: _ctxEmail()
        };

        notify("A registar...", "info");
        google.script.run
            .withFailureHandler(err => notify("Erro ao gravar: " + (err || ""), "error"))
            .withSuccessHandler(res => {
                if (res && res.success) {
                    cancelarRegisto();
                    nav('logistica');
                    loadCC();
                    notify("Registo guardado!", "success");
                } else {
                    notify("Erro ao gravar: " + (res ? res.error : "Resposta inválida"), "error");
                }
            })
            .saveLogisticaConsolidated(dados);
    }

    function cancelarRegisto() {
        document.getElementById('artigosContainer').innerHTML = '';
        document.getElementById('validationForm').classList.add('hidden');

        const fData = document.getElementById('f_data');
        const fForn = document.getElementById('f_fornecedor');
        const fTipo = document.getElementById('f_tipo');
        const fStock = document.getElementById('f_stock');
        const fJaPaga = document.getElementById('f_ja_paga');
        const fDataPag = document.getElementById('f_data_pagamento');
        const fBloco = document.getElementById('f_bloco_data_pagamento');

        if (fData) fData.value = '';
        if (fForn) fForn.value = '';
        if (fTipo) fTipo.value = 'Entrada';
        if (fStock) fStock.checked = true;
        if (fJaPaga) fJaPaga.value = '0';
        if (fDataPag) fDataPag.value = '';
        if (fBloco) fBloco.classList.add('hidden');

        const totalDoc = document.getElementById('validationTotalDoc');
        if (totalDoc) totalDoc.value = '';
        const multiInput = document.getElementById('multiOCRInput');
        if (multiInput) multiInput.value = '';
    }

    // ==========================================
    // EXPORTAÇÃO DE INTERFACE PARA WINDOW
    // ==========================================
    
    // Exportar todas as funções que são chamadas pela UI
    window.setCcFlow = setCcFlow;
    window.loadCC = loadCC;
    window.filterCcTab = filterCcTab;
    window.setMode = setMode;
    window.handleInvoiceUpload = handleInvoiceUpload;
    window.submitLogistica = submitLogistica;
    window.updateValidationTotal = updateValidationTotal;
    window.toggleCCSelection = toggleCCSelection;
    window.toggleBulkAction = toggleBulkAction;
    window.toggleCCSelectionByIndex = toggleCCSelectionByIndex;
    window.updateBulkBar = updateBulkBar;
    window.openBulkPaymentModal = window.openBulkPaymentModal;
    window.closeBulkPaymentModal = closeBulkPaymentModal;
    window.confirmBulkPayment = confirmBulkPayment;
    window.enviarAvisoIndividual = enviarAvisoIndividual;
    window.enviarAvisosMassivos = enviarAvisosMassivos;
    window.pay = pay;
    window.validatePaymentAmount = validatePaymentAmount;
    window.confirmPayment = confirmPayment;
    window.handleMultiUpload = handleMultiUpload;
    window.renderMultiModule = renderMultiModule;
    window.renderManualForm = renderManualForm;
    window.addInvoiceLineRow = addInvoiceLineRow;
    window.removeInvoiceLineRow = removeInvoiceLineRow;
    window.renderInvoiceLines = renderInvoiceLines;
    window.getInvoiceLinesFromModal = getInvoiceLinesFromModal;
    window.getInvoiceModalTotal = getInvoiceModalTotal;
    window.onIvTipoChange = onIvTipoChange;
    window.toggleIvDataPagamento = toggleIvDataPagamento;
    window.toggleFDataPagamento = toggleFDataPagamento;
    window.toggleMDataPagamento = toggleMDataPagamento;
    window.getManualFormArtigos = getManualFormArtigos;
    window.addManualRow = addManualRow;
    window.addArtigoRow = addArtigoRow;
    window.getValidationFormArtigos = getValidationFormArtigos;
    window.validateDate = validateDate;
    window.validateDatalist = validateDatalist;
    window.validateDatalistFornecedorCliente = validateDatalistFornecedorCliente;
    window.updateTipoLancamento = updateTipoLancamento;
    window.loadFornecedoresDatalist = loadFornecedoresDatalist;
    window.loadClientesDatalist = loadClientesDatalist;
    window.submitManualLog = submitManualLog;
    window.confirmInvoiceValidation = confirmInvoiceValidation;
    window.cancelInvoiceValidation = cancelInvoiceValidation;
    window.cancelarRegisto = cancelarRegisto;
    window.openCCItensModalByIndex = openCCItensModalByIndex;
    window.groupCCByDoc = groupCCByDoc;
    window.renderCCList = renderCCList;
    window.compressImageForOCR = compressImageForOCR;
    
})();
