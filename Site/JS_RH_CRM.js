(function() {
    'use strict';
    
    // ==========================================
    // 1. VARIÁVEIS GLOBAIS E UTILITÁRIOS
    // ==========================================
    let _editingStaffEmail = null;
    var _auditClienteId = null;
    var _auditFornecedorId = null;

    /** Parser PT-PT: aceita números com vírgula decimal e ponto como separador de milhares */
    function parsePTFloat(val) {
        if (val === null || val === undefined || val === '') return 0;
        if (typeof val === 'number') return val;
        const str = String(val).replace(/\s/g, '');
        const cleanStr = str.replace(/\./g, '').replace(',', '.');
        const num = parseFloat(cleanStr);
        return isNaN(num) ? 0 : num;
    }

    // ==========================================
    // 2. RECURSOS HUMANOS E SALÁRIOS
    // ==========================================

    function loadRH() {
        const list = document.getElementById('staffListRh') || document.getElementById('staffList');
        if (!list) return;
        list.innerHTML = '<div class="text-center py-20 opacity-30 text-[10px] font-black uppercase tracking-widest animate-pulse">A carregar colaboradores...</div>';

        google.script.run
            .withFailureHandler(err => {
                list.innerHTML = '<div class="p-10 bg-rose-50 dark:bg-rose-900/20 text-rose-500 rounded-[2rem] text-center border border-rose-100 dark:border-rose-800"><p class="font-black text-xs">' + err + '</p></div>';
            })
            .withSuccessHandler(res => {
                list.innerHTML = '';
                staffCache = res.staff || [];
                if (res.users) { window._usersCache = res.users; window.usersCache = res.users; }

                const ativos = staffCache.filter(s => s.status === 'Ativo');
                const numColaboradores = ativos.length;
                const diasUteis = (typeof getWorkingDaysInMonth === 'function') ? getWorkingDaysInMonth(new Date().getMonth() + 1, new Date().getFullYear()) : 22;
                const tsuPct = 23.75;

                // Recalcula custoMensalReal a partir dos componentes
                const calcCustoReal = (s) => {
                    const base = parsePTFloat(s.vencimento) || 0;
                    const premios = parsePTFloat(s.premios) || 0;
                    const subAlim = parsePTFloat(s.subAlim) || 0;
                    const seguro = parsePTFloat(s.seguro) || 0;
                    const provFerias = parsePTFloat(s.provFerias) || (base / 12);
                    const provNatal = parsePTFloat(s.provNatal) || (base / 12);
                    const tsu = parsePTFloat(s.tsuPct) || tsuPct;
                    const baseIncidenciaTSU = base + premios + provFerias + provNatal;
                    const tsuValor = baseIncidenciaTSU * (tsu / 100);
                    return base + premios + (subAlim * diasUteis) + seguro + provFerias + provNatal + tsuValor;
                };

                // Provisões: cap para evitar valores inflacionados na sheet
                const provRescCap = (s) => { const v = parsePTFloat(s.vencimento) || 0; const p = parsePTFloat(s.provRescisao) || 0; return (p > v * 1.5) ? (v * 0.055) : p; };
                const provFormCap = (s) => { const v = parsePTFloat(s.vencimento) || 0; const p = parsePTFloat(s.provFormacao) || 0; return (p > v * 0.5) ? (v * 0.02) : p; };

                const totalCustoMensal = ativos.reduce((a, s) => a + (parsePTFloat(s.vencimento) > 0 ? calcCustoReal(s) : parsePTFloat(s.custoMensalReal)), 0);
                const totalProvisoes = ativos.reduce((a, s) => a + provRescCap(s) + provFormCap(s), 0);
                const totalComProvisao = totalCustoMensal + totalProvisoes;

                if (staffCache.length === 0) {
                    list.innerHTML = '<div class="text-center py-16"><p class="text-slate-400 dark:text-slate-500 font-bold text-sm">Nenhum colaborador registado.</p><p class="text-[10px] text-slate-300 dark:text-slate-600 mt-1">Use o botão + para adicionar.</p></div>';
                } else {
                    staffCache.forEach(s => {
                        const custo = parsePTFloat(s.custoMensalReal || s.custoTotal);
                        const custoStr = custo.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
                        const bloqueado = (s.status === 'Bloqueado' || s.status === 'Inativo');
                        const statusColor = s.status === 'Ativo' ? 'bg-emerald-100 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 border-emerald-200 dark:border-emerald-800' : (s.status === 'Pendente' ? 'bg-amber-100 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 border-amber-200 dark:border-amber-800' : 'bg-rose-100 dark:bg-rose-900/30 text-rose-600 dark:text-rose-400 border-rose-200 dark:border-rose-800');
                        const idEsc = (s.id || '').replace(/'/g, "\\'");
                        const nomeEsc = (s.nome || 'N/A').replace(/'/g, "\\'");
                        const semProvisoes = (!s.provFerias && !s.provNatal);
                        const sugestaoIA = semProvisoes && s.vencimento > 0 ? '<div class="mt-2 p-2 bg-cyan-50 dark:bg-cyan-900/20 rounded-lg border border-cyan-100 dark:border-cyan-800"><p class="text-[8px] font-black text-cyan-600 dark:text-cyan-400 uppercase">Sugestão IA</p><p class="text-[9px] text-cyan-700 dark:text-cyan-300 mt-0.5">Férias/Natal: ' + (s.vencimento / 12).toFixed(0) + '€/mês · Rescisão: ' + (s.vencimento * 0.055).toFixed(0) + '€ · Formação: ' + (s.vencimento * 0.02).toFixed(0) + '€</p></div>' : '';

                        list.innerHTML += '<div class="theme-card p-5 rounded-2xl border shadow-sm mb-3 animate-slide-up transition-all hover:shadow-md ' + (bloqueado ? 'opacity-60' : '') + '">'
                            + '<div class="flex justify-between items-start">'
                            + '<div class="flex items-center gap-3">'
                            + '<div class="h-11 w-11 bg-slate-100 dark:bg-slate-700 rounded-2xl flex items-center justify-center font-black text-flowly-primary text-lg">' + (s.nome || '?').charAt(0).toUpperCase() + '</div>'
                            + '<div><p class="font-black text-sm text-slate-900 dark:text-slate-50">' + (s.nome || '—') + '</p>'
                            + '<p class="text-[10px] text-slate-400 dark:text-slate-500 font-bold uppercase tracking-widest">' + (s.cargo || 'Colaborador') + '</p></div>'
                            + '</div>'
                            + '<span class="px-2.5 py-0.5 ' + statusColor + ' text-[8px] font-black rounded-full border">' + (s.status || '—') + '</span>'
                            + '</div>'
                            + (function () {
                                const isPendente = s.status === 'Pendente';
                                const isInativo = (s.status === 'Inativo' || s.status === 'Bloqueado');
                                const fator = s.fatorProporcional != null ? s.fatorProporcional : (isPendente ? 0 : 100);
                                const custoRecalc = (parsePTFloat(s.vencimento) > 0) ? calcCustoReal(s) : parsePTFloat(s.custoMensalReal || s.custoBase);
                                const custoBaseVal = custoRecalc;
                                const custoEfetivo = Math.round(custoRecalc * (fator / 100) * 100) / 100;
                                const custoDisplay = (isPendente ? custoBaseVal : (s.status === 'Ativo' ? custoBaseVal : custoEfetivo)).toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';

                                let label = 'Custo Mensal Real';
                                let bgClass = 'from-flowly-midnight to-slate-800';
                                let propTag = '';

                                if (isPendente) { label = 'Custo Projetado (Aguarda Ativação)'; bgClass = 'from-amber-600 to-amber-500'; }
                                else if (isInativo && fator > 0 && fator < 100) { label = 'Custo Proporcional (' + fator + '% do mês)'; bgClass = 'from-rose-700 to-rose-500'; propTag = '<p class="text-[7px] opacity-70 mt-0.5">Base: ' + custoBaseVal.toLocaleString('pt-PT', { minimumFractionDigits: 2 }) + ' €/mês</p>'; }
                                else if (isInativo && fator === 0) { label = 'Sem Custo (Saiu)'; bgClass = 'from-slate-500 to-slate-400'; }
                                else if (isInativo) { label = 'Custo Mensal (Inativo)'; bgClass = 'from-rose-700 to-rose-500'; }

                                const dataSaidaTag = s.dataSaida ? '<p class="text-[7px] opacity-70 mt-0.5">Saída: ' + s.dataSaida + '</p>' : '';

                                return '<div class="mt-3 bg-gradient-to-r ' + bgClass + ' p-3 rounded-2xl text-white flex justify-between items-center">'
                                    + '<div><p class="text-[8px] font-bold opacity-60 uppercase tracking-widest">' + label + '</p>'
                                    + '<p class="font-black text-lg tracking-tight">' + custoDisplay + '</p>' + propTag + dataSaidaTag + '</div>'
                                    + '<div class="text-right text-[9px] opacity-70">'
                                    + '<p>Base: ' + (s.vencimento || 0).toFixed(0) + '€ · TSU: ' + (s.tsuPct || 23.75) + '%</p>'
                                    + '<p>Alim: ' + (s.subAlim || 0).toFixed(2) + '€/dia · Seg: ' + (s.seguro || 0).toFixed(0) + '€ · Prém: ' + (s.premios || 0).toFixed(0) + '€</p></div></div>';
                            })()
                            + sugestaoIA
                            + '<div class="flex gap-2 mt-3">'
                            + '<button onclick="openStaffForm(\'' + idEsc + '\')" class="flex-1 py-2 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-xl text-[10px] font-bold hover:bg-slate-200 dark:hover:bg-slate-600 transition"><i data-lucide="edit-3" class="w-3 h-3 inline mr-1"></i>Editar</button>'
                            + (s.status === 'Ativo' ? '<button onclick="deactivateStaffClick(\'' + idEsc + '\',\'' + nomeEsc + '\')" class="py-2 px-3 bg-flowly-warning/10 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 rounded-xl text-[10px] font-bold hover:bg-amber-100 dark:hover:bg-amber-900/50 transition" title="Desativar"><i data-lucide="user-x" class="w-3 h-3 inline"></i></button>' : '')
                            + '<button onclick="deleteStaffClick(\'' + idEsc + '\',\'' + nomeEsc + '\')" class="py-2 px-3 bg-rose-50 dark:bg-rose-900/30 text-rose-600 dark:text-rose-400 rounded-xl text-[10px] font-bold hover:bg-rose-100 dark:hover:bg-rose-900/50 transition" title="Eliminar"><i data-lucide="trash-2" class="w-3 h-3 inline"></i></button>'
                            + '</div></div>';
                    });
                }

                const rhTotalColab = document.getElementById('rhTotalColaboradores');
                if (rhTotalColab) rhTotalColab.innerText = String(numColaboradores);
                const rhTotalSalarios = document.getElementById('rhTotalSalarios');
                if (rhTotalSalarios) rhTotalSalarios.innerText = totalCustoMensal.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
                const rhTotal = document.getElementById('rhTotalCostDis');
                if (rhTotal) rhTotal.innerText = totalComProvisao.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';

                if (window.lucide) lucide.createIcons();
            }).getMasterData(_ctxEmail());
    }

    function calcStaffPreview() {
        try {
            const gv = (id) => { const el = document.getElementById(id); return parsePTFloat(el ? el.value : '') || 0; };
            const base = gv('sf_venc');
            const premios = gv('sf_premios');
            const tsu = 23.75;
            const subAlim = gv('sf_alim');
            const seguro = gv('sf_seg');

            const preview = document.getElementById('sfCalcPreview');
            if (base <= 0) { if (preview) preview.classList.add('hidden'); return; }
            if (preview) preview.classList.remove('hidden');

            const diasUteis = (typeof getWorkingDaysInMonth === 'function') ? getWorkingDaysInMonth(new Date().getMonth() + 1, new Date().getFullYear()) : 22;

            const provFerias = base / 12;
            const provNatal = base / 12;
            const baseIncidenciaTSU = base + premios + provFerias + provNatal;
            const tsuValor = baseIncidenciaTSU * (tsu / 100);
            const custoMensalReal = base + premios + (subAlim * diasUteis) + seguro + provFerias + provNatal + tsuValor;

            const provRescisao = base * 0.055;
            const provFormacao = base * 0.02;
            const provLongo = provRescisao + provFormacao;
            const totalComProvisao = custoMensalReal + provLongo;

            const set = (id, val) => { const el = document.getElementById(id); if (el) el.innerText = val; };
            set('sfPrevFerias', provFerias.toFixed(0) + '€');
            set('sfPrevNatal', provNatal.toFixed(0) + '€');
            set('sfPrevRescisao', provRescisao.toFixed(0) + '€');
            set('sfPrevFormacao', provFormacao.toFixed(0) + '€');
            set('sfPrevCusto', custoMensalReal.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');
            set('sfPrevProvLongo', provLongo.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');
            set('sfPrevCustoTotal', totalComProvisao.toLocaleString('pt-PT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');

            const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val; };
            setVal('sf_prov_ferias', provFerias.toFixed(2));
            setVal('sf_prov_natal', provNatal.toFixed(2));
            setVal('sf_prov_rescisao', provRescisao.toFixed(2));
            setVal('sf_prov_formacao', provFormacao.toFixed(2));
        } catch (e) { console.error('calcStaffPreview error:', e); }
    }

    function saveStaff() {
        const idField = document.getElementById('sf_id');
        const idVal = idField ? idField.value : "";
        const gv = (elId) => { const el = document.getElementById(elId); return el ? el.value : ''; };
        const gnum = (elId) => { const el = document.getElementById(elId); return parsePTFloat(el ? el.value : '') || 0; };

        const p = {
            id: (idVal === "" || idVal === "null" || idVal === undefined) ? "NOVO" : idVal,
            nome: gv('sf_nome'),
            email: gv('sf_email'),
            nif: gv('sf_nif'),
            cargo: gv('sf_cargo'),
            vencimento: gnum('sf_venc'),
            subAlim: gnum('sf_alim'),
            seguro: gnum('sf_seg'),
            tsuPct: 23.75,
            status: gv('sf_status') || 'Ativo',
            provFerias: gnum('sf_prov_ferias'),
            provNatal: gnum('sf_prov_natal'),
            provRescisao: gnum('sf_prov_rescisao'),
            provFormacao: gnum('sf_prov_formacao'),
            admissao: gv('sf_admissao'),
            diasContrato: 0,
            premios: gnum('sf_premios'),
            permissions: collectPermissionsFromContainer('staffFormColabModulesContainer')
        };

        if (!p.nome) return notify("Preencha o Nome do colaborador.", "warning");
        if (!p.email || !p.email.includes('@')) return notify("Preencha um Email válido.", "warning");
        if (!p.cargo) return notify("Selecione uma Função/Cargo da lista.", "warning");
        if (!p.vencimento || p.vencimento <= 0) return notify("Preencha o Vencimento Base.", "warning");
        if (!p.admissao) return notify("Preencha a Data de Admissão.", "warning");

        if (!p.provFerias && p.vencimento > 0) {
            p.provFerias = parseFloat((p.vencimento / 12).toFixed(2));
            p.provNatal = parseFloat((p.vencimento / 12).toFixed(2));
            p.provRescisao = parseFloat((p.vencimento * 0.055).toFixed(2));
            p.provFormacao = parseFloat((p.vencimento * 0.02).toFixed(2));
        }

        notify("A guardar...", "info");

        google.script.run
            .withFailureHandler(err => notify("Erro no Servidor: " + err, "error"))
            .withSuccessHandler(res => {
                if (res.success) {
                    if (res.isNew && res.emailSent) {
                        notify("Colaborador criado (Pendente). Email enviado — ficará Ativo após definir a palavra-passe.", "success");
                    } else if (res.isNew && !res.emailSent) {
                        notify("Colaborador criado (Pendente). Não foi possível enviar o email — use 'Reenviar Convite' nos Acessos.", "warning");
                    } else {
                        notify("Colaborador atualizado com sucesso!", "success");
                    }
                    document.getElementById('staffForm').style.display = 'none';
                    loadRH();
                    loadUsers();
                    if (typeof invalidateViewCache === 'function') invalidateViewCache('admin');
                } else {
                    notify("Atenção: " + res.error, "error");
                }
            }).saveStaffData(p, _ctxEmail());
    }

    function openStaffForm(id = null, isInvite = false) {
        const form = document.getElementById('staffForm');
        if (!form) return;
        const sv = (elId, val) => { const el = document.getElementById(elId); if (el) el.value = val || ''; };

        sv('sf_id', id);
        sv('sf_nome', ''); sv('sf_email', ''); sv('sf_nif', ''); sv('sf_cargo', '');
        sv('sf_venc', ''); sv('sf_alim', ''); sv('sf_seg', '');
        sv('sf_status', isInvite ? 'Pendente' : 'Ativo');
        sv('sf_prov_ferias', ''); sv('sf_prov_natal', ''); sv('sf_prov_rescisao', ''); sv('sf_prov_formacao', '');
        sv('sf_admissao', ''); sv('sf_premios', '');

        const defPerms = { dashboard: true, cc: false, logistica: true, ia: false, rh: false, admin: false };
        var currentColabPermissions = defPerms;

        if (id) {
            const s = staffCache.find(x => x.id == id);
            if (s) {
                sv('sf_nome', s.nome); sv('sf_email', s.email); sv('sf_nif', s.nif); sv('sf_cargo', s.cargo);
                sv('sf_venc', s.vencimento); sv('sf_alim', s.subAlim); sv('sf_seg', s.seguro);
                sv('sf_status', s.status);
                if (s.admissao) {
                    let formattedDate = "";
                    if (String(s.admissao).includes('/')) {
                        const parts = String(s.admissao).split('/');
                        if (parts.length === 3) formattedDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                    } else {
                        const d = new Date(s.admissao);
                        if (!isNaN(d)) formattedDate = d.toISOString().split('T')[0];
                    }
                    sv('sf_admissao', formattedDate);
                } else {
                    sv('sf_admissao', '');
                }
                sv('sf_premios', s.premios);

                const users = (window._usersCache || window.usersCache || []);
                const u = users.find(x => x.email === s.email);
                currentColabPermissions = (u && u.permissions) ? u.permissions : defPerms;
                calcStaffPreview();
            }
        }

        var clientCfg = window._clientConfig || {};
        renderColabFeatures('staffFormColabModulesContainer', currentColabPermissions, clientCfg);
        const preview = document.getElementById('sfCalcPreview');
        if (preview) preview.classList.add('hidden');
        form.style.display = 'flex';
        form.classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    }

    // ==========================================
    // 3. ACESSOS E PERMISSÕES (ADMIN/USERS)
    // ==========================================

    function loadUsers() {
        const list = document.getElementById('staffList');
        if (!list) return;
        list.innerHTML = '<div class="text-center py-24 opacity-30 text-[10px] font-black uppercase tracking-[0.3em] animate-pulse">A carregar colaboradores...</div>';

        google.script.run
            .withFailureHandler(err => {
                list.innerHTML = `<div class="p-10 bg-rose-50 dark:bg-rose-900/20 text-rose-500 rounded-[3rem] text-center border border-rose-100 dark:border-rose-800"><p class="font-black text-xs">${(err && err.toString) ? err.toString() : String(err)}</p></div>`;
                if (window.lucide) lucide.createIcons();
            })
            .withSuccessHandler(res => {
                if (!res) {
                    list.innerHTML = '<p class="text-center py-10 text-slate-500 dark:text-slate-400 font-bold">Sem dados disponíveis.</p>';
                    if (window.lucide) lucide.createIcons();
                    return;
                }
                list.innerHTML = '';
                const staff = (res.staff && Array.isArray(res.staff)) ? res.staff : [];
                const users = (res.users && Array.isArray(res.users)) ? res.users : [];
                window._staffCache = staff;
                window._usersCache = users;
                window._clientConfig = res.clientConfig || res.planConfig || {};
                staffCache = staff;

                const userByEmail = {};
                (users || []).forEach(u => { if (u && u.email) userByEmail[u.email] = u; });

                if (res.isMaster && staff.length === 0 && !res.sheetId) {
                    list.innerHTML = '<p class="text-center py-10 text-flowly-warning dark:text-amber-400 font-bold">Selecione um cliente no dropdown de visão para gerir os acessos.</p>';
                } else if (staff.length === 0) {
                    list.innerHTML = '<p class="text-center py-10 text-slate-500 dark:text-slate-400 font-bold">Nenhum colaborador. Use o módulo RH para adicionar ou convidar.</p>';
                } else {
                    staff.forEach(s => {
                        const u = userByEmail[s.email] || {};
                        const perms = u.permissions || { dashboard: false, cc: false, logistica: true, ia: false, rh: false };
                        const userStatus = u.status || 'Sem conta';

                        const statusColor = s.status === 'Ativo' ? 'bg-emerald-100 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 border-emerald-200 dark:border-emerald-800' : (s.status === 'Pendente' ? 'bg-amber-100 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 border-amber-200 dark:border-amber-800' : 'bg-rose-100 dark:bg-rose-900/30 text-rose-600 dark:text-rose-400 border-rose-200 dark:border-rose-800');
                        const regStatusColor = userStatus === 'Ativo' ? 'text-flowly-success' : (userStatus === 'Suspenso' ? 'text-rose-500' : 'text-amber-500');
                        const regStatusLabel = userStatus === 'Ativo' ? 'Conta Ativa' : (userStatus === 'Suspenso' ? 'Conta Suspensa' : 'Aguarda Ativação');

                        const permTags = [
                            perms.dashboard ? '<span class="px-2 py-0.5 bg-violet-100 dark:bg-violet-900/30 text-violet-600 dark:text-violet-400 text-[7px] font-black rounded-lg">DASHBOARD</span>' : '',
                            perms.cc ? '<span class="px-2 py-0.5 bg-indigo-100 dark:bg-indigo-900/30 text-indigo-600 dark:text-indigo-400 text-[7px] font-black rounded-lg">TESOURARIA</span>' : '',
                            perms.ia ? '<span class="px-2 py-0.5 bg-amber-100 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 text-[7px] font-black rounded-lg">IA</span>' : '',
                            perms.logistica ? '<span class="px-2 py-0.5 bg-cyan-100 dark:bg-cyan-900/30 text-cyan-600 dark:text-cyan-400 text-[7px] font-black rounded-lg">LOGÍSTICA</span>' : '',
                            perms.rh ? '<span class="px-2 py-0.5 bg-emerald-100 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 text-[7px] font-black rounded-lg">RH</span>' : '',
                            perms.admin ? '<span class="px-2 py-0.5 bg-rose-100 dark:bg-rose-900/30 text-rose-600 dark:text-rose-400 text-[7px] font-black rounded-lg">ACESSOS</span>' : ''
                        ].filter(Boolean).join(' ') || '<span class="text-[8px] text-slate-400 dark:text-slate-500">Sem permissões</span>';

                        const emailEsc = (s.email || '').replace(/'/g, "\\'");
                        const nomeEsc = (s.nome || 'N/A').replace(/'/g, "\\'");
                        const idEsc = (s.id || '').replace(/'/g, "\\'");

                        list.innerHTML += `
                    <div class="theme-card p-5 rounded-[2rem] border shadow-sm mb-3">
                        <div class="flex justify-between items-start">
                            <div class="flex items-center gap-4">
                                <div class="h-12 w-12 bg-slate-50 dark:bg-slate-700 rounded-2xl flex items-center justify-center font-black text-flowly-primary text-lg">${(s.nome || '?').charAt(0).toUpperCase()}</div>
                                <div>
                                    <span class="font-black text-sm text-slate-900 dark:text-slate-50">${s.nome || '—'}</span>
                                    <span class="block text-[10px] text-slate-400 dark:text-slate-500">${s.email || '—'}</span>
                                    <div class="flex items-center gap-2 mt-1">
                                        <span class="px-2 py-0.5 ${statusColor} text-[8px] font-black rounded-full border">${s.status || '—'}</span>
                                        <span class="text-[8px] font-bold ${regStatusColor}">● ${regStatusLabel}</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="flex flex-wrap gap-1 mt-3 mb-3">${permTags}</div>
                        <div class="flex gap-2 mt-2 pt-3 border-t border-slate-100 dark:border-slate-700">
                            ${s.email ? `<button onclick="openStaffPermissionsModal('${emailEsc}','${nomeEsc}')" class="flex-1 py-2 bg-cyan-50 dark:bg-cyan-900/30 text-flowly-primary rounded-xl text-[10px] font-bold hover:bg-cyan-100 dark:hover:bg-cyan-900/50 transition flex items-center justify-center gap-1"><i data-lucide="settings-2" class="w-3 h-3"></i> Modificar Acessos</button>` : ''}
                            ${s.email ? `<button onclick="resendStaffInviteClick('${emailEsc}')" class="py-2 px-3 bg-flowly-warning/10 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 rounded-xl text-[10px] font-bold hover:bg-amber-100 dark:hover:bg-amber-900/50 transition flex items-center gap-1" title="Reenviar link de ativação"><i data-lucide="mail" class="w-3 h-3"></i> Reenviar Link</button>` : ''}
                            ${s.status === 'Ativo' ? `<button onclick="deactivateStaffClick('${idEsc}','${nomeEsc}')" class="py-2 px-3 bg-flowly-warning/10 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 rounded-xl text-[10px] font-bold hover:bg-amber-100 dark:hover:bg-amber-900/50 transition" title="Desativar"><i data-lucide="user-x" class="w-3 h-3"></i></button>` : ''}
                            <button onclick="deleteStaffClick('${idEsc}','${nomeEsc}')" class="py-2 px-3 bg-rose-50 dark:bg-rose-900/30 text-rose-600 dark:text-rose-400 rounded-xl text-[10px] font-bold hover:bg-rose-100 dark:hover:bg-rose-900/50 transition" title="Eliminar"><i data-lucide="trash-2" class="w-3 h-3"></i></button>
                        </div>
                    </div>`;
                    });
                }
                if (window.lucide) lucide.createIcons();
            }).getMasterData(_ctxEmail());
    }

    function openStaffPermissionsModal(email, nome) {
        _editingStaffEmail = email;
        document.getElementById('modalStaffPermEmail').textContent = email;
        const users = (window._usersCache || []);
        const u = users.find(x => x.email === email);
        const colabData = { permissions: (u && u.permissions) || { dashboard: false, cc: false, logistica: true, ia: false, rh: false, admin: false, exportacao: false } };
        renderColabFeatures('colabModulesContainer', colabData.permissions);
        document.getElementById('modalStaffPermissions').classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    }

    function saveStaffPermissions() {
        if (!_editingStaffEmail) return;
        const container = document.getElementById('colabModulesContainer');
        const perms = { dashboard: false, cc: false, logistica: false, ia: false, rh: false, admin: false, exportacao: false };
        if (container) {
            container.querySelectorAll('input[type="checkbox"][data-feature-id]').forEach(function (cb) {
                perms[cb.dataset.featureId] = !!cb.checked;
            });
        }
        google.script.run.withSuccessHandler(() => {
            document.getElementById('modalStaffPermissions').classList.add('hidden');
            notify('Permissões guardadas.', 'success');
            loadUsers();
        }).withFailureHandler(err => notify(err, 'error')).saveUserPermissions(_editingStaffEmail, perms, _ctxEmail());
        _editingStaffEmail = null;
    }

    function resendStaffInviteClick(email) {
        notify('A reenviar convite...', 'info');
        google.script.run.withSuccessHandler(res => {
            if (res.success) notify('Convite reenviado para ' + email, 'success');
            else notify(res.error || 'Erro', 'error');
        }).withFailureHandler(err => notify(err, 'error')).resendStaffInvite(email, _ctxEmail());
    }

    function deactivateStaffClick(id, nome) {
        showConfirm('Desativar o acesso de ' + nome + '? O colaborador ficará no histórico mas não poderá fazer login.', () => {
            notify('A desativar...', 'info');
            google.script.run.withSuccessHandler(res => {
                if (res.success) { notify('Acesso desativado.', 'success'); loadUsers(); loadRH(); }
                else notify(res.error || 'Erro', 'error');
            }).withFailureHandler(err => notify(err, 'error')).deactivateStaffMember(id, _ctxEmail());
        });
    }

    function deleteStaffClick(id, nome) {
        showConfirm('Eliminar permanentemente ' + nome + '? Esta ação não pode ser revertida.', () => {
            notify('A eliminar...', 'info');
            google.script.run.withSuccessHandler(res => {
                if (res.success) { notify('Colaborador eliminado.', 'success'); loadUsers(); loadRH(); }
                else notify(res.error || 'Erro', 'error');
            }).withFailureHandler(err => notify(err, 'error')).deleteStaffMember(id, _ctxEmail());
        });
    }

    function openInviteStaffModal() {
        document.getElementById('inviteStaffNome').value = '';
        document.getElementById('inviteStaffEmail').value = '';
        document.getElementById('inviteStaffCargo').value = '';
        document.getElementById('modalInviteStaff').classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    }

    function submitInviteStaff() {
        const nome = document.getElementById('inviteStaffNome').value.trim();
        const email = document.getElementById('inviteStaffEmail').value.trim();
        const cargo = document.getElementById('inviteStaffCargo').value.trim();
        if (!email || !email.includes('@')) return notify('Email inválido.', 'error');
        document.getElementById('modalInviteStaff').classList.add('hidden');
        notify('A enviar convite...', 'info');

        google.script.run.withSuccessHandler(res => {
            if (res.success) { notify(res.message || 'Convite enviado!', 'success'); loadUsers(); }
            else notify(res.error || 'Erro', 'error');
        }).withFailureHandler(err => notify(err, 'error')).inviteStaff({ nome: nome || email, email: email, cargo: cargo }, _ctxEmail());
    }

    // ==========================================
    // 4. CRM CLIENTES & FORNECEDORES
    // ==========================================

    function maskNifInput(e) {
        var v = (e.target.value || '').replace(/\D/g, '').slice(0, 9);
        e.target.value = v;
    }

    function maskTelefoneInput(e) {
        var v = (e.target.value || '').replace(/\D/g, '').slice(0, 9);
        if (v.length <= 3) e.target.value = v ? '+351 ' + v : '';
        else if (v.length <= 6) e.target.value = '+351 ' + v.slice(0, 3) + ' ' + v.slice(3);
        else e.target.value = '+351 ' + v.slice(0, 3) + ' ' + v.slice(3, 6) + ' ' + v.slice(6);
    }

    function initCrmFormMasks() {
        if (window._crmFormMasksInit) return;
        window._crmFormMasksInit = true;
        var ids = ['clienteFormNif', 'clienteFormTelefone', 'fornecedorFormNif', 'fornecedorFormTelefone'];
        ids.forEach(function (id) {
            var el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('input', id.indexOf('Telefone') !== -1 ? maskTelefoneInput : maskNifInput);
        });
    }

    function loadClientesCRM() {
        var tbody = document.getElementById('crmClientesTableBody');
        if (!tbody) return;
        tbody.innerHTML = '<tr><td colspan="6" class="px-4 py-8 text-center text-slate-400 animate-pulse">A carregar...</td></tr>';

        google.script.run
            .withFailureHandler(function (err) {
                tbody.innerHTML = '<tr><td colspan="6" class="px-4 py-8 text-center text-red-500">Erro ao carregar. ' + (err || '') + '</td></tr>';
            })
            .withSuccessHandler(function (list) {
                if (!list || list.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="px-4 py-12 text-center text-slate-500">Nenhum cliente registado. Clique em "Novo Cliente" para adicionar.</td></tr>';
                    return;
                }
                var esc = function (s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
                tbody.innerHTML = list.map(function (c) {
                    return '<tr class="bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700/50 transition">' +
                        '<td class="px-4 py-3 font-bold text-flowly-midnight dark:text-slate-50">' + esc(c.nomeEmpresa) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400">' + esc(c.emailMasked) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400">' + esc(c.nifMasked) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400">' + esc(c.telefoneMasked) + '</td>' +
                        '<td class="px-4 py-3"><span class="px-2 py-1 rounded-full text-[9px] font-black ' + (c.status === 'Ativo' ? 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400' : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400') + '">' + esc(c.status) + '</span></td>' +
                        '<td class="px-4 py-3"><div class="flex gap-2">' +
                        '<button type="button" onclick="openAuditModal(\'' + esc(c.idCliente) + '\')" class="p-2 rounded-xl text-flowly-warning hover:bg-flowly-warning/10 dark:hover:bg-amber-900/20 transition" title="Desmascarar"><i data-lucide="eye" class="w-4 h-4"></i></button>' +
                        '<button type="button" onclick="openEditClienteModal(\'' + esc(c.idCliente) + '\')" class="p-2 rounded-xl text-flowly-primary hover:bg-cyan-50 dark:hover:bg-cyan-900/20 transition" title="Editar"><i data-lucide="pencil" class="w-4 h-4"></i></button>' +
                        '<button type="button" data-crm-id="' + esc(c.idCliente) + '" data-crm-name="' + esc(c.nomeEmpresa) + '" onclick="confirmApagarCliente(this.dataset.crmId, this.dataset.crmName)" class="p-2 rounded-xl text-flowly-danger hover:bg-red-50 dark:hover:bg-red-900/20 transition" title="Eliminar"><i data-lucide="trash-2" class="w-4 h-4"></i></button>' +
                        '</div></td></tr>';
                }).join('');
                if (window.lucide) lucide.createIcons();
            }).getClientesCRM(_ctxEmail());
    }

    function openModalClienteForm() {
        _auditClienteId = null;
        switchCrmTab('clientes');
        document.getElementById('clienteFormIdCliente').value = '';
        document.getElementById('modalClienteFormTitle').textContent = 'Novo Cliente';
        var sub = document.getElementById('modalClienteFormSubtitle');
        if (sub) sub.textContent = 'Preencha os dados e confirme o consentimento RGPD.';

        document.getElementById('clienteFormNome').value = '';
        document.getElementById('clienteFormNif').value = '';
        document.getElementById('clienteFormEmail').value = '';
        document.getElementById('clienteFormTelefone').value = '';
        document.getElementById('clienteFormMorada').value = '';
        document.getElementById('clienteFormStatus').value = 'Ativo';
        document.getElementById('rgpdConsent').checked = false;

        document.getElementById('modalClienteForm').classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    }

    function openEditClienteModal(idCliente) {
        _auditClienteId = null;
        switchCrmTab('clientes');
        document.getElementById('clienteFormIdCliente').value = idCliente;
        document.getElementById('modalClienteFormTitle').textContent = 'Editar Cliente';
        var sub = document.getElementById('modalClienteFormSubtitle');
        if (sub) sub.textContent = 'Preencha os dados e confirme o consentimento RGPD.';
        document.getElementById('rgpdConsent').checked = true;
        document.getElementById('rgpdConsent').required = false;

        google.script.run.withSuccessHandler(function (data) {
            if (data) {
                document.getElementById('clienteFormNome').value = data.nomeEmpresa || '';
                document.getElementById('clienteFormNif').value = data.nif || '';
                document.getElementById('clienteFormEmail').value = data.email || '';
                document.getElementById('clienteFormTelefone').value = data.telefone || '';
                document.getElementById('clienteFormMorada').value = data.morada || '';
                document.getElementById('clienteFormStatus').value = data.status || 'Ativo';
            }
            document.getElementById('modalClienteForm').classList.remove('hidden');
            if (window.lucide) lucide.createIcons();
        }).withFailureHandler(function () { notify('Erro ao carregar dados do cliente.', 'error'); }).getClienteById(idCliente, _ctxEmail());
    }

    function submitClienteForm() {
        var cb = document.getElementById('rgpdConsent');
        if (!cb || !cb.checked) {
            notify('Deve confirmar o consentimento RGPD para prosseguir.', 'error');
            return;
        }

        var idCliente = document.getElementById('clienteFormIdCliente').value.trim();
        var payload = {
            nomeEmpresa: document.getElementById('clienteFormNome').value.trim(),
            nif: document.getElementById('clienteFormNif').value.trim(),
            email: document.getElementById('clienteFormEmail').value.trim(),
            telefone: document.getElementById('clienteFormTelefone').value.trim(),
            morada: document.getElementById('clienteFormMorada').value.trim(),
            status: document.getElementById('clienteFormStatus').value || 'Ativo'
        };

        if (!payload.nomeEmpresa) { notify('Nome da empresa obrigatório.', 'error'); return; }
        document.getElementById('modalClienteForm').classList.add('hidden');

        if (idCliente) {
            google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao guardar.', 'error'); loadClientesCRM(); })
                .withSuccessHandler(function () { notify('Cliente atualizado.', 'success'); loadClientesCRM(); if (typeof loadClientesDatalist === 'function') loadClientesDatalist(); })
                .editCliente(idCliente, payload, _ctxEmail());
        } else {
            google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao criar.', 'error'); loadClientesCRM(); })
                .withSuccessHandler(function () { notify('Cliente criado.', 'success'); loadClientesCRM(); if (typeof loadClientesDatalist === 'function') loadClientesDatalist(); })
                .addCliente(payload, _ctxEmail());
        }
    }

    function confirmApagarCliente(idCliente, nomeEmpresa) {
        var label = (nomeEmpresa || idCliente || 'este cliente');
        if (!confirm('Tem a certeza que deseja eliminar o cliente "' + label + '"? Esta ação é irreversível (Direito ao Esquecimento - RGPD).')) return;

        google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao eliminar.', 'error'); })
            .withSuccessHandler(function () { notify('Cliente eliminado.', 'success'); loadClientesCRM(); })
            .apagarCliente(idCliente, _ctxEmail());
    }

    function switchCrmTab(tabName) {
        var tabClientes = document.getElementById('crmTabClientes');
        var tabFornecedores = document.getElementById('crmTabFornecedores');
        var btnClientes = document.getElementById('crmTabBtnClientes');
        var btnFornecedores = document.getElementById('crmTabBtnFornecedores');
        if (!tabClientes || !tabFornecedores) return;

        if (tabName === 'fornecedores') {
            tabClientes.classList.add('hidden');
            tabFornecedores.classList.remove('hidden');
            if (btnClientes) { btnClientes.classList.remove('bg-flowly-primary/20', 'text-flowly-primary', 'dark:text-cyan-300', 'border-flowly-primary'); btnClientes.classList.add('text-slate-500', 'dark:text-slate-400', 'border-transparent'); }
            if (btnFornecedores) { btnFornecedores.classList.add('bg-flowly-primary/20', 'text-flowly-primary', 'dark:text-cyan-300', 'border-flowly-primary'); btnFornecedores.classList.remove('text-slate-500', 'dark:text-slate-400', 'border-transparent'); }
        } else {
            tabClientes.classList.remove('hidden');
            tabFornecedores.classList.add('hidden');
            if (btnClientes) { btnClientes.classList.add('bg-flowly-primary/20', 'text-flowly-primary', 'dark:text-cyan-300', 'border-flowly-primary'); btnClientes.classList.remove('text-slate-500', 'dark:text-slate-400', 'border-transparent'); }
            if (btnFornecedores) { btnFornecedores.classList.remove('bg-flowly-primary/20', 'text-flowly-primary', 'dark:text-cyan-300', 'border-flowly-primary'); btnFornecedores.classList.add('text-slate-500', 'dark:text-slate-400', 'border-transparent'); }
        }
    }

    function switchCrmViewTab(tabName) {
        var contentClientes = document.getElementById('crmContentClientes');
        var contentFornecedores = document.getElementById('crmContentFornecedores');
        var btnClientes = document.getElementById('crmViewBtnClientes');
        var btnFornecedores = document.getElementById('crmViewBtnFornecedores');
        if (!contentClientes || !contentFornecedores) return;

        if (tabName === 'fornecedores') {
            contentClientes.classList.add('hidden');
            contentFornecedores.classList.remove('hidden');
            if (btnClientes) { btnClientes.classList.remove('bg-flowly-primary', 'text-white'); btnClientes.classList.add('bg-slate-200', 'dark:bg-slate-600', 'text-slate-600', 'dark:text-slate-300'); }
            if (btnFornecedores) { btnFornecedores.classList.add('bg-flowly-primary', 'text-white'); btnFornecedores.classList.remove('bg-slate-200', 'dark:bg-slate-600', 'text-slate-600', 'dark:text-slate-300'); }
        } else {
            contentClientes.classList.remove('hidden');
            contentFornecedores.classList.add('hidden');
            if (btnClientes) { btnClientes.classList.add('bg-flowly-primary', 'text-white'); btnClientes.classList.remove('bg-slate-200', 'dark:bg-slate-600', 'text-slate-600', 'dark:text-slate-300'); }
            if (btnFornecedores) { btnFornecedores.classList.remove('bg-flowly-primary', 'text-white'); btnFornecedores.classList.add('bg-slate-200', 'dark:bg-slate-600', 'text-slate-600', 'dark:text-slate-300'); }
        }
        if (window.lucide) lucide.createIcons();
    }

    function loadFornecedoresCRM() {
        var tbody = document.getElementById('crmFornecedoresTableBody');
        if (!tbody) return;
        tbody.innerHTML = '<tr><td colspan="7" class="px-4 py-8 text-center text-slate-400 animate-pulse">A carregar...</td></tr>';

        google.script.run
            .withFailureHandler(function (err) {
                tbody.innerHTML = '<tr><td colspan="7" class="px-4 py-12 text-center text-red-500">Erro ao carregar. ' + (err || '') + '</td></tr>';
            })
            .withSuccessHandler(function (list) {
                if (!list || list.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="7" class="px-4 py-12 text-center text-slate-500">Nenhum fornecedor registado. Clique em "Novo Fornecedor" para adicionar.</td></tr>';
                    return;
                }
                var esc = function (s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
                tbody.innerHTML = list.map(function (f) {
                    return '<tr class="bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700/50 transition" data-fornecedor-id="' + esc(f.idFornecedor) + '">' +
                        '<td class="px-4 py-3 font-bold text-flowly-midnight dark:text-slate-50">' + esc(f.nomeEmpresa || f.nome) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400 forn-email">' + esc(f.emailMasked) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400 forn-nif">' + esc(f.nifMasked) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400 forn-telefone">' + esc(f.telefoneMasked) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400">' + esc(f.categoria) + '</td>' +
                        '<td class="px-4 py-3 text-slate-600 dark:text-slate-400">' + esc(f.condicoesPagamento) + '</td>' +
                        '<td class="px-4 py-3"><div class="flex gap-2">' +
                        '<button type="button" onclick="openAuditModalFornecedor(\'' + esc(f.idFornecedor) + '\')" class="p-2 rounded-xl text-flowly-warning hover:bg-flowly-warning/10 dark:hover:bg-amber-900/20 transition" title="Desmascarar"><i data-lucide="eye" class="w-4 h-4"></i></button>' +
                        '<button type="button" onclick="openEditFornecedorModal(\'' + esc(f.idFornecedor) + '\')" class="p-2 rounded-xl text-flowly-primary hover:bg-cyan-50 dark:hover:bg-cyan-900/20 transition" title="Editar"><i data-lucide="pencil" class="w-4 h-4"></i></button>' +
                        '<button type="button" data-crm-id="' + esc(f.idFornecedor) + '" data-crm-name="' + esc(f.nomeEmpresa || f.nome) + '" onclick="confirmApagarFornecedor(this.dataset.crmId, this.dataset.crmName)" class="p-2 rounded-xl text-flowly-danger hover:bg-red-50 dark:hover:bg-red-900/20 transition" title="Eliminar"><i data-lucide="trash-2" class="w-4 h-4"></i></button>' +
                        '</div></td></tr>';
                }).join('');
                if (window.lucide) lucide.createIcons();
            }).getFornecedoresCRM(_ctxEmail());
    }

    function openModalFornecedorForm(idFornecedor) {
        switchCrmTab('fornecedores');
        document.getElementById('fornecedorFormIdFornecedor').value = idFornecedor || '';
        document.getElementById('modalClienteFormTitle').textContent = idFornecedor ? 'Editar Fornecedor' : 'Novo Fornecedor';
        var sub = document.getElementById('modalClienteFormSubtitle');
        if (sub) sub.textContent = 'Preencha os dados do fornecedor.';

        document.getElementById('fornecedorFormNome').value = '';
        document.getElementById('fornecedorFormNif').value = '';
        document.getElementById('fornecedorFormEmail').value = '';
        document.getElementById('fornecedorFormTelefone').value = '';
        document.getElementById('fornecedorFormCategoria').value = '';
        document.getElementById('fornecedorFormCondicoesPagamento').value = '';
        document.getElementById('modalClienteForm').classList.remove('hidden');

        if (idFornecedor) {
            google.script.run.withSuccessHandler(function (data) {
                if (data) {
                    document.getElementById('fornecedorFormNome').value = data.nomeEmpresa || data.nome || '';
                    document.getElementById('fornecedorFormNif').value = data.nif || '';
                    document.getElementById('fornecedorFormEmail').value = data.email || '';
                    document.getElementById('fornecedorFormTelefone').value = data.telefone || '';
                    document.getElementById('fornecedorFormCategoria').value = data.categoria || '';
                    document.getElementById('fornecedorFormCondicoesPagamento').value = data.condicoesPagamento || '';
                }
                if (window.lucide) lucide.createIcons();
            }).withFailureHandler(function () { notify('Erro ao carregar dados do fornecedor.', 'error'); }).getFornecedorById(idFornecedor, _ctxEmail());
        } else {
            if (window.lucide) lucide.createIcons();
        }
    }

    function openEditFornecedorModal(idFornecedor) {
        openModalFornecedorForm(idFornecedor);
    }

    function submitFornecedorForm() {
        var idFornecedor = document.getElementById('fornecedorFormIdFornecedor').value.trim();
        var payload = {
            nomeEmpresa: document.getElementById('fornecedorFormNome').value.trim(),
            nif: document.getElementById('fornecedorFormNif').value.trim(),
            email: document.getElementById('fornecedorFormEmail').value.trim(),
            telefone: document.getElementById('fornecedorFormTelefone').value.trim(),
            categoria: document.getElementById('fornecedorFormCategoria').value.trim(),
            condicoesPagamento: document.getElementById('fornecedorFormCondicoesPagamento').value.trim()
        };

        if (!payload.nomeEmpresa) { notify('Nome/Empresa obrigatório.', 'error'); return; }
        document.getElementById('modalClienteForm').classList.add('hidden');

        if (idFornecedor) {
            google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao guardar.', 'error'); loadFornecedoresCRM(); })
                .withSuccessHandler(function () { notify('Fornecedor atualizado.', 'success'); loadFornecedoresCRM(); if (typeof loadFornecedoresDatalist === 'function') loadFornecedoresDatalist(); })
                .editFornecedor(idFornecedor, payload, _ctxEmail());
        } else {
            google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao criar.', 'error'); loadFornecedoresCRM(); })
                .withSuccessHandler(function () { notify('Fornecedor criado.', 'success'); loadFornecedoresCRM(); if (typeof loadFornecedoresDatalist === 'function') loadFornecedoresDatalist(); })
                .addFornecedor(payload, _ctxEmail());
        }
    }

    function confirmApagarFornecedor(idFornecedor, nomeEmpresa) {
        var label = (nomeEmpresa || idFornecedor || 'este fornecedor');
        if (!confirm('Tem a certeza que deseja eliminar o fornecedor "' + label + '"? Esta ação é irreversível.')) return;
        google.script.run.withFailureHandler(function (err) { notify(err || 'Erro ao eliminar.', 'error'); })
            .withSuccessHandler(function () { notify('Fornecedor eliminado.', 'success'); loadFornecedoresCRM(); if (typeof loadFornecedoresDatalist === 'function') loadFornecedoresDatalist(); })
            .apagarFornecedor(idFornecedor, _ctxEmail());
    }

    // ==========================================
    // 5. AUDITORIA RGPD (DESMASCARAR DADOS)
    // ==========================================

    function openAuditModal(idCliente) {
        _auditClienteId = idCliente;
        _auditFornecedorId = null;
        document.getElementById('auditMotivoSelect').value = 'Emissão de Faturação';
        document.getElementById('auditDetalhesInput').value = '';
        document.getElementById('auditDetalhesWrap').classList.add('hidden');
        document.getElementById('modalAuditFormStep').classList.remove('hidden');
        document.getElementById('modalAuditResultStep').classList.add('hidden');
        document.getElementById('modalAuditDesmascarar').classList.remove('hidden');

        document.getElementById('auditMotivoSelect').onchange = function () {
            document.getElementById('auditDetalhesWrap').classList.toggle('hidden', this.value !== 'Outro');
        };
        if (window.lucide) lucide.createIcons();
    }

    function openAuditModalFornecedor(idFornecedor) {
        _auditClienteId = null;
        _auditFornecedorId = idFornecedor;
        document.getElementById('auditMotivoSelect').value = 'Emissão de Faturação';
        document.getElementById('auditDetalhesInput').value = '';
        document.getElementById('auditDetalhesWrap').classList.add('hidden');
        document.getElementById('modalAuditFormStep').classList.remove('hidden');
        document.getElementById('modalAuditResultStep').classList.add('hidden');
        document.getElementById('modalAuditDesmascarar').classList.remove('hidden');

        document.getElementById('auditMotivoSelect').onchange = function () {
            document.getElementById('auditDetalhesWrap').classList.toggle('hidden', this.value !== 'Outro');
        };
        if (window.lucide) lucide.createIcons();
    }

    function submitAuditDesmascarar() {
        var motivo = document.getElementById('auditMotivoSelect').value;
        var detalhes = document.getElementById('auditDetalhesInput').value.trim();
        var motivoCompleto = motivo;
        if (motivo === 'Outro') {
            if (!detalhes) { notify('Detalhes do motivo obrigatórios quando seleciona "Outro".', 'error'); return; }
            motivoCompleto = detalhes;
        }

        if (_auditClienteId) {
            google.script.run.withFailureHandler(function (err) {
                notify(err || 'Erro ao desbloquear.', 'error');
            }).withSuccessHandler(function (res) {
                if (!res || !res.success) { notify(res ? res.error : 'Erro.', 'error'); return; }
                document.getElementById('modalAuditFormStep').classList.add('hidden');
                document.getElementById('modalAuditResultStep').classList.remove('hidden');
                var cancelWrap = document.getElementById('modalAuditCancelWrap');
                if (cancelWrap) cancelWrap.classList.add('hidden');

                document.getElementById('auditResultEmail').textContent = res.email || '-';
                document.getElementById('auditResultNif').textContent = res.nif || '-';
                document.getElementById('auditResultTelefone').textContent = res.telefone || '-';
                if (window.lucide) lucide.createIcons();
            }).desbloquearClienteComAuditoria(_auditClienteId, motivoCompleto, _ctxEmail());
            return;
        }

        if (_auditFornecedorId) {
            google.script.run.withFailureHandler(function (err) {
                notify(err || 'Erro ao desbloquear.', 'error');
            }).withSuccessHandler(function (res) {
                if (!res || !res.success) { notify(res ? res.error : 'Erro.', 'error'); return; }
                document.getElementById('modalAuditFormStep').classList.add('hidden');
                document.getElementById('modalAuditResultStep').classList.remove('hidden');
                var cancelWrap = document.getElementById('modalAuditCancelWrap');
                if (cancelWrap) cancelWrap.classList.add('hidden');

                document.getElementById('auditResultEmail').textContent = res.email || '-';
                document.getElementById('auditResultNif').textContent = res.nif || '-';
                document.getElementById('auditResultTelefone').textContent = res.telefone || '-';
                if (window.lucide) lucide.createIcons();
            }).desbloquearFornecedorComAuditoria(_auditFornecedorId, motivoCompleto, _ctxEmail());
            return;
        }
    }

    function closeAuditModalAndClear() {
        _auditClienteId = null;
        _auditFornecedorId = null;
        document.getElementById('modalAuditDesmascarar').classList.add('hidden');
        document.getElementById('modalAuditFormStep').classList.remove('hidden');
        document.getElementById('modalAuditResultStep').classList.add('hidden');
        var cancelWrap = document.getElementById('modalAuditCancelWrap');
        if (cancelWrap) cancelWrap.classList.remove('hidden');
    }

    // ==========================================
    // EXPORTAÇÃO GLOBAL PARA WINDOW
    // ==========================================
    
    // Funções chamadas pela UI
    window.loadRH = loadRH;
    window.saveStaff = saveStaff;
    window.openStaffForm = openStaffForm;
    window.loadUsers = loadUsers;
    window.submitClienteForm = submitClienteForm;
    window.openAuditModal = openAuditModal;
    window.calcStaffPreview = calcStaffPreview;
    window.openStaffPermissionsModal = openStaffPermissionsModal;
    window.saveStaffPermissions = saveStaffPermissions;
    window.resendStaffInviteClick = resendStaffInviteClick;
    window.deactivateStaffClick = deactivateStaffClick;
    window.deleteStaffClick = deleteStaffClick;
    window.openInviteStaffModal = openInviteStaffModal;
    window.submitInviteStaff = submitInviteStaff;
    window.maskNifInput = maskNifInput;
    window.maskTelefoneInput = maskTelefoneInput;
    window.initCrmFormMasks = initCrmFormMasks;
    window.loadClientesCRM = loadClientesCRM;
    window.openModalClienteForm = openModalClienteForm;
    window.openEditClienteModal = openEditClienteModal;
    window.confirmApagarCliente = confirmApagarCliente;
    window.switchCrmTab = switchCrmTab;
    window.switchCrmViewTab = switchCrmViewTab;
    window.loadFornecedoresCRM = loadFornecedoresCRM;
    window.openModalFornecedorForm = openModalFornecedorForm;
    window.openEditFornecedorModal = openEditFornecedorModal;
    window.submitFornecedorForm = submitFornecedorForm;
    window.confirmApagarFornecedor = confirmApagarFornecedor;
    window.openAuditModalFornecedor = openAuditModalFornecedor;
    window.submitAuditDesmascarar = submitAuditDesmascarar;
    window.closeAuditModalAndClear = closeAuditModalAndClear;

})();
