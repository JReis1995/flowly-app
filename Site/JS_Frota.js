(function() {
    // ==========================================
    // 1. VARIÁVEIS GLOBAIS (FROTA E CUSTOS)
    // ==========================================
    let _pendingOcrCustoB64 = null;
    let _pendingOcrCustoData = null;

    /** Margem em dias para alertas de IPO e Seguro (alinhado com backend). */
    var MARGEM_ALERTA_IPO_SEGURO = 30;

    // ==========================================
    // 2. NAVEGAÇÃO DE TABS DA FROTA
    // ==========================================

    // Alterna entre as tabs do módulo Frota (Veículos, Custos, Vasilhame)
    function switchFrotaTab(contentId) {
        const contents = ['contentVeiculos', 'contentCustos', 'contentVasilhame'];
        const tabs = ['tabVeiculos', 'tabCustos', 'tabVasilhame'];

        contents.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.toggle('hidden', id !== contentId);
        });

        tabs.forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                const isActive = (id === 'tabVeiculos' && contentId === 'contentVeiculos') ||
                    (id === 'tabCustos' && contentId === 'contentCustos') ||
                    (id === 'tabVasilhame' && contentId === 'contentVasilhame');
                if (isActive) {
                    btn.classList.add('bg-flowly-primary', 'text-white');
                    btn.classList.remove('bg-slate-100', 'text-slate-500');
                } else {
                    btn.classList.remove('bg-flowly-primary', 'text-white');
                    btn.classList.add('bg-slate-100', 'text-slate-500');
                }
            }
        });

        if (contentId === 'contentVeiculos') {
            var tbody = document.getElementById('tbodyVeiculos');
            var msgVazio = document.getElementById('msgVeiculosVazio');
            if (frotaVeiculosCache && tbody) {
                _renderVeiculosFromCache(frotaVeiculosCache, tbody, msgVazio);
            } else {
                loadVeiculos();
            }
        }

        if (contentId === 'contentCustos') {
            if (typeof loadFornecedoresDatalist === 'function') loadFornecedoresDatalist();
            var veiculosDisponiveis = frotaVeiculosCache || cachedVeiculos;
            if (!veiculosDisponiveis || veiculosDisponiveis.length === 0) {
                google.script.run.withSuccessHandler(function (veiculos) {
                    cachedVeiculos = (veiculos || []).filter(function (v) { return (v.status || 'Ativo') === 'Ativo'; });
                    frotaVeiculosCache = cachedVeiculos;
                    loadCustosVeiculos();
                    loadCustos();
                    initFormCustos();
                }).withFailureHandler(function () {
                    loadCustosVeiculos();
                    loadCustos();
                    initFormCustos();
                }).getVeiculos(_ctxEmail());
            } else {
                loadCustosVeiculos();
                loadCustos();
                initFormCustos();
            }
        }

        if (contentId === 'contentVasilhame') {
            if (typeof loadVasilhameData === 'function') loadVasilhameData();
        }

        if (window.lucide) lucide.createIcons();
    }

    // ==========================================
    // 3. VEÍCULOS E ALERTAS
    // ==========================================

    function diasAteData(dataStr) {
        if (!dataStr || typeof dataStr !== 'string') return null;
        var m = dataStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!m) return null;
        var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
        if (isNaN(d.getTime())) return null;
        var hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        d.setHours(0, 0, 0, 0);
        return Math.ceil((d - hoje) / 86400000);
    }

    function getAlertaBadgeClasses(dias) {
        if (dias === null) return 'text-slate-500 dark:text-slate-400';
        if (dias < 0) return 'text-flowly-danger dark:text-red-400 bg-red-100 dark:bg-red-900/30 px-2 py-1 rounded-full text-xs font-bold';
        if (dias <= MARGEM_ALERTA_IPO_SEGURO) return 'text-flowly-danger dark:text-red-400 bg-red-100 dark:bg-red-900/30 px-2 py-1 rounded-full text-xs font-bold';
        if (dias <= 60) return 'text-yellow-600 dark:text-yellow-400 bg-yellow-100 dark:bg-yellow-900/30 px-2 py-1 rounded-full text-xs font-bold';
        return 'text-green-600 dark:text-green-400 bg-green-100 dark:bg-green-900/30 px-2 py-1 rounded-full text-xs';
    }

    function getLavagemBadgeClasses(diasDesde) {
        if (diasDesde === null || diasDesde > 7) return 'text-orange-600 dark:text-orange-400 bg-orange-100 dark:bg-orange-900/30 px-2 py-1 rounded-full text-xs font-bold';
        return 'text-green-600 dark:text-green-400 bg-green-100 dark:bg-green-900/30 px-2 py-1 rounded-full text-xs';
    }

    function diasDesdeData(dataStr) {
        var dias = diasAteData(dataStr);
        if (dias === null) return null;
        return Math.max(0, -dias);
    }

    function formatarDataExibicao(str) {
        if (!str) return '—';
        var m = String(str).match(/^(\d{4})-(\d{2})-(\d{2})/);
        return m ? m[3] + '/' + m[2] + '/' + m[1] : str;
    }

    function loadVeiculos(force) {
        var tbody = document.getElementById('tbodyVeiculos');
        var msgVazio = document.getElementById('msgVeiculosVazio');
        if (!tbody) return;

        if (!force && frotaVeiculosCache) {
            _renderVeiculosFromCache(frotaVeiculosCache, tbody, msgVazio);
            return;
        }

        tbody.innerHTML = '';
        if (msgVazio) msgVazio.classList.add('hidden');

        google.script.run.withSuccessHandler(function (veiculos) {
            cachedVeiculos = (veiculos || []).filter(function (v) { return (v.status || 'Ativo') === 'Ativo'; });
            frotaVeiculosCache = cachedVeiculos;
            if (!cachedVeiculos.length) {
                if (msgVazio) msgVazio.classList.remove('hidden');
                return;
            }
            _renderVeiculosFromCache(cachedVeiculos, tbody, msgVazio);
        }).withFailureHandler(function (err) {
            if (msgVazio) {
                msgVazio.textContent = 'Erro ao carregar veículos: ' + (err || 'Desconhecido');
                msgVazio.classList.remove('hidden');
            }
        }).getVeiculos(_ctxEmail());
    }

    function _renderVeiculosFromCache(veiculos, tbody, msgVazio) {
        if (!tbody) return;
        tbody.innerHTML = '';
        if (msgVazio) msgVazio.classList.add('hidden');

        if (!veiculos || !veiculos.length) {
            if (msgVazio) msgVazio.classList.remove('hidden');
            var banner = document.getElementById('frotaAlertasBanner');
            if (banner) banner.classList.add('hidden');
            return;
        }

        var totalAlertas = 0;
        veiculos.forEach(function (v) {
            var diasInspecao = diasAteData(v.dataInspecao);
            var diasSeguro = diasAteData(v.validadeSeguro);
            var diasRevisao = diasAteData(v.proximaRevisao);
            var diasDesdeLavagem = v.ultimaLavagem ? diasDesdeData(v.ultimaLavagem) : 999;

            var clsInspecao = getAlertaBadgeClasses(diasInspecao);
            var clsSeguro = getAlertaBadgeClasses(diasSeguro);
            var clsRevisao = getAlertaBadgeClasses(diasRevisao);
            var clsLavagem = getLavagemBadgeClasses(diasDesdeLavagem);

            var iconIPO = (diasInspecao !== null && diasInspecao <= MARGEM_ALERTA_IPO_SEGURO) ? ' <span class="inline-block" title="Inspeção em breve!">⚠️</span> ' : '';
            var iconSeguro = (diasSeguro !== null && diasSeguro <= MARGEM_ALERTA_IPO_SEGURO) ? ' <span class="inline-block" title="Seguro a vencer!">⚠️</span> ' : '';
            var iconLavagem = (diasDesdeLavagem === null || diasDesdeLavagem > 7) ? ' <span class="inline-block" title="Viatura por lavar há mais de 7 dias!">💧</span> ' : '';

            var temAlertaCritico = (diasInspecao !== null && diasInspecao <= MARGEM_ALERTA_IPO_SEGURO) ||
                (diasSeguro !== null && diasSeguro <= MARGEM_ALERTA_IPO_SEGURO) ||
                (diasDesdeLavagem === null || diasDesdeLavagem > 7);

            if (temAlertaCritico) totalAlertas++;

            var tr = document.createElement('tr');
            tr.className = 'hover:bg-slate-50 dark:hover:bg-slate-800/50 transition' + (temAlertaCritico ? ' border-l-4 border-l-red-500 bg-red-50/30 dark:bg-red-900/10' : '');

            var combCat = (v.combustivel || '—') + (v.combustivel && v.categoria ? ' - ' + v.categoria : '');
            var matriculaEsc = (v.matricula || '').replace(/"/g, '&quot;');

            tr.innerHTML = '<td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300 font-medium">' + (v.matricula || '—') + '</td>' +
                '<td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300">' + (v.marcaModelo || '—') + (v.ano ? ' (' + v.ano + ')' : '') + '</td>' +
                '<td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300">' + combCat + '</td>' +
                '<td class="px-4 py-3 text-sm"><span class="' + clsInspecao + '">' + iconIPO + formatarDataExibicao(v.dataInspecao) + '</span></td>' +
                '<td class="px-4 py-3 text-sm"><span class="' + clsSeguro + '">' + iconSeguro + formatarDataExibicao(v.validadeSeguro) + '</span></td>' +
                '<td class="px-4 py-3 text-sm"><span class="' + clsLavagem + '">' + iconLavagem + (v.ultimaLavagem ? formatarDataExibicao(v.ultimaLavagem) : 'Nunca') + '</span></td>' +
                '<td class="px-4 py-3 text-sm text-center"><button type="button" data-matricula="' + matriculaEsc + '" onclick="registarLavagemClick(this.getAttribute(\'data-matricula\'))" class="inline-flex items-center justify-center gap-1 px-2.5 py-2 rounded-lg text-sm font-bold bg-amber-100 dark:bg-amber-900/30 text-flowly-warning dark:text-amber-400 hover:bg-amber-200 dark:hover:bg-amber-800/50 transition shadow-sm" title="Registar lavagem"><span>✨</span></button></td>' +
                '<td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300"><span class="' + clsRevisao + '">' + formatarDataExibicao(v.proximaRevisao) + '</span></td>' +
                '<td class="px-4 py-3 text-sm text-right"><button type="button" data-matricula="' + matriculaEsc + '" onclick="registarLavagemClick(this.getAttribute(\'data-matricula\'))" class="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold bg-flowly-primary/10 text-flowly-primary hover:bg-flowly-primary hover:text-white transition shadow-sm" title="Registar que lavou o carro"><span>✨</span> Lavei o Carro</button></td>' +
                '<td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300 text-right"><button type="button" data-id="' + (v.id || '').replace(/"/g, '&quot;') + '" onclick="editVeiculo(this.getAttribute(\'data-id\'))" class="p-2 text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition" title="Editar"><i data-lucide="pencil" class="w-4 h-4"></i></button>' +
                '<button type="button" data-id="' + (v.id || '').replace(/"/g, '&quot;') + '" onclick="confirmDeleteVeiculo(this.getAttribute(\'data-id\'))" class="p-2 text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition ml-1" title="Eliminar"><i data-lucide="trash-2" class="w-4 h-4"></i></button></td>';
            tbody.appendChild(tr);
        });

        if (totalAlertas > 0) {
            var banner = document.getElementById('frotaAlertasBanner');
            if (banner) {
                banner.innerHTML = '<span class="font-bold">⚠️ ' + totalAlertas + ' viatura(s) com alertas críticos</span> — IPO, Seguro ou Lavagem requerem atenção.';
                banner.classList.remove('hidden');
            }
        } else {
            var banner = document.getElementById('frotaAlertasBanner');
            if (banner) banner.classList.add('hidden');
        }

        if (window.lucide) lucide.createIcons();
    }

    function openModalVeiculo() {
        document.getElementById('v_id').value = '';
        document.getElementById('v_matricula').value = '';
        document.getElementById('v_marcaModelo').value = '';
        document.getElementById('v_ano').value = '';
        document.getElementById('v_combustivel').value = '';
        document.getElementById('v_categoria').value = '';
        document.getElementById('v_lotacao').value = '';
        document.getElementById('v_dataInspecao').value = '';
        document.getElementById('v_validadeSeguro').value = '';
        document.getElementById('v_proximaRevisao').value = '';
        document.getElementById('modalVeiculoTitle').textContent = 'Adicionar Veículo';

        document.getElementById('modalVeiculo').classList.remove('hidden');
        document.getElementById('modalVeiculo').classList.add('flex');
        if (window.lucide) lucide.createIcons();
    }

    function closeModalVeiculo() {
        document.getElementById('modalVeiculo').classList.add('hidden');
        document.getElementById('modalVeiculo').classList.remove('flex');
    }

    function saveVeiculoFromModal() {
        var data = {
            id: document.getElementById('v_id').value.trim() || undefined,
            matricula: document.getElementById('v_matricula').value.trim(),
            marcaModelo: document.getElementById('v_marcaModelo').value.trim(),
            ano: document.getElementById('v_ano').value || '',
            combustivel: document.getElementById('v_combustivel').value || '',
            categoria: document.getElementById('v_categoria').value || '',
            lotacao: document.getElementById('v_lotacao').value.trim(),
            dataInspecao: document.getElementById('v_dataInspecao').value || '',
            validadeSeguro: document.getElementById('v_validadeSeguro').value || '',
            proximaRevisao: document.getElementById('v_proximaRevisao').value || ''
        };

        if (!data.matricula) { notify('Matrícula é obrigatória.', 'error'); return; }

        google.script.run.withSuccessHandler(function (res) {
            if (res && res.success) {
                notify('Veículo guardado com sucesso.', 'success');
                closeModalVeiculo();
                loadVeiculos(true);
            } else {
                notify(res && res.error ? res.error : 'Erro ao guardar.', 'error');
            }
        }).withFailureHandler(function (err) {
            notify('Erro: ' + (err || 'Desconhecido'), 'error');
        }).saveVeiculo(data, _ctxEmail());
    }

    function editVeiculo(id) {
        google.script.run.withSuccessHandler(function (veiculos) {
            var v = (veiculos || []).find(function (x) { return x.id === id; });
            if (!v) { notify('Veículo não encontrado.', 'error'); return; }

            document.getElementById('v_id').value = v.id || '';
            document.getElementById('v_matricula').value = v.matricula || '';
            document.getElementById('v_marcaModelo').value = v.marcaModelo || '';
            document.getElementById('v_ano').value = v.ano || '';
            document.getElementById('v_combustivel').value = v.combustivel || '';
            document.getElementById('v_categoria').value = v.categoria || '';
            document.getElementById('v_lotacao').value = v.lotacao || '';
            document.getElementById('v_dataInspecao').value = v.dataInspecao || '';
            document.getElementById('v_validadeSeguro').value = v.validadeSeguro || '';
            document.getElementById('v_proximaRevisao').value = v.proximaRevisao || '';

            document.getElementById('modalVeiculoTitle').textContent = 'Editar Veículo';
            document.getElementById('modalVeiculo').classList.remove('hidden');
            document.getElementById('modalVeiculo').classList.add('flex');

            if (window.lucide) lucide.createIcons();
        }).withFailureHandler(function () {
            notify('Erro ao carregar dados.', 'error');
        }).getVeiculos(_ctxEmail());
    }

    function registarLavagemClick(matricula) {
        if (!matricula) { notify('Matrícula inválida.', 'error'); return; }
        google.script.run.withSuccessHandler(function (res) {
            if (res && res.success) {
                notify('Lavagem registada com sucesso!', 'success');
                loadVeiculos(true);
            } else {
                notify(res && res.error ? res.error : 'Erro ao registar lavagem.', 'error');
            }
        }).withFailureHandler(function (err) {
            notify('Erro: ' + (err || 'Desconhecido'), 'error');
        }).registarLavagem(matricula, _ctxEmail());
    }

    function confirmDeleteVeiculo(id) {
        if (!confirm('Tem a certeza que deseja eliminar este veículo?')) return;
        google.script.run.withSuccessHandler(function (res) {
            if (res && res.success) {
                notify('Veículo eliminado.', 'success');
                loadVeiculos(true);
            } else {
                notify(res && res.error ? res.error : 'Erro ao eliminar.', 'error');
            }
        }).withFailureHandler(function (err) {
            notify('Erro: ' + (err || 'Desconhecido'), 'error');
        }).deleteVeiculo(id, _ctxEmail());
    }

    // ==========================================
    // 4. REGISTO DE CUSTOS (APP DO MOTORISTA E OCR)
    // ==========================================

    function toggleCustoLitros() {
        var tipo = document.getElementById('custo_tipo');
        var wrap = document.getElementById('wrapCustoLitros');
        if (!tipo || !wrap) return;
        var isAbastecimento = tipo.value === 'Abastecimento';
        wrap.classList.toggle('hidden', !isAbastecimento);
        var litrosEl = document.getElementById('custo_litros');
        if (litrosEl) litrosEl.required = isAbastecimento;
        if (!isAbastecimento && litrosEl) litrosEl.value = '';
    }

    function toggleCustoKm() {
        var tipo = document.getElementById('custo_tipo');
        var wrap = document.getElementById('wrapCustoKm');
        if (!tipo || !wrap) return;
        var isPortagens = tipo.value === 'Portagens';
        wrap.classList.toggle('hidden', isPortagens);
    }

    function handleTipoDespesaChange(event) {
        var tipo = document.getElementById('custo_tipo');
        if (!tipo) return;
        var valor = tipo.value || '';
        var isPortagens = valor === 'Portagens';

        var step1Buttons = document.getElementById('step1-buttons');
        var wrapFile = document.getElementById('wrapCustoFile');
        var step1 = document.getElementById('step1Custos');
        var step2 = document.getElementById('step2Custos');

        var wrapFornecedor = document.getElementById('wrapCustoFornecedor');
        var wrapData = document.getElementById('wrapCustoData');
        var wrapKm = document.getElementById('wrapCustoKm');
        var wrapLitros = document.getElementById('wrapCustoLitros');
        var kmEl = document.getElementById('custo_km');
        var litrosEl = document.getElementById('custo_litros');
        var wrapValorIva = document.getElementById('wrapCustoValorIva');

        if (isPortagens) {
            if (step1Buttons) step1Buttons.classList.add('hidden');
            if (wrapFile) wrapFile.classList.add('hidden');
            if (step1) step1.classList.add('hidden');
            if (step2) step2.classList.remove('hidden');
            if (wrapFornecedor) wrapFornecedor.classList.add('hidden');
            if (wrapData) wrapData.classList.add('hidden');
            if (wrapKm) wrapKm.classList.add('hidden');
            if (wrapLitros) wrapLitros.classList.add('hidden');
            if (wrapValorIva) wrapValorIva.classList.add('hidden');

            if (kmEl) kmEl.removeAttribute('required');
            if (litrosEl) litrosEl.removeAttribute('required');

            var dataEl = document.getElementById('custo_data');
            if (dataEl) {
                var hoje = new Date();
                dataEl.value = hoje.getFullYear() + '-' + String(hoje.getMonth() + 1).padStart(2, '0') + '-' + String(hoje.getDate()).padStart(2, '0');
            }
        } else {
            if (step1Buttons) step1Buttons.classList.remove('hidden');
            if (wrapFile) wrapFile.classList.add('hidden');
            if (step1) step1.classList.remove('hidden');
            if (step2) step2.classList.add('hidden');
            if (wrapFornecedor) wrapFornecedor.classList.remove('hidden');
            if (wrapData) wrapData.classList.remove('hidden');
            if (wrapValorIva) wrapValorIva.classList.toggle('hidden', valor !== 'Abastecimento');

            toggleCustoLitros();
            toggleCustoKm();

            if (kmEl) kmEl.setAttribute('required', 'required');
            if (litrosEl && valor === 'Abastecimento') litrosEl.setAttribute('required', 'required');
            else if (litrosEl) litrosEl.removeAttribute('required');
        }
        if (window.lucide) lucide.createIcons();
    }

    function mostrarInputTalaoEAnalisar() {
        var matricula = document.getElementById('custo_matricula').value.trim();
        var tipo = document.getElementById('custo_tipo').value.trim();
        if (!matricula) { notify('Selecione um veículo.', 'error'); return; }
        if (!tipo) { notify('Selecione o tipo de registo.', 'error'); return; }
        if (tipo === 'Portagens') { notify('Para Portagens, use "Inserir Manualmente".', 'info'); return; }

        var faturaUpload = document.getElementById('faturaUpload');
        if (faturaUpload) faturaUpload.click();
    }

    function initFormCustos() {
        var step1 = document.getElementById('step1Custos');
        var step2 = document.getElementById('step2Custos');
        var wrapFile = document.getElementById('wrapCustoFile');

        if (step1) step1.classList.remove('hidden');
        if (step2) step2.classList.add('hidden');
        if (wrapFile) wrapFile.classList.add('hidden');

        _pendingOcrCustoB64 = null;
        _pendingOcrCustoData = null;

        var fileEl = document.getElementById('custo_file');
        if (fileEl) { fileEl.value = ''; fileEl.onchange = null; }

        var faturaUpload = document.getElementById('faturaUpload');
        if (faturaUpload) {
            faturaUpload.value = '';
            faturaUpload.onchange = function () {
                if (faturaUpload.files && faturaUpload.files[0]) analisarTalaoCusto();
            };
        }

        var valorPagoEl = document.getElementById('custo_valor_pago');
        if (valorPagoEl) valorPagoEl.value = '';
        var fornecedorEl = document.getElementById('custo_fornecedor');
        if (fornecedorEl) fornecedorEl.value = '';
        var dataEl = document.getElementById('custo_data');
        if (dataEl) dataEl.value = '';
        var kmEl = document.getElementById('custo_km');
        if (kmEl) kmEl.value = '';
        var litrosEl = document.getElementById('custo_litros');
        if (litrosEl) litrosEl.value = '';
        var obsEl = document.getElementById('custo_observacoes');
        if (obsEl) obsEl.value = '';
        var valorIvaEl = document.getElementById('custo_valor_iva');
        if (valorIvaEl) valorIvaEl.value = '';

        var wrapValorIva = document.getElementById('wrapCustoValorIva');
        if (wrapValorIva) wrapValorIva.classList.add('hidden');

        toggleCustoLitros();
        toggleCustoKm();
        handleTipoDespesaChange();

        if (window.lucide) lucide.createIcons();
    }

    function resetFormCustos() {
        initFormCustos();
    }

    function resetFormularioCustos() {
        var ids = ['custo_fornecedor', 'custo_data', 'custo_valor_pago', 'custo_valor_iva', 'custo_km', 'custo_litros', 'custo_observacoes'];
        ids.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.value = '';
        });
        var statusEl = document.getElementById('custo_status_pagamento');
        if (statusEl) statusEl.value = 'Pago';

        var faturaUpload = document.getElementById('faturaUpload');
        if (faturaUpload) faturaUpload.value = '';
        var custoFile = document.getElementById('custo_file');
        if (custoFile) { custoFile.value = ''; custoFile.onchange = null; }

        _pendingOcrCustoData = null;
        _pendingOcrCustoB64 = null;

        var step2 = document.getElementById('step2Custos');
        var step1 = document.getElementById('step1Custos');
        var step1Buttons = document.getElementById('step1-buttons');
        if (step2) step2.classList.add('hidden');
        if (step1) step1.classList.remove('hidden');
        if (step1Buttons) step1Buttons.classList.remove('hidden');

        var wrapFile = document.getElementById('wrapCustoFile');
        if (wrapFile) wrapFile.classList.add('hidden');
        var resumoEl = document.getElementById('ocrResumoCustos');
        if (resumoEl) { resumoEl.classList.add('hidden'); resumoEl.innerHTML = ''; }

        if (faturaUpload) {
            faturaUpload.onchange = function () {
                if (faturaUpload.files && faturaUpload.files[0]) analisarTalaoCusto();
            };
        }
        toggleCustoLitros();
        toggleCustoKm();
        if (window.lucide) lucide.createIcons();
    }

    function cancelarRegistoCusto() {
        resetFormularioCustos();
    }

    function inserirManualmenteCusto() {
        var matricula = document.getElementById('custo_matricula').value.trim();
        var tipo = document.getElementById('custo_tipo').value.trim();
        if (!matricula) { notify('Selecione um veículo.', 'error'); return; }
        if (!tipo) { notify('Selecione o tipo de registo.', 'error'); return; }
        _pendingOcrCustoData = null;
        showStep2Custos(null);
    }

    function analisarTalaoCusto() {
        var fileEl = document.getElementById('faturaUpload');
        if (!fileEl || !fileEl.files || !fileEl.files[0]) fileEl = document.getElementById('custo_file');

        var matricula = document.getElementById('custo_matricula').value.trim();
        var tipo = document.getElementById('custo_tipo').value.trim();

        if (!matricula) { notify('Selecione um veículo.', 'error'); return; }
        if (!tipo) { notify('Selecione o tipo de registo.', 'error'); return; }
        if (!fileEl || !fileEl.files || !fileEl.files[0]) { notify('Carregue o talão ou fatura.', 'error'); return; }

        var file = fileEl.files[0];
        if (!file.type.match(/^image\//) && file.type !== 'application/pdf') { notify('Formato não suportado. Use imagem ou PDF.', 'error'); return; }

        notify('A analisar documento com IA...', 'info');
        var loadingUI = document.getElementById('loadingUI');
        if (loadingUI) loadingUI.classList.remove('hidden');

        var reader = new FileReader();
        reader.onload = function (ev) {
            var dataUrl = ev.target.result;
            if (file.type.match(/^image\//)) {
                // Assumindo que compressImageForOCR está disponível via JS_CC_Logistica
                if (typeof compressImageForOCR === 'function') {
                    compressImageForOCR(dataUrl, 1200, 0.7).then(function (compressed) {
                        _pendingOcrCustoB64 = compressed;
                        google.script.run.withFailureHandler(function (err) {
                            if (loadingUI) loadingUI.classList.add('hidden');
                            notify('Erro OCR: ' + (err || 'Desconhecido'), 'error');
                        }).withSuccessHandler(function (ocrRes) {
                            if (loadingUI) loadingUI.classList.add('hidden');
                            if (!ocrRes.success) { notify('Erro na AI: ' + (ocrRes.error || ''), 'error'); return; }
                            _pendingOcrCustoData = ocrRes.data || {};
                            showStep2Custos(ocrRes);
                        }).extractDocument(compressed);
                    });
                } else {
                    if (loadingUI) loadingUI.classList.add('hidden');
                    notify('Função de compressão não encontrada.', 'error');
                }
            } else {
                notify('PDF: use uma captura de ecrã da primeira página. Carregue uma imagem.', 'warning');
                if (loadingUI) loadingUI.classList.add('hidden');
            }
        };
        if (file.type.match(/^image\//)) reader.readAsDataURL(file);
    }

    function showStep2Custos(ocrRes) {
        var isManual = !ocrRes || !ocrRes.data;
        var data = (ocrRes && ocrRes.data) ? ocrRes.data : {};

        var cab = data.cabecalho || {};
        var fornecedor = cab.fornecedor || data.fornecedor || '';
        var dataStr = cab.data || data.data || '';
        var linhas = data.linhas || [];
        var valorTotal = parseFloat(data.valorTotal || data.valor_total) || 0;

        if (valorTotal === 0 && linhas.length) {
            linhas.forEach(function (l) {
                var q = parseFloat(l.quantidade) || 1;
                var p = parseFloat(l.preco_custo) || 0;
                valorTotal += q * p;
            });
        }

        var resumoEl = document.getElementById('ocrResumoCustos');
        if (resumoEl) {
            if (isManual) {
                resumoEl.innerHTML = '<div class="font-bold text-slate-900 dark:text-white">Registo manual</div><div class="text-slate-600 dark:text-slate-400">Preencha os campos abaixo.</div>';
                resumoEl.classList.remove('hidden');
            } else {
                var valorIvaStr = (data.valorIva != null || data.valor_iva != null) ? (parseFloat(data.valorIva || data.valor_iva) || 0).toFixed(2) + ' €' : '—';
                var litrosStr = (data.litros != null && data.litros !== '') ? (parseFloat(data.litros) || '—') : '—';
                resumoEl.innerHTML = '<div class="font-bold text-slate-900 dark:text-white">Dados extraídos (pode editar):</div>' +
                    '<div>Fornecedor: ' + (fornecedor || '—') + '</div>' +
                    '<div>Data: ' + (dataStr || '—') + '</div>' +
                    '<div>Valor Total: ' + (valorTotal > 0 ? valorTotal.toFixed(2) + ' €' : '—') + '</div>' +
                    (valorIvaStr !== '—' ? '<div>Valor IVA: ' + valorIvaStr + '</div>' : '') +
                    (litrosStr !== '—' ? '<div>Litros: ' + litrosStr + '</div>' : '');
                resumoEl.classList.remove('hidden');
            }
        }

        var valorIva = parseFloat(data.valorIva || data.valor_iva) || 0;
        var litrosOcr = parseFloat(data.litros) || (data.litros !== undefined && data.litros !== null ? parseFloat(data.litros) : 0);

        _pendingOcrCustoData = isManual ? null : Object.assign({}, data, { valorTotal: valorTotal, fornecedor: fornecedor, dataDoc: dataStr, valorIva: valorIva, litros: litrosOcr });

        var step1 = document.getElementById('step1Custos');
        var step2 = document.getElementById('step2Custos');
        if (step1) step1.classList.add('hidden');
        if (step2) step2.classList.remove('hidden');

        var fornecedorEl = document.getElementById('custo_fornecedor');
        var dataEl = document.getElementById('custo_data');
        var valorPagoEl = document.getElementById('custo_valor_pago');
        var valorIvaEl = document.getElementById('custo_valor_iva');
        var kmEl = document.getElementById('custo_km');
        var litrosEl = document.getElementById('custo_litros');
        var wrapValorIva = document.getElementById('wrapCustoValorIva');

        if (fornecedorEl) fornecedorEl.value = fornecedor || '';
        if (dataEl) {
            if (dataStr) {
                var m = String(dataStr).match(/^(\d{4})-(\d{2})-(\d{2})/);
                dataEl.value = m ? dataStr : ((m = String(dataStr).match(/(\d{2})\/(\d{2})\/(\d{4})/)) ? m[3] + '-' + m[2] + '-' + m[1] : '');
            } else {
                var hoje = new Date();
                dataEl.value = hoje.getFullYear() + '-' + String(hoje.getMonth() + 1).padStart(2, '0') + '-' + String(hoje.getDate()).padStart(2, '0');
            }
        }
        if (valorPagoEl) valorPagoEl.value = (valorTotal > 0 ? valorTotal.toFixed(2) : '');
        if (valorIvaEl) valorIvaEl.value = (valorIva > 0 ? valorIva.toFixed(2) : '');

        var tipoAtual = document.getElementById('custo_tipo');
        if (wrapValorIva) wrapValorIva.classList.toggle('hidden', !(tipoAtual && tipoAtual.value === 'Abastecimento'));

        if (kmEl) kmEl.value = '';
        if (litrosEl) litrosEl.value = (litrosOcr > 0 ? litrosOcr.toString() : '');

        toggleCustoLitros();
        toggleCustoKm();
        if (window.lucide) lucide.createIcons();
    }

    function confirmarRegistarDespesa() {
        var matricula = document.getElementById('custo_matricula').value.trim();
        var tipo = document.getElementById('custo_tipo').value.trim();
        var isPortagens = tipo === 'Portagens';

        var kmEl = document.getElementById('custo_km');
        var km = kmEl ? kmEl.value.trim() : '';

        var litrosEl = document.getElementById('custo_litros');
        var litros = litrosEl ? litrosEl.value : '';

        var valorPagoEl = document.getElementById('custo_valor_pago');
        var valorPagoStr = valorPagoEl ? valorPagoEl.value.trim() : '';

        var fornecedorEl = document.getElementById('custo_fornecedor');
        var fornecedor = fornecedorEl ? fornecedorEl.value.trim() : '';

        var dataEl = document.getElementById('custo_data');
        var dataDoc = dataEl ? dataEl.value.trim() : '';

        var obsEl = document.getElementById('custo_observacoes');
        var observacoes = obsEl ? obsEl.value.trim() : '';

        if (!matricula) { notify('Selecione um veículo.', 'error'); return; }
        if (!isPortagens && (!km || isNaN(parseFloat(km)))) { notify('Km Atuais é obrigatório.', 'error'); return; }
        if (tipo === 'Abastecimento' && (!litros || isNaN(parseFloat(litros)))) { notify('Litros é obrigatório para Abastecimento.', 'error'); return; }

        var valorTotal = parseFloat(valorPagoStr) || 0;
        if (valorTotal <= 0) {
            var ocrData = _pendingOcrCustoData || {};
            valorTotal = ocrData.valorTotal || 0;
        }
        if (valorTotal <= 0) { notify('Valor Pago é obrigatório.', 'error'); return; }

        if (isPortagens && !fornecedor) fornecedor = 'Via Verde';

        var valorIvaEl = document.getElementById('custo_valor_iva');
        var valorIva = valorIvaEl ? (parseFloat(valorIvaEl.value) || 0) : 0;
        if (valorIva <= 0 && _pendingOcrCustoData && _pendingOcrCustoData.valorIva) valorIva = parseFloat(_pendingOcrCustoData.valorIva) || 0;

        var statusPagamentoEl = document.getElementById('custo_status_pagamento');
        var statusPagamento = statusPagamentoEl ? (statusPagamentoEl.value || 'Pago') : 'Pago';

        var payload = {
            matricula: matricula,
            tipo: tipo,
            kmAtuais: isPortagens ? '' : km,
            litros: tipo === 'Abastecimento' ? litros : '',
            valorTotal: valorTotal,
            valorIva: valorIva,
            fornecedor: fornecedor,
            dataDoc: dataDoc,
            observacoes: observacoes,
            statusPagamento: statusPagamento,
            ocrData: _pendingOcrCustoData || {},
            imageBase64: _pendingOcrCustoB64 || null
        };

        google.script.run.withSuccessHandler(function (res) {
            if (res && res.success) {
                notify('Despesa registada com sucesso.', 'success');
                resetFormularioCustos();
                loadCustos();
            } else {
                notify(res && res.error ? res.error : 'Erro ao registar.', 'error');
            }
        }).withFailureHandler(function (err) {
            notify('Erro: ' + (err || 'Desconhecido'), 'error');
        }).saveCustoFrota(payload, _ctxEmail());
    }

    function loadCustosVeiculos() {
        var sel = document.getElementById('custo_matricula');
        if (!sel) return;

        var opts = sel.querySelectorAll('option:not([value=""])');
        opts.forEach(function (o) { o.remove(); });

        var veiculos = frotaVeiculosCache || cachedVeiculos || [];
        veiculos.forEach(function (v) {
            var opt = document.createElement('option');
            opt.value = v.matricula || '';
            opt.textContent = v.matricula || '—';
            sel.appendChild(opt);
        });
    }

    function loadCustos() {
        var lista = document.getElementById('listaCustos');
        var msgVazio = document.getElementById('msgCustosVazio');
        if (!lista) return;

        lista.innerHTML = '';
        if (msgVazio) msgVazio.classList.add('hidden');

        google.script.run.withSuccessHandler(renderCustosFrota).withFailureHandler(function () {
            if (msgVazio) { msgVazio.textContent = 'Erro ao carregar registos.'; msgVazio.classList.remove('hidden'); }
        }).getCustosFrota(_ctxEmail(), 20);
    }

    function renderCustosFrota(dados) {
        var lista = document.getElementById('listaCustos');
        var msgVazio = document.getElementById('msgCustosVazio');
        if (!lista) return;

        if (!dados || dados.length === 0) {
            if (msgVazio) msgVazio.classList.remove('hidden');
            return;
        }

        dados.forEach(function (item) {
            var valor = item.valor != null ? item.valor : (item.custoTotal != null ? item.custoTotal : 0);
            var custoStr = (valor != null && !isNaN(valor) && valor > 0) ? valor.toFixed(2) + ' €' : '—';
            var consumoStr = (item.tipo === 'Abastecimento' || item.tipo === 'Gasóleo') && item.consumoMedio && item.consumoMedio !== 'N/A' ? item.consumoMedio + ' L/100km' : '';
            var linkFoto = item.linkFoto || '';
            var iconLink = linkFoto ? '<a href="' + linkFoto.replace(/"/g, '&quot;') + '" target="_blank" rel="noopener" class="inline-flex items-center justify-center w-8 h-8 rounded-lg bg-cyan-100 dark:bg-cyan-900/30 text-cyan-600 dark:text-cyan-400 hover:bg-cyan-200 dark:hover:bg-cyan-800/50 transition ml-2" title="Abrir fatura">📎</a>' : '';

            var div = document.createElement('div');
            div.className = 'flex items-center justify-between gap-3 p-3 rounded-xl bg-slate-50 dark:bg-slate-700/50 border border-slate-100 dark:border-slate-600';
            div.innerHTML = '<div class="flex-1 min-w-0 flex items-center"><span class="font-medium text-slate-900 dark:text-white">' + (item.matricula || '—') + '</span><span class="text-sm text-slate-700 dark:text-slate-200 ml-2">' + (item.tipo || '') + '</span>' + iconLink + '</div>' +
                '<div class="text-right flex-shrink-0"><span class="font-bold text-slate-900 dark:text-white">' + custoStr + '</span>' +
                (consumoStr ? '<br><span class="text-xs text-slate-700 dark:text-slate-200">' + consumoStr + '</span>' : '') + '</div>';
            lista.appendChild(div);
        });
    }

    // ==========================================
    // EXPORT FUNCTIONS TO GLOBAL SCOPE
    // ==========================================
    window.switchFrotaTab = switchFrotaTab;
    window.diasAteData = diasAteData;
    window.getAlertaBadgeClasses = getAlertaBadgeClasses;
    window.getLavagemBadgeClasses = getLavagemBadgeClasses;
    window.diasDesdeData = diasDesdeData;
    window.formatarDataExibicao = formatarDataExibicao;
    window.loadVeiculos = loadVeiculos;
    window._renderVeiculosFromCache = _renderVeiculosFromCache;
    window.openModalVeiculo = openModalVeiculo;
    window.closeModalVeiculo = closeModalVeiculo;
    window.saveVeiculoFromModal = saveVeiculoFromModal;
    window.editVeiculo = editVeiculo;
    window.registarLavagemClick = registarLavagemClick;
    window.confirmDeleteVeiculo = confirmDeleteVeiculo;
    window.toggleCustoLitros = toggleCustoLitros;
    window.toggleCustoKm = toggleCustoKm;
    window.handleTipoDespesaChange = handleTipoDespesaChange;
    window.mostrarInputTalaoEAnalisar = mostrarInputTalaoEAnalisar;
    window.initFormCustos = initFormCustos;
    window.resetFormCustos = resetFormCustos;
    window.resetFormularioCustos = resetFormularioCustos;
    window.cancelarRegistoCusto = cancelarRegistoCusto;
    window.inserirManualmenteCusto = inserirManualmenteCusto;
    window.analisarTalaoCusto = analisarTalaoCusto;
    window.showStep2Custos = showStep2Custos;
    window.confirmarRegistarDespesa = confirmarRegistarDespesa;
    window.loadCustosVeiculos = loadCustosVeiculos;
    window.loadCustos = loadCustos;
    window.renderCustosFrota = renderCustosFrota;

})();
