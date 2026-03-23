/**
 * Flowly 360 - API Wrapper
 * Converte google.script.run para chamadas fetch REST API
 * Autor: Senior Fullstack Architect
 */

// 1) Constante com URL do WebApp Google Apps Script
const GAS_URL = "https://script.google.com/macros/s/AKfycbwbE1xesPxxbZsFJIUPG4ToDlc62Z019r9AThGE1kqMLQkXnjQY3Qsz0P45dNYLIqwITQ/exec";

/**
 * Flowly API - Wrapper para comunicação com Google Apps Script via fetch
 * Implementa chamadas RESTful POST com CORS e cache desativado
 * 
 * @param {string} functionName - Nome da função a executar no backend GAS
 * @param {...any} args - Argumentos a passar para a função
 * @returns {Promise} - Promise que resolve com o resultado ou rejeita com erro
 */
const flowlyAPI = {
    /**
     * Método principal de chamada à API
     * @param {string} functionName - Nome da função GAS a executar
     * @param {...any} args - Parâmetros para a função
     * @returns {Promise} - Promise com resultado da chamada
     */
    call: function(functionName, ...args) {
        return new Promise((resolve, reject) => {
            // Validação de parâmetros obrigatórios
            if (!functionName || typeof functionName !== 'string') {
                reject(new Error('Nome da função é obrigatório e deve ser uma string'));
                return;
            }

            if (!GAS_URL || GAS_URL === "TEU_URL_DO_WEBAPP_AQUI") {
                reject(new Error('GAS_URL não configurado. Defina o URL do seu WebApp Google Apps Script.'));
                return;
            }

            // Preparação do payload conforme especificado
            const payload = {
                function: functionName,
                parameters: args
            };

            // Configuração da requisição fetch
            const fetchOptions = {
                method: 'POST',
                mode: 'cors',
                cache: 'no-cache',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify(payload)
            };

            // Execução da chamada fetch
            fetch(GAS_URL, fetchOptions)
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`HTTP Error: ${response.status} ${response.statusText}`);
                    }
                    return response.text(); // GAS retorna texto, não JSON
                })
                .then(responseText => {
                    // Tentativa de parse JSON da resposta
                    let result;
                    try {
                        result = JSON.parse(responseText);
                    } catch (parseError) {
                        // Se não for JSON, retorna como texto simples
                        result = responseText;
                    }

                    // Verificação de erro na resposta do servidor
                    if (result && typeof result === 'object' && result.error) {
                        reject(new Error(result.error));
                        return;
                    }

                    // Sucesso - resolve com o resultado
                    resolve(result);
                })
                .catch(error => {
                    // Tratamento de erros de rede ou outros
                    console.error('Flowly API Error:', error);
                    reject(error);
                });
        });
    },

    /**
     * Método auxiliar para chamadas com timeout
     * @param {string} functionName - Nome da função GAS a executar
     * @param {number} timeoutMs - Timeout em milissegundos
     * @param {...any} args - Parâmetros para a função
     * @returns {Promise} - Promise com timeout configurado
     */
    callWithTimeout: function(functionName, timeoutMs = 30000, ...args) {
        return Promise.race([
            this.call(functionName, ...args),
            new Promise((_, reject) => 
                setTimeout(() => reject(new Error(`Timeout após ${timeoutMs}ms`)), timeoutMs)
            )
        ]);
    },

    /**
     * Método auxiliar para chamadas batch (múltiplas funções)
     * @param {Array} calls - Array de objetos {functionName, args: []}
     * @returns {Promise} - Promise que resolve com array de resultados
     */
    batchCall: function(calls) {
        if (!Array.isArray(calls)) {
            return Promise.reject(new Error('calls deve ser um array'));
        }

        const promises = calls.map(call => {
            if (!call.functionName) {
                return Promise.reject(new Error('Cada chamada deve ter functionName'));
            }
            return this.call(call.functionName, ...(call.args || []));
        });

        return Promise.all(promises);
    }
};

// Export para uso em módulos (se disponível)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { flowlyAPI, GAS_URL };
}

// Disponibilização global para uso direto no browser
if (typeof window !== 'undefined') {
    window.flowlyAPI = flowlyAPI;
    window.GAS_URL = GAS_URL;
}

/**
 * Exemplos de uso:
 * 
 * // Chamada simples
 * flowlyAPI.call('getDashboardData', 'user@example.com')
 *   .then(result => console.log('Sucesso:', result))
 *   .catch(error => console.error('Erro:', error));
 * 
 * // Chamada com timeout
 * flowlyAPI.callWithTimeout('processLargeData', 10000, data)
 *   .then(result => console.log('Processado:', result))
 *   .catch(error => console.error('Timeout ou erro:', error));
 * 
 * // Chamada batch
 * flowlyAPI.batchCall([
 *   { functionName: 'getUsers', args: [] },
 *   { functionName: 'getRecords', args: ['2024-01-01'] }
 * ]).then(results => {
 *   console.log('Users:', results[0]);
 *   console.log('Records:', results[1]);
 * });
 */
