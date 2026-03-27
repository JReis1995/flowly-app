/**
 * SYS_Router.js - Sistema de Routing da Web App Flowly 360
 * Centraliza o tratamento de URLs, Handshake OAuth e Redirecionamentos.
 */

// --- SERVIÇO JSON API PRINCIPAL ---
function doGet(e) {
  const params = (e && e.parameter) || {};
  const code = (params.code || "").toString();
  const state = (params.state || "").toString();
  const page = (params.page || "").toString().toLowerCase();
  const action = (params.action || "").toString().toLowerCase();
  const sessionId = (params.session_id || "").toString();
  const token = (params.token || "").toString();
  const format = (params.format || "").toString().toLowerCase();

  // 1. HANDSHAKE OAUTH SAGE CLOUD
  // Se houver 'code' e 'state', estamos a receber o callback da Sage
  if (code && state) {
    return handleSageCallback(code, state);
  }

  // 2. ENDPOINT JSON API - Se format=json ou não especificar página/action
  if (format === "json" || (!page && !action && !sessionId && !token)) {
    return serveJsonApi(params);
  }

  // 3. REDIRECIONAMENTO STRIPE SUCCESS (mantém HTML para compatibilidade)
  if (page === "success" && sessionId) {
    const result = completePurchase(sessionId);
    const jsonResponse = {
      success: true,
      type: "stripe_success",
      paymentSuccess: result.success,
      creditsAdded: result.success ? (result.credits || 0) : 0,
      message: result.success ? "Pagamento processado com sucesso" : "Erro no processamento do pagamento"
    };
    return createJsonResponse(jsonResponse);
  }

  // 4. SERVE O TEMPLATE PRINCIPAL (index.html como esqueleto, com includes por modulo)
  const template = HtmlService.createTemplateFromFile('index');
  template.setupToken = "";
  template.actionSetPassword = false;
  template.setPasswordToken = "";
  template.paymentSuccess = false;
  template.creditsAdded = 0;
  template.initialPage = "";

  if (page === "setup-password" && token) {
    template.setupToken = token;
    template.initialPage = "setup-password";
  } else if (action === "setpassword" && token) {
    template.actionSetPassword = true;
    template.setPasswordToken = token;
  } else if (action === "register" && token) {
    template.setupToken = token;
    template.initialPage = "register";
  } else if (action === "recover" && token) {
    template.actionSetPassword = true;
    template.setPasswordToken = token;
  }

  return template.evaluate()
    .setTitle('Flowly 360 Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Processa o pagamento final da Stripe, valida a sessão e adiciona créditos.
 */
function completePurchase(sessionId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const stripeSecret = props.getProperty('STRIPE_SECRET_KEY');
    if (!stripeSecret) throw new Error("STRIPE_SECRET_KEY não configurada.");

    // Evita processamento duplo (Idempotência básica)
    const checkKey = "STRIPE_PROC_" + sessionId;
    if (props.getProperty(checkKey)) {
      return { success: true, credits: 0, note: "Já processado" };
    }

    const url = "https://api.stripe.com/v1/checkout/sessions/" + sessionId;
    const options = {
      method: "get",
      headers: { "Authorization": "Bearer " + stripeSecret },
      muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(resp.getContentText());

    if (resp.getResponseCode() !== 200 || !json || json.payment_status !== 'paid') {
      return { success: false, error: "Pagamento não confirmado ou sessão inválida." };
    }

    // Extração de metadados definidos no createStripeCheckout (em MOD_SaaS.js)
    const metadata = json.metadata || {};
    const email = metadata.userEmail || metadata.email;
    const credits = parseInt(metadata.credits || "0", 10);

    if (!email || isNaN(credits) || credits <= 0) {
      return { success: false, error: "Dados da compra incompletos nos metadados." };
    }

    // Adiciona os créditos à Base de Dados Mestre
    // Nota: funçao definida em MOD_AI.js
    addCreditsToMasterDB(email, credits);
    
    // Marca como processado
    props.setProperty(checkKey, "true");

    return { success: true, credits: credits };
  } catch (e) {
    Logger.log("Erro em completePurchase: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Trata o Handshake do OAuth 2.0 da Sage Cloud.
 */
function handleSageCallback(code, state) {
  try {
    const props = PropertiesService.getScriptProperties();
    const storedState = props.getProperty("SAGE_OAUTH_STATE") || "";
    
    if (state !== storedState) {
      return HtmlService.createHtmlOutput("<p>Erro de validação OAuth (CSRF). Tente novamente.</p>");
    }
    
    const clientId = props.getProperty("SAGE_CLIENT_ID") || "";
    const clientSecret = props.getProperty("SAGE_CLIENT_SECRET") || "";
    const redirectUri = ScriptApp.getService().getUrl();
    
    if (!clientId || !clientSecret) {
      return HtmlService.createHtmlOutput("<p>Credenciais Sage (Client ID/Secret) não configuradas.</p>");
    }

    const tokenUrl = "https://oauth.accounting.sage.com/token";
    const payload = {
      grant_type: "authorization_code",
      code: code,
      redirect_uri: redirectUri,
      client_id: clientId,
      client_secret: clientSecret
    };

    const resp = UrlFetchApp.fetch(tokenUrl, {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(resp.getContentText());
    if (json.error) {
      props.deleteProperty("SAGE_OAUTH_STATE");
      return HtmlService.createHtmlOutput("<p>Erro Sage OAuth: " + (json.error_description || json.error) + "</p>");
    }

    if (json.access_token) {
      props.setProperty("SAGE_ACCESS_TOKEN", json.access_token);
      if (json.refresh_token) props.setProperty("SAGE_REFRESH_TOKEN", json.refresh_token);
    }

    props.deleteProperty("SAGE_OAUTH_STATE");
    
    // Redireciona para o URL limpo da App
    const cleanUrl = redirectUri.split("?")[0];
    return HtmlService.createHtmlOutput("<script>window.location.href='" + cleanUrl + "';</script>");
    
  } catch (err) {
    return HtmlService.createHtmlOutput("<p>Erro crítico no Handshake Sage: " + err.toString() + "</p>");
  }
}

/**
 * Função utilitária para incluir ficheiros HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Cria uma resposta JSON com headers adequados.
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Endpoint principal da API JSON - serve dados do utilizador e da app.
 */
function serveJsonApi(params) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const impersonateTarget = params.impersonate || "";
    
    // Obter contexto do cliente
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.success) {
      return createJsonResponse({
        success: false,
        error: "Contexto do cliente não encontrado",
        type: "error"
      });
    }

    // Obter dados do utilizador
    const userData = {
      email: impersonateTarget || userEmail,
      isAdmin: ctx.isAdmin,
      isMaster: ctx.isMaster,
      clientName: ctx.clientName,
      plan: ctx.plan,
      modules: ctx.modules || {},
      aiCredits: ctx.aiCredits || 0
    };

    // Obter dados da app (configurações básicas)
    const appData = {
      name: "Flowly 360 Pro",
      version: "2.0",
      baseUrl: ScriptApp.getService().getUrl(),
      timestamp: new Date().toISOString(),
      features: {
        ai: ctx.modules?.ia || false,
        logistics: ctx.modules?.logistica || false,
        hr: ctx.modules?.rh || false,
        cc: ctx.modules?.cc || false,
        admin: ctx.modules?.admin || false,
        dashboard: ctx.modules?.dashboard || false
      }
    };

    // Obter dados do dashboard se solicitado
    let dashboardData = null;
    if (params.dashboard === "true") {
      try {
        dashboardData = getDashboardData(impersonateTarget, null, null, null, null, false, null);
      } catch (e) {
        dashboardData = { success: false, error: e.toString() };
      }
    }

    const response = {
      success: true,
      type: "api_data",
      user: userData,
      app: appData,
      dashboard: dashboardData,
      message: "API Flowly 360 - Dados carregados com sucesso"
    };

    return createJsonResponse(response);

  } catch (e) {
    Logger.log("Erro em serveJsonApi: " + e.toString());
    return createJsonResponse({
      success: false,
      error: e.toString(),
      type: "error",
      message: "Erro ao carregar dados da API"
    });
  }
}

// --- SERVIÇO POST API ---
function doPost(e) {
  const params = (e && e.parameter) || {};
  const format = (params.format || "").toString().toLowerCase();
  
  // Se format=json, tratar como API
  if (format === "json") {
    return serveJsonApi(params);
  }
  
  // Para outras requisições POST, devolver erro
  return createJsonResponse({
    success: false,
    error: "Método POST não suportado para este endpoint",
    type: "error"
  });
}

/** 
 * Inicia o fluxo OAuth 2.0 da Sage Cloud. 
 * Retorna o URL de autorização para o frontend.
 */
function startOAuthFlow() {
  try {
    const clientId = PropertiesService.getScriptProperties().getProperty("SAGE_CLIENT_ID") || "";
    if (!clientId) return { success: false, error: "SAGE_CLIENT_ID não configurado." };
    
    const redirectUri = ScriptApp.getService().getUrl();
    const state = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty("SAGE_OAUTH_STATE", state);
    
    const authUrl = "https://www.sageone.com/oauth2/auth/central" +
      "?filter=apiv3.1" +
      "&country=pt" + // Ajustado para PT se necessário
      "&response_type=code" +
      "&client_id=" + encodeURIComponent(clientId) +
      "&redirect_uri=" + encodeURIComponent(redirectUri) +
      "&scope=full_access" +
      "&state=" + encodeURIComponent(state);
      
    return { success: true, authUrl: authUrl };
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  }
}