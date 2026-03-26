/**
 * SYS_Router.js - Sistema de Routing da Web App Flowly 360
 * Centraliza o tratamento de URLs, Handshake OAuth e Redirecionamentos.
 */

// --- SERVIÇO HTML PRINCIPAL ---
function doGet(e) {
  var params = (e && e.parameter) || {};
  var code = (params.code || "").toString();
  var state = (params.state || "").toString();
  var page = (params.page || "").toString().toLowerCase();
  var action = (params.action || "").toString().toLowerCase();
  var sessionId = (params.session_id || "").toString();
  var token = (params.token || "").toString();

  // 1. HANDSHAKE OAUTH SAGE CLOUD
  // Se houver 'code' e 'state', estamos a receber o callback da Sage
  if (code && state) {
    return handleSageCallback(code, state);
  }

  // 2. PREPARAÇÃO DO TEMPLATE INDEX.HTML
  var template = HtmlService.createTemplateFromFile('Template');
  
  // Inicialização segura de variáveis de template (evita erros de "is not defined")
  template.setupToken = "";
  template.actionSetPassword = false;
  template.setPasswordToken = "";
  template.paymentSuccess = false;
  template.creditsAdded = 0;
  template.initialPage = "";

  // 3. TRATAMENTO DE REDIRECIONAMENTO DE SUCESSO DA STRIPE
  if (page === "success" && sessionId) {
    try {
      var result = completePurchase(sessionId);
      template.paymentSuccess = result.success;
      template.creditsAdded = result.success ? (result.credits || 0) : 0;
    } catch (err) {
      template.paymentSuccess = false;
      template.creditsAdded = 0;
    }
  }
  // 4. TRATAMENTO DE ATIVAÇÃO DE CONTA / DEFINIÇÃO DE PASSWORD
  else if (page === "setup-password" && token) {
    template.setupToken = token;
    template.initialPage = "setup-password";
  }
  else if (action === "setpassword" && token) {
    template.actionSetPassword = true;
    template.setPasswordToken = token;
  }
  else if (action === "register" && token) {
    template.setupToken = token;
    template.initialPage = "register";
  }
  else if (action === "recover" && token) {
    template.actionSetPassword = true;
    template.setPasswordToken = token;
  }

  // 5. RENDERIZAÇÃO FINAL
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
    var props = PropertiesService.getScriptProperties();
    var stripeSecret = props.getProperty('STRIPE_SECRET_KEY');
    if (!stripeSecret) throw new Error("STRIPE_SECRET_KEY não configurada.");

    // Evita processamento duplo (Idempotência básica)
    var checkKey = "STRIPE_PROC_" + sessionId;
    if (props.getProperty(checkKey)) {
      return { success: true, credits: 0, note: "Já processado" };
    }

    var url = "https://api.stripe.com/v1/checkout/sessions/" + sessionId;
    var options = {
      method: "get",
      headers: { "Authorization": "Bearer " + stripeSecret },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(resp.getContentText());

    if (resp.getResponseCode() !== 200 || !json || json.payment_status !== 'paid') {
      return { success: false, error: "Pagamento não confirmado ou sessão inválida." };
    }

    // Extração de metadados definidos no createStripeCheckout (em MOD_SaaS.js)
    var metadata = json.metadata || {};
    var email = metadata.userEmail || metadata.email;
    var credits = parseInt(metadata.credits || "0", 10);

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
    var props = PropertiesService.getScriptProperties();
    var storedState = props.getProperty("SAGE_OAUTH_STATE") || "";
    
    if (state !== storedState) {
      return HtmlService.createHtmlOutput("<p>Erro de validação OAuth (CSRF). Tente novamente.</p>");
    }
    
    var clientId = props.getProperty("SAGE_CLIENT_ID") || "";
    var clientSecret = props.getProperty("SAGE_CLIENT_SECRET") || "";
    var redirectUri = ScriptApp.getService().getUrl();
    
    if (!clientId || !clientSecret) {
      return HtmlService.createHtmlOutput("<p>Credenciais Sage (Client ID/Secret) não configuradas.</p>");
    }

    var tokenUrl = "https://oauth.accounting.sage.com/token";
    var payload = {
      grant_type: "authorization_code",
      code: code,
      redirect_uri: redirectUri,
      client_id: clientId,
      client_secret: clientSecret
    };

    var resp = UrlFetchApp.fetch(tokenUrl, {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    });
    
    var json = JSON.parse(resp.getContentText());
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
    var cleanUrl = redirectUri.split("?")[0];
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
 * Inicia o fluxo OAuth 2.0 da Sage Cloud. 
 * Retorna o URL de autorização para o frontend.
 */
function startOAuthFlow() {
  try {
    var clientId = PropertiesService.getScriptProperties().getProperty("SAGE_CLIENT_ID") || "";
    if (!clientId) return { success: false, error: "SAGE_CLIENT_ID não configurado." };
    
    var redirectUri = ScriptApp.getService().getUrl();
    var state = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty("SAGE_OAUTH_STATE", state);
    
    var authUrl = "https://www.sageone.com/oauth2/auth/central" +
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
