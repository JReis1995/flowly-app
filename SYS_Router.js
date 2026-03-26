/**
 * SYS_Router.js - Sistema de Routing da Web App Flowly 360
 * Centraliza o tratamento de URLs, Handshake OAuth e Redirecionamentos.
 */

// --- SERVIÇO HTML PRINCIPAL ---
function doGet(e) {
  const params = (e && e.parameter) || {};
  const code = (params.code || "").toString();
  const state = (params.state || "").toString();
  const page = (params.page || "").toString().toLowerCase();
  const action = (params.action || "").toString().toLowerCase();
  const sessionId = (params.session_id || "").toString();
  const token = (params.token || "").toString();

  // 1. HANDSHAKE OAUTH SAGE CLOUD
  // Se houver 'code' e 'state', estamos a receber o callback da Sage
  if (code && state) {
    return handleSageCallback(code, state);
  }

  // 2. PREPARAÇÃO DO TEMPLATE INDEX.HTML
  const template = HtmlService.createTemplateFromFile('Template');
  
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
      const result = completePurchase(sessionId);
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
  // Caso B: Ativação via Convite Admin ou Recuperação (setpassword)
  else if (action === "setpassword" && token) {
    template.actionSetPassword = true;
    template.setPasswordToken = token;
  }
  // Caso C: Convite de Registo Inicial
  else if (action === "register" && token) {
    template.setupToken = token;
    template.initialPage = "register";
  }
  // Caso D: Recuperação de Password
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
