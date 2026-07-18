// ============================================================
//  UNIVERSAL AI ENGINE — AI Smart Inventory System
//  Supports ALL major AI providers via a single interface
//  Version 3.0 | June 2026
//
//  SUPPORTED PROVIDERS (just change AI_PROVIDER + AI_API_KEY):
//  ─────────────────────────────────────────────────────────
//  OpenAI-Compatible (same format, different URL):
//    • openai       → gpt-4o, gpt-4-turbo, gpt-3.5-turbo
//    • groq         → llama-3.1-70b, mixtral-8x7b (VERY FAST, FREE tier)
//    • openrouter   → 100+ models in one key (openrouter.ai)
//    • together     → togethercomputer.com models
//    • mistral      → mistral-large, mistral-medium
//    • deepseek     → deepseek-chat, deepseek-coder
//    • perplexity   → pplx-70b-online (with internet access!)
//    • fireworks    → fireworks.ai models
//    • anyscale     → meta-llama/Llama-3 etc
//    • octoai       → octoai.cloud models
//    • ollama       → any local model via your own server
//    • custom       → any OpenAI-compatible endpoint
//  ─────────────────────────────────────────────────────────
//  Native Format (separate handler):
//    • gemini       → gemini-2.5-flash, gemini-2.5-pro, gemini-2.0 (FREE)
//    • anthropic    → claude-3.5-sonnet, claude-3-haiku
//    • huggingface  → any HuggingFace inference endpoint
//  ─────────────────────────────────────────────────────────
//  Auto Mode:
//    • auto         → tries all configured keys in priority order
// ============================================================

/**
 * Universal AI Provider Integration
 * Handles interactions with Groq, Gemini, Claude, Local LLMs (Ollama)
 */

function _safeParseJSON(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error(`API returned invalid JSON (Possible Timeout/Cloudflare 524): ${text.substring(0, 150)}`);
  }
}

// ─── AI PROVIDER CONFIGURATION ───────────────────────────────
// Edit this section in Agent Settings sidebar, or directly here.
// Priority order for AUTO mode: first non-empty key wins.

const AI_PROVIDERS_CONFIG = {
  // ── OpenAI ──────────────────────────────────────────────
  openai: {
    name: "OpenAI",
    url: "https://api.openai.com/v1/chat/completions",
    format: "openai",
    defaultModel: "gpt-4o-mini",
    models: ["gpt-4o", "gpt-4o-mini", "o1", "o1-mini", "o1-preview", "gpt-4-turbo"],
    propKey: "AI_OPENAI_KEY",
    modelProp: "AI_OPENAI_MODEL",
    docs: "https://platform.openai.com/api-keys",
  },

  // ── Groq (ULTRA FAST, free tier) ────────────────────────
  groq: {
    name: "Groq",
    url: "https://api.groq.com/openai/v1/chat/completions",
    format: "openai",
    defaultModel: "llama-3.3-70b-versatile",
    models: ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "deepseek-r1-distill-llama-70b", "mixtral-8x7b-32768", "gemma2-9b-it"],
    propKey: "AI_GROQ_KEY",
    modelProp: "AI_GROQ_MODEL",
    docs: "https://console.groq.com/keys",
    note: "FREE tier available — very fast inference"
  },

  // ── OpenRouter (100+ models, one key) ───────────────────
  openrouter: {
    name: "OpenRouter",
    url: "https://openrouter.ai/api/v1/chat/completions",
    format: "openai",
    defaultModel: "meta-llama/llama-3.1-8b-instruct:free",
    models: [
      "openai/gpt-4o",
      "openai/gpt-4o-mini",
      "anthropic/claude-3.5-sonnet",
      "meta-llama/llama-3.1-70b-instruct",
      "meta-llama/llama-3.1-8b-instruct:free",
      "google/gemini-flash-1.5",
      "mistralai/mistral-large",
      "deepseek/deepseek-chat",
      "qwen/qwen-2.5-72b-instruct",
      "nousresearch/hermes-3-llama-3.1-70b",
    ],
    propKey: "AI_OPENROUTER_KEY",
    modelProp: "AI_OPENROUTER_MODEL",
    docs: "https://openrouter.ai/keys",
    note: "Access 100+ models with one key. Some models are FREE.",
    extraHeaders: {
      "HTTP-Referer": "https://ai-smart-inventory.app",
      "X-Title": "AI Smart Inventory System",
    }
  },

  // ── Google Gemini (FREE tier) ────────────────────────────
  gemini: {
    name: "Google Gemini",
    url: "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
    format: "gemini",
    defaultModel: "gemini-2.5-flash",
    models: ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.5-flash", "gemini-2.5-pro"],
    propKey: "AI_GEMINI_KEY",
    modelProp: "AI_GEMINI_MODEL",
    docs: "https://aistudio.google.com/apikey",
    note: "FREE tier: 15 req/min, 1M tokens/day"
  },

  // ── Anthropic (Claude) ───────────────────────────────────
  anthropic: {
    name: "Anthropic Claude",
    url: "https://api.anthropic.com/v1/messages",
    format: "anthropic",
    defaultModel: "claude-3-haiku-20240307",
    models: ["claude-3-5-sonnet-20241022", "claude-3-haiku-20240307", "claude-3-opus-20240229"],
    propKey: "AI_ANTHROPIC_KEY",
    modelProp: "AI_ANTHROPIC_MODEL",
    docs: "https://console.anthropic.com/",
  },

  // ── Mistral AI ───────────────────────────────────────────
  mistral: {
    name: "Mistral AI",
    url: "https://api.mistral.ai/v1/chat/completions",
    format: "openai",
    defaultModel: "mistral-small-latest",
    models: ["mistral-large-latest", "mistral-medium-latest", "mistral-small-latest", "open-mixtral-8x22b"],
    propKey: "AI_MISTRAL_KEY",
    modelProp: "AI_MISTRAL_MODEL",
    docs: "https://console.mistral.ai/api-keys",
  },

  // ── DeepSeek ─────────────────────────────────────────────
  deepseek: {
    name: "DeepSeek",
    url: "https://api.deepseek.com/chat/completions",
    format: "openai",
    defaultModel: "deepseek-chat",
    models: ["deepseek-chat", "deepseek-coder"],
    propKey: "AI_DEEPSEEK_KEY",
    modelProp: "AI_DEEPSEEK_MODEL",
    docs: "https://platform.deepseek.com/api_keys",
    note: "Very affordable pricing"
  },

  // ── Together AI ──────────────────────────────────────────
  together: {
    name: "Together AI",
    url: "https://api.together.xyz/v1/chat/completions",
    format: "openai",
    defaultModel: "meta-llama/Meta-Llama-3.1-70B-Instruct-Turbo",
    models: [
      "meta-llama/Meta-Llama-3.1-70B-Instruct-Turbo",
      "meta-llama/Meta-Llama-3.1-8B-Instruct-Turbo",
      "mistralai/Mixtral-8x7B-Instruct-v0.1",
      "Qwen/Qwen2.5-72B-Instruct-Turbo",
    ],
    propKey: "AI_TOGETHER_KEY",
    modelProp: "AI_TOGETHER_MODEL",
    docs: "https://api.together.ai/settings/api-keys",
  },

  // ── Perplexity (with INTERNET access!) ──────────────────
  perplexity: {
    name: "Perplexity AI",
    url: "https://api.perplexity.ai/chat/completions",
    format: "openai",
    defaultModel: "llama-3.1-sonar-small-128k-online",
    models: ["llama-3.1-sonar-large-128k-online", "llama-3.1-sonar-small-128k-online"],
    propKey: "AI_PERPLEXITY_KEY",
    modelProp: "AI_PERPLEXITY_MODEL",
    docs: "https://www.perplexity.ai/settings/api",
    note: "Has internet access — can look up real-time info"
  },

  // ── Fireworks AI ─────────────────────────────────────────
  fireworks: {
    name: "Fireworks AI",
    url: "https://api.fireworks.ai/inference/v1/chat/completions",
    format: "openai",
    defaultModel: "accounts/fireworks/models/llama-v3p1-70b-instruct",
    models: ["accounts/fireworks/models/llama-v3p1-70b-instruct", "accounts/fireworks/models/mixtral-8x7b-instruct"],
    propKey: "AI_FIREWORKS_KEY",
    modelProp: "AI_FIREWORKS_MODEL",
    docs: "https://fireworks.ai/account/api-keys",
  },

  // ── Hugging Face ─────────────────────────────────────────
  huggingface: {
    name: "Hugging Face",
    url: "https://api-inference.huggingface.co/models/{model}/v1/chat/completions",
    format: "openai",
    defaultModel: "meta-llama/Meta-Llama-3-8B-Instruct",
    models: ["meta-llama/Meta-Llama-3-8B-Instruct", "mistralai/Mistral-7B-Instruct-v0.3"],
    propKey: "AI_HF_KEY",
    modelProp: "AI_HF_MODEL",
    docs: "https://huggingface.co/settings/tokens",
    note: "Use any open-source model on HuggingFace for free"
  },

  // ── Ollama (LOCAL model on your own server) ──────────────
  ollama: {
    name: "Ollama (Local)",
    url: "{baseUrl}/api/chat",    // e.g. http://your-server:11434
    format: "ollama",
    defaultModel: "llama3.1",
    models: ["llama3.1", "llama3.1:70b", "mistral", "codellama", "gemma2"],
    propKey: "AI_OLLAMA_URL",     // stores base URL instead of API key
    modelProp: "AI_OLLAMA_MODEL",
    docs: "https://ollama.com",
    note: "Run any model locally — no API costs"
  },

  // ── Custom OpenAI-Compatible Endpoint ────────────────────
  custom: {
    name: "Custom (OpenAI-Compatible)",
    url: "{baseUrl}/v1/chat/completions",
    format: "openai",
    defaultModel: "default",
    models: [],
    propKey: "AI_CUSTOM_URL",     // stores full base URL
    modelProp: "AI_CUSTOM_MODEL",
    docs: "",
    note: "Any API that follows OpenAI format: LM Studio, vLLM, text-generation-webui, etc."
  },

  // ── 9router (Local/Proxy AI Router) ────────────────────────
  "9router": {
    name: "9router (Local Proxy)",
    url: "{baseUrl}/v1/chat/completions",
    format: "openai",
    defaultModel: "default",
    models: ["default", "llama3", "gpt-4o", "gemini-2.5-pro"],
    propKey: "AI_9ROUTER_URL",     // stores the Ngrok/Cloudflare URL
    modelProp: "AI_9ROUTER_MODEL",
    docs: "https://github.com/9router/9router",
    note: "Gunakan Ngrok/Cloudflare Tunnel jika dijalankan di localhost!"
  },
};

// Priority order for AUTO mode
const AI_AUTO_PRIORITY = ["gemini", "groq", "openrouter", "openai", "anthropic", "mistral", "deepseek", "together", "perplexity", "fireworks", "huggingface", "ollama", "custom", "9router"];

// ─── UNIVERSAL AI CALLER ─────────────────────────────────────
/**
 * Call the AI with automatic provider selection.
 * Uses the configured provider from Script Properties, or AUTO mode.
 *
 * @param {string} prompt - The prompt to send
 * @param {string} [systemPrompt] - Optional system prompt
 * @returns {string} AI response text
 */
function callAI(prompt, systemPrompt) {
  const props = _getScriptProps();
  const providerName = (props.getProperty("AI_PROVIDER") || "auto").toLowerCase();

  if (providerName === "auto") {
    return _callAIAuto(prompt, systemPrompt);
  }

  const config = AI_PROVIDERS_CONFIG[providerName];
  if (!config) throw new Error(`Unknown AI provider: ${providerName}`);

  return _callProvider(config, prompt, systemPrompt);
}

/**
 * Auto mode: tries providers in priority order until one succeeds
 */
function _callAIAuto(prompt, systemPrompt) {
  const errors = [];
  for (const providerName of AI_AUTO_PRIORITY) {
    const config = AI_PROVIDERS_CONFIG[providerName];
    const key = _getScriptProps().getProperty(config.propKey);
    if (!key) continue; // skip unconfigured providers
    try {
      const result = _callProvider(config, prompt, systemPrompt);
      Logger.log(`[AI Auto] Used: ${config.name}`);
      return result;
    } catch (err) {
      errors.push(`${config.name}: ${err.message}`);
      Logger.log(`[AI Auto] ${config.name} failed: ${err.message}`);
    }
  }
  throw new Error("All AI providers failed:\n" + errors.join("\n"));
}

/**
 * Call a specific provider based on its config
 */
function _callProvider(config, prompt, systemPrompt) {
  switch (config.format) {
    case "openai":     return _callOpenAICompatible(config, prompt, systemPrompt);
    case "gemini":     return _callGeminiProvider(config, prompt, systemPrompt);
    case "anthropic":  return _callAnthropicProvider(config, prompt, systemPrompt);
    case "ollama":     return _callOllamaProvider(config, prompt, systemPrompt);
    default: throw new Error(`Unknown format: ${config.format}`);
  }
}

// ─── OPENAI-COMPATIBLE HANDLER ────────────────────────────────
// Works with: OpenAI, Groq, OpenRouter, Together, Mistral,
//             DeepSeek, Perplexity, Fireworks, HuggingFace,
//             and ANY other OpenAI-compatible API
function _callOpenAICompatible(config, prompt, systemPrompt) {
  const props = _getScriptProps();
  let apiKey = (props.getProperty(config.propKey) || "").replace(/\s+/g, "");
  const model = (props.getProperty(config.modelProp) || config.defaultModel).replace(/\s+/g, "");

  // For custom endpoint, apiKey prop stores the base URL
  let url = config.url;
  if (config.name === "Custom (OpenAI-Compatible)") {
    url = (props.getProperty(config.propKey) || "").replace(/\s+/g, "");
    apiKey = (props.getProperty("AI_CUSTOM_API_KEY") || "").replace(/\s+/g, "");
    url = url.replace(/\/$/, "") + "/v1/chat/completions";
  } else if (config.name === "9router (Local Proxy)") {
    url = (props.getProperty(config.propKey) || "").replace(/\/$/, "").replace(/\s+/g, "") + "/v1/chat/completions";
    apiKey = (props.getProperty("AI_CUSTOM_API_KEY") || "").replace(/\s+/g, "");
  } else if (config.name === "Hugging Face") {
    url = url.replace("{model}", model);
  }

  const messages = [];
  if (systemPrompt) messages.push({ role: "system", content: systemPrompt });
  messages.push({ role: "user", content: prompt });

  const headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + apiKey,
    "Bypass-Tunnel-Reminder": "true",
    "ngrok-skip-browser-warning": "69420"
  };

  // Extra headers for specific providers (e.g. OpenRouter)
  if (config.extraHeaders) {
    Object.assign(headers, config.extraHeaders);
  }

  const payload = {
    model: model,
    messages: messages,
    temperature: 0.1,
    max_tokens: 1024,
  };

  // Force JSON response where supported (OpenAI, Groq, Together, 9router, etc.)
  const providerKey = config.propKey.replace("AI_", "").replace("_KEY", "").toLowerCase();
  if (["openai", "groq", "together", "deepseek", "9router_url", "custom_url"].includes(providerKey)) {
    payload.response_format = { type: "json_object" };
  }

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = resp.getResponseCode();
  const text   = resp.getContentText();
  const json   = _safeParseJSON(text);

  if (status !== 200 || json.error) {
    const errMsg = json.error?.message || json.error || `HTTP ${status}: ${text.substring(0, 200)}`;
    throw new Error(errMsg);
  }

  return json.choices[0].message.content;
}

// ─── GOOGLE GEMINI HANDLER ────────────────────────────────────
function _callGeminiProvider(config, prompt, systemPrompt) {
  const props  = _getScriptProps();
  const apiKey = (props.getProperty(config.propKey) || "").replace(/\s+/g, "");
  const model  = (props.getProperty(config.modelProp) || config.defaultModel).replace(/\s+/g, "");
  const url    = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  let combinedText = "";
  if (systemPrompt) combinedText += "System: " + systemPrompt + "\n\n";
  combinedText += prompt;

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: combinedText }] }],
      generationConfig: { 
        temperature: 0.1, 
        maxOutputTokens: 1024,
        responseMimeType: "application/json"
      },
      safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      ],
    }),
    muteHttpExceptions: true,
  });

  const json = _safeParseJSON(resp.getContentText());
  if (json.error) throw new Error(json.error.message);
  let fullText = "";
  if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts) {
    const parts = json.candidates[0].content.parts;
    const finalPart = parts.find(p => !p.thought) || parts[parts.length - 1];
    if (finalPart && finalPart.text) {
      fullText = finalPart.text;
    }
  }
  return fullText.trim();
}

// ─── ANTHROPIC CLAUDE HANDLER ────────────────────────────────
function _callAnthropicProvider(config, prompt, systemPrompt) {
  const props  = _getScriptProps();
  const apiKey = (props.getProperty(config.propKey) || "").replace(/\s+/g, "");
  const model  = (props.getProperty(config.modelProp) || config.defaultModel).replace(/\s+/g, "");

  const body = {
    model: model,
    max_tokens: 1024,
    messages: [{ role: "user", content: prompt }],
  };
  if (systemPrompt) body.system = systemPrompt;

  const resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const json = _safeParseJSON(resp.getContentText());
  if (json.error) throw new Error(json.error.message);
  return json.content[0].text;
}

// ─── OLLAMA (LOCAL) HANDLER ───────────────────────────────────
function _callOllamaProvider(config, prompt, systemPrompt) {
  const props   = _getScriptProps();
  const baseUrl = (props.getProperty(config.propKey) || "http://localhost:11434").replace(/\/$/, "").replace(/\s+/g, "");
  const model   = (props.getProperty(config.modelProp) || config.defaultModel).replace(/\s+/g, "");

  const messages = [];
  if (systemPrompt) messages.push({ role: "system", content: systemPrompt });
  messages.push({ role: "user", content: prompt });

  const resp = UrlFetchApp.fetch(baseUrl + "/api/chat", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ model, messages, stream: false }),
    muteHttpExceptions: true,
    });
  
  const json = _safeParseJSON(resp.getContentText());
  if (json.error) throw new Error(json.error);
  return json.message.content;
}

// ─── SETTINGS SIDEBAR (ENHANCED) ─────────────────────────────
function showAIProviderSettings() {
  const props = _getScriptProps();
  const current = props.getProperty("AI_PROVIDER") || "auto";

  const providerRows = Object.entries(AI_PROVIDERS_CONFIG).map(([key, cfg]) => {
    const savedKey = props.getProperty(cfg.propKey) || "";
    const savedModel = props.getProperty(cfg.modelProp) || cfg.defaultModel;
    const modelOptions = cfg.models.map(m => `<option value="${m}">${m}</option>`).join("");
    const isOllama = key === "ollama" || key === "custom" || key === "9router";
    const isActive = current === key;

    return `
      <div class="provider-card ${isActive ? "active" : ""}" id="card-${key}" onclick="selectProvider('${key}')">
        <div class="prov-header">
          <div class="prov-radio ${isActive ? "radio-on" : ""}"></div>
          <span class="prov-name">${cfg.name}</span>
          ${cfg.note ? `<span class="badge">${cfg.note}</span>` : ""}
        </div>
        <div class="prov-body ${isActive ? "" : "hidden"}" id="body-${key}">
          <label>${isOllama ? "Base URL / Server URL" : "API Key"}</label>
          <input type="text" id="key-${key}" value="${savedKey}"
            placeholder="${isOllama ? (key === "ollama" ? "http://your-server:11434" : (key === "9router" ? "https://ai.ngrok.io" : "http://localhost:1234")) : "Paste your API key here"}" />
          ${key === "custom" || key === "9router" ? `<label>Custom API Key (Bearer)</label><input type="text" id="customApiKey_${key}" value="${props.getProperty("AI_CUSTOM_API_KEY") || ""}" placeholder="Optional auth token" />` : ""}
          <label>Model</label>
            ${cfg.models.length > 0
              ? `<input type="text" id="model-${key}" value="${savedModel}" list="list-${key}" placeholder="Select or type model..." style="width:100%; padding:8px; background:#0F172A; border:1px solid #334155; border-radius:6px; color:#F1F5F9;" />
                 <datalist id="list-${key}">${modelOptions}</datalist>`
              : `<input type="text" id="model-${key}" value="${savedModel}" placeholder="Enter model name" />`
            }
          ${cfg.docs ? `<a href="${cfg.docs}" target="_blank" class="docs-link">📖 Get API Key →</a>` : ""}
        </div>
      </div>`;
  }).join("");

  const autoRow = `
    <div class="provider-card ${current === "auto" ? "active" : ""}" id="card-auto" onclick="selectProvider('auto')" style="border-color:#6366f1;">
      <div class="prov-header">
        <div class="prov-radio ${current === "auto" ? "radio-on" : ""}" style="background:#6366f1;"></div>
        <span class="prov-name">⚡ AUTO MODE (Recommended)</span>
        <span class="badge" style="background:#6366f1;">Tries all configured keys in order</span>
      </div>
      <div class="prov-body ${current === "auto" ? "" : "hidden"}" id="body-auto">
        <p style="color:#94A3B8;font-size:12px;line-height:1.5;">Auto mode tries each API key you've configured, in this order:<br>${AI_AUTO_PRIORITY.join(" → ")}<br><br>Just configure as many keys as you have below — AI will use the first working one.</p>
      </div>
    </div>`;

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; }
        body { background: #0F172A; color: #E2E8F0; padding: 14px; font-size: 13px; }
        h2 { color: #10B981; font-size: 15px; margin-bottom: 14px; display: flex; align-items: center; gap: 8px; }
        .provider-card { background: #1E293B; border: 1px solid #334155; border-radius: 10px; margin-bottom: 8px; cursor: pointer; transition: border-color 0.2s; overflow: hidden; }
        .provider-card.active { border-color: #10B981; }
        .prov-header { padding: 12px 14px; display: flex; align-items: center; gap: 10px; }
        .prov-radio { width: 14px; height: 14px; border-radius: 50%; border: 2px solid #475569; flex-shrink: 0; transition: background 0.2s; }
        .prov-radio.radio-on { background: #10B981; border-color: #10B981; }
        .prov-name { font-weight: 700; font-size: 13px; flex: 1; }
        .badge { font-size: 10px; background: #064E3B; color: #6EE7B7; padding: 2px 7px; border-radius: 20px; white-space: nowrap; }
        .prov-body { padding: 0 14px 14px; border-top: 1px solid #334155; }
        .hidden { display: none; }
        label { display: block; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.7px; color: #64748B; margin-top: 10px; margin-bottom: 4px; }
        input, select { width: 100%; padding: 8px 10px; background: #0F172A; border: 1px solid #334155; border-radius: 6px; color: #F1F5F9; font-size: 12px; font-family: monospace; outline: none; }
        input:focus, select:focus { border-color: #10B981; }
        select { font-family: 'Segoe UI', sans-serif; }
        .docs-link { display: inline-block; margin-top: 8px; font-size: 11px; color: #10B981; text-decoration: none; }
        .btn-row { display: flex; gap: 8px; margin-top: 14px; }
        .btn { flex: 1; padding: 11px; border: none; border-radius: 7px; font-size: 13px; font-weight: 700; cursor: pointer; }
        .btn-primary { background: #10B981; color: white; }
        .btn-secondary { background: #1E293B; color: #94A3B8; border: 1px solid #334155; }
        .btn:hover { opacity: 0.9; }
        .test-result { margin-top: 10px; padding: 8px 12px; border-radius: 6px; font-size: 12px; display: none; }
        .test-ok { background: #052E16; border: 1px solid #166534; color: #86EFAC; }
        .test-err { background: #450A0A; border: 1px solid #991B1B; color: #FCA5A5; }
        .section-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #6366F1; margin: 16px 0 8px; padding-left: 4px; }
      </style>
    </head>
    <body>
      <h2>🤖 Universal AI Engine</h2>
      <p style="color:#94A3B8;font-size:12px;margin-bottom:14px;">Configure one or more AI providers. System uses the first working key automatically.</p>

      <div class="section-title">⚡ Mode</div>
      ${autoRow}

      <div class="section-title">🔑 Providers</div>
      ${providerRows}

      <div class="btn-row">
        <button class="btn btn-secondary" onclick="testAI()">🔌 Test Active Provider</button>
        <button class="btn btn-primary" onclick="saveAll()">💾 Save All Settings</button>
      </div>
      <div class="test-result" id="testResult"></div>

      <script>
        let selectedProvider = '${current}';

        function selectProvider(key) {
          document.querySelectorAll('.provider-card').forEach(c => {
            c.classList.remove('active');
            c.querySelector('.prov-radio').classList.remove('radio-on');
          });
          document.querySelectorAll('.prov-body').forEach(b => b.classList.add('hidden'));
          const card = document.getElementById('card-' + key);
          card.classList.add('active');
          card.querySelector('.prov-radio').classList.add('radio-on');
          document.getElementById('body-' + key).classList.remove('hidden');
          selectedProvider = key;
        }

        function collectData() {
          const data = { provider: selectedProvider };
          ${Object.entries(AI_PROVIDERS_CONFIG).map(([k]) => `
            data['key_${k}'] = (document.getElementById('key-${k}') || {value:''}).value;
            data['model_${k}'] = (document.getElementById('model-${k}') || {value:''}).value;
          `).join("")}
          const activeCustomKey = document.getElementById('customApiKey_' + selectedProvider);
          if (activeCustomKey) {
            data['customApiKey'] = activeCustomKey.value;
          } else {
            data['customApiKey'] = (document.getElementById('customApiKey_custom') || {value:''}).value || (document.getElementById('customApiKey_9router') || {value:''}).value;
          }
          return data;
        }

        function saveAll() {
          google.script.run.withSuccessHandler(() => {
            alert('✅ All AI settings saved!');
            google.script.host.close();
          }).saveUniversalAISettings(collectData());
        }

        function testAI() {
          const res = document.getElementById('testResult');
          res.style.display = 'block';
          res.className = 'test-result';
          res.textContent = '⏳ Testing connection...';
          google.script.run
            .withSuccessHandler(msg => {
              res.className = 'test-result test-ok';
              res.textContent = '✅ ' + msg;
            })
            .withFailureHandler(err => {
              res.className = 'test-result test-err';
              res.textContent = '❌ ' + err.message;
            })
            .testAIConnection(selectedProvider, collectData());
        }
      </script>
    </body>
    </html>
  `).setTitle("Universal AI Engine").setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── SAVE AI SETTINGS ─────────────────────────────────────────
function saveUniversalAISettings(data) {
  const props = _getScriptProps();
  const bulkProps = {};
  
  bulkProps["AI_PROVIDER"] = data.provider || "auto";

  Object.entries(AI_PROVIDERS_CONFIG).forEach(([key, cfg]) => {
    const keyVal   = data["key_" + key]   || "";
    const modelVal = data["model_" + key] || cfg.defaultModel;
    if (keyVal)   bulkProps[cfg.propKey]   = keyVal;
    if (modelVal) bulkProps[cfg.modelProp] = modelVal;
  });

  if (data.customApiKey) bulkProps["AI_CUSTOM_API_KEY"] = data.customApiKey;
  
  props.setProperties(bulkProps);
}

// ─── TEST AI CONNECTION ────────────────────────────────────────
function testAIConnection(providerName, formData) {
  if (providerName === "9router") {
    try {
      const baseUrl = (_getScriptProps().getProperty("AI_9ROUTER_URL") || "").replace(/\/$/, "").replace(/\s+/g, "");
      const apiKey = (_getScriptProps().getProperty("AI_CUSTOM_API_KEY") || "").replace(/\s+/g, "");
      const resp = UrlFetchApp.fetch(baseUrl + "/v1/models", {
        headers: {
          "Authorization": "Bearer " + apiKey,
          "Bypass-Tunnel-Reminder": "true",
          "ngrok-skip-browser-warning": "69420"
        },
        muteHttpExceptions: true
      });
      return "SUCCESS! 9router responded with: " + resp.getContentText().substring(0, 100) + "...";
    } catch (e) {
      throw new Error("Proof fetch failed: " + e.message);
    }
  }

  // Save settings first
  if (formData) saveUniversalAISettings(formData);

  const testPrompt = 'Reply with exactly this JSON: {"status":"ok","message":"AI connected successfully"}';
  try {
    const providerToTest = providerName === "auto" ? null : providerName;
    let result;
    if (providerToTest) {
      const cfg = AI_PROVIDERS_CONFIG[providerToTest];
      if (!cfg) throw new Error("Unknown provider: " + providerToTest);
      result = _callProvider(cfg, testPrompt, null);
    } else {
      result = _callAIAuto(testPrompt, null);
    }
    
    // Some models return noisy text (e.g. repeating the prompt). Extract the JSON block:
    let cleanJson = result.replace(/```json\n?|\n?```/g, "").trim();
    const match = cleanJson.match(/\{[\s\S]*\}/);
    if (match) {
      cleanJson = match[0];
    }

    const parsed = _safeParseJSON(cleanJson);
    const usedProvider = providerName === "auto" ? "(auto)" : AI_PROVIDERS_CONFIG[providerName]?.name;
    return `${parsed.message} — Provider: ${usedProvider}`;
  } catch (err) {
    throw new Error(err.message);
  }
}

// ─── SHOW ALL AVAILABLE PROVIDERS ────────────────────────────
function showProviderList() {
  const props = _getScriptProps();
  const lines = ["🤖 AVAILABLE AI PROVIDERS\n"];
  Object.entries(AI_PROVIDERS_CONFIG).forEach(([key, cfg]) => {
    const hasKey = !!(props.getProperty(cfg.propKey));
    lines.push(`${hasKey ? "✅" : "⬜"} ${cfg.name} (${key})`);
    if (cfg.note) lines.push(`    ↳ ${cfg.note}`);
    if (!hasKey && cfg.docs) lines.push(`    ↳ Get key: ${cfg.docs}`);
  });
  SpreadsheetApp.getUi().alert("AI Providers", lines.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─── BACKWARD COMPAT: old _callGemini still works ────────────
// Existing AGENT.gs code calls _callGemini() — keep it working
function _callGemini(prompt, apiKey) {
  const props = _getScriptProps();
  const key   = (apiKey || props.getProperty("AI_GEMINI_KEY") || props.getProperty("GEMINI_KEY") || "").replace(/\s+/g, "");
  if (!key) {
    // Try universal AI if no Gemini key
    return callAI(prompt, null);
  }
  const model = (props.getProperty("AI_GEMINI_MODEL") || "gemini-2.5-flash").replace(/\s+/g, "");
  const url   = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  const resp  = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { 
        temperature: 0.1, 
        maxOutputTokens: 1024,
        responseMimeType: "application/json"
      },
    }),
    muteHttpExceptions: true,
    });
  const json = _safeParseJSON(resp.getContentText());
  if (json.error) throw new Error(json.error.message);
  let fullText = "";
  if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts) {
    const parts = json.candidates[0].content.parts;
    const finalPart = parts.find(p => !p.thought) || parts[parts.length - 1];
    if (finalPart && finalPart.text) {
      fullText = finalPart.text;
    }
  }
  return fullText.trim();
}
function test9RouterConnection() { const url = "https://vast-webs-occur.loca.lt/v1/models"; const resp = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer 123456", "Bypass-Tunnel-Reminder": "true" } }); console.log("RESPONSE FROM GOOGLE SHEETS: " + resp.getContentText()); return resp.getContentText(); }
