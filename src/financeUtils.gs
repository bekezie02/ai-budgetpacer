const plaidClientId = properties.getProperty('PLAID_CLIENT_ID')
const plaidSecret = properties.getProperty('PLAID_SECRET')
const plaidURL = properties.getProperty('PLAID_URL')
const MY_EMAIL = Session.getEffectiveUser().getEmail();

// =========================================================================
// CORE PLAID OPERATIONS
// =========================================================================
function runStep1() {
  try {
    let user = properties.getProperty('PLAID_USER');
    if (!user) {
      user = Utilities.getUuid();
      properties.setProperty('PLAID_USER', user);
    }
    const payload = {
      "client_id": plaidClientId,
      "secret": plaidSecret,
      "client_name": "AI Personal Finance Tracker",
      "user": { "client_user_id": user },
      "products": ["transactions"],
      "country_codes": ["US"],
      "language": "en",
      "transactions": {
      "days_requested": 243
    },
      "hosted_link": {}
    };

    const response = UrlFetchApp.fetch(`${plaidURL}/link/token/create`, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200) {
      CacheService.getScriptCache().put('CURRENT_LINK_TOKEN', data.link_token, 1500);
      return data.hosted_link_url;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function runStep2() {
  try {
    const linkToken = CacheService.getScriptCache().get('CURRENT_LINK_TOKEN');
    if (!linkToken) return { status: "error", message: "Token expired out of cache memory." };

    const getResponse = UrlFetchApp.fetch(`${plaidURL}/link/token/get`, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify({ "client_id": plaidClientId, "secret": plaidSecret, "link_token": linkToken }),
      "muteHttpExceptions": true
    });

    const getData = JSON.parse(getResponse.getContentText());
    if (getResponse.getResponseCode() !== 200) return { status: "error", message: getResponse.getContentText() };
    if (!getData.link_sessions || getData.link_sessions.length === 0) return { status: "error", message: "No session array history." };

    const latestSession = getData.link_sessions[0];

    if (!latestSession.results || !latestSession.results.item_add_results || latestSession.results.item_add_results.length === 0) {
      return { status: "error", message: "Login window closed or user did not finish bank authorization in time." };
    }

    const publicToken = latestSession.results.item_add_results[0].public_token;

    // FIXED: was referencing an undefined BASE_URL — this is the same Plaid
    // base URL used everywhere else (plaidURL).
    const exchangeResponse = UrlFetchApp.fetch(`${plaidURL}/item/public_token/exchange`, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify({ "client_id": plaidClientId, "secret": plaidSecret, "public_token": publicToken }),
      "muteHttpExceptions": true
    });

    const exchangeData = JSON.parse(exchangeResponse.getContentText());

    if (exchangeResponse.getResponseCode() === 200) {
      let itemIds = properties.getProperty('PLAID_ITEM_IDS');
      if (!itemIds) itemIds = "[]";
      itemIds = JSON.parse(itemIds);
      if (!itemIds.includes(exchangeData.item_id)) itemIds.push(exchangeData.item_id);
      properties.setProperty('PLAID_ITEM_IDS', JSON.stringify(itemIds));
      // FIXED: only set the access token for the item this exchange was
      // actually for — the old loop assigned this run's new token to every
      // item missing a stored token, which breaks as soon as more than one
      // item has ever been linked.
      properties.setProperty(`PLAID_ACCESSS_TOKEN_${exchangeData.item_id}`, exchangeData.access_token);

      return { status: "success" };
    } else {
      return { status: "error", message: `Plaid exchange failed: ${exchangeResponse.getContentText()}` };
    }
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function getAllRequest(requests) {
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    if (responses.length <= 0) {
      return { success: false, error: "empty list" };
    }

    return responses.map(res => {
      const code = res.getResponseCode();
      const text = res.getContentText();
      if (code === 200) {
        return { success: true, code: code, data: JSON.parse(text) };
      }
      return { success: false, code: code, error: text };
    });
  } catch (e) {
    console.log("Fetch error: " + e.message);
    return [{ success: false, error: e.message }];
  }
}

function fetchNextSyncPage(accessToken, cursor) {
  const payload = {
    "client_id": plaidClientId,
    // FIXED: was referencing an undefined global PLAID_SECRET — the actual
    // const defined above is plaidSecret.
    "secret": plaidSecret,
    "access_token": accessToken,
    "cursor": cursor,
    // Smaller than Plaid's default (100) so each page finishes — and its
    // cursor gets saved — well within Apps Script's execution time limit,
    // even during a large backfill. More, smaller pages instead of fewer,
    // bigger ones that risk never finishing.
    "count": 50
  };

  const requestOptions = {
    "url": `${plaidURL}/transactions/sync`,
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const results = getAllRequest([requestOptions]);
  return results[0];
}

// NOTE: getTransactions()/processTransactions() from the original draft are
// superseded by runDailyFinanceSync() in financeCore.gs, which handles
// added/modified/removed, categorization, ledgering, and fixed-pool
// subtraction. They're removed here to avoid two competing sync paths.

/**
 * Sends a prompt to the LLM and returns its raw text response in `data`.
 * Callers are responsible for JSON.parse'ing `data` if they asked the model
 * for structured output.
 *
 * options (all optional, all default to prior behavior so nothing that
 * already calls llm(prompt) is affected):
 *   - maxCompletionTokens: override the 1200-token completion budget.
 *   - reasoningEffort: "low" | "medium" | "high" — passed through to Groq
 *     as reasoning_effort. openai/gpt-oss-120b is a reasoning model that
 *     spends completion tokens on hidden chain-of-thought before writing
 *     its visible answer, and those reasoning tokens count against
 *     max_completion_tokens even though you never see them — at the
 *     default "medium" effort, a short, simple task (e.g. writing a
 *     150-250 word summary) can burn most of a 1200-token budget on
 *     invisible thinking and get cut off mid-answer (finish_reason:
 *     "length") before ever finishing the actual response. Passing "low"
 *     for tasks that don't need deep reasoning fixes that at the source
 *     instead of just raising the ceiling.
 */
function llm(prompt, options) {
  options = options || {};
  const inputTokens = Math.ceil(prompt.length / 4);
  console.log("Prompt chars:", prompt.length, "~tokens:", inputTokens);

  // Fixed completion budget. The old version derived this as
  // (8000 - inputTokens - 250), which meant every extra observation added to
  // the prompt ate directly into the model's room to respond. Once the
  // prompt grew large (e.g. after webScraper results were embedded into the
  // observation history), this collapsed to its 100-token floor — nowhere
  // near enough to emit a full task JSON block, causing the model to fall
  // back to degenerate output like {"nextAction":{"tool":null}}.
  // TPM ("tokens per minute") is a rate limit on your Groq account, not a
  // per-request context budget, so it shouldn't be used to shrink this.
  const MAX_COMPLETION_TOKENS = options.maxCompletionTokens || 1200;

  const groqKey = properties.getProperty("GROQ_API_KEY");
  const groqUrl = properties.getProperty("GROQ_URL");

  const payload = {
    model: "openai/gpt-oss-20b",
    messages: [{ role: "user", content: prompt }],
    tool_choice: "none",
    temperature: 0,
    max_completion_tokens: MAX_COMPLETION_TOKENS,
  };
  if (options.reasoningEffort) {
    payload.reasoning_effort = options.reasoningEffort;
  }

  const options_ = {
    method: "post",
    headers: { Authorization: `Bearer ${groqKey}` },
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const MAX_RETRIES = 5;

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      console.log("Calling LLM...");
      // Only throttle on retries — no need to pre-emptively delay the first attempt.
      if (attempt > 0) Utilities.sleep(5000);

      const response = UrlFetchApp.fetch(groqUrl, options_);
      const status = response.getResponseCode();
      const text = response.getContentText();

      const headers = response.getHeaders();
      const remainingTokens = headers["x-ratelimit-remaining-tokens"];
      const resetTokens = headers["x-ratelimit-reset-tokens"];
      if (remainingTokens !== undefined || resetTokens !== undefined) {
        console.log(`Groq rate-limit — remaining tokens: ${remainingTokens}, resets in: ${resetTokens}`);
      }

      if (status === 429) {
        const delay = Math.min(15000 * Math.pow(2, attempt), 60000);
        console.warn(`Rate limited. Retry ${attempt + 1}/${MAX_RETRIES} in ${delay / 1000}s`);
        Utilities.sleep(delay);
        continue;
      }

      const data = JSON.parse(text);
      console.log("Parsed JSON");

      if (status >= 400 || data.error) {
        return {
          success: false,
          error: data?.error?.message || text
        };
      }

      const finishReason = data?.choices?.[0]?.finish_reason;
      console.log("finish_reason:", finishReason);
      if (finishReason && finishReason !== "stop") {
        // "length" = hit max_completion_tokens (or a live account-level TPM
        // ceiling below what we estimated); anything else (e.g.
        // "content_filter") is a different problem entirely — either way,
        // this is worth knowing explicitly rather than only inferring it
        // from a downstream JSON.parse failure.
        console.warn(`llm(): response did not finish normally (finish_reason="${finishReason}") — output may be truncated.`);
      }

      return {
        success: true,
        data: data.choices[0].message.content
      };
    } catch (e) {
      return {
        success: false,
        error: "Network error: " + e.toString()
      };
    }
  }

  return {
    success: false,
    error: "Rate limit exceeded after maximum retries."
  };
}
