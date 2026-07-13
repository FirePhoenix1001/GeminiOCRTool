import { DEFAULT_MODEL } from './config.js';

let logMsg = console.log;

export function setLogger(logger) {
    if (typeof logger === 'function') {
        logMsg = logger;
    }
}

// Parse keys for rotation
function getAvailableKeys() {
    const keys = [];
    const savedKeysStr = localStorage.getItem('geminiocr-api-keys');
    if (savedKeysStr) {
        try {
            const list = JSON.parse(savedKeysStr);
            list.forEach((item, index) => {
                const k = item.key ? item.key.trim() : '';
                if (!k) return;
                const name = item.type === 'openai' ? `OpenAI Key ${index + 1}` : `Gemini Key ${index + 1}`;
                keys.push({
                    name: name,
                    key: k,
                    type: item.type
                });
            });
        } catch (e) {
            console.error('Failed to parse API keys', e);
        }
    }
    
    // Fallback/Backward compatibility
    if (keys.length === 0) {
        const primaryGemini = localStorage.getItem('gemini-key') || '';
        const primaryOpenai = localStorage.getItem('openai-key') || '';
        if (primaryGemini) {
            keys.push({ name: 'Gemini Primary', key: primaryGemini, type: 'gemini' });
        }
        if (primaryOpenai) {
            keys.push({ name: 'OpenAI Primary', key: primaryOpenai, type: 'openai' });
        }
    }

    return keys;
}

// Sleep helper for backoff
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Page Visibility guard: pauses execution when the browser tab/display is inactive.
// This prevents API calls from being made while the system is in sleep/display-off mode,
// which would cause "Failed to fetch" or undefined errors.
async function waitForPageVisible() {
    if (typeof document === 'undefined') return; // SSR/Node guard
    if (document.visibilityState === 'visible') return;

    logMsg('💤 偵測到頁面處於背景/螢幕關閉狀態，暫停處理中... (恢復後將自動繼續)');
    
    await new Promise(resolve => {
        const onVisible = () => {
            if (document.visibilityState === 'visible') {
                document.removeEventListener('visibilitychange', onVisible);
                resolve();
            }
        };
        document.addEventListener('visibilitychange', onVisible);
    });

    // After waking up, wait a moment for the network stack to fully recover
    logMsg('☀️ 頁面已恢復前景，等待 3 秒讓網路連線穩定後繼續...');
    await sleep(3000);
}

let currentKeyIndex = 0;

// Call API with Key Rotation and Exponential Backoff Retry
export async function callAIApi(promptRule, inputType, base64Image = null) {
    const keys = getAvailableKeys();
    if (keys.length === 0) {
        throw new Error('未設定任何 API Key，請前往「金鑰與模型設定」進行設定。');
    }

    const modelName = localStorage.getItem('model-name') || DEFAULT_MODEL;
    const customModel = localStorage.getItem('custom-model') || '';
    const finalModel = modelName === 'custom' ? customModel : modelName;

    let attempts = 0;
    const maxAttempts = keys.length;

    while (attempts < maxAttempts) {
        // Ensure index is within range in case keys list changed
        if (currentKeyIndex >= keys.length) {
            currentKeyIndex = 0;
        }
        const currentKeyInfo = keys[currentKeyIndex];
        logMsg(`🔑 目前調用金鑰: 【${currentKeyInfo.name}】(${currentKeyInfo.type === 'openai' ? 'OpenAI' : 'Gemini'})`);

        let retryCount = 0;
        const maxRetries = 3; // Retry up to 3 times (4 total attempts)
        const baseDelay = 2000; // Base delay: 2 seconds

        let success = false;
        let result = null;
        let lastError = null;

        while (retryCount <= maxRetries) {
            try {
                // Guard: wait if page is hidden (display off / system sleep)
                await waitForPageVisible();

                if (currentKeyInfo.type === 'openai' || finalModel.startsWith('gpt') || finalModel.startsWith('o1')) {
                    result = await fetchOpenai(currentKeyInfo.key, finalModel, promptRule, inputType, base64Image);
                } else {
                    result = await fetchGemini(currentKeyInfo.key, finalModel, promptRule, inputType, base64Image);
                }
                success = true;
                break;
            } catch (err) {
                lastError = err;
                const errStr = (err.message || '').toLowerCase();
                const isHighDemand = errStr.includes('high demand') || errStr.includes('spikes in demand');
                
                // Detect network-level transient errors (browser fetch failures)
                const isNetworkError = errStr.includes('failed to fetch') ||
                                       errStr.includes('networkerror') ||
                                       errStr.includes('network') ||
                                       errStr.includes('timeout') ||
                                       errStr.includes('aborted') ||
                                       errStr.includes('econnreset') ||
                                       errStr.includes('cors') ||
                                       errStr === '' || errStr === 'undefined';

                // Detect retryable errors (Busy servers, rate limits, overloads)
                const isRetryable = isNetworkError ||
                                    errStr.includes('429') || 
                                    errStr.includes('resource_exhausted') || 
                                    errStr.includes('demand') || 
                                    errStr.includes('overloaded') || 
                                    errStr.includes('try again') || 
                                    errStr.includes('temporary') || 
                                    errStr.includes('busy') ||
                                    errStr.includes('limit') ||
                                    errStr.includes('quota');

                if (isRetryable) {
                    // Check if message specifies a retry duration (e.g. Please retry in 21.255346312s or 49.936108ms)
                    const retryMatch = /Please retry in (\d+(?:\.\d+)?)(ms|s)/i.exec(err.message);
                    if (retryMatch) {
                        const val = parseFloat(retryMatch[1]);
                        const unit = retryMatch[2].toLowerCase();
                        const parsedSeconds = unit === 'ms' ? val / 1000 : val;
                        const waitTime = parsedSeconds + 5;

                        if (waitTime > 20 && keys.length > 1) {
                            logMsg(`⚠️ [限流重試] 偵測到超額限制，預估需等待 ${waitTime.toFixed(1)} 秒 (超過 20 秒)。直接切換 API 金鑰...`);
                            break; // Break the retry loop to trigger key rotation
                        } else {
                            if (retryCount < maxRetries) {
                                retryCount++;
                                logMsg(`⚠️ [限流重試] 偵測到超額限制，預估需等待 ${waitTime.toFixed(1)} 秒 (小於等於 20 秒)。`);
                                logMsg(`⏱️ 依指示等待 ${waitTime.toFixed(1)} 秒後進行第 ${retryCount} 次重試...`);
                                await sleep(waitTime * 1000);
                                continue;
                            } else {
                                break;
                            }
                        }
                    }

                    // High demand or other retryable errors
                    const currentMaxRetries = isHighDemand ? 5 : (isNetworkError ? 4 : maxRetries);
                    if (retryCount < currentMaxRetries) {
                        retryCount++;
                        const delay = isNetworkError
                            ? (baseDelay * Math.pow(2, retryCount - 1) + Math.random() * 2000) // Longer jitter for network errors
                            : (baseDelay * Math.pow(2, retryCount - 1) + Math.random() * 1000);
                        logMsg(`⚠️ [${isNetworkError ? '網路重試' : '限流重試'}] ${isNetworkError ? '網路連線異常' : '伺服器忙碌或限制'}: ${err.message || '(未知錯誤)'}`);
                        logMsg(`⏱️ 啟用指數型退避等待，將於 ${(delay / 1000).toFixed(1)} 秒後進行第 ${retryCount} 次重試...`);
                        await sleep(delay);
                    } else {
                        break;
                    }
                } else {
                    break;
                }
            }
        }

        if (success) {
            return result;
        } else {
            const errStr = (lastError.message || '').toLowerCase();
            const isNetworkErr = errStr.includes('failed to fetch') ||
                                 errStr.includes('networkerror') ||
                                 errStr.includes('network') ||
                                 errStr.includes('timeout') ||
                                 errStr === '' || errStr === 'undefined';
            const triggersRotation = isNetworkErr ||
                                     errStr.includes('429') || 
                                     errStr.includes('resource_exhausted') || 
                                     errStr.includes('limit') || 
                                     errStr.includes('key') || 
                                     errStr.includes('quota') ||
                                     errStr.includes('demand') ||
                                     errStr.includes('overloaded') ||
                                     errStr.includes('temporary') ||
                                     errStr.includes('busy') ||
                                     errStr.includes('try again');
            
            if (triggersRotation && keys.length > 1) {
                logMsg(`⚠️ 金鑰【${currentKeyInfo.name}】${isNetworkErr ? '網路連線異常' : '連線超額'}且重試失敗: ${lastError.message || '(未知錯誤)'}`, 'error');
                currentKeyIndex = (currentKeyIndex + 1) % keys.length;
                logMsg(`🔄 嘗試切換至下一組金鑰...`);
                attempts++;
                continue;
            } else {
                throw lastError; // Re-throw fatal error
            }
        }
    }

    // All keys exhausted — perform a final cooldown retry cycle
    logMsg(`⚠️ 所有金鑰均已嘗試失敗，啟動冷卻期後最終重試...`, 'warning');
    logMsg(`⏱️ 等待 30 秒冷卻後重新嘗試所有金鑰...`);
    await sleep(30000);

    // One final attempt cycling through all keys (no further cooldown)
    for (let i = 0; i < keys.length; i++) {
        const keyIdx = (currentKeyIndex + i) % keys.length;
        const keyInfo = keys[keyIdx];
        logMsg(`🔑 [最終重試] 嘗試金鑰: 【${keyInfo.name}】`);
        try {
            await waitForPageVisible(); // Guard against display-off
            let finalResult;
            if (keyInfo.type === 'openai' || finalModel.startsWith('gpt') || finalModel.startsWith('o1')) {
                finalResult = await fetchOpenai(keyInfo.key, finalModel, promptRule, inputType, base64Image);
            } else {
                finalResult = await fetchGemini(keyInfo.key, finalModel, promptRule, inputType, base64Image);
            }
            currentKeyIndex = keyIdx; // Update to the successful key
            logMsg(`✅ [最終重試] 金鑰【${keyInfo.name}】成功！`);
            return finalResult;
        } catch (finalErr) {
            logMsg(`⚠️ [最終重試] 金鑰【${keyInfo.name}】失敗: ${finalErr.message || '(未知錯誤)'}`);
        }
    }

    throw new Error('所有設定的金鑰已耗盡或無法成功連線 API。');
}

// Fetch Google Gemini
async function fetchGemini(key, model, promptRule, inputType, base64Image) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
    
    let contents = [];
    if (inputType === 'text') {
        contents = [
            {
                parts: [{ text: promptRule }]
            }
        ];
    } else {
        contents = [
            {
                parts: [
                    {
                        inlineData: {
                            mimeType: 'image/png',
                            data: base64Image
                        }
                    },
                    {
                        text: promptRule
                    }
                ]
            }
        ];
    }

    const res = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ contents: contents })
    });

    if (!res.ok) {
        const errBody = await res.json().catch(() => ({}));
        const errMsg = errBody.error?.message || `HTTP ${res.status}`;
        throw new Error(errMsg);
    }

    const data = await res.json();
    
    try {
        if (data.candidates && data.candidates[0].content.parts) {
            return data.candidates[0].content.parts.map(p => p.text).join('');
        }
    } catch (e) {
        console.error('Gemini text extraction failed', e);
    }

    throw new Error('無法從 Gemini API 回傳內容中解析文本結構。');
}

// Fetch OpenAI GPT
async function fetchOpenai(key, model, promptRule, inputType, base64Image) {
    const url = 'https://api.openai.com/v1/chat/completions';
    
    let messages = [];
    if (inputType === 'text') {
        messages = [
            { role: 'user', content: promptRule }
        ];
    } else {
        messages = [
            { role: 'system', content: promptRule },
            {
                role: 'user',
                content: [
                    { type: 'text', text: promptRule },
                    {
                        type: 'image_url',
                        image_url: {
                            url: `data:image/png;base64,${base64Image}`
                        }
                    }
                ]
            }
        ];
    }

    const res = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${key}`
        },
        body: JSON.stringify({
            model: model,
            messages: messages
        })
    });

    if (!res.ok) {
        const errBody = await res.json().catch(() => ({}));
        const errMsg = errBody.error?.message || `HTTP ${res.status}`;
        throw new Error(errMsg);
    }

    const data = await res.json();
    return data.choices[0].message.content;
}
