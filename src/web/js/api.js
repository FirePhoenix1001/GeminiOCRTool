// api.js - GeminiOCR API Communication Module

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

let currentKeyIndex = 0;

// Call API with Key Rotation and Exponential Backoff Retry
export async function callAIApi(promptRule, inputType, base64Image = null) {
    const keys = getAvailableKeys();
    if (keys.length === 0) {
        throw new Error('未設定任何 API Key，請前往「金鑰與模型設定」進行設定。');
    }

    const modelName = localStorage.getItem('model-name') || 'gemini-3-flash-preview';
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
                // Detect retryable errors (Busy servers, rate limits, overloads)
                const isRetryable = errStr.includes('429') || 
                                    errStr.includes('resource_exhausted') || 
                                    errStr.includes('demand') || 
                                    errStr.includes('overloaded') || 
                                    errStr.includes('try again') || 
                                    errStr.includes('temporary') || 
                                    errStr.includes('busy') ||
                                    errStr.includes('limit');

                if (isRetryable && retryCount < maxRetries) {
                    retryCount++;
                    const delay = baseDelay * Math.pow(2, retryCount - 1) + Math.random() * 1000;
                    logMsg(`⚠️ [限流重試] 伺服器忙碌或限制: ${err.message}`);
                    logMsg(`⏱️ 啟用指數型退避等待，將於 ${(delay / 1000).toFixed(1)} 秒後進行第 ${retryCount} 次重試...`);
                    await sleep(delay);
                } else {
                    break;
                }
            }
        }

        if (success) {
            return result;
        } else {
            const errStr = (lastError.message || '').toLowerCase();
            const triggersRotation = errStr.includes('429') || errStr.includes('resource_exhausted') || errStr.includes('limit') || errStr.includes('key') || errStr.includes('quota');
            
            if (triggersRotation && keys.length > 1) {
                logMsg(`⚠️ 金鑰【${currentKeyInfo.name}】連線超額且重試失敗: ${lastError.message}`, 'error');
                currentKeyIndex = (currentKeyIndex + 1) % keys.length;
                logMsg(`🔄 嘗試切換至下一組金鑰...`);
                attempts++;
                continue;
            } else {
                throw lastError; // Re-throw fatal error
            }
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
