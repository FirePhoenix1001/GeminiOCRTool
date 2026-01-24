import base64

def setup_gpt(client_content, model_name, rule_text):
    global client, model, rule
    client = client_content
    model = model_name
    rule = rule_text

def gpt_identify(pic_path):
    # base64 編碼圖片
    with open(pic_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "system",
                "content": rule
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": rule},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{img_b64}"
                        }
                    }
                ]
            }
        ]
    )


    return response.choices[0].message.content