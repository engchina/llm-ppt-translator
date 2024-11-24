import os
import time

import gradio as gr
import oci
from dotenv import load_dotenv, find_dotenv
from openai import OpenAI
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE

# read local .env file
_ = load_dotenv(find_dotenv())

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"], base_url=os.environ["OPENAI_BASE_URL"])


def translate_text(text, target_lang, model_name="gpt-4"):
    max_attempts = 5  # 最大尝试次数

    if model_name == "gpt-4":
        for attempt in range(max_attempts):
            try:
                completion = client.chat.completions.create(
                    model=os.environ["OPENAI_MODEL_NAME"],
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant that translates text."},
                        {"role": "user",
                         "content": f"Translate to {target_lang}, maintain the original tone and style, output translation only: {text}"}
                    ]
                )
                print(f"{text=}")
                print(f"translated: {completion.choices[0].message.content}\n")
                return completion.choices[0].message.content
            except Exception as e:
                if attempt < max_attempts - 1:
                    print(f"Attempt {attempt + 1} failed with error: {str(e)}. Retrying...")
                    time.sleep(1)
                else:
                    print(f"All {max_attempts} attempts failed with error: {str(e)}")
                    return text
    else:
        for attempt in range(max_attempts):
            try:
                # Initialize OCI GenAI client
                config = oci.config.from_file('~/.oci/config', os.environ.get("CONFIG_PROFILE"))
                endpoint = "https://inference.generativeai.us-chicago-1.oci.oraclecloud.com"

                generative_ai_inference_client = oci.generative_ai_inference.GenerativeAiInferenceClient(
                    config=config,
                    service_endpoint=endpoint,
                    retry_strategy=oci.retry.NoneRetryStrategy(),
                    timeout=(10, 240)
                )

                # Prepare chat request
                chat_detail = oci.generative_ai_inference.models.ChatDetails()
                chat_request = oci.generative_ai_inference.models.CohereChatRequest()

                chat_request.preamble_override = "You are a helpful assistant that translates text."
                chat_request.message = f"Translate to {target_lang}, maintain the original tone and style, output translation only: {text}"

                # Set other parameters
                chat_request.max_tokens = 2000
                chat_request.temperature = 0
                chat_request.frequency_penalty = 0
                chat_request.top_p = 0.75
                chat_request.top_k = 0
                chat_request.is_stream = False

                chat_detail.serving_mode = oci.generative_ai_inference.models.OnDemandServingMode(
                    model_id=model_name
                )
                chat_detail.chat_request = chat_request
                chat_detail.compartment_id = os.environ.get("COMPARTMENT_ID")

                # Make the API call
                chat_response = generative_ai_inference_client.chat(chat_detail)

                print(f"{text=}")
                print(f"translated: {chat_response.data.chat_response.text}\n")
                return chat_response.data.chat_response.text
            except Exception as e:
                if attempt < max_attempts - 1:
                    print(f"Attempt {attempt + 1} failed with error: {str(e)}. Retrying...")
                    time.sleep(1)
                else:
                    print(f"All {max_attempts} attempts failed with error: {str(e)}")
                    return text


def translate_ppt(model_name, input_ppt, target_lang):
    # 读取输入PPT
    ppt = Presentation(input_ppt)

    # 遍历每一张幻灯片
    for slide_index, slide in enumerate(ppt.slides, start=1):
        print(f'Translate slide {slide_index}/{len(ppt.slides)}')
        print('-------------------------------------------')
        for shape in slide.shapes:
            # 跳过页脚部分
            if shape.is_placeholder and shape.placeholder_format.type in [
                PP_PLACEHOLDER_TYPE.FOOTER,
                PP_PLACEHOLDER_TYPE.SLIDE_NUMBER,
                PP_PLACEHOLDER_TYPE.DATE,
            ]:
                continue

            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        original_text = cell.text_frame
                        if original_text and original_text.strip():
                            translated_text = translate_text(original_text, target_lang, model_name)
                            cell.text = translated_text
            elif shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        original_text = run.text
                        if original_text and original_text.strip():
                            translated_text = translate_text(original_text, target_lang, model_name)
                            run.text = translated_text

        if slide.has_notes_slide:
            original_text = slide.notes_slide.notes_text_frame.text
            if original_text and original_text.strip():
                translated_text = translate_text(original_text, target_lang, model_name)
                slide.notes_slide.notes_text_frame.text = translated_text

    # 保存翻译后的PPT
    output_ppt = f"{input_ppt.name.split('.')[0]}_{target_lang}.{input_ppt.name.split('.')[1]}"
    ppt.save(os.path.join("/tmp", output_ppt))
    print("Translation completed.")
    return output_ppt


model_type = gr.Radio(
    choices=[
        "cohere.command-r-08-2024",
        "cohere.command-r-plus-08-2024",
        "gpt-4"
    ],
    label="Model Type",
    value="gpt-4"
)
input_ppt = gr.File(label="Upload PPTX file", file_types=[".pptx"])
target_lang = gr.Dropdown(choices=["English", "Japanese", "Chinese"], label="Target Language", value="Japanese")
output_ppt = gr.File(label="Download Translated PPTX")

gr.Interface(
    fn=translate_ppt,
    inputs=[model_type, input_ppt, target_lang],
    outputs=output_ppt,
    title="PPT Translator",
    description="Upload a PPTX file and specify target language to get a translated PPTX file."
).launch(server_port=8000)
