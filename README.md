# LLM PPT Translator

## UI Preview

![image-20241126081838873](images/image-20241126081838873.png)



## Installation

```
conda create -n llm-ppt-translator python=3.11 -y
conda activate llm-ppt-translator
```

```
pip install -r requirements.txt
# pip list --format=freeze > requirements.txt
```

## Configration

```
cp .env.example .env
```

Modify the following parameters:

- For OCI GenAI
```
COMPARTMENT_ID=ocid1.compartment.oc1..******
CONFIG_PROFILE=DEFAULT
```

- For OpenAI API
```
OPENAI_API_KEY=sk-******
OPENAI_BASE_URL=http://xxx.xxx.xxx.xxx:8000/v1
OPENAI_MODEL_NAME=gpt-4
```

## Run

```
python main.py
```

## Access

Open [http://127.0.0.1:8080](http://127.0.0.1:8080)

## Limitation

- Not support SmartART

## Welcome to WeChat

![微信图片_20241126082444](images/微信图片_20241126082444.jpg)

## Translated Result Samples

![image-20241126081952034](images/image-20241126081952034.png)

![image-20241126082008343](images/image-20241126082008343.png)

![image-20241126082024391](images/image-20241126082024391.png)

![image-20241126082041272](images/image-20241126082041272.png)

![image-20241126082055254](images/image-20241126082055254.png)

![image-20241126082119970](images/image-20241126082119970.png)

![image-20241126082131561](images/image-20241126082131561.png)

![image-20241126082146464](images/image-20241126082146464.png)

![image-20241126082157871](images/image-20241126082157871.png)

![image-20241126083706873](images/image-20241126083706873.png)



![image-20241126082226906](images/image-20241126082226906.png)

![image-20241126083728314](images/image-20241126083728314.png)

