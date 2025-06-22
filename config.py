from langchain_openai import ChatOpenAI

api_key = "sk-tavkanqmfymbpcncmjbkpskoqnzscwolinkoptpbigopbbdj"

model = ChatOpenAI(
    model_name="https://deepseek.deepseek.cn/api/v1",
    temperature=0.5,
    openai_api_key=api_key
)