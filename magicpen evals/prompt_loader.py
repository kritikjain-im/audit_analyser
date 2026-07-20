import re
from config import PROMPT_TEMPLATE_PATH
from utils import normalize_value

def load_template(path: str = PROMPT_TEMPLATE_PATH) -> tuple:
    """
    Loads and splits the prompt template into (system_prompt_template, user_prompt_template).
    """
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
        
    parts = content.split("User Prompt:")
    system_part = parts[0].replace("System Prompt:", "").strip()
    user_part = parts[1].strip() if len(parts) > 1 else ""
    return system_part, user_part

def extract_seller_info(user_details_str: str) -> tuple:
    """
    Extracts name and company from the seller's user details.
    """
    name_match = re.search(r"\bName\s*=\s*([^|#\-\n\r\t]+?)(?=\b\w+\s*=|$)", user_details_str)
    company_match = re.search(r"\bCompany\s*=\s*([^|#\-\n\r\t]+?)(?=\b\w+\s*=|$)", user_details_str)
    
    seller_name = name_match.group(1).strip() if name_match else "Aditya"
    seller_company = company_match.group(1).strip() if company_match else "IndiaMART"
    return seller_name, seller_company

def extract_buyer_name(ongoing_conv_str: str) -> str:
    """
    Extracts the buyer name from the ongoing conversation context.
    """
    match = re.search(r"buyer named\s+([A-Za-z\s]+?)(?=\s+from|$)", ongoing_conv_str)
    if match:
        return match.group(1).strip()
    return "Buyer"

def format_prompts(system_template: str, user_template: str, row: dict) -> tuple:
    """
    Extracts row variables and constructs both System and User prompts.
    
    Returns:
        tuple: (system_prompt, user_prompt)
    """
    user_details = normalize_value(row.get("User Details"))
    ongoing_conversation = normalize_value(row.get("Ongoing Conversation"))
    product_details = normalize_value(row.get("Product Details"))
    past_conversation = normalize_value(row.get("Past Conversation"))
    
    seller_name, seller_company = extract_seller_info(user_details)
    buyer_name = extract_buyer_name(ongoing_conversation)
    
    # Format system prompt
    system_prompt = system_template.replace("{{seller_name}}", seller_name)
    system_prompt = system_prompt.replace("{{seller_company}}", seller_company)
    
    # Format user prompt
    user_prompt = user_template
    user_prompt = user_prompt.replace("{{seller_detail}}", user_details)
    user_prompt = user_prompt.replace("{{prdDetail}}", product_details)
    user_prompt = user_prompt.replace("{{convintronew}}", past_conversation)
    user_prompt = user_prompt.replace("{{origConvDetail}}", ongoing_conversation)
    user_prompt = user_prompt.replace("{{seller_name}}", seller_name)
    user_prompt = user_prompt.replace("{{buyer_name}}", buyer_name)
    
    return system_prompt, user_prompt
