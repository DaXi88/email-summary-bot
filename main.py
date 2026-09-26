# main.py

# ==============================================================================
# 导入必要的库
# ==============================================================================
import os
import imaplib
import email
import re
from email.header import make_header, decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta, timezone
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import openai
import json
import markdown2
import time

# ==============================================================================
# 全局常量与配置
# ==============================================================================

SYSTEM_PROMPT = """
# 角色
你是一名专业的邮件分析助手，任务是根据下方提供的邮件JSON数据，生成一段Markdown格式的摘要报告。

# 任务指令
1. 仔细阅读提供的邮件JSON数据。
2. **不要**生成顶层标题或总览信息，只专注于处理邮件列表。
3. 按顺序逐一处理`{{emails}}`数组中的每一封邮件，并为每封邮件提取以下信息：
   - **发件人**：提取`from_sender`。
   - **主题**：提取`subject`。
   - **摘要**：根据`body_preview`概括核心内容。
   - **关键行动点**：识别具体任务，或填写"无"。
4. 严格按照“输出格式要求”生成内容。邮件序号**必须从 {{start_index}} 开始**。

# 输出格式要求
#### 邮件 {{start_index}}：[第一封邮件的主题]
- **发件人**：[发件人信息]
- **摘要**：[简洁概括]
- **行动点**：[具体行动或"无"]

---

#### 邮件 {{start_index + 1}}：[第二封邮件的主题]
- **发件人**：[发件人信息]
- **摘要**：[简洁概括]
- **行动点**：[具体行动或"无"]

---
... (以此类推，处理完批次内的所有邮件)

# 特别说明
- 社团邮件标记：如果邮件内容是关于社团活动，请在主题末尾添加 `[社团邮件]`。
- 志愿者招募标记：如果邮件内容是关于志愿者招募（volunteer recruitment），请在主题末尾添加 `[志愿者招募]`。

# 待分析的邮件数据
{{emails}}
"""

# 从GitHub Secrets安全地加载环境变量
IMAP_EMAIL = os.environ.get("IMAP_EMAIL")
IMAP_AUTH_CODE = os.environ.get("IMAP_AUTH_CODE")
IMAP_SERVER = os.environ.get("IMAP_SERVER")
TARGET_FOLDER = os.environ.get("TARGET_FOLDER")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY") # 这里不用改，因为Secrets名就是这个
BASE_URL = os.environ.get("BASE_URL", "https://api.deepseek.com/v1") # 新增：读取你的Agens地址
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
SENDER_AUTH_CODE = os.environ.get("SENDER_AUTH_CODE")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")
SMTP_SERVER = os.environ.get("SMTP_SERVER")
SMTP_PORT = int(os.environ.get("SMTP_PORT", 465))
LLM_MODEL = os.environ.get("LLM_MODEL", "agnes-2.5-flash") # 修改：默认模型换成你的模型
LLM_TEMPERATURE = float(os.environ.get("LLM_TEMPERATURE", 0))
LLM_MAX_TOKENS = int(os.environ.get("LLM_MAX_TOKENS", 4096))

# ==============================================================================
# 核心功能函数
# ==============================================================================

def _decode_part_payload(part):
    payload = part.get_payload(decode=True)
    if payload is None:
        return ""
    charsets = [part.get_content_charset(), "utf-8", "gb18030"]
    for charset in charsets:
        if not charset:
            continue
        try:
            return payload.decode(charset)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="ignore")

def _strip_html_tags(html_text):
    if not html_text:
        return ""
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html_text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def _extract_body_preview(msg, max_len=1500):
    plain_candidates = []
    html_candidates = []
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        content_disposition = (part.get("Content-Disposition") or "").lower()
        if "attachment" in content_disposition:
            continue
        content_type = part.get_content_type()
        if content_type not in ("text/plain", "text/html"):
            continue
        decoded_text = _decode_part_payload(part)
        if not decoded_text:
            continue
        if content_type == "text/plain":
            plain_candidates.append(decoded_text)
        elif content_type == "text/html":
            html_candidates.append(decoded_text)

    if plain_candidates:
        body = "\n".join(plain_candidates)
    elif html_candidates:
        body = _strip_html_tags("\n".join(html_candidates))
    else:
        body = ""

    body = re.sub(r"\s+", " ", body).strip()
    return body[:max_len]

def get_emails_from_target_date(target_date):
    mail_list = []
    beijing_tz = timezone(timedelta(hours=8))
    try:
        conn = imaplib.IMAP4_SSL(IMAP_SERVER)
        # 去除授权码中可能存在的空格
        clean_auth_code = IMAP_AUTH_CODE.replace(" ", "")
        conn.login(IMAP_EMAIL, clean_auth_code)
        conn.select(f'"{TARGET_FOLDER}"')
        
        fetch_since_dt = target_date - timedelta(days=2)
        fetch_since_str = fetch_since_dt.strftime("%d-%b-%Y")
        search_query = f'(SINCE "{fetch_since_str}")'
        
        status, messages = conn.search(None, search_query)
        if status != "OK":
            print(f"IMAP search failed for query: {search_query}")
            conn.logout()
            return []
            
        email_ids = messages[0].split()
        
        for email_id in reversed(email_ids):
            _, msg_data = conn.fetch(email_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            try:
                date_header = msg.get("Date")
                if not date_header: continue
                
                email_dt_original = parsedate_to_datetime(date_header)
                if email_dt_original.tzinfo is None:
                    email_dt_in_beijing = email_dt_original.replace(tzinfo=timezone.utc).astimezone(beijing_tz)
                else:
                    email_dt_in_beijing = email_dt_original.astimezone(beijing_tz)

                if email_dt_in_beijing.date() != target_date.date():
                    continue

                subject = str(make_header(decode_header(msg.get("Subject", ""))))
                if not subject.strip():
                    subject = "(无主题)"

                from_ = str(make_header(decode_header(msg.get("From", ""))))
                if not from_.strip():
                    from_ = "(未知发件人)"

                body_preview = _extract_body_preview(msg)
                mail_list.append({ "from_sender": from_, "subject": subject, "body_preview": body_preview })
            except Exception as e:
                print(f"解析邮件 {email_id.decode()} 时出错: {e}")
                continue
                
        conn.logout()
        print(f"成功从'{TARGET_FOLDER}'文件夹获取并过滤出 {len(mail_list)} 封邮件。")
        return mail_list
    except Exception as e:
        print(f"获取邮件失败: {e}")
        return []

def _extract_status_code(exception):
    status_code = getattr(exception, "status_code", None)
    if status_code is not None:
        return status_code
    response = getattr(exception, "response", None)
    if response is not None:
        return getattr(response, "status_code", None)
    return None

def _is_retryable_exception(exception):
    status_code = _extract_status_code(exception)
    if status_code == 429 or (status_code is not None and 500 <= status_code < 600):
        return True
    exception_text = str(exception).lower()
    timeout_signals = ["timeout", "timed out", "readtimeout", "connecttimeout", "网络超时"]
    return any(signal in exception_text for signal in timeout_signals)

def _build_batch_error_block(start_index, batch_size, exception):
    batch_end = start_index + batch_size - 1
    exception_type = type(exception).__name__
    status_code = _extract_status_code(exception)
    structured_error = {
        "batch_range": f"{start_index}-{batch_end}",
        "exception_type": exception_type,
        "status_code": status_code,
        "message": str(exception),
    }
    return (
        "---\n\n"
        f"#### 处理邮件 {start_index} 到 {batch_end} 时失败\n"
        "```json\n"
        f"{json.dumps(structured_error, ensure_ascii=False, indent=2)}\n"
        "```\n\n"
        "---"
    )

def summarize_single_batch(client, email_batch, start_index, max_retries=3, base_delay=2):
    emails_json_str = json.dumps(email_batch, ensure_ascii=False, indent=2)
    prompt_filled = SYSTEM_PROMPT.replace("{{emails}}", emails_json_str)
    prompt_filled = prompt_filled.replace("{{start_index}}", str(start_index))
    
    for attempt in range(max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=LLM_MODEL,
                temperature=LLM_TEMPERATURE,
                max_tokens=LLM_MAX_TOKENS,
                messages=[{"role": "system", "content": prompt_filled}]
            )
            return {
                "success": True,
                "content": response.choices[0].message.content,
                "error": None,
            }
        except Exception as e:
            should_retry = attempt < max_retries and _is_retryable_exception(e)
            if should_retry:
                retry_delay = base_delay * (2 ** attempt)
                print(f"处理批次 (起始序号 {start_index}) 第 {attempt + 1} 次调用失败，将在 {retry_delay}s 后重试: {e}")
                time.sleep(retry_delay)
                continue
            print(f"处理批次 (起始序号 {start_index}) 最终失败: {e}")
            return {
                "success": False,
                "content": _build_batch_error_block(start_index, len(email_batch), e),
                "error": {
                    "exception_type": type(e).__name__,
                    "message": str(e),
                },
            }

def summarize_with_llm(email_list, batch_size=25, max_retries=3, base_delay=2):
    if not email_list:
        return "### 每日邮件汇总\n**总览：共 0 封邮件**\n\n--- \n\n今日没有收到新邮件。"
        
    client = openai.OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url=BASE_URL # 使用环境变量中的 BASE_URL
    )
    
    total_emails = len(email_list)
    report_parts = []
    failed_batches = 0
    
    print(f"开始分批总结 {total_emails} 封邮件，每批最多 {batch_size} 封...")
    
    for i in range(0, total_emails, batch_size):
        batch = email_list[i:i + batch_size]
        start_index = i + 1
        print(f"  正在处理第 {i//batch_size + 1} 批 (邮件 {start_index} 到 {min(i+batch_size, total_emails)})...")
        
        batch_result = summarize_single_batch(
            client,
            batch,
            start_index,
            max_retries=max_retries,
            base_delay=base_delay,
        )
        if not batch_result["success"]:
            failed_batches += 1
        report_parts.append(batch_result["content"])

    overview = (
        f"### 每日邮件汇总\n"
        f"**总览：共 {total_emails} 封邮件，处理失败批次 {failed_batches} 个**\n\n"
        "---"
    )
    return "\n".join([overview] + report_parts)

def send_email_notification(summary_md, date_for_subject):
    if not SENDER_EMAIL or not SENDER_AUTH_CODE or not RECEIVER_EMAIL:
        print("发送邮件所需的环境变量不完整，跳过发送。")
        return

    html_content = markdown2.markdown(summary_md, extras=["tables", "fenced-code-blocks"])
    message = MIMEText(html_content, 'html', 'utf-8')
    
    subject_str = f"每日邮件总结 - {date_for_subject.strftime('%Y-%m-%d')}"
    message['Subject'] = Header(subject_str, 'utf-8')
    message['From'] = SENDER_EMAIL
    message['To'] = RECEIVER_EMAIL

    try:
        # 修复 465 端口必须用 SMTP_SSL 的问题
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
        server.login(SENDER_EMAIL, SENDER_AUTH_CODE)
        server.sendmail(SENDER_EMAIL, [RECEIVER_EMAIL], message.as_string())
        server.quit()
        print(f"成功发送邮件总结到 {RECEIVER_EMAIL}！")
    except Exception as e:
        print(f"发送邮件失败: {e}")

# ==============================================================================
# 主执行入口
# ==============================================================================
if __name__ == "__main__":
    required_vars = ["IMAP_EMAIL", "IMAP_AUTH_CODE", "TARGET_FOLDER", "DEEPSEEK_API_KEY", 
                     "SENDER_EMAIL", "SENDER_AUTH_CODE", "RECEIVER_EMAIL", "SMTP_SERVER", "SMTP_PORT", "BASE_URL"]
    if not all(os.environ.get(var) for var in required_vars):
        print("错误：一个或多个必要的环境变量未设置。")
        exit(1)

    print(f"任务启动于 (UTC): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")
    
    beijing_timezone = timezone(timedelta(hours=8))
    beijing_now = datetime.now(beijing_timezone)
    
    target_day = beijing_now - timedelta(days=1)
    print(f"将要总结的日期是 (北京时间): {target_day.strftime('%Y-%m-%d')}")
    
    emails = get_emails_from_target_date(target_day)
    summary_report = summarize_with_llm(emails)
    send_email_notification(summary_report, target_day)
    
    print(f"任务执行完毕于 (UTC): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}")
