
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

PRODUCT_FILE = Path(
    os.getenv(
        "NIVORA_PRODUCT_FILE",
        str(BASE_DIR / "data" / "Final_master_product_list_with_azure_image_urls.xlsx"),
    )
)

KNOWLEDGE_FILE = Path(
    os.getenv(
        "NIVORA_KNOWLEDGE_FILE",
        str(BASE_DIR / "data" / "nivora_knowledge_base.txt"),
    )
)
COMPANY_INFO_FILE = KNOWLEDGE_FILE
FAQ_FILE = KNOWLEDGE_FILE

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

app = FastAPI(
    title="Nivora Customer Chatbot API",
    description="FastAPI backend for Nivora customer support. Use Swagger UI at /docs to test endpoints.",
    version="3.1.0",
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, float) and pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in {"nan", "none", "null"}:
        return ""

    return value


def norm(value: Any) -> str:
    """Normalize text so capitalization and punctuation do not matter."""
    return re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()


def price_value(value: Any) -> float | None:
    match = re.search(r"\d[\d,]*(?:\.\d+)?", clean(value))

    if not match:
        return None

    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def read_text_file(path: Path) -> str:
    try:
        if path.exists():
            return path.read_text(encoding="utf-8-sig").strip()
    except OSError:
        pass

    return ""


# ============================================================
# COMPANY INFORMATION
# ============================================================

def company_info() -> str:
    info = read_text_file(COMPANY_INFO_FILE)

    if info:
        start_marker = "=== BEGIN COMPANY INFORMATION ==="
        end_marker = "=== END COMPANY INFORMATION ==="
        if start_marker in info and end_marker in info:
            info = info.split(start_marker, 1)[1].split(end_marker, 1)[0]
        return info.strip()[:12000]

    return (
        "No verified company profile is currently configured. "
        "Do not invent company history, leadership details, "
        "contact information, or policies. Direct customers "
        "to the official Nivora support channel."
    )


def answer_company_question(message: str) -> dict[str, Any] | None:
    """
    Answer common company questions from the FAQ or company
    information file. Matching is case-insensitive.
    """
    text = norm(message)

    company_terms = (
        "nivora", "company", "founder", "founders", "ceo",
        "chief executive", "director", "incorporated",
        "incorporation", "history", "headquarters",
        "what do you do", "who are you", "about"
    )

    if not any(term in text for term in company_terms):
        return None

    entries = load_faq_entries()

    # Leadership questions always use Q4 when present.
    if re.search(
        r"\b(ceo|chief executive|founder|founders|leadership)\b",
        text,
    ):
        entry = next(
            (item for item in entries if item["number"] == "4"),
            None,
        )

        if entry:
            return {
                "answer": entry["answer"],
                "products": [],
                "source": "faq_q4",
            }

    return None


# ============================================================
# PRODUCT CATALOGUE
# ============================================================

def load_catalogue() -> tuple[list[dict[str, Any]], str | None]:
    if not PRODUCT_FILE.exists():
        return [], f"Product workbook not found: {PRODUCT_FILE}"

    try:
        from openpyxl import load_workbook

        book = load_workbook(
            PRODUCT_FILE,
            read_only=False,
            data_only=True,
        )

        sheet = book[book.sheetnames[0]]
        headers = [clean(cell.value) for cell in sheet[1]]

        records: list[dict[str, Any]] = []

        for row in sheet.iter_rows(min_row=2, values_only=False):
            item = {
                headers[i]: row[i].value
                for i in range(min(len(headers), len(row)))
                if headers[i]
            }

            if "Image URL" in headers:
                cell = row[headers.index("Image URL")]

                target = (
                    cell.hyperlink.target
                    if cell.hyperlink
                    else clean(cell.value)
                )

                item["Image URL"] = (
                    target
                    if str(target).startswith(("https://", "http://"))
                    else ""
                )

            if not (
                clean(item.get("Item ID"))
                or clean(item.get("Item Code"))
                or clean(item.get("Item Name"))
            ):
                continue

            item["_price"] = (
                price_value(item.get("Selling price"))
                or price_value(item.get("Price"))
            )

            fields = [
                "Item Code",
                "Category",
                "Sub category",
                "Item type",
                "Item Subtype 1",
                "Item Subtype 2",
                "Item Name",
                "Item Size",
                "Item Color",
                "Item Description",
                "Supplier Name",
            ]

            item["_search"] = norm(
                " ".join(clean(item.get(key)) for key in fields)
            )

            item["_code_norm"] = norm(item.get("Item Code"))

            records.append(item)

        book.close()

        return records, None

    except Exception as exc:
        return [], f"Could not load product workbook: {exc}"


PRODUCTS, CATALOGUE_ERROR = load_catalogue()


def public_product(p: dict[str, Any]) -> dict[str, Any]:
    return {
        "item_id": clean(p.get("Item ID")),
        "item_code": clean(p.get("Item Code")),
        "name": (
            clean(p.get("Item Name"))
            or clean(p.get("Item Description"))
            or "Unnamed product"
        ),
        "description": clean(p.get("Item Description")),
        "category": clean(p.get("Category")),
        "subcategory": clean(p.get("Sub category")),
        "size": clean(p.get("Item Size")),
        "color": clean(p.get("Item Color")),
        "supplier": clean(p.get("Supplier Name")),
        "price": p.get("_price"),
        "price_display": (
            clean(p.get("Selling price"))
            or clean(p.get("Price"))
            or "Price on request"
        ),
        "quantity": clean(p.get("Quantity")),
        "image_url": clean(p.get("Image URL")),
    }


def search_catalogue(
    query: str,
    limit: int = 5,
    max_price: float | None = None,
) -> list[dict[str, Any]]:
    q = norm(query)

    if not q:
        return []

    ignored_words = {
        "please", "show", "find", "search", "me", "some", "any",
        "the", "a", "an", "for", "of", "do", "you", "have", "can",
        "could", "would", "like", "want", "need", "looking", "tell",
        "about", "what", "which", "is", "are", "price", "cost", "how",
        "much", "product", "products", "item", "items", "buy", "get",
        "under", "below",
    }

    tokens = [
        word for word in q.split()
        if word not in ignored_words
    ]

    if not tokens:
        return []

    scored = []

    for product in PRODUCTS:
        price = product.get("_price")

        if max_price is not None:
            if price is None or price > max_price:
                continue

        haystack = product.get("_search", "")
        words = set(haystack.split())

        score = 0.0

        if q == product.get("_code_norm"):
            score += 100

        if q in haystack:
            score += 8

        for token in tokens:
            if token in words:
                score += 3
            elif len(token) >= 3 and any(
                word.startswith(token) for word in words
            ):
                score += 1

        if score >= 3:
            scored.append((score, product))

    scored.sort(
        key=lambda item: (
            -item[0],
            clean(item[1].get("Item Name")),
        )
    )

    return [
        public_product(product)
        for _, product in scored[:limit]
    ]


# ============================================================
# LOAD ALL 67 FAQ QUESTIONS AND ANSWERS
# ============================================================

FAQ_STOP_WORDS = {
    "a", "an", "the", "is", "are", "was", "were", "be",
    "been", "being", "do", "does", "did", "can", "could",
    "would", "should", "what", "when", "where", "who",
    "whom", "which", "why", "how", "i", "me", "my", "we",
    "our", "you", "your", "it", "this", "that", "to", "of",
    "for", "from", "in", "on", "at", "and", "or", "if",
    "please", "tell", "about", "give", "explain", "want",
    "know", "nivora", "customer", "customers", "company",
}


def load_faq_entries() -> list[dict[str, str]]:
    """
    Parse numbered FAQ questions and answers from the consolidated
    data/nivora_knowledge_base.txt file.
    """
    content = read_text_file(FAQ_FILE)

    if not content:
        return []

    pattern = re.compile(
        r"(?ms)^\s*Q(\d+)\.\s*(.*?)\s*^\s*A:\s*(.*?)"
        r"(?=^\s*Q\d+\.|^\s*END OF FAQ|"
        r"^\s*IMPLEMENTATION NOTE|\Z)"
    )

    entries = []

    for match in pattern.finditer(content):
        question = re.sub(r"\s+", " ", match.group(2)).strip()
        answer = re.sub(r"\s+", " ", match.group(3)).strip()

        if question and answer:
            entries.append({
                "number": match.group(1),
                "question": question,
                "answer": answer,
            })

    return entries


FAQ_ENTRIES = load_faq_entries()


def faq_tokens(text: str) -> set[str]:
    return {
        word
        for word in norm(text).split()
        if word not in FAQ_STOP_WORDS and len(word) > 1
    }


def get_faq(number: str) -> dict[str, str] | None:
    return next(
        (
            entry for entry in FAQ_ENTRIES
            if entry["number"] == number
        ),
        None,
    )


def faq_result(entry: dict[str, str]) -> dict[str, str]:
    return {
        "answer": entry["answer"],
        "source": f"faq_q{entry['number']}",
    }


def find_faq_answer(message: str) -> dict[str, str] | None:
    """
    Match questions using normalized lowercase text, aliases,
    topic rules, and token similarity.
    """
    query = norm(message)

    if not query or not FAQ_ENTRIES:
        return None

    # Short forms and common variations are expanded.
    expansions = {
        "ceo": {
            "ceo", "founder", "founders", "leadership",
            "director", "chief", "executive", "officer",
        },
        "refund": {
            "refund", "refunds", "reimbursement",
            "money", "repayment",
        },
        "return": {
            "return", "returns", "returning", "send", "back",
        },
        "process": {
            "process", "steps", "procedure", "workflow", "how",
        },
        "cancel": {
            "cancel", "cancellation", "cancelled", "canceled",
        },
        "delivery": {
            "delivery", "shipping", "dispatch", "shipment",
        },
        "damaged": {
            "damaged", "damage", "defective", "broken",
            "incorrect", "missing", "issue", "report",
            "reporting", "faulty",
        },
        "exchange": {
            "exchange", "replacement", "replace", "swap",
        },
        "privacy": {
            "privacy", "personal", "private", "data", "security",
        },
        "contact": {
            "contact", "phone", "email", "support", "help",
        },
        "price": {
            "price", "prices", "cost", "selling",
        },
        "stock": {
            "stock", "availability", "available", "inventory",
        },
        "tracking": {
            "tracking", "track", "status", "order",
        },
        "payment": {
            "payment", "pay", "paid", "cash", "cod",
        },
        "warranty": {
            "warranty", "repair", "guarantee",
        },
    }

    # Exact match after normalization ignores capitalization
    # and punctuation.
    for entry in FAQ_ENTRIES:
        if query == norm(entry["question"]):
            return faq_result(entry)

    # CEO / founder questions always map to Q4, whether the
    # user types CEO, ceo, CeO, or a longer question.
    if re.search(
        r"\b(ceo|chief executive officer|founder|founders)\b",
        query,
    ):
        entry = get_faq("4")

        if entry:
            return faq_result(entry)

    # Company name / purpose.
    if any(
        phrase in query
        for phrase in (
            "full official legal company name",
            "legal company name",
            "what does nivora do",
            "what is nivora",
            "about nivora",
            "what is your company",
        )
    ):
        target = "1" if any(
            phrase in query
            for phrase in ("legal company name", "company name")
        ) else "2"

        entry = get_faq(target)

        if entry:
            return faq_result(entry)

    # Refund process: distinguish it from timelines, methods,
    # eligibility, and delayed refund questions.
    if (
        "refund" in query
        and any(
            word in query
            for word in ("process", "steps", "procedure", "workflow")
        )
        and not any(
            word in query
            for word in (
                "timeline", "timelines", "when", "payment method",
                "cash on delivery", "delayed", "rejected", "status",
            )
        )
    ):
        entry = get_faq("19")

        if entry:
            return faq_result(entry)

    # Return process.
    if (
        any(word in query for word in ("return", "returning"))
        and any(word in query for word in ("process", "steps", "how"))
        and "refund" not in query
    ):
        entry = get_faq("15")

        if entry:
            return faq_result(entry)

    # Strong topic-specific routing prevents similar FAQs from
    # being mixed together.
    topic_rules = [
        (
            r"\b(return window|return period|how long.*return)\b",
            "12",
        ),
        (
            r"\b(return.*eligible|eligible.*return|return exclusions)\b",
            "13",
        ),
        (
            r"\b(refund.*eligible|refund eligibility|refund approved)\b",
            "18",
        ),
        (
            r"\b(refund timeline|refund timelines|how long.*refund|"
            r"when.*refund|refund.*take)\b",
            "20",
        ),
        (
            r"\b(refund method|refund payment|cash on delivery|cod refund)\b",
            "21",
        ),
        (
            r"\b(delayed refund|refund delayed|refund rejected|"
            r"refund not received|refund not visible)\b",
            "23",
        ),
        (
            r"\b(live refund status|check refund status|refund status)\b",
            "24",
        ),
        (
            r"\b(cancel.*order|order.*cancel|cancellation)\b",
            "33",
        ),
        (
            r"\b(track.*order|order status|delivery tracking)\b",
            "67",
        ),
        (
            r"\b(damaged item|defective item|broken item|"
            r"incorrect item|missing item|faulty product)\b",
            "26",
        ),
        (
            r"\b(exchange|replace.*product|replacement)\b",
            "29",
        ),
        (
            r"\b(payment methods|how can i pay|payment options|"
            r"upi payment)\b",
            "21",
        ),
        (
            r"\b(privacy policy|terms and conditions|privacy terms)\b",
            "59",
        ),
        (
            r"\b(customer support|support email|support phone|"
            r"contact nivora|contact customer care)\b",
            "6",
        ),
        (
            r"\b(official website|website link|nivora website)\b",
            "5",
        ),
        (
            r"\b(found|founded|incorporated|company history|"
            r"when was nivora established)\b",
            "3",
        ),
        (
            r"\b(product categories|what categories)\b",
            "7",
        ),
        (
            r"\b(product catalogue|product catalog|source of product prices)\b",
            "8",
        ),
        (
            r"\b(live order status|live order tracking|"
            r"can you track my order)\b",
            "67",
        ),
    ]

    for pattern, number in topic_rules:
        if re.search(pattern, query):
            entry = get_faq(number)

            if entry:
                return faq_result(entry)

    # General fuzzy matching using expanded topic words.
    query_words = faq_tokens(query)
    expanded_query = set(query_words)

    for word in query_words:
        for group in expansions.values():
            if word in group:
                expanded_query.update(group)

    best_entry = None
    best_score = 0.0

    for entry in FAQ_ENTRIES:
        question_words = faq_tokens(entry["question"])
        expanded_question = set(question_words)

        for word in question_words:
            for group in expansions.values():
                if word in group:
                    expanded_question.update(group)

        overlap = expanded_query & expanded_question

        if not overlap:
            continue

        score = (
            2 * len(overlap)
        ) / max(
            1,
            len(expanded_query) + len(expanded_question),
        )

        normalized_question = norm(entry["question"])

        if query in normalized_question:
            score += 0.45

        if len(expanded_query) == 1 and len(overlap) == 1:
            score = max(score, 0.48)

        if score > best_score:
            best_score = score
            best_entry = entry

    if best_entry and best_score >= 0.22:
        return faq_result(best_entry)

    return None


# ============================================================
# API REQUEST MODEL AND ENDPOINTS
# ============================================================

class ChatRequest(BaseModel):
    message: str = Field(min_length=1, max_length=2000)
    history: list[dict[str, str]] = Field(default_factory=list)


@app.get("/", tags=["system"])
def root():
    return {
        "service": "Nivora Customer Chatbot API",
        "status": "ok",
        "documentation": "/docs",
        "health": "/api/health",
    }


@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "catalogue_loaded": bool(PRODUCTS),
        "product_count": len(PRODUCTS),
        "catalogue_error": CATALOGUE_ERROR,
        "faq_loaded": bool(FAQ_ENTRIES),
        "faq_count": len(FAQ_ENTRIES),
        "company_info_loaded": COMPANY_INFO_FILE.exists(),
    }


@app.get("/api/categories")
def categories():
    result: dict[str, int] = {}

    for product in PRODUCTS:
        category = clean(product.get("Category"))

        if category:
            result[category] = result.get(category, 0) + 1

    return {
        "categories": [
            {"name": name, "product_count": count}
            for name, count in sorted(result.items())
        ]
    }


@app.get("/api/products")
def products(q: str = "", limit: int = 10):
    return {
        "products": search_catalogue(
            q,
            max(1, min(limit, 30)),
        )
    }


@app.post("/api/reload-catalogue")
def reload_catalogue():
    global PRODUCTS, CATALOGUE_ERROR

    PRODUCTS, CATALOGUE_ERROR = load_catalogue()

    return {
        "catalogue_loaded": bool(PRODUCTS),
        "product_count": len(PRODUCTS),
        "error": CATALOGUE_ERROR,
    }


# ============================================================
# INTENT DETECTION
# ============================================================

def detect_intent(message: str) -> str:
    text = norm(message)

    if re.search(
        r"\b(hi|hello|hey|hiya|how are you|good morning|"
        r"good afternoon|good evening|good night)\b",
        text,
    ):
        return "greeting"

    if re.search(
        r"\b(thanks|thank you|thankyou|thx|cheers)\b",
        text,
    ):
        return "thanks"

    if re.search(
        r"\b(return|returns|refund|refunds|exchange|exchanges|"
        r"returning|replacement|replace a product)\b",
        text,
    ):
        return "return_policy"

    if re.search(
        r"\b(ceo|founder|founders|company history|history of nivora|"
        r"about nivora|who are you|your company|about the company|"
        r"director|chief executive officer)\b",
        text,
    ):
        return "company_info"

    if re.search(
        r"\b(delivery policy|shipping policy|warranty|payment policy|"
        r"cancel my order|order status|track my order|track order)\b",
        text,
    ):
        return "customer_service"

    if re.search(
        r"\b(product|products|item|items|price|prices|cost|buy|purchase|"
        r"available|availability|stock|under|below|cheaper|size|colour|"
        r"color|code|catalogue|catalog|show me|looking for|need a|want a)\b",
        text,
    ):
        return "product_search"

    return "general"


# ============================================================
# OFFLINE FALLBACK ANSWERS
# ============================================================

def fallback_answer(
    intent: str,
    message: str,
    found: list[dict[str, Any]],
) -> dict[str, Any]:

    if intent == "greeting":
        return {
            "answer": (
                "Hey! 👋 Welcome to Nivora. I'm here to help with "
                "products, listed prices, company information, "
                "and customer-service questions. What would you like to know?"
            ),
            "products": [],
            "source": "friendly_fallback",
        }

    if intent == "thanks":
        return {
            "answer": (
                "You're very welcome! 😊 Is there anything else I can help you with?"
            ),
            "products": [],
            "source": "friendly_fallback",
        }

    if intent == "return_policy":
        return {
            "answer": (
                "I can help explain the available return and refund information. "
                "However, Nivora's official return window, eligibility rules, "
                "and refund timelines have not been confirmed in the connected "
                "information. Please contact Nivora's official support channel "
                "to confirm the applicable process."
            ),
            "products": [],
            "source": "verified_policy_needed",
        }

    if intent == "company_info":
        return {
            "answer": (
                "I don't have enough verified information to answer that company "
                "question yet. Please check Nivora's official corporate channels."
            ),
            "products": [],
            "source": "verified_company_info_needed",
        }

    if intent == "customer_service":
        return {
            "answer": (
                "I can't verify live order status or confirm shipping, warranty, "
                "or cancellation outcomes from this chat. Please use Nivora's "
                "official support channel for confirmation."
            ),
            "products": [],
            "source": "verified_policy_needed",
        }

    if intent == "product_search" and found:
        return {
            "answer": (
                "I found catalogue matches that may help. The product cards show "
                "the details and prices listed in the workbook. Please confirm "
                "current availability before ordering."
            ),
            "products": found,
            "source": "catalogue_search",
        }

    if intent == "product_search":
        return {
            "answer": (
                "I couldn't find a close match in the current catalogue. Try a "
                "product name, item code, size, category, or supplier."
            ),
            "products": [],
            "source": "catalogue_search",
        }

    return {
        "answer": (
            "I'm here to help! 😊 You can ask about Nivora, listed products "
            "and prices, company information, or customer-service topics."
        ),
        "products": [],
        "source": "friendly_fallback",
    }


# ============================================================
# MAIN CHAT ENDPOINT
# ============================================================

@app.post("/api/chat")
def chat(payload: ChatRequest):
    message = payload.message.strip()

    if not message:
        raise HTTPException(
            status_code=400,
            detail="Please enter a message.",
        )

    intent = detect_intent(message)

    # Answer from local FAQ first. This works without OpenAI credits.
    faq_match = find_faq_answer(message)

    if faq_match:
        return {
            "answer": faq_match["answer"],
            "products": [],
            "source": faq_match["source"],
        }

    # Greetings and thanks should not trigger catalogue searches.
    max_price = None

    budget = re.search(
        r"(?:under|below|less than|up to|upto|within|maximum|max)"
        r"\s*(?:₹|rs\.?|inr)?\s*([\d,]+(?:\.\d+)?)",
        message.lower(),
    )

    if budget:
        max_price = price_value(budget.group(1))

    found = (
        search_catalogue(message, limit=5, max_price=max_price)
        if intent == "product_search"
        else []
    )

    # Optional AI response. The app still works if the API is unavailable.
    if OPENAI_API_KEY:
        try:
            from openai import OpenAI

            client = OpenAI(
                api_key=OPENAI_API_KEY,
                timeout=20.0,
                max_retries=0,
            )

            history = []

            for item in payload.history[-12:]:
                role = item.get("role", "")
                content = clean(item.get("content"))

                if role in {"user", "assistant"} and content:
                    history.append({
                        "role": role,
                        "content": content[:2000],
                    })

            system = f"""
You are Nivora's friendly customer assistant.

Answer using only verified company information and the supplied catalogue
matches. Never invent policies, refund timelines, company facts, prices,
stock status, contact information, or order status.

For company leadership questions, use the local FAQ answer when available.
Never reveal private phone numbers, personal email addresses, home addresses,
passwords, OTPs, CVVs, full payment-card numbers, or API keys.

For returns and refunds, do not present draft workflows as official policy.
Do not claim a cancellation, refund, order update, or support ticket has
been completed unless a connected integration confirms it.

VERIFIED COMPANY INFORMATION:
{company_info()}

RELEVANT PRODUCT CATALOGUE MATCHES:
{found}
"""

            messages = [
                {"role": "system", "content": system},
                *history,
                {"role": "user", "content": message},
            ]

            response = client.chat.completions.create(
                model=OPENAI_MODEL,
                messages=messages,
                temperature=0.3,
            )

            answer = (
                response.choices[0].message.content or ""
            ).strip()

            if answer:
                return {
                    "answer": answer,
                    "products": (
                        found if intent == "product_search" else []
                    ),
                    "source": (
                        "ai_assistant_and_catalogue"
                        if found
                        else "ai_assistant"
                    ),
                }

        except Exception:
            # API quota, connectivity, or model errors fall back
            # to local answers instead of breaking the chatbot.
            pass

    return fallback_answer(intent, message, found)


# ============================================================
# LOCAL TESTING
# ============================================================

@app.get("/api/faq-status")
def faq_status():
    """Check whether all numbered FAQ entries were loaded."""
    return {
        "faq_file": str(FAQ_FILE),
        "faq_loaded": bool(FAQ_ENTRIES),
        "faq_count": len(FAQ_ENTRIES),
        "faq_numbers": [
            entry["number"] for entry in FAQ_ENTRIES
        ],
    }
