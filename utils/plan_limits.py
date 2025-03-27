from models.user import PlanType
from datetime import datetime, timedelta
import redis
import os
from dotenv import load_dotenv

# 環境変数の読み込み
load_dotenv()

# Redis接続
redis_client = redis.Redis(
    host=os.getenv("REDIS_HOST", "localhost"),
    port=int(os.getenv("REDIS_PORT", 6379)),
    db=0,
    decode_responses=True
)

# プラン別の制限
PLAN_LIMITS = {
    PlanType.FREE: {
        "daily_conversions": 5,
        "max_pages": 1,
        "storage_days": 0,
        "has_vision_api": False,
        "has_analytics": False,
        "show_ads": True
    },
    PlanType.PRO: {
        "monthly_conversions": 500,
        "max_pages": 3,
        "storage_days": 0,
        "has_vision_api": True,
        "has_analytics": True,
        "show_ads": False
    },
    PlanType.PREMIUM: {
        "monthly_conversions": 2000,
        "max_pages": 6,
        "storage_days": 30,
        "has_vision_api": True,
        "has_analytics": True,
        "show_ads": False
    }
}

def get_conversion_count_key(user_id: int, plan_type: PlanType) -> str:
    """変換回数カウントのRedisキーを生成"""
    if plan_type == PlanType.FREE:
        return f"conv:daily:{user_id}:{datetime.now().strftime('%Y-%m-%d')}"
    else:
        return f"conv:monthly:{user_id}:{datetime.now().strftime('%Y-%m')}"

def can_convert_pdf(user_id: int, plan_type: PlanType, pdf_pages: int) -> tuple[bool, str]:
    """PDFの変換が可能かチェック"""
    limits = PLAN_LIMITS[plan_type]
    
    # ページ数制限チェック
    if pdf_pages > limits["max_pages"]:
        return False, f"PDFのページ数が制限を超えています（上限: {limits['max_pages']}ページ）"
    
    # 変換回数制限チェック
    count_key = get_conversion_count_key(user_id, plan_type)
    current_count = int(redis_client.get(count_key) or 0)
    
    if plan_type == PlanType.FREE:
        if current_count >= limits["daily_conversions"]:
            return False, "本日の変換回数制限に達しました"
        # 24時間後に期限切れ
        if not redis_client.exists(count_key):
            redis_client.setex(count_key, 86400, 1)
        else:
            redis_client.incr(count_key)
    else:
        monthly_limit = limits["monthly_conversions"]
        if current_count >= monthly_limit:
            return False, "今月の変換回数制限に達しました"
        # 月末まで保持
        if not redis_client.exists(count_key):
            next_month = datetime.now() + timedelta(days=32)
            next_month = next_month.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            ttl = int((next_month - datetime.now()).total_seconds())
            redis_client.setex(count_key, ttl, 1)
        else:
            redis_client.incr(count_key)
    
    return True, ""

def should_show_ads(plan_type: PlanType) -> bool:
    """広告表示の有無を判定"""
    return PLAN_LIMITS[plan_type]["show_ads"]

def can_use_vision_api(plan_type: PlanType) -> bool:
    """Vision APIの使用可否を判定"""
    return PLAN_LIMITS[plan_type]["has_vision_api"]

def can_use_analytics(plan_type: PlanType) -> bool:
    """分析機能の使用可否を判定"""
    return PLAN_LIMITS[plan_type]["has_analytics"]

def get_storage_days(plan_type: PlanType) -> int:
    """ファイル保存期間を取得"""
    return PLAN_LIMITS[plan_type]["storage_days"] 