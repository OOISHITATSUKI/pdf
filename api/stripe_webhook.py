import stripe
import os
from dotenv import load_dotenv
from fastapi import APIRouter, Request, HTTPException
from database import get_db
from models.user import User, PlanType
from sqlalchemy.orm import Session

# 環境変数の読み込み
load_dotenv()
stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
webhook_secret = os.getenv("STRIPE_WEBHOOK_SECRET")

router = APIRouter()

# プラン価格IDとプランタイプのマッピング
PRICE_ID_TO_PLAN = {
    "price_xxx": PlanType.PRO,  # プロプランの価格ID
    "price_yyy": PlanType.PREMIUM,  # プレミアムプランの価格ID
}

@router.post("/webhook")
async def stripe_webhook(request: Request):
    try:
        # リクエストボディを取得
        payload = await request.body()
        sig_header = request.headers.get("stripe-signature")
        
        try:
            # Webhookイベントを検証
            event = stripe.Webhook.construct_event(
                payload, sig_header, webhook_secret
            )
        except ValueError as e:
            raise HTTPException(status_code=400, detail="Invalid payload")
        except stripe.error.SignatureVerificationError as e:
            raise HTTPException(status_code=400, detail="Invalid signature")

        # イベントタイプに応じた処理
        if event["type"] == "checkout.session.completed":
            session = event["data"]["object"]
            customer_id = session["customer"]
            price_id = session["line_items"]["data"][0]["price"]["id"]
            
            # データベースセッションを取得
            db = next(get_db())
            
            try:
                # ユーザーを検索・更新
                user = db.query(User).filter(User.stripe_customer_id == customer_id).first()
                if user and price_id in PRICE_ID_TO_PLAN:
                    user.plan_type = PRICE_ID_TO_PLAN[price_id]
                    db.commit()
            finally:
                db.close()

        return {"status": "success"}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) 