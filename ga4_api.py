from flask import Flask, jsonify
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
import os
import json

app = Flask(__name__)

# 環境変数から設定を取得
PROPERTY_ID = os.environ.get('GA4_PROPERTY_ID')
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')

@app.route('/')
def home():
    return jsonify({
        "status": "GA4 API is running",
        "endpoints": {
            "/ga4/sessions": "Get sessions data for the last 7 days",
            "/health": "Health check"
        }
    })

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/ga4/sessions')
def get_sessions():
    try:
        # 環境変数のチェック
        if not PROPERTY_ID or not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "環境変数が設定されていません。GA4_PROPERTY_ID と SERVICE_ACCOUNT_JSON を設定してください。"
            }), 500
        
        # サービスアカウント情報をパース
        service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
        
        # 認証情報を作成
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/analytics.readonly']
        )
        
        # クライアント作成
        client = BetaAnalyticsDataClient(credentials=credentials)
        
        # レポートリクエスト
        request = RunReportRequest(
            property=f"properties/{PROPERTY_ID}",
            date_ranges=[{"start_date": "7daysAgo", "end_date": "today"}],
            metrics=[{"name": "sessions"}]
        )
        
        # APIを実行
        response = client.run_report(request)
        
        # セッション数を取得
        sessions = 0
        if response.rows:
            sessions = response.rows[0].metric_values[0].value
        
        return jsonify({
            "success": True,
            "sessions": int(sessions),
            "message": f"過去7日間のセッション数: {sessions}",
            "period": "7daysAgo to today"
        })
    
    except json.JSONDecodeError as e:
        return jsonify({
            "success": False,
            "error": "SERVICE_ACCOUNT_JSON の形式が正しくありません",
            "details": str(e)
        }), 500
    
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e),
            "hint": "GA4のプロパティIDとサービスアカウントの設定を確認してください"
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
