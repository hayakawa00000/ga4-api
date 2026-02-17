from flask import Flask, jsonify, request
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import RunReportRequest
from google.oauth2 import service_account
import os
import json
from datetime import datetime

app = Flask(__name__)

# 環境変数からサービスアカウント情報を取得
SERVICE_ACCOUNT_JSON = os.environ.get('SERVICE_ACCOUNT_JSON')

@app.route('/')
def home():
    return jsonify({
        "status": "GA4 API is running",
        "endpoints": {
            "/ga4/sessions": "Get sessions data",
            "/health": "Health check"
        },
        "usage": {
            "basic": "/ga4/sessions?property_id=123456789",
            "with_dates": "/ga4/sessions?property_id=123456789&start_date=2024-01-01&end_date=2024-01-31"
        }
    })

@app.route('/health')
def health():
    return jsonify({"status": "ok"})

@app.route('/ga4/sessions')
def get_sessions():
    try:
        # URLパラメータを取得
        property_id = request.args.get('property_id')
        start_date = request.args.get('start_date', '7daysAgo')  # デフォルト: 7日前
        end_date = request.args.get('end_date', 'today')  # デフォルト: 今日
        
        if not property_id:
            return jsonify({
                "success": False,
                "error": "property_id パラメータが必要です",
                "usage": "/ga4/sessions?property_id=123456789&start_date=2024-01-01&end_date=2024-01-31"
            }), 400
        
        # 日付フォーマットの検証（YYYY-MM-DD形式の場合）
        if start_date not in ['7daysAgo', '30daysAgo', 'yesterday', 'today']:
            try:
                datetime.strptime(start_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "start_date は YYYY-MM-DD 形式で指定してください（例: 2024-01-01）"
                }), 400
        
        if end_date not in ['today', 'yesterday']:
            try:
                datetime.strptime(end_date, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    "success": False,
                    "error": "end_date は YYYY-MM-DD 形式で指定してください（例: 2024-01-31）"
                }), 400
        
        # 環境変数のチェック
        if not SERVICE_ACCOUNT_JSON:
            return jsonify({
                "success": False,
                "error": "SERVICE_ACCOUNT_JSON が設定されていません"
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
        request_obj = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[{
                "start_date": start_date,
                "end_date": end_date
            }],
            metrics=[{"name": "sessions"}]
        )
        
        # APIを実行
        response = client.run_report(request_obj)
        
        # セッション数を取得
        sessions = 0
        if response.rows:
            sessions = response.rows[0].metric_values[0].value
        
        return jsonify({
            "success": True,
            "property_id": property_id,
            "sessions": int(sessions),
            "start_date": start_date,
            "end_date": end_date,
            "message": f"期間 {start_date} 〜 {end_date} のセッション数: {sessions}"
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
            "property_id": property_id if 'property_id' in locals() else None
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
